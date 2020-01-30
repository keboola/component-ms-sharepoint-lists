import logging
from dataclasses import asdict
from dataclasses import dataclass
from typing import List

import requests
from kbc.client_base import HttpClientBase
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from ms_graph import exceptions
from ms_graph.dataobjects import SharepointList


@dataclass
class BatchRequest:
    id: str
    url: str
    method: str
    body: dict = None
    headers: dict = None


class Client(HttpClientBase):
    OAUTH_LOGIN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    MAX_RETRIES = 9
    BASE_URL = 'https://graph.microsoft.com/v1.0/'
    SYSTEM_LIST_COLUMNS = ["ComplianceAssetId",
                           "ContentType",
                           # "Modified",
                           # "Created",
                           # "Author",
                           # "Editor",
                           "Attachments",
                           "Edit",
                           "LinkTitleNoMenu",
                           "LinkTitle",
                           "DocIcon",
                           "ItemChildCount",
                           "FolderChildCount",
                           "AppAuthor",
                           "AppEditor"]

    def __init__(self, refresh_token, client_secret, client_id, scope):
        HttpClientBase.__init__(self, base_url=self.BASE_URL, max_retries=self.MAX_RETRIES, backoff_factor=0.3,
                                status_forcelist=(429, 503, 500, 502, 504, 507))
        # refresh always on init
        self.__refresh_token = refresh_token
        self.__clien_secret = client_secret
        self.__client_id = client_id
        self.__scope = scope
        access_token = self.refresh_token()

        # set auth header
        self._auth_header = {"Authorization": 'Bearer ' + access_token,
                             "Content-Type": "application/json"}

    def __response_hook(self, res, *args, **kwargs):
        # refresh token if expired
        if res.status_code == 401:
            token = self.refresh_token()
            # update auth header
            self._auth_header = {"Authorization": 'Bearer ' + token,
                                 "Content-Type": "application/json"}
            # reset header
            res.request.headers['Authorization'] = 'Bearer ' + token
            s = requests.Session()
            # retry request
            return self.requests_retry_session(session=s).send(res.request)

    def refresh_token(self):
        data = {"client_id": self.__client_id,
                "client_secret": self.__clien_secret,
                "refresh_token": self.__refresh_token,
                "grant_type": "refresh_token",
                "scope": self.__scope}
        r = requests.post(url=self.OAUTH_LOGIN_URL, data=data)
        parsed = self._parse_response(r, 'login')
        return parsed['access_token']

    def requests_retry_session(self, session=None):
        session = session or requests.Session()
        retry = Retry(
            total=self.max_retries,
            read=self.max_retries,
            connect=self.max_retries,
            backoff_factor=self.backoff_factor,
            status_forcelist=self.status_forcelist,
            method_whitelist=('GET', 'POST', 'PATCH', 'UPDATE', 'DELETE')
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        # append response hook
        session.hooks['response'].append(self.__response_hook)
        return session

    def create_list(self, site_id, lst_object: SharepointList):
        endpoint = f'/sites/{site_id}/lists'
        url = self.base_url + endpoint
        data = asdict(lst_object)
        return self._parse_response(self.post_raw(url=url, json=data), 'create list')

    def _get_paged_result_pages(self, endpoint, parameters):

        has_more = True
        next_url = self.base_url + endpoint
        while has_more:

            resp = self.get_raw(next_url, params=parameters)
            req_response = self._parse_response(resp, endpoint)

            if req_response.get('@odata.nextLink'):
                has_more = True
                next_url = req_response['@odata.nextLink']
            else:
                has_more = False

            yield req_response

    def _delete_raw(self, *args, **kwargs):
        s = requests.Session()
        headers = kwargs.pop('headers', {})
        headers.update(self._auth_header)
        s.headers.update(headers)
        s.auth = self._auth

        # set default params
        params = kwargs.pop('params', {})

        if params is None:
            params = {}

        if self._default_params:

            all_pars = {**params, **self._default_params}
            kwargs.update({'params': all_pars})

        else:

            kwargs.update({'params': params})

        r = self.requests_retry_session(session=s).request('DELETE', *args, **kwargs)
        return r

    def make_batch_request(self, batch_requests: List[dict], r_type=''):
        endpoint = '/$batch'
        rq_url = self.base_url + endpoint

        data = {"requests": batch_requests}

        resp = self.post_raw(rq_url, json=data)
        r = self._parse_response(resp, f'batch: {r_type}')
        return self._get_failed_batch_resp(r)

    def get_site_by_relative_url(self, hostname, site_path):
        """

        :param hostname: e.g. mytenant.sharepoint.com
        :param site_path: e.g. /site/MyTeamSite
        :return:
        """
        url = self.base_url + f'/sites/{hostname}:/{site_path}'
        resp = self._parse_response(self.get_raw(url), 'sites')
        return resp

    def get_site_lists(self, site_id):
        endpoint = f'/sites/{site_id}/lists'
        lists = []
        for l in self._get_paged_result_pages(endpoint, {}):
            lists.extend(l['value'])
        return lists

    def get_site_list_by_name(self, site_id, list_name):
        """

        :param site_id: site id
        :param list_name: unique list name (case sensitive)
        :return: list object
        """
        lists = self.get_site_lists(site_id)
        # ms removes -
        list_name = list_name.replace('-', '')
        res_list = [l for l in lists if l['name'] == list_name]

        return res_list[0] if res_list else None

    def get_site_list_columns(self, site_id, list_id, include_system=False,
                              expand_par='columns(select=name, description, displayName)'):
        """
        Gets array of columns available in the specified list.

        :param site_id:
        :param list_id:
        :param include_system:
        :param expand_par:
        :return:
        """
        endpoint = f'/sites/{site_id}/lists/{list_id}'
        parameters = {'expand': expand_par}

        columns = []
        for l in self._get_paged_result_pages(endpoint, parameters):
            columns.extend(l['columns'])

        if not include_system:
            columns = [c for c in columns if
                       c['name'] not in self.SYSTEM_LIST_COLUMNS and not c['name'].startswith('_')]

        self._dedupe_header(columns)
        return columns

    def get_site_list_fields(self, site_id, list_id, expand='fields'):
        endpoint = f'/sites/{site_id}/lists/{list_id}/items'
        params = {'expand': expand}
        for r in self._get_paged_result_pages(endpoint, params):
            yield [f['fields'] for f in r['value']]

    def delete_list_item(self, site_id, list_id, item_id):
        endpoint = f'/sites/{site_id}/lists/{list_id}/items/{item_id}'
        url = self.base_url + endpoint
        r = self._delete_raw(url=url)
        self._parse_response(r, endpoint)

    def delete_list_items(self, site_id, list_id, item_ids, batch_limit=20):

        batch = []
        failed = []
        batch_index = 0
        for ri, item_id in enumerate(item_ids):
            endpoint = f'/sites/{site_id}/lists/{list_id}/items/{item_id}'
            batch.append(asdict(BatchRequest(str(ri), endpoint, 'DELETE')))
            batch_index += 1
            if batch_index >= batch_limit:
                batch_index = 0
                f = self.make_batch_request(batch, 'Delete items')
                failed.extend(f)
                batch.clear()
            # last batch
        if batch:
            f = self.make_batch_request(batch, 'Delete items')
            failed.extend(f)

        # retry failed one by one. Retry strategy applied
        if failed:
            logging.warning(f'Some requests failed ({failed}), retrying. ')

        failed_idx = []

        for fid, f in enumerate(failed):
            try:
                self.delete_list_item(site_id, list_id, item_ids[int(f['id'])])
            except exceptions.NotFound:
                logging.warning(f'Item {item_ids[int(f["id"])]} already deleted.')

            failed_idx.append(fid)

        return [f for i, f in enumerate(failed) if i not in failed_idx]

    def create_list_item(self, site_id, list_id, fields):
        """

        :param site_id:
        :param list_id:
        :param fields: Dictionary with fields. {key: value}
        :return:
        """
        endpoint = f'/sites/{site_id}/lists/{list_id}/items'

        data = {'fields': fields}
        url = self.base_url + endpoint
        rs = self.post_raw(url=url, json=data)
        return self._parse_response(rs, 'create list item')

    def build_create_list_item_batch_request(self, rq_id, site_id, list_id, fields):
        """

        :param site_id:
        :param list_id:
        :param fields: Dictionary with fields. {key: value}
        :return:
        """
        endpoint = f'/sites/{site_id}/lists/{list_id}/items'

        data = {'fields': fields}
        headers = {'Content-Type': 'application/json'}

        return asdict(BatchRequest(rq_id, endpoint, 'POST', data, headers))

    def _parse_response(self, response, endpoint):
        status_code = response.status_code
        if 'application/json' in response.headers['Content-Type']:
            r = response.json()
        else:
            r = response.text
        if status_code in (200, 201, 202):
            return r
        elif status_code == 204:
            return None
        elif status_code == 400:
            raise exceptions.BadRequest(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 401:
            raise exceptions.Unauthorized(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 403:
            raise exceptions.Forbidden(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 404:
            raise exceptions.NotFound(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 405:
            raise exceptions.MethodNotAllowed(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 406:
            raise exceptions.NotAcceptable(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 409:
            raise exceptions.Conflict(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 410:
            raise exceptions.Gone(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 411:
            raise exceptions.LengthRequired(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 412:
            raise exceptions.PreconditionFailed(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 413:
            raise exceptions.RequestEntityTooLarge(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 415:
            raise exceptions.UnsupportedMediaType(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 416:
            raise exceptions.RequestedRangeNotSatisfiable(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 422:
            raise exceptions.UnprocessableEntity(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 429:
            raise exceptions.TooManyRequests(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 500:
            raise exceptions.InternalServerError(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 501:
            raise exceptions.NotImplemented(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 503:
            raise exceptions.ServiceUnavailable(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 504:
            raise exceptions.GatewayTimeout(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 507:
            raise exceptions.InsufficientStorage(f'Calling endpoint {endpoint} failed', r)
        elif status_code == 509:
            raise exceptions.BandwidthLimitExceeded(f'Calling endpoint {endpoint} failed', r)
        else:
            raise exceptions.UnknownError(f'Calling endpoint {endpoint} failed', r)

    def _get_failed_batch_resp(self, response):
        failed = []
        for r in response['responses']:
            if r['status'] >= 300:
                failed.append(r)
        return failed

    def _dedupe_header(self, columns):
        col_keys = dict()
        dup_headers = set()
        for col in columns:
            if col['displayName'] in col_keys:
                dup_headers.add(col['displayName'])
                col['displayName'] = col['displayName'] + '_' + col['name']
            else:
                col_keys[col['displayName']] = col
        # update first value names as well
        for c in dup_headers:
            col_keys[c]['displayName'] = col_keys[c]['displayName'] + '_' + col_keys[c]['name']
