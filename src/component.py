'''
Template Component main class.

'''

import csv
import json
import logging
import sys

from kbc.env_handler import KBCEnvHandler

from ms_graph.client import Client
from ms_graph.dataobjects import get_col_def_name, get_col_definition, TextColumn, SharepointList, ColumnDefinition
from ms_graph.exceptions import BaseError

# global constants'
BATCH_LIMIT = 20
KEY_COLUMN_SETUP = 'column_setup'
OAUTH_APP_SCOPE = 'offline_access Files.Read Sites.ReadWrite.All'
# configuration variables
KEY_BASE_HOST = 'base_host_name'
KEY_SITE_REL_PATH = 'site_url_rel_path'
KEY_LIST_NAME = 'list_name'
KEY_CREATE_NEW = 'create_new'
KEY_TITLE_COL = 'title_column'
KEY_SRC_NAME = 'name'

# #### Keep for debug
KEY_DEBUG = 'debug'
MANDATORY_PARS = [KEY_BASE_HOST, KEY_LIST_NAME, KEY_SITE_REL_PATH]
MANDATORY_IMAGE_PARS = []


class Component(KBCEnvHandler):

    def __init__(self, debug=False):
        KBCEnvHandler.__init__(self, MANDATORY_PARS, log_level=logging.DEBUG if debug else logging.INFO)
        # override debug from config
        if self.cfg_params.get(KEY_DEBUG):
            debug = True
        if debug:
            logging.getLogger().setLevel(logging.DEBUG)
        logging.info('Loading configuration...')

        try:
            self.validate_config(MANDATORY_PARS)

        except ValueError as ex:
            logging.exception(ex)
            exit(1)

        authorization_data = json.loads(self.get_authorization().get('#data'))
        token = authorization_data.get('refresh_token')
        if not token:
            raise Exception('Missing access token in authorization data!')

        self.client = Client(refresh_token=token, client_id=self.get_authorization()['appKey'],
                             client_secret=self.get_authorization()['#appSecret'], scope=OAUTH_APP_SCOPE)

    def run(self):
        '''
        Main execution code
        '''
        params = self.cfg_params  # noqa

        try:

            in_tables = self.configuration.get_input_tables()

            if len(in_tables) == 0:
                logging.error('There is no table specified on the input mapping! You must provide one input table!')
                exit(1)

            in_table = in_tables[0]

            site = self.client.get_site_by_relative_url(params[KEY_BASE_HOST], params[KEY_SITE_REL_PATH])
            if not site.get('id'):
                raise RuntimeError(
                    f'No site with given url: '
                    f'{"/".join([params[KEY_BASE_HOST], params[KEY_SITE_REL_PATH]])} found.')

            # get existing list
            sh_list = self.client.get_site_list_by_name(site['id'], params[KEY_LIST_NAME])

            list_dsc = ''
            if params.get('list_description'):
                list_dsc = params['list_description'][0]

            table_pars = params.get(KEY_CREATE_NEW, {})
            title_col_mapping = table_pars[0][KEY_TITLE_COL] if table_pars else None

            if table_pars and not sh_list:
                # create new list
                table_pars = table_pars[0]
                title_col_mapping = table_pars[KEY_TITLE_COL]
                sh_list = self._create_new_list(site['id'], params[KEY_LIST_NAME], list_dsc, table_pars,
                                                in_table)
            else:
                if not sh_list:
                    raise RuntimeError(
                        f'No list named "{params[KEY_LIST_NAME]}" found on site : '
                        f'{"/".join([params[KEY_BASE_HOST], params[KEY_SITE_REL_PATH]])} .')
                elif params.get(KEY_CREATE_NEW, {}):
                    logging.warning(f'The list "{params[KEY_LIST_NAME]}" already exists. The "new list" '
                                    f'configuration will be ignored and the existing list updated.')

            logging.info('Getting list details...')
            list_columns = self.client.get_site_list_columns(site['id'], sh_list['id'],
                                                             expand_par='columns')

            non_existent_cols = self.validate_table_cols(list_columns, in_table, title_col_mapping)
            if non_existent_cols:
                logging.warning(
                    f'Some columns: {non_existent_cols} were not found in the destination list. They will be ignored!')

            # emtpy the list first
            logging.warning('Removing all existing items..')
            self._empty_list(site['id'], sh_list)

            logging.info('Writing table items.')
            self.write_table(site['id'], sh_list['id'], in_table, non_existent_cols,
                             title_col_mapping)

            logging.info('Export finished!')

        except BaseError as ex:
            logging.exception(ex)
            exit(1)

    def _empty_list(self, site_id, sh_lst):
        for fl in self.client.get_site_list_fields(site_id, sh_lst['id'], expand='fields'):
            f = self.client.delete_list_items(site_id, sh_lst['id'], [f['id'] for f in fl])
            if f:
                raise RuntimeError(f"Some records couldn't be deleted: {f}")

    def write_table(self, site_id, list_id, in_table, nonexistent_cols, title_col):
        with open(in_table['full_path'], mode='r',
                  encoding='utf-8') as in_file:
            reader = csv.DictReader(in_file, lineterminator='\n')

            batch = []
            failed = []
            batch_index = 0
            for ri, line in enumerate(reader):
                if title_col:
                    # creating new list, have col mapping
                    line['Title'] = line.pop(title_col[KEY_SRC_NAME])
                    if title_col[KEY_SRC_NAME] in nonexistent_cols:
                        nonexistent_cols.remove(title_col[KEY_SRC_NAME])

                self._cleanup_record_fields(line, nonexistent_cols)
                br = self.client.build_create_list_item_batch_request(ri, site_id, list_id, line)
                batch.append(br)
                batch_index += 1
                if batch_index >= BATCH_LIMIT:
                    batch_index = 0
                    f = self.client.make_batch_request(batch, 'Create items')
                    failed.extend(f)
                    batch.clear()
            # last batch
            if batch:
                f = self.client.make_batch_request(batch, 'Create items')
                failed.extend(f)

        if failed:
            raise RuntimeError(f'Write finished with error. Some records failed: {failed}')

    def validate_table_cols(self, list_columns, in_table, title_col_mapping=None):
        src_cols = list()
        with open(in_table['full_path'], mode='r',
                  encoding='utf-8') as in_file:
            reader = csv.DictReader(in_file, lineterminator='\n')
            src_cols = reader.fieldnames

        dst_cols = [c['name'] for c in list_columns]
        required_dst_cols = [c['name'] for c in list_columns if c['required']]
        nonexisting_cols = [src_col for src_col in src_cols if src_col not in dst_cols]
        if title_col_mapping:
            # just in case the col does not exist
            src_name = title_col_mapping[KEY_SRC_NAME]
            required_dst_cols.append(src_name)
            src_cols.append('Title')
        missing_required = [req_col for req_col in required_dst_cols if req_col not in src_cols]

        if missing_required:
            raise ValueError(f'Some required columns are missing in the source table: {missing_required}')

        return nonexisting_cols

    def _cleanup_record_fields(self, line, nonexistent_cols):
        for c in nonexistent_cols:
            line.pop(c)

    def _create_new_list(self, site_id, list_name, list_desc, table_pars, in_table):
        title_col = table_pars[KEY_TITLE_COL]
        column_pars = table_pars[KEY_COLUMN_SETUP]

        default_cols = self.validate_table_cols(column_pars, in_table, title_col)
        # validate title col
        if title_col[KEY_SRC_NAME] not in default_cols:
            raise ValueError(f'Specified title column "{title_col[KEY_SRC_NAME]}" is missing in the source table.')
        default_cols.remove(title_col[KEY_SRC_NAME])

        lst_def = self._build_table_def(list_name, list_desc, table_pars, default_cols)

        # create list
        res = self.client.create_list(site_id, lst_def)
        logging.debug(f'List created: {res}')
        return res

    def _build_table_def(self, list_name, list_desc, table_pars, default_cols):
        col_def = list()
        for cpar in table_pars[KEY_COLUMN_SETUP]:
            # only text cols for now
            params = {"name": cpar[KEY_SRC_NAME],
                      "displayName": cpar['display_name'],
                      "description": cpar.get('description', ''),
                      get_col_def_name(cpar['col_type']): get_col_definition(cpar['col_type'])}

            cdef = ColumnDefinition(**params)
            col_def.append(cdef)
        # build default text cols
        for c in default_cols:
            cdef = ColumnDefinition(name=c,
                                    displayName=c,
                                    description='',
                                    text=TextColumn())
            col_def.append(cdef)

        return SharepointList(list_name, col_def)


"""
        Main entrypoint
"""
if __name__ == "__main__":
    if len(sys.argv) > 1:
        debug_arg = sys.argv[1]
    else:
        debug_arg = False
    try:
        comp = Component(debug_arg)
        comp.run()
    except Exception as e:
        logging.exception(e)
        exit(1)
