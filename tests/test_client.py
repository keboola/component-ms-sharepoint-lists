import unittest

import mock

from ms_graph.client import Client


def _token_response():
    response = mock.Mock(status_code=200, headers={'Content-Type': 'application/json'})
    response.json.return_value = {'access_token': 'access-token', 'refresh_token': 'new-refresh-token'}
    return response


class TestClientAuthority(unittest.TestCase):

    @mock.patch('ms_graph.client.requests.post', return_value=_token_response())
    def test_default_authority_url(self, post):
        Client(refresh_token='refresh-token', client_secret='secret', client_id='id', scope='scope')
        self.assertEqual(post.call_args[1]['url'], 'https://login.microsoftonline.com/common/oauth2/v2.0/token')

    @mock.patch('ms_graph.client.requests.post', return_value=_token_response())
    def test_custom_authority_url(self, post):
        authority_url = 'https://login.microsoftonline.com/00000000-0000-0000-0000-000000000000'
        Client(refresh_token='refresh-token', client_secret='secret', client_id='id', scope='scope',
               authority_url=authority_url)
        self.assertEqual(post.call_args[1]['url'], authority_url + '/oauth2/v2.0/token')


if __name__ == "__main__":
    unittest.main()
