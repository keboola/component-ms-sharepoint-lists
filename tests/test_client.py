import unittest

from ms_graph.client import Client


class TestBuildTokenUrl(unittest.TestCase):
    """image_parameters.oneDriveAuthorityUrl support, see CFTL-833."""

    def test_no_override_uses_common_endpoint(self):
        # Default behaviour must not change when the stack sets nothing.
        self.assertEqual('https://login.microsoftonline.com/common/oauth2/v2.0/token',
                         Client.build_token_url(None))
        self.assertEqual(Client.OAUTH_LOGIN_URL, Client.build_token_url(None))

    def test_empty_override_uses_common_endpoint(self):
        self.assertEqual(Client.OAUTH_LOGIN_URL, Client.build_token_url(''))

    def test_override_uses_tenant_endpoint(self):
        self.assertEqual(
            'https://login.microsoftonline.com/11111111-2222-3333-4444-555555555555/oauth2/v2.0/token',
            Client.build_token_url('https://login.microsoftonline.com/11111111-2222-3333-4444-555555555555'))

    def test_override_with_trailing_slash(self):
        self.assertEqual(
            'https://login.microsoftonline.com/11111111-2222-3333-4444-555555555555/oauth2/v2.0/token',
            Client.build_token_url('https://login.microsoftonline.com/11111111-2222-3333-4444-555555555555/'))


if __name__ == "__main__":
    unittest.main()
