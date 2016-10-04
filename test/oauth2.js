import { expect } from 'chai';
import {OAuth2} from '../src/oauth2';

describe('OAuth2', function test() {
  describe('when configured', () => {
    // This should fail when some required parameters are missing e.g. tokenUri
    it('with some parameters', () => {
      const authentication = OAuth2({
        clientId: 'test_client_id',
        clientSecret: 'test_client_secret',
        redirectUri: 'http://localhost',
        authorizeUri: 'http://localhost'
      });

      expect(authentication).to.be.a('object');
      expect(authentication.requestToken).to.be.a('function');
    });
  });

  describe('when request token', () => {
    it('should return undefined and log when the token uri is missing', () => {
      const authentication = OAuth2({
        clientId: 'test_client_id',
        clientSecret: 'test_client_secret',
        redirectUri: 'http://localhost',
        authorizeUri: 'http://localhost',
        realm: 'test_realm',
        resource: 'test_resource'
      });

      authentication.requestToken('123').then((result) => {
        expect(result).to.be.undefined;
      });
    });
  });
});
