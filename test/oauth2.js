import {expect} from 'chai';
import {OAuth2} from '../src/oauth2';
import log from './mocks/logger';

describe('OAuth2', function test() {
  describe('when configured', () => {
    it('should be configured with only some parameters provided', () => {
      const authentication = OAuth2({
        clientId: 'test_client_id',
        clientSecret: 'test_client_secret',
        redirectUri: 'http://localhost',
        authorizeUri: 'http://localhost',
        log
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
        resource: 'test_resource',
        log
      });

      return authentication.requestToken('123')
        .then(() => {
          throw new Error('Must never reach here');
        }, (err) => {
          expect(err.toString()).to.contain('options.uri is a required argument');
        });
    });
  });
});
