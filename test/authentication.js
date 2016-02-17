import { expect } from 'chai';
// import nock from 'nock';
const authenticationRewire = require('../src/authentication');
const Authentication = authenticationRewire.Authentication;
const host = 'https://ohaiprismatik.sharepoint.com/wut';

describe('Authentication', function test() {
  this.timeout(40000);

  describe('when configured', () => {
    it('correctly', () => {
      let username = '';
      let password = '';
      let authentication = Authentication(username, password, host);
      expect(authentication).to.be.a('object');
      expect(authentication.request).to.be.a('function');
    });
  });
});
