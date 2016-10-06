import { expect } from 'chai';

const authenticationRewire = require('../src/authentication');
const Authentication = authenticationRewire.Authentication;
const host = 'https://safetyculture.sharepoint.com/IntegrationAutomatedTestSite';

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
