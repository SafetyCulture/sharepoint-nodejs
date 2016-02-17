import { expect } from 'chai';
// import nock from 'nock';
const filesRequire = require('../src/files');
const Files = filesRequire.Files;
import sinon from 'sinon';

describe('Files', function test() {
  this.timeout(40000);

  describe('responds', () => {
    it('correctly', () => {
      const get = sinon.stub().returns(Promise.resolve({ null }));

      let files = Files({ get });
      expect(files).to.be.a('object');
    });
  });
});
