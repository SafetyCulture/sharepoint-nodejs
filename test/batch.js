import { expect } from 'chai';
import nock from 'nock';
const spbAPIRewire = require('../src/batch');
const Batch = spbAPIRewire.Batch;
const host = 'https://safetyculture.sharepoint.com/IntegrationAutomatedTestSite';
const batchEndpoint = `/_api/$batch`;
const mockedAuth = { FedAuth: '123', rtFa: '123', requestDigest: '123' };

describe('Batch-Mocked', function test() {
  this.timeout(40000);

  before(() => {
    // Mock uuid.v4() to always return the same for testing purposes
    spbAPIRewire.__Rewire__('uuid', {
      v4: () => '1'
    });
  });

  after(() => {
    spbAPIRewire.__ResetDependency__('uuid');
  });

  describe('when configured incorrectly', () => {
    it('should error on instantiation without host', () => {
      try {
        return new Batch();
      } catch (e) {
        expect(e.message).to.equal('Batch requires host string');
      }
    });

    it('should error on instantiation without auth', () => {
      try {
        return new Batch(host);
      } catch (e) {
        expect(e.message).to.equal('Batch requires auth object');
      }
    });
  });

  describe('when configured correctly', () => {
    let sp;

    beforeEach(() => {
      sp = new Batch(host, mockedAuth);
    });

    afterEach(() => {
      nock.cleanAll();
    });

    it('should config the url correctly on any request', () => {
      nock(host)
        .post(batchEndpoint)
        .reply(200, '');

      return sp.run('test')
      .then(() => {
        expect(nock.isDone()).to.be.true;
      });
    });

    it('should config the headers correctly on any request', () => {
      nock(host, {
        reqheaders: {
          'Cookie': 'FedAuth=123;rtFa=123;',
          'X-RequestDigest': '123',
          'Accept': 'application/json;odata=verbose',
          'Content-Type': `multipart/mixed; boundary=${sp.batchBoundary}`
        }
      })
      .post(batchEndpoint)
      .reply(200, '');

      return sp.run()
      .then(() => {
        expect(nock.isDone()).to.be.true;
      });
    });

    describe('#run', () => {
      it('should batch requests and send to sharepoint', () => {
        const expectedResult =
          `--batch_1\r\nContent-Type: multipart/mixed; boundary=changeset_1\r\nHost: safetyculture.sharepoint.com\r\nContent-Length: 370\r\nContent-Transfer-Encoding: binary\r\n\r\n--changeset_1\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\n\r\nPOST https://safetyculture.sharepoint.com/IntegrationAutomatedTestSite/_api/web/lists HTTP/1.1\r\nContent-Type: application/json;odata=verbose\r\n\r\n{"__metadata":{"type":"SP.List"},"AllowContentTypes":true,"BaseTemplate":100,"ContentTypesEnabled":true,"Description":"test","Title":"test"}\r\n\r\n--changeset_1--\r\n\r\n--batch_1\r\nContent-Type: multipart/mixed; boundary=changeset_1\r\nHost: safetyculture.sharepoint.com\r\nContent-Length: 323\r\nContent-Transfer-Encoding: binary\r\n\r\n--changeset_1\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\n\r\nPOST https://safetyculture.sharepoint.com/IntegrationAutomatedTestSite/_api/web/lists/GetByTitle('test')/fields HTTP/1.1\r\nContent-Type: application/json;odata=verbose\r\n\r\n{"__metadata":{"type":"SP.Field"},"Title":"test","FieldTypeKind":1}\r\n\r\n--changeset_1--\r\n\r\n--batch_1--`;

        sp.createList('test', 'test', { 'test': 1 });

        nock(host)
          .post(batchEndpoint, expectedResult)
          .reply(200, '');

        return sp.run()
        .then(() => {
          expect(nock.isDone()).to.be.true;
        });
      });

      it('should resolve with an array of responses', () => {
        const exampleListId = '25f7c47f-12bb-429f-b99c-fbfb59c8f3d8';
        const exampleResponse =
          `--batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178
          Content-Type: application/http
          Content-Transfer-Encoding: binary

          HTTP/1.1 201 Created
          CONTENT-TYPE: application/atom+xml;type=entry;charset=utf-8
          ETAG: "1"
          LOCATION:${host}/_api/Web/Lists/test

          <?xml version="1.0" encoding="utf-8"?><entry xml:base="${host}/_api/" xmlns="http://www.w3.org/2005/Atom" xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" xmlns:georss="http://www.georss.org/georss" xmlns:gml="http://www.opengis.net/gml" m:etag="&quot;1&quot;"><id>${exampleListId}</id></entry>
          --batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178
          Content-Type: application/http
          Content-Transfer-Encoding: binary

          --batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178
          Content-Type: application/http
          Content-Transfer-Encoding: binary

          HTTP/1.1 200 OK
          CONTENT-TYPE: application/json;odata=verbose;charset=utf-8

          {"d":{"__metadata":{"id":"${exampleListId}"}}}
          --batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178--`;

        nock(host)
          .post(batchEndpoint)
          .reply(200, exampleResponse);

        return sp.run()
        .then(res => {
          expect(res[0].entry.id[0]).to.equal(exampleListId);
          expect(res[1].d.__metadata.id).to.equal(exampleListId);
        });
      });

      it('should reset the instance requests', () => {
        sp.createList('test', 'test', { 'test': 1 });

        nock(host)
          .post(batchEndpoint)
          .reply(200, '');

        return sp.run()
        .then(() => {
          expect(sp.requests).to.be.empty;
        });
      });
    });

    describe('#addChangeset', () => {
      it('should push a set of changes into request', () => {
        const testObject = { test: 'test' };
        const otherObject = { other: 'other' };

        sp.addChangeset([
          sp.createListItem('test', testObject),
          sp.createListItem('test', otherObject)
        ]);

        expect(sp.requests[0]).to.include;

        expect(sp.requests[0]).to.include(JSON.stringify(testObject));
        expect(sp.requests[0]).to.include(JSON.stringify(otherObject));
      });
    });

    describe('#createList', () => {
      it('should add a changeset that creates a list then adds each field', () => {
        const createListRequest =
          `{"__metadata":{"type":"SP.List"},"AllowContentTypes":true,"BaseTemplate":100,"ContentTypesEnabled":true,"Description":"test","Title":"test"}`;

        const createFieldRequest =
          `{"__metadata":{"type":"SP.Field"},"Title":"test","FieldTypeKind":1}`;

        sp.createList('test', 'test', { 'test': 1 });

        expect(sp.requests[0]).to.include(createListRequest);
        expect(sp.requests[1]).to.include(createFieldRequest);
      });
    });

    describe('#_addGet', () => {
      it('should add a get request to requests', () => {
        const getRequest = `GET ${host}/_api/web/test`;

        sp._addGet('/test');

        expect(sp.requests[0]).to.include(getRequest);
      });
    });

    describe('#_change', () => {
      it('should return a function', () => {
        const result = sp._change();

        expect(result).to.be.a('function');
      });

      it('should generate a change request with the returned function', () => {
        const expected = '--1\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\nContent-ID: 1\r\n\r\ntest';
        const fn = sp._change('test', 1);
        const result = fn(1);

        expect(result).to.equal(expected);
      });
    });

    describe('#_create', () => {
      it('should return a create change body', () => {
        const expected =
        `POST ${host}/_api/web/test HTTP/1.1\r\nContent-Type: application/json;odata=verbose\r\n\r\n{"test":"test"}\r\n`;

        const result = sp._create('/test', {test: 'test'});

        expect(result).to.equal(expected);
      });
    });

    describe('#_update', () => {
      it('should return an update change body', () => {
        const expected =
        `PATCH https://safetyculture.sharepoint.com/IntegrationAutomatedTestSite/_api/web/test HTTP/1.1\r\nContent-Type: application/json;odata=verbose\r\nAccept: application/json;odata=verbose\r\nIf-Match: "1"\r\n\r\n{"test":"test"}\r\n`;

        const result = sp._update('/test', {test: 'test'});

        expect(result).to.equal(expected);
      });
    });

    describe('#_delete', () => {
      it('should return an update change body', () => {
        const expected =
        `DELETE https://safetyculture.sharepoint.com/IntegrationAutomatedTestSite/_api/web/test HTTP/1.1\r\nIf-Match: *\r\n`;

        const result = sp._delete('/test');

        expect(result).to.equal(expected);
      });
    });

    describe('#getList', () => {
      it('should add a changeset that gets a list', () => {
        const getRequest = `GET ${host}/_api/web/lists/GetByTitle(\'test\')`;

        sp.getList('test');

        expect(sp.requests[0]).to.include(getRequest);
      });
    });

    describe('#createListItem', () => {
      const testItem = { test: 'test' };

      it('should return a function', () => {
        const result = sp.createListItem('test', testItem);
        expect(result).to.be.a('function');
      });

      it('should return a function that generates the right request', () => {
        const fn = sp.createListItem('test', testItem);
        const result = fn('test');
        expect(result).to.include(JSON.stringify(testItem));
      });
    });

    describe('#deleteListItem', () => {
      it('should return a function', () => {
        const result = sp.deleteListItem('test', 1);
        expect(result).to.be.a('function');
      });

      it('should return a function that generates the right request', () => {
        const fn = sp.deleteListItem('test', 1);
        const result = fn('test');
        expect(result).to.include(`DELETE ${host}/_api/web/lists/GetByTitle('test')/items(1)`);
      });
    });
  });
});
