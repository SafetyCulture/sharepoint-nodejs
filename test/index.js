import { expect } from 'chai';
import nock from 'nock';
const spRewire = require('../../src/sharepoint/index');
const SharePoint = spRewire.default;
import { format, title } from '../utils';

const username = 'testusername';
const password = 'testpassword';
const host = 'https://ohaiprismatik.sharepoint.com/wut';
const batchEndpoint = `/_api/$batch`;

describe('SharePoint-Mock', function test() {
  this.timeout(40000);

  before(() => {
    // mock spAuth to always log in
    spRewire.__Rewire__('spAuth', (opts, fn) => fn(null, {
      cookies: {
        FedAuth: '123',
        rtFa: '123'
      },
      requestDigest: '123'
    }));
  });

  after(() => {
    spRewire.__ResetDependency__('spAuth');
  });

  describe('#ensureLists', () => {
    it('should set ensured to true if both lists exist', () => {
      const sp = new SharePoint({ username, password, host });

      nock(host)
        .post(batchEndpoint)
        .reply(200, '');

      return sp.ensureLists()
      .then(() => {
        expect(sp.ensured).to.be.true;
      });
    });

    const exampleResponseNotFound =
      `--batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178
      Content-Type: application/http
      Content-Transfer-Encoding: binary

      HTTP/1.1 404 ERROR
      CONTENT-TYPE: application/json;odata=verbose;charset=utf-8

      {"error":{"message": "does not exist"}}

      --batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178
      Content-Type: application/http
      Content-Transfer-Encoding: binary

      HTTP/1.1 200 OK
      CONTENT-TYPE: application/json;odata=verbose;charset=utf-8

      {"d":{"__metadata":{"id":"2"}}}
      --batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178--`;

    const exampleCreateResponse =
      `--batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178
      Content-Type: application/http
      Content-Transfer-Encoding: binary

      HTTP/1.1 200 OK
      CONTENT-TYPE: application/json;odata=verbose;charset=utf-8

      { "entry": { "category": [ {"$": { "term": "SP.List" } } ] } }

      --batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178
      Content-Type: application/http
      Content-Transfer-Encoding: binary

      HTTP/1.1 200 OK
      CONTENT-TYPE: application/json;odata=verbose;charset=utf-8

      { "entry": { "category": [ {"$": { "term": "SP.List" } } ], "content": [{"m:properties": [{"d:Id": [{"_": 1}]}]}] } }
      --batchresponse_4cc2e5b3-bf5b-42d7-9de3-9a381313d178--`;

    it('should create lists if they do not exist, and link them', () => {
      const auditsList = 'SafetyCulture Audits';
      const sp = new SharePoint({ username, password, host });

      nock(host)
        // Query
        .post(batchEndpoint)
        .reply(200, exampleResponseNotFound)
        // CreateLists
        .post(batchEndpoint)
        .reply(200, exampleCreateResponse)
        .post(format(title(auditsList) + '/fields/addfield'), {
          'parameters': {
            '__metadata': { 'type': 'SP.FieldCreationInformation' },
            'Title': 'Items',
            'FieldTypeKind': 7,
            'LookupListId': 1,
            'LookupFieldName': 'Item_x0020_Id'
          }
        })
        .reply(200, { d: { Id: 1 } })
        .post(format(title(auditsList) + `/fields('1')`), {
          '__metadata': { 'type': 'SP.FieldLookup' },
          'AllowMultipleValues': true
        })
        .reply(200);


      return sp.ensureLists()
      .then(() => {
        expect(nock.isDone()).to.be.true;
      });
    });

    it('custom list name', () => {
      const auditsList = 'QA Results';
      const sp = new SharePoint({ username, password, host, auditsList });

      nock(host)
        // Query
        .post(batchEndpoint)
        .reply(200, exampleResponseNotFound)
        // CreateLists
        .post(batchEndpoint)
        .reply(200, exampleCreateResponse)
        .post(format(title(auditsList) + '/fields/addfield'), {
          'parameters': {
            '__metadata': { 'type': 'SP.FieldCreationInformation' },
            'Title': 'Items',
            'FieldTypeKind': 7,
            'LookupListId': 1,
            'LookupFieldName': 'Item_x0020_Id'
          }
        })
        .reply(200, { d: { Id: 1 } })
        .post(format(title(auditsList) + `/fields('1')`), {
          '__metadata': { 'type': 'SP.FieldLookup' },
          'AllowMultipleValues': true
        })
        .reply(200);


      return sp.ensureLists()
      .then(() => {
        expect(nock.isDone()).to.be.true;
      });
    });
  });

  describe('#createOrUpdate', () => {
    // Assume lists ensured
    let sp = new SharePoint({ username, password, host });
    const auditsList = 'SafetyCulture Audits';
    sp.ensured = true;

    it('should upload an audit with items to SharePoint', () => {
      const testAudit = { 'Audit Id': 1 };
      const testItem = { 'Item Id': 1 };

      nock(host)
        .get(format(title(auditsList) + `/items?$filter=Audit_x0020_Id eq '1'`))
        .reply(200, { d: { results: [] } })
        // Create items
        .post(batchEndpoint)
        .reply(200)
        // Create audit
        .post(batchEndpoint)
        .reply(200);

      return sp.createOrUpdate(testAudit, [testItem])
      .then(() => {
        expect(nock.isDone()).to.be.true;
      });
    });

    it('should delete an audit and items on SharePoint if they already exist', () => {
      const testAudit = { 'Audit Id': 1 };
      const testItem = { 'Item Id': 1 };

      nock(host)
        .get(format(title(auditsList) + `/items?$filter=Audit_x0020_Id eq '1'`))
        .reply(200, { d: { results: [ { Id: 1, ItemsId: { results: [1] } } ] } })
        // the item delete
        .post(batchEndpoint)
        .reply(200)
        // the audit delete
        .post(batchEndpoint)
        .reply(200)
        // Create items
        .post(batchEndpoint)
        .reply(200)
        // Create audit
        .post(batchEndpoint)
        .reply(200);

      return sp.createOrUpdate(testAudit, [testItem])
      .then(() => {
        expect(nock.isDone()).to.be.true;
      });
    });
  });
});
