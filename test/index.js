import { expect } from 'chai';
import nock from 'nock';

const spAPIRewire = require('../src/index');
const SharePoint = spAPIRewire.SharePoint;

import { getAuthHeaders } from '../src/misc';

import { listURI as title } from '../src/lists';
import { format } from './utils';

const host = 'https://safetyculture.sharepoint.com/IntegrationAutomatedTestSite';
const mockedAuth = { FedAuth: '123', rtFa: '123', requestDigest: '123' };

describe('SharePoint-Mocked', function test() {
  this.timeout(40000);

  describe('when configured incorrectly', () => {
    it('should error on instantiation without host', () => {
      try {
        return new SharePoint();
      } catch (e) {
        expect(e.message).to.equal('SharePoint requires host string');
      }
    });

    it('should error on instantiation without auth', () => {
      try {
        return new SharePoint(host);
      } catch (e) {
        expect(e.message).to.equal('SharePoint requires auth object');
      }
    });
  });

  describe('when configured correctly', () => {
    const sp = new SharePoint(host, mockedAuth);

    afterEach(() => {
      nock.cleanAll();
    });

    it('should config the url correctly on any request', () => {
      nock(host)
          .get(format(title('dogs') + '/items?'))
          .reply(200, { d: { } });


      return sp.getListItems('dogs')
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
          'Content-Type': 'application/json;odata=verbose'
        }
      })
      .get(format(title('test') + '/items?'))
      .reply(200, { d: {} });

      return sp.getListItems('test')
      .then(() => {
        expect(nock.isDone()).to.be.true;
      });
    });

    describe('#linkLists', () => {
      it('should link two lists', () => {
        const fieldName = 'test';
        const lookupListId = 'aaaa-bbbb-cccc';
        const lookupFieldName = 'Some Field';

        nock(host)
          .post(format(title('test') + '/fields/addfield'), {
            'parameters': {
              '__metadata': { 'type': 'SP.FieldCreationInformation' },
              'Title': fieldName,
              'FieldTypeKind': 7,
              'LookupListId': lookupListId,
              'LookupFieldName': lookupFieldName.replace(/ /g, '_x0020_')
            }
          })
          .reply(200, { d: { Id: 1 } });

        return sp.linkLists('test', lookupListId, fieldName, lookupFieldName, false)
        .then(() => {
          expect(nock.isDone()).to.be.true;
        });
      });

      it('should link two lists with multiple values', () => {
        const fieldName = 'test';
        const lookupListId = 'aaaa-bbbb-cccc';
        const lookupFieldName = 'Some Field';

        nock(host)
          .post(format(title('test') + '/fields/addfield'), {
            'parameters': {
              '__metadata': { 'type': 'SP.FieldCreationInformation' },
              'Title': fieldName,
              'FieldTypeKind': 7,
              'LookupListId': lookupListId,
              'LookupFieldName': lookupFieldName.replace(/ /g, '_x0020_')
            }
          })
          .reply(200, { d: { Id: 1 } })
          .post(format(`${title('test')}/fields('1')`), {
            '__metadata': { 'type': 'SP.FieldLookup' },
            'AllowMultipleValues': true
          })
          .reply(200);

        return sp.linkLists('test', lookupListId, fieldName, lookupFieldName, true)
        .then(() => {
          expect(nock.isDone()).to.be.true;
        });
      });
    });

    describe('#getListItems', () => {
      it('should return all list items', () => {
        const testItems = [
          { id: 1 },
          { id: 2 }
        ];

        nock(host)
          .get(format(title('test') + '/items?'))
          .reply(200, { d: { results: testItems } });

        return sp.getListItems('test')
        .then(items => {
          expect(items).to.deep.equal(testItems);
          expect(nock.isDone()).to.be.true;
        });
      });

      it('should return all filtered list items', () => {
        const testItems = [
          { id: 1 }
        ];

        nock(host)
          .get(format(`${title('test')}/items?$filter=id eq '1'`))
          .reply(200, { d: { results: testItems } });

        return sp.getListItems('test', { where: { id: 1 } })
        .then(items => {
          expect(items).to.deep.equal(testItems);
          expect(nock.isDone()).to.be.true;
        });
      });
    });
  });

  describe('getAuthHeaders', () => {
    const auth = {token: '1234'};
    it('should return Authorization header if token present', () => {
      expect(getAuthHeaders(auth)).to.deep.equal({'Authorization': `Bearer ${auth.token}`});
    });

    it('should return Cookie header if no token', () => {
      expect(getAuthHeaders(mockedAuth)).to.deep.equal({'Cookie': 'FedAuth=123;rtFa=123;',
                                                      'X-RequestDigest': '123'});
    });
  });

  describe('#formatResponse', () => {
    const formatResponse = spAPIRewire.__GetDependency__('formatResponse');

    it('should replace all "_x0020_" in a response with spaces', () => {
      const response = {
        Test_x0020_Field: 'test',
        Test_x0020_Array: [
          { Test_x0020_Item: 'test'}
        ],
        Test_x0020_Object: {
          test: 'test'
        }
      };

      const expectedResult = {
        'Test Field': 'test',
        'Test Array': [
          { 'Test Item': 'test' }
        ],
        'Test Object': { test: 'test' }
      };

      expect(formatResponse(response)).to.deep.equal(expectedResult);
    });
  });
});
