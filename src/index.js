import _ from 'lodash';
import querystring from 'querystring';
import axios from 'axios';

import { listURI } from './lists';
import { Files } from './files';

// re-export
export { FIELD_TYPES } from './fields';
export { LIST_TEMPLATES, listURI, listType } from './lists';
export { Batch } from './batch';
export { Authentication } from './authentication';

// Small helper to replace spaces in keys with '_x0020_' within an object
const fillSpaces = data =>
  _.mapKeys(data, (val, key) => key.replace(/ /g, '_x0020_'));

/**
* Formats a response to replace '_x0020_' with spaces.
* @param {object} res Response to deep replace
* @returns {object} Formatted response
*/
function formatResponse(res) {
  return _.transform(res, (result, val, key) => {
    let newVal = val;
    const newKey = _.isString(key) ? key.replace(/_x0020_/g, ' ') : key;

    if (_.isArray(val)) newVal = _.map(val, formatResponse);
    if (_.isObject(val)) newVal = formatResponse(val);

    result[newKey] = newVal;
  });
}

/**
* SharePoint class
* @param {string} host Host for SharePoint
* @param {object} auth Auth object for SharePoint
* @returns {object} SharePoint instance
*/
export class SharePoint {
  constructor(host, auth) {
    if (!host) throw new Error('SharePoint requires host string');
    if (!auth) throw new Error('SharePoint requires auth object');

    this._axios = this._configureInterceptors(axios.create(), { host, auth });
    this.files = Files(this);
  }

  /**
  * Returns axios instance configured with auth details
  * @param {object} instance Axios instance
  * @param {object} options Options to pass to spAuth
  * @returns {object} Axios instance
  */
  _configureInterceptors(instance, { host, auth }) {
    instance.interceptors.request.use(config => {
      config.url = `${host}/_api/web${config.url}`;
      config.headers = _.assign({}, config.headers, {
        'Cookie': `FedAuth=${auth.FedAuth};rtFa=${auth.rtFa};`,
        'X-RequestDigest': auth.requestDigest,
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      });

      config.timeout = 360000;

      return config;
    });

    instance.interceptors.response.use(res => {
      if (res.data && res.data.d) {
        res.data.d = formatResponse(res.data.d);
      }
      return res;
    });

    return instance;
  }

  /**
  * Link two fields together within SharePoint
  * @param {string} list Target list title
  * @param {integer} lookupListId List to link's SharePoint id
  * @param {string} fieldName Name of the field on target list
  * @param {string} lookupFieldName Name of the field on lookup list
  * @param {boolean} multiValues True for many links per entry, false for one to one
  * @returns {Promise} Resolves on success, rejects with error from SP
  */
  linkLists(list, lookupListId, fieldName, lookupFieldName, multiValues) {
    return this._axios.post(`${listURI(list)}/fields/addfield`, {
      'parameters': {
        '__metadata': { 'type': 'SP.FieldCreationInformation' },
        'Title': fieldName,
        'FieldTypeKind': 7,
        'LookupListId': lookupListId,
        'LookupFieldName': lookupFieldName.replace(/ /g, '_x0020_')
      }
    })
    .then(res => {
      if (!multiValues) return Promise.resolve();

      return this._axios.post(`${listURI(list)}/fields('${res.data.d.Id}')`, {
        '__metadata': { 'type': 'SP.FieldLookup' },
        'AllowMultipleValues': true
      }, { headers: { 'X-HTTP-Method': 'MERGE' } });
    });
  }


  create(resource, body) {
    return this._axios.post(`${resource}`, body, {});
  }

  update(resource, body) {
    let headers = {headers: {'IF-MATCH': 'etag or "*"',
                             'X-HTTP-Method': 'MERGE'}};
    return this._axios.post(`${resource}`,
                            body,
                            headers);
  }

  createItem(list, item) {
    return this.create(`${listURI(list)}/items`, item);
  }

  updateItem(list, itemId, item) {
    let url = `${listURI(list)}/items/(${itemId})`;
    return this.update(url, item);
  }

  getListType(list) {
    return this._axios.post(`${listURI(list)}?$select=ListItemEntityTypeFullName`, {}, {})
                      .then((response) => {
                        return response.data.d.ListItemEntityTypeFullName;
                      });
  }

  /**
  * Queries for list items on SharePoint
  * @param {string} list Target list title
  * @param {object} options
  * @param {object} options.where Mapping of field to their queried value
  * @returns {Promise} Resolves with Items on success, rejects with error from SP
  */
  getListItems(list, { where } = {}) {
    let query = {};

    if (where) {
      query.$filter =
        _.map(fillSpaces(where), (value, field) => `${field} eq '${value}'`)
        .join(' and ');
    }

    const qs = querystring.stringify(query);
    return this._axios.get(`${listURI(list)}/items?${qs}`)
                      .then(res => res.data.d.results);
  }
}
