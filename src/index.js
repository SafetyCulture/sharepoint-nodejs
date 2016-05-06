import _ from 'lodash';
import querystring from 'querystring';
import axios from 'axios';
import rp from 'request-promise';

import { listURI, fillSpaces } from './lists';
import { Files } from './files';
import { USER_AGENT, formatResponse, getAuthHeaders } from './misc';

// re-export
export { FIELD_TYPES } from './fields';
export { LIST_TEMPLATES, listURI, listType, fillSpaces } from './lists';
export { Batch } from './batch';
export { Authentication } from './authentication';
export { OAuth2 } from './oauth2';

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

    this.auth = auth;
    this.host = host;
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
        'Accept': 'application/json;odata=verbose',
        'User-Agent': USER_AGENT,
        'Content-Type': 'application/json;odata=verbose'
      }, getAuthHeaders(auth));

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
    let headers = _.merge(getAuthHeaders(this.auth), {
      'Accept': 'application/json;odata=verbose',
      'User-Agent': USER_AGENT,
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE',
      'Content-Type': 'application/json;odata=verbose'});

    let options = {
      headers: headers,
      method: 'POST',
      body: body,
      resolveWithFullResponse: true,
      json: true,
      uri: `${this.host}/_api/web${resource}`
    };

    return rp(options);
  }

  delete(resource) {
    const headers = {
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    };
    this._axios.post(resource, {}, headers);
  }

  createItem(list, item) {
    return this.create(`${listURI(list)}/items`, item);
  }

  updateItem(list, itemId, item) {
    return this.update(`${listURI(list)}/items(${itemId})`, fillSpaces(item));
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
