import url from 'url';
import { assign, map, identity } from 'lodash';
import axios from 'axios';
import uuid from 'node-uuid';
import { parseString } from 'xml2js';

import { LIST_TEMPLATES, STD_NEWLINE_REGEX, AFFIXED_NEWLINE_REGEX, listURI, fillSpaces } from './lists';
import { getAuthHeaders } from './misc';

/**
* Batch class
* @param {string} host Host for SharePoint
* @param {object} auth Auth object for SharePoint
* @returns {object} SharePoint Batch instance
*/
export class Batch {
  constructor(host, auth) {
    if (!host) throw new Error('Batch requires host string');
    if (!auth) throw new Error('Batch requires auth object');

    // remove trailing slash
    this.host = host.replace(/\/$/, '');
    this.hostname = url.parse(host).hostname;
    this.batchBoundary = `batch_${uuid.v4()}`;
    this.requests = [];
    this.auth = auth;
  }

  /**
  * Runs batch
  * @returns {Promise} Resolves with responses array in order of changes, rejects with error from SP
  */
  run() {
    this.requests.push(`--${this.batchBoundary}--`);
    const data = this.requests.join('\r\n');

    return axios.post(`${this.host}/_api/$batch`, data, {
      headers: assign({}, getAuthHeaders(this.auth), {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': `multipart/mixed; boundary=${this.batchBoundary}`
      }),
      timeout: 360000
    })
    .then(res => {
      return Promise.all(res.data.split(`--batchresponse`).map(batchRes => {
        return new Promise((resolve, reject) => {
          const xmlStart = batchRes.indexOf('<?xml');
          if (xmlStart >= 0) {
            parseString(batchRes.slice(xmlStart), (err, result) => {
              if (err) return reject(err);
              resolve(result);
            });
          } else {
            const bodyStart = batchRes.indexOf('{');
            if (bodyStart > 0) return resolve(JSON.parse(batchRes.slice(bodyStart)));
            resolve();
          }
        });
      }))
      .then(responses => {
        this.requests = [];

        return responses.filter(identity);
      });
    });
  }

  /**
  * Adds a list of changes to one changeset and then adds to requests
  * @param {array} changes Array of change functions to add to a single changeset
  */
  addChangeset(changes) {
    const changesetBoundary = `changeset_${uuid.v4()}`;
    const changesBody =
      map(changes, change => change(changesetBoundary)).join('\r\n');

    this.requests.push([
      `--${this.batchBoundary}`,
      `Content-Type: multipart/mixed; boundary=${changesetBoundary}`,
      `Host: ${this.hostname}`,
      `Content-Length: ${changesBody.length}`,
      `Content-Transfer-Encoding: binary`,
      ``,
      changesBody,
      `--${changesetBoundary}--`,
      ``
    ].join('\r\n'));
  }

  /**
  * Adds a get request to requests
  * @param {string} resource Endpoint for resource to get
  */
  _addGet(resource) {
    this.requests.push([
      `--${this.batchBoundary}`,
      `Content-Type: application/http`,
      `Content-Transfer-Encoding: binary`,
      ``,
      `GET ${this.host}/_api/web${resource} HTTP/1.1`,
      `Host: ${this.hostname}`,
      `Accept: application/json;odata=verbose`,
      ``
    ].join(`\r\n`));
  }

  /**
  * Generates a change
  * @param {object} body Resource body to create
  * @param {integer} id Id to reference change in further changes
  * @returns {function} Change function to pass to addChangeset
  */
  _change(body, id) {
    return changesetBoundary => {
      let headers = [
        `--${changesetBoundary}`,
        `Content-Type: application/http`,
        `Content-Transfer-Encoding: binary`
      ];

      if (id) headers.push(`Content-ID: ${id}`);

      headers.push(``);
      headers.push(body);
      return headers.join(`\r\n`);
    };
  }

  /**
  * Generates a create body
  * @param {string} resource Resource endpoint to create
  * @param {object} body Resource body to create
  * @returns {string} Request body to create resource
  */
  _create(resource, body) {
    return [
      `POST ${this.host}/_api/web${resource} HTTP/1.1`,
      `Content-Type: application/json;odata=verbose`,
      ``,
      JSON.stringify(fillSpaces(body)),
      ``
    ].join(`\r\n`);
  }

  _delete(resource, body) {
    return [
      `POST ${this.host}/_api/web${resource} HTTP/1.1`,
      `Content-Type: application/json;odata=verbose`,
      `IF-MATCH: etag or "*"`,
      `X-HTTP-Method: DELETE`,
      ``,
      JSON.stringify(fillSpaces(body)),
      ``
    ].join(`\r\n`);
  }

  /**
  * Generates an update body
  * @param {string} resource Resource endpoint to update
  * @param {object} body Updated resource
  * @returns {string} Request body to update resource
  */
  _update(resource, body) {
    return [
      `PATCH ${this.host}/_api/web${resource} HTTP/1.1`,
      `Content-Type: application/json;odata=verbose`,
      `Accept: application/json;odata=verbose`,
      `If-Match: "1"`,
      ``,
      JSON.stringify(fillSpaces(body)),
      ``
    ].join(`\r\n`);
  }

  /**
  * Generates a delete body
  * @param {string} resource Resource endpoint to delete
  * @returns {string} Request body to delete resource
  */
  _delete(resource) {
    return [
      `DELETE ${this.host}/_api/web${resource} HTTP/1.1`,
      `If-Match: *`,
      ``
    ].join(`\r\n`);
  }

  /**
  * Creates a list on sharepoint
  * @param {string} title Sharepoint title of the list
  * @param {string} description Sharepoint description of the list
  * @param {object} fields Mapping of fieldname to fieldtype for sharepoint
  */
  createList(title, description, fields, baseTemplate = LIST_TEMPLATES.STANDARD) {
    this.addChangeset([
      this._change(this._create('/lists', {
        '__metadata': { 'type': 'SP.List' },
        'AllowContentTypes': true,
        'BaseTemplate': baseTemplate,
        'ContentTypesEnabled': true,
        'Description': description,
        'Title': title
      }))
    ]);

    this.addChangeset(map(fields, (fieldType, field) =>
      this._change(this._create(`${listURI(title)}/fields`, {
        '__metadata': { 'type': 'SP.Field' },
        'Title': field.replace(AFFIXED_NEWLINE_REGEX, '').replace(STD_NEWLINE_REGEX, ' '),
        'FieldTypeKind': fieldType
      }))
    ));
  }

  /**
  * Get list by title on SharePoint
  * @param {string} title Sharepoint title of the list
  */
  getList(title) {
    this._addGet(`${listURI(title)}`);
  }

  /**
  * Creates a list item on SharePoint
  * @param {string} list Target list title
  * @param {object} item Item to add to list
  * @returns {function} change Function to pass into addChangeset() array
  */
  createListItem(list, item) {
    return this._change(this._create(`${listURI(list)}/items`, item));
  }


  /**
  * Deletes a list item on SharePoint
  * @param {string} list Target list title
  * @param {string} id Item id to delete
  * @returns {function} change Function to pass to into addChangeset() array
  */
  deleteListItem(list, id) {
    return this._change(this._delete(`${listURI(list)}/items(${id})`));
  }
}
