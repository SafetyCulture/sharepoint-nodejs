'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.Batch = undefined;

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _url = require('url');

var _url2 = _interopRequireDefault(_url);

var _lodash = require('lodash');

var _axios = require('axios');

var _axios2 = _interopRequireDefault(_axios);

var _nodeUuid = require('node-uuid');

var _nodeUuid2 = _interopRequireDefault(_nodeUuid);

var _xml2js = require('xml2js');

var _lists = require('./lists');

var _misc = require('./misc');

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

/**
* Batch class
* @param {string} host Host for SharePoint
* @param {object} auth Auth object for SharePoint
* @returns {object} SharePoint Batch instance
*/
var Batch = exports.Batch = function () {
  function Batch(host, auth) {
    _classCallCheck(this, Batch);

    if (!host) throw new Error('Batch requires host string');
    if (!auth) throw new Error('Batch requires auth object');

    // remove trailing slash
    this.host = host.replace(/\/$/, '');
    this.hostname = _url2.default.parse(host).hostname;
    this.batchBoundary = 'batch_' + _nodeUuid2.default.v4();
    this.requests = [];
    this.auth = auth;
  }

  /**
  * Runs batch
  * @returns {Promise} Resolves with responses array in order of changes, rejects with error from SP
  */


  _createClass(Batch, [{
    key: 'run',
    value: function run() {
      var _this = this;

      this.requests.push('--' + this.batchBoundary + '--');
      var data = this.requests.join('\r\n');

      return _axios2.default.post(this.host + '/_api/$batch', data, {
        headers: (0, _lodash.assign)({}, (0, _misc.getAuthHeaders)(this.auth), {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'multipart/mixed; boundary=' + this.batchBoundary
        }),
        timeout: 360000
      }).then(function (res) {
        return Promise.all(res.data.split('--batchresponse').map(function (batchRes) {
          return new Promise(function (resolve, reject) {
            var xmlStart = batchRes.indexOf('<?xml');
            if (xmlStart >= 0) {
              (0, _xml2js.parseString)(batchRes.slice(xmlStart), function (err, result) {
                if (err) return reject(err);
                resolve(result);
              });
            } else {
              var bodyStart = batchRes.indexOf('{');
              if (bodyStart > 0) return resolve(JSON.parse(batchRes.slice(bodyStart)));
              resolve();
            }
          });
        })).then(function (responses) {
          _this.requests = [];

          return responses.filter(_lodash.identity);
        });
      });
    }

    /**
    * Adds a list of changes to one changeset and then adds to requests
    * @param {array} changes Array of change functions to add to a single changeset
    */

  }, {
    key: 'addChangeset',
    value: function addChangeset(changes) {
      var changesetBoundary = 'changeset_' + _nodeUuid2.default.v4();
      var changesBody = (0, _lodash.map)(changes, function (change) {
        return change(changesetBoundary);
      }).join('\r\n');

      this.requests.push(['--' + this.batchBoundary, 'Content-Type: multipart/mixed; boundary=' + changesetBoundary, 'Host: ' + this.hostname, 'Content-Length: ' + changesBody.length, 'Content-Transfer-Encoding: binary', '', changesBody, '--' + changesetBoundary + '--', ''].join('\r\n'));
    }

    /**
    * Adds a get request to requests
    * @param {string} resource Endpoint for resource to get
    */

  }, {
    key: '_addGet',
    value: function _addGet(resource) {
      this.requests.push(['--' + this.batchBoundary, 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '', 'GET ' + this.host + '/_api/web' + resource + ' HTTP/1.1', 'Host: ' + this.hostname, 'Accept: application/json;odata=verbose', ''].join('\r\n'));
    }

    /**
    * Generates a change
    * @param {object} body Resource body to create
    * @param {integer} id Id to reference change in further changes
    * @returns {function} Change function to pass to addChangeset
    */

  }, {
    key: '_change',
    value: function _change(body, id) {
      return function (changesetBoundary) {
        var headers = ['--' + changesetBoundary, 'Content-Type: application/http', 'Content-Transfer-Encoding: binary'];

        if (id) headers.push('Content-ID: ' + id);

        headers.push('');
        headers.push(body);
        return headers.join('\r\n');
      };
    }

    /**
    * Generates a create body
    * @param {string} resource Resource endpoint to create
    * @param {object} body Resource body to create
    * @returns {string} Request body to create resource
    */

  }, {
    key: '_create',
    value: function _create(resource, body) {
      return ['POST ' + this.host + '/_api/web' + resource + ' HTTP/1.1', 'Content-Type: application/json;odata=verbose', '', JSON.stringify((0, _lists.fillSpaces)(body)), ''].join('\r\n');
    }
  }, {
    key: '_delete',
    value: function _delete(resource, body) {
      return ['POST ' + this.host + '/_api/web' + resource + ' HTTP/1.1', 'Content-Type: application/json;odata=verbose', 'IF-MATCH: etag or "*"', 'X-HTTP-Method: DELETE', '', JSON.stringify((0, _lists.fillSpaces)(body)), ''].join('\r\n');
    }

    /**
    * Generates an update body
    * @param {string} resource Resource endpoint to update
    * @param {object} body Updated resource
    * @returns {string} Request body to update resource
    */

  }, {
    key: '_update',
    value: function _update(resource, body) {
      return ['PATCH ' + this.host + '/_api/web' + resource + ' HTTP/1.1', 'Content-Type: application/json;odata=verbose', 'Accept: application/json;odata=verbose', 'If-Match: "1"', '', JSON.stringify((0, _lists.fillSpaces)(body)), ''].join('\r\n');
    }

    /**
    * Generates a delete body
    * @param {string} resource Resource endpoint to delete
    * @returns {string} Request body to delete resource
    */

  }, {
    key: '_delete',
    value: function _delete(resource) {
      return ['DELETE ' + this.host + '/_api/web' + resource + ' HTTP/1.1', 'If-Match: *', ''].join('\r\n');
    }

    /**
    * Creates a list on sharepoint
    * @param {string} title Sharepoint title of the list
    * @param {string} description Sharepoint description of the list
    * @param {object} fields Mapping of fieldname to fieldtype for sharepoint
    */

  }, {
    key: 'createList',
    value: function createList(title, description, fields) {
      var _this2 = this;

      var baseTemplate = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : _lists.LIST_TEMPLATES.STANDARD;

      this.addChangeset([this._change(this._create('/lists', {
        '__metadata': { 'type': 'SP.List' },
        'AllowContentTypes': true,
        'BaseTemplate': baseTemplate,
        'ContentTypesEnabled': true,
        'Description': description,
        'Title': title
      }))]);

      this.addChangeset((0, _lodash.map)(fields, function (fieldType, field) {
        return _this2._change(_this2._create((0, _lists.listURI)(title) + '/fields', {
          '__metadata': { 'type': 'SP.Field' },
          'Title': field.replace(_lists.AFFIXED_NEWLINE_REGEX, '').replace(_lists.STD_NEWLINE_REGEX, ' '),
          'FieldTypeKind': fieldType
        }));
      }));
    }

    /**
    * Get list by title on SharePoint
    * @param {string} title Sharepoint title of the list
    */

  }, {
    key: 'getList',
    value: function getList(title) {
      this._addGet('' + (0, _lists.listURI)(title));
    }

    /**
    * Creates a list item on SharePoint
    * @param {string} list Target list title
    * @param {object} item Item to add to list
    * @returns {function} change Function to pass into addChangeset() array
    */

  }, {
    key: 'createListItem',
    value: function createListItem(list, item) {
      return this._change(this._create((0, _lists.listURI)(list) + '/items', item));
    }

    /**
    * Deletes a list item on SharePoint
    * @param {string} list Target list title
    * @param {string} id Item id to delete
    * @returns {function} change Function to pass to into addChangeset() array
    */

  }, {
    key: 'deleteListItem',
    value: function deleteListItem(list, id) {
      return this._change(this._delete((0, _lists.listURI)(list) + '/items(' + id + ')'));
    }
  }]);

  return Batch;
}();