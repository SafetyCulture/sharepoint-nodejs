'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _lodash = require('lodash');

var _lodash2 = _interopRequireDefault(_lodash);

var _querystring = require('querystring');

var _querystring2 = _interopRequireDefault(_querystring);

var _axios = require('axios');

var _axios2 = _interopRequireDefault(_axios);

var _path = require('path');

var _path2 = _interopRequireDefault(_path);

var _fs = require('fs');

var _fs2 = _interopRequireDefault(_fs);

var _bluebird = require('bluebird');

var _bluebird2 = _interopRequireDefault(_bluebird);

var _lists = require('./lists');

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var readFile = _bluebird2.default.promisify(_fs2.default.readFile);

// String builder for list URIS
var folderUrl = function folderUrl(title) {
  return 'GetFolderByServerRelativeUrl(\'' + title + '\')';
};

// Small helper to replace spaces in keys with '_x0020_' within an object
var fillSpaces = function fillSpaces(data) {
  return _lodash2.default.mapKeys(data, function (val, key) {
    return key.replace(/ /g, '_x0020_');
  });
};

/**
* Formats a response to replace '_x0020_' with spaces.
* @param {object} res Response to deep replace
* @returns {object} Formatted response
*/
function formatResponse(res) {
  return _lodash2.default.transform(res, function (result, val, key) {
    var newVal = val;
    var newKey = _lodash2.default.isString(key) ? key.replace(/_x0020_/g, ' ') : key;

    if (_lodash2.default.isArray(val)) newVal = _lodash2.default.map(val, formatResponse);
    if (_lodash2.default.isObject(val)) newVal = formatResponse(val);

    result[newKey] = newVal;
  });
}

/**
* SharePointAPI class
* @param {string} host Host for SharePoint
* @param {object} auth Auth object for SharePoint
* @returns {object} SharePoint instance
*/

var SharePointAPI = function () {
  function SharePointAPI(host, auth) {
    _classCallCheck(this, SharePointAPI);

    if (!host) throw new Error('SharePointAPI requires host string');
    if (!auth) throw new Error('SharePointAPI requires auth object');

    this._axios = this._configureInterceptors(_axios2.default.create(), { host: host, auth: auth });
  }

  /**
  * Returns axios instance configured with auth details
  * @param {object} instance Axios instance
  * @param {object} options Options to pass to spAuth
  * @returns {object} Axios instance
  */


  _createClass(SharePointAPI, [{
    key: '_configureInterceptors',
    value: function _configureInterceptors(instance, _ref) {
      var host = _ref.host;
      var auth = _ref.auth;

      instance.interceptors.request.use(function (config) {
        config.url = host + '/_api/web' + config.url;
        config.headers = _lodash2.default.assign({}, config.headers, {
          'Cookie': 'FedAuth=' + auth.FedAuth + ';rtFa=' + auth.rtFa + ';',
          'X-RequestDigest': auth.requestDigest,
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        });

        config.timeout = 360000;

        return config;
      });

      instance.interceptors.response.use(function (res) {
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

  }, {
    key: 'linkLists',
    value: function linkLists(list, lookupListId, fieldName, lookupFieldName, multiValues) {
      var _this = this;

      return this._axios.post((0, _lists.listURI)(list) + '/fields/addfield', {
        'parameters': {
          '__metadata': { 'type': 'SP.FieldCreationInformation' },
          'Title': fieldName,
          'FieldTypeKind': 7,
          'LookupListId': lookupListId,
          'LookupFieldName': lookupFieldName.replace(/ /g, '_x0020_')
        }
      }).then(function (res) {
        if (!multiValues) return _bluebird2.default.resolve();

        return _this._axios.post((0, _lists.listURI)(list) + '/fields(\'' + res.data.d.Id + '\')', {
          '__metadata': { 'type': 'SP.FieldLookup' },
          'AllowMultipleValues': true
        }, { headers: { 'X-HTTP-Method': 'MERGE' } });
      });
    }

    /**
     * Upload a file to Sharepoint
     *
     * @param {string} list The name of the list to upload the file too
     * @param {string} filePath The location of the file to upload
     * @param {string} folderName The name of the folder in the lsit to upload too
     * @returns {Promise} Resolves on success, rejects with error from SP
     */

  }, {
    key: 'uploadFile',
    value: function uploadFile(list, fileLocation) {
      var _this2 = this;

      var folderName = arguments.length <= 2 || arguments[2] === undefined ? null : arguments[2];

      var fileName = _path2.default.basename(fileLocation);
      var folder = folderName ? folderUrl(folderName) : 'RootFolder';

      return readFile(fileLocation).then(function (data) {
        var headers = { 'content-length': data.length };
        return _this2._axios.post((0, _lists.listURI)(list) + '/' + folder + '/Files/Add(url=\'' + fileName + '\', overwrite=true)?$expand=ListItemAllFields', data, { headers: headers }).then(function (response) {
          return response.data;
        });
      });
    }

    /**
     * Add file to a list item
     *
     * @param {string} list The name of the list to upload the file too
     * @param {string} itemId The id of the item to attach the file too
     * @param {string} filePath The location of the file to upload
     * @returns {Promise} Resolves on success, rejects with error from SP
     */

  }, {
    key: 'attachFileToItem',
    value: function attachFileToItem(list, itemId, fileLocation) {
      var _this3 = this;

      var fileName = _path2.default.basename(fileLocation);

      return readFile(fileLocation).then(function (data) {
        var headers = { 'content-length': data.length };
        return _this3._axios.post((0, _lists.listURI)(list) + '/items({' + itemId + ')/Files/Add(FileName=\'' + fileName + '\', overwrite=true)', data, { headers: headers });
      });
    }
  }, {
    key: 'create',
    value: function create(resource, body) {
      return this._axios.post('' + resource, body, {});
    }
  }, {
    key: 'update',
    value: function update(resource, body) {
      var headers = { headers: { 'IF-MATCH': 'etag or "*"',
          'X-HTTP-Method': 'MERGE' } };
      return this._axios.post('' + resource, body, headers);
    }
  }, {
    key: 'createItem',
    value: function createItem(list, item) {
      return this.create((0, _lists.listURI)(list) + '/items', item);
    }
  }, {
    key: 'updateItem',
    value: function updateItem(list, itemId, item) {
      var url = (0, _lists.listURI)(list) + '/items/(' + itemId + ')';
      return this.update(url, item);
    }
  }, {
    key: 'getListType',
    value: function getListType(list) {
      return this._axios.post((0, _lists.listURI)(list) + '?$select=ListItemEntityTypeFullName', {}, {}).then(function (response) {
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

  }, {
    key: 'getListItems',
    value: function getListItems(list) {
      var _ref2 = arguments.length <= 1 || arguments[1] === undefined ? {} : arguments[1];

      var where = _ref2.where;

      var query = {};

      if (where) {
        query.$filter = _lodash2.default.map(fillSpaces(where), function (value, field) {
          return field + ' eq \'' + value + '\'';
        }).join(' and ');
      }

      var qs = _querystring2.default.stringify(query);
      return this._axios.get((0, _lists.listURI)(list) + '/items?' + qs).then(function (res) {
        return res.data.d.results;
      });
    }
  }]);

  return SharePointAPI;
}();

exports.default = SharePointAPI;