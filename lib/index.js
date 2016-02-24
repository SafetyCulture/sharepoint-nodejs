'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.SharePoint = exports.Authentication = exports.Batch = exports.listType = exports.listURI = exports.LIST_TEMPLATES = exports.FIELD_TYPES = undefined;

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _fields = require('./fields');

Object.defineProperty(exports, 'FIELD_TYPES', {
  enumerable: true,
  get: function get() {
    return _fields.FIELD_TYPES;
  }
});

var _lists = require('./lists');

Object.defineProperty(exports, 'LIST_TEMPLATES', {
  enumerable: true,
  get: function get() {
    return _lists.LIST_TEMPLATES;
  }
});
Object.defineProperty(exports, 'listURI', {
  enumerable: true,
  get: function get() {
    return _lists.listURI;
  }
});
Object.defineProperty(exports, 'listType', {
  enumerable: true,
  get: function get() {
    return _lists.listType;
  }
});

var _batch = require('./batch');

Object.defineProperty(exports, 'Batch', {
  enumerable: true,
  get: function get() {
    return _batch.Batch;
  }
});

var _authentication = require('./authentication');

Object.defineProperty(exports, 'Authentication', {
  enumerable: true,
  get: function get() {
    return _authentication.Authentication;
  }
});

var _lodash = require('lodash');

var _lodash2 = _interopRequireDefault(_lodash);

var _querystring = require('querystring');

var _querystring2 = _interopRequireDefault(_querystring);

var _axios = require('axios');

var _axios2 = _interopRequireDefault(_axios);

var _files = require('./files');

var _misc = require('./misc');

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

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
* SharePoint class
* @param {string} host Host for SharePoint
* @param {object} auth Auth object for SharePoint
* @returns {object} SharePoint instance
*/

var SharePoint = exports.SharePoint = function () {
  function SharePoint(host, auth) {
    _classCallCheck(this, SharePoint);

    if (!host) throw new Error('SharePoint requires host string');
    if (!auth) throw new Error('SharePoint requires auth object');

    this._axios = this._configureInterceptors(_axios2.default.create(), { host: host, auth: auth });
    this.files = (0, _files.Files)(this);
  }

  /**
  * Returns axios instance configured with auth details
  * @param {object} instance Axios instance
  * @param {object} options Options to pass to spAuth
  * @returns {object} Axios instance
  */


  _createClass(SharePoint, [{
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
          'User-Agent': _misc.USER_AGENT,
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
        if (!multiValues) return Promise.resolve();

        return _this._axios.post((0, _lists.listURI)(list) + '/fields(\'' + res.data.d.Id + '\')', {
          '__metadata': { 'type': 'SP.FieldLookup' },
          'AllowMultipleValues': true
        }, { headers: { 'X-HTTP-Method': 'MERGE' } });
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

  return SharePoint;
}();