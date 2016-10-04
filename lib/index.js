'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.SharePoint = exports.OAuth2 = exports.Authentication = exports.Batch = exports.sharepointEscapeChars = exports.fillSpaces = exports.libraryType = exports.listType = exports.listURI = exports.AFFIXED_NEWLINE_REGEX = exports.STD_NEWLINE_REGEX = exports.LIST_TEMPLATES = exports.FIELD_TYPES = undefined;

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
Object.defineProperty(exports, 'STD_NEWLINE_REGEX', {
  enumerable: true,
  get: function get() {
    return _lists.STD_NEWLINE_REGEX;
  }
});
Object.defineProperty(exports, 'AFFIXED_NEWLINE_REGEX', {
  enumerable: true,
  get: function get() {
    return _lists.AFFIXED_NEWLINE_REGEX;
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
Object.defineProperty(exports, 'libraryType', {
  enumerable: true,
  get: function get() {
    return _lists.libraryType;
  }
});
Object.defineProperty(exports, 'fillSpaces', {
  enumerable: true,
  get: function get() {
    return _lists.fillSpaces;
  }
});
Object.defineProperty(exports, 'sharepointEscapeChars', {
  enumerable: true,
  get: function get() {
    return _lists.sharepointEscapeChars;
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

var _oauth = require('./oauth2');

Object.defineProperty(exports, 'OAuth2', {
  enumerable: true,
  get: function get() {
    return _oauth.OAuth2;
  }
});

var _lodash = require('lodash');

var _querystring = require('querystring');

var _querystring2 = _interopRequireDefault(_querystring);

var _axios = require('axios');

var _axios2 = _interopRequireDefault(_axios);

var _requestPromise = require('request-promise');

var _requestPromise2 = _interopRequireDefault(_requestPromise);

var _files = require('./files');

var _misc = require('./misc');

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

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

    this.auth = auth;
    this.host = host;
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
        config.headers = (0, _lodash.assign)({}, config.headers, {
          'Accept': 'application/json;odata=verbose',
          'User-Agent': _misc.USER_AGENT,
          'Content-Type': 'application/json;odata=verbose'
        }, (0, _misc.getAuthHeaders)(auth));

        config.timeout = 360000;

        return config;
      });

      instance.interceptors.response.use(function (res) {
        if (res.data && res.data.d) {
          res.data.d = (0, _misc.formatResponse)(res.data.d);
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
          'LookupFieldName': (0, _lists.sharepointEscapeChars)(lookupFieldName)
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
      var headers = (0, _lodash.merge)((0, _misc.getAuthHeaders)(this.auth), {
        'Accept': 'application/json;odata=verbose',
        'User-Agent': _misc.USER_AGENT,
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
        'Content-Type': 'application/json;odata=verbose' });

      var options = {
        headers: headers,
        method: 'POST',
        body: body,
        resolveWithFullResponse: true,
        json: true,
        uri: this.host + '/_api/web' + resource
      };

      return (0, _requestPromise2.default)(options);
    }
  }, {
    key: 'delete',
    value: function _delete(resource) {
      var headers = {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      };
      this._axios.post(resource, {}, headers);
    }
  }, {
    key: 'createItem',
    value: function createItem(list, item) {
      return this.create((0, _lists.listURI)(list) + '/items', item);
    }
  }, {
    key: 'updateItem',
    value: function updateItem(list, itemId, item) {
      return this.update((0, _lists.listURI)(list) + '/items(' + itemId + ')', (0, _lists.fillSpaces)(item));
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
      var _ref2 = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};

      var where = _ref2.where;

      var query = {};

      if (where) {
        query.$filter = (0, _lodash.map)((0, _lists.fillSpaces)(where), function (value, field) {
          return field + ' eq \'' + value + '\'';
        }).join(' and ');
      }

      var qs = _querystring2.default.stringify(query);
      return this._axios.get((0, _lists.listURI)(list) + '/items?' + qs).then(function (res) {
        return res.data.d.results;
      });
    }
  }, {
    key: 'getDefaultView',
    value: function getDefaultView(list) {
      return this._axios.get((0, _lists.listURI)(list) + '/DefaultView');
    }
  }, {
    key: 'addViewField',
    value: function addViewField(list, view, field) {
      return this._axios.post((0, _lists.listURI)(list) + '/views(guid\'' + view + '\')/ViewFields/AddViewField(\'' + (0, _lists.sharepointEscapeChars)(field) + '\')');
    }
  }]);

  return SharePoint;
}();