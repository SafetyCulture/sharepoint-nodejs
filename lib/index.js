'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.SharePoint = exports.FIELD_TYPES = exports.ITEMS_LIST = exports.AUDITS_LIST = undefined;

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _lodash = require('lodash');

var _lodash2 = _interopRequireDefault(_lodash);

var _sharepointAuth = require('sharepoint-auth');

var _sharepointAuth2 = _interopRequireDefault(_sharepointAuth);

var _api = require('./api');

var _api2 = _interopRequireDefault(_api);

var _batchApi = require('./batchApi');

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var AUDITS_LIST = exports.AUDITS_LIST = 'SafetyCulture Audits';
var ITEMS_LIST = exports.ITEMS_LIST = 'SafetyCulture Items';

// Common SharePoint Field Types
// https://msdn.microsoft.com/en-au/library/office/microsoft.sharepoint.client.fieldtype.aspx
var FIELD_TYPES = exports.FIELD_TYPES = {
  INVALID: 0,
  INTEGER: 1,
  TEXT: 2,
  NOTE: 3,
  DATE_TIME: 4,
  COUNTER: 5,
  CHOICE: 6,
  LOOKUP: 7,
  BOOLEAN: 8,
  NUMBER: 9,
  CURRENCY: 10,
  URL: 11
};

// SharePoint fields to field types.
var AUDIT_FIELDS = {
  'Audit Id': FIELD_TYPES.TEXT,
  'Score': FIELD_TYPES.NUMBER,
  'Total Score': FIELD_TYPES.NUMBER,
  'Score Percentage': FIELD_TYPES.NUMBER,
  'Duration': FIELD_TYPES.NUMBER,
  'Date Modified': FIELD_TYPES.DATE_TIME,
  'Date Started': FIELD_TYPES.DATE_TIME,
  'Date Completed': FIELD_TYPES.DATE_TIME,
  'SafetyCulture Owner': FIELD_TYPES.TEXT,
  'SafetyCulture Author': FIELD_TYPES.TEXT,
  'Device Id': FIELD_TYPES.TEXT,
  'Template Name': FIELD_TYPES.TEXT,
  'Template Description': FIELD_TYPES.NOTE
};

var ITEM_FIELDS = {
  'Parent Id': FIELD_TYPES.TEXT,
  'Item Id': FIELD_TYPES.TEXT,
  'Label': FIELD_TYPES.TEXT,
  'Type': FIELD_TYPES.TEXT,
  'Score': FIELD_TYPES.NUMBER,
  'Max Score': FIELD_TYPES.NUMBER,
  'Percentage': FIELD_TYPES.NUMBER,
  'C Score': FIELD_TYPES.NUMBER,
  'C Max Score': FIELD_TYPES.NUMBER,
  'C Score Percentage': FIELD_TYPES.NUMBER,
  'O Weighting': FIELD_TYPES.NUMBER,
  'R Id': FIELD_TYPES.TEXT,
  'R Type': FIELD_TYPES.TEXT,
  'R Label': FIELD_TYPES.TEXT,
  'R Short Label': FIELD_TYPES.TEXT,
  'R Colour': FIELD_TYPES.TEXT,
  'R Image': FIELD_TYPES.TEXT,
  'R Enable Score': FIELD_TYPES.TEXT,
  'R Score': FIELD_TYPES.NUMBER
};

/**
* SharePoint class
* @param {object} options Listed options
* @param {string} options.username Username for SharePoint
* @param {string} options.password Password for SharePoint
* @param {string} options.host Host for SharePoint
* @returns {object} SharePoint instance
*/

var SharePoint = exports.SharePoint = function () {
  function SharePoint(_ref) {
    var username = _ref.username;
    var password = _ref.password;
    var host = _ref.host;
    var _ref$auditsList = _ref.auditsList;
    var auditsList = _ref$auditsList === undefined ? AUDITS_LIST : _ref$auditsList;
    var _ref$itemsList = _ref.itemsList;
    var itemsList = _ref$itemsList === undefined ? ITEMS_LIST : _ref$itemsList;
    var _ref$auditFields = _ref.auditFields;
    var auditFields = _ref$auditFields === undefined ? AUDIT_FIELDS : _ref$auditFields;
    var _ref$itemFields = _ref.itemFields;
    var itemFields = _ref$itemFields === undefined ? ITEM_FIELDS : _ref$itemFields;
    var _ref$disableItems = _ref.disableItems;
    var disableItems = _ref$disableItems === undefined ? false : _ref$disableItems;

    _classCallCheck(this, SharePoint);

    this.options = { auth: { username: username, password: password }, host: host };
    this.auditsList = auditsList;
    this.itemsList = itemsList;
    this.auditFields = auditFields;
    this.itemFields = itemFields;
    this.disableItems = disableItems;
  }

  /**
  * Returns authorization needed to make requests to SharePoint
  * @param {object} options Options to pass to spAuth
  * @returns {Promise} Resolves with auth object, rejects with error from SP
  */


  _createClass(SharePoint, [{
    key: '_getAuth',
    value: function _getAuth() {
      var _this = this;

      return new Promise(function (resolve, reject) {
        if (_this._auth) return resolve(_this._auth);
        (0, _sharepointAuth2.default)(_this.options, function (err, res) {
          if (err) return reject(err);
          _this._auth = {
            FedAuth: res.cookies.FedAuth,
            rtFa: res.cookies.rtFa,
            requestDigest: res.requestDigest
          };
          resolve(_this._auth);
        });
      });
    }

    /**
    * Ensure both required lists are created, otherwise creates and links them
    * @returns {Promise} Resolves if lists exist, rejects when something goes wrong
    */

  }, {
    key: 'ensureLists',
    value: function ensureLists() {
      var _this2 = this;

      if (this.ensured) return Promise.resolve();

      return this._getAuth().then(function (auth) {
        var api = new _api2.default(_this2.options.host, auth);
        var batchApi = new _batchApi.SharePointBatchAPI(_this2.options.host, auth);

        batchApi.getList(_this2.itemsList);
        batchApi.getList(_this2.auditsList);

        return batchApi.run().then(function (getRes) {
          if (!_lodash2.default.every(getRes, function (res) {
            return _lodash2.default.has(res, 'd');
          })) {
            batchApi.createList(_this2.auditsList, _this2.auditsList, _this2.auditFields, _batchApi.LIST_TYPES.LIBRARY);

            if (!_this2.disableItems) {
              batchApi.createList(_this2.itemsList, _this2.itemsList, _this2.itemFields, _batchApi.LIST_TYPES.STANDARD);
            }

            return batchApi.run().then(function (createRes) {
              if (_lodash2.default.some(createRes, function (r) {
                return _lodash2.default.has(r, 'error');
              })) throw createRes;

              if (!_this2.disableItems) {
                var lists = _lodash2.default.filter(createRes, function (r) {
                  return r.entry.category[0].$.term === 'SP.List';
                });
                var itemListId = lists[1].entry.content[0]['m:properties'][0]['d:Id'][0]._;

                // Link with normal api as batching can't hold reference in PATCH requests
                return api.linkLists(_this2.auditsList, itemListId, 'Items', 'Item Id', true);
              }
            });
          }
        }).then(function () {
          return _this2.ensured = true;
        });
      });
    }

    /**
    * Creates or Updates an audit on SharePoint
    * @param {object} audit The audit to Create/Update
    * @param {array} items Array of items to add
    * @returns {Promise} Resolves with audit object, rejects with error from SP
    */

  }, {
    key: 'createOrUpdate',
    value: function createOrUpdate(audit, items) {
      var _this3 = this;

      return this._getAuth().then(function (auth) {
        var api = new _api2.default(_this3.options.host, auth);
        var batchApi = new _batchApi.SharePointBatchAPI(_this3.options.host, auth);

        return _this3.ensureLists().then(function () {
          return api.getListItems(_this3.auditsList, {
            where: { 'Audit Id': audit['Audit Id'] }
          });
        }).then(function (audits) {
          if (audits[0] && audits[0].ItemsId) {
            var itemIds = audits[0].ItemsId.results;

            batchApi.addChangeset(itemIds.map(function (itemId) {
              return batchApi.deleteListItem(_this3.itemsList, itemId);
            }));

            return batchApi.run().then(function (responses) {
              if (_lodash2.default.some(responses, function (r) {
                return _lodash2.default.has(r, 'error');
              })) throw JSON.stringify(responses);
              batchApi.addChangeset([batchApi.deleteListItem(_this3.auditsList, audits[0].Id)]);

              return batchApi.run();
            });
          }
        }).then(function (responses) {
          if (_lodash2.default.some(responses, function (r) {
            return _lodash2.default.has(r, 'error');
          })) throw JSON.stringify(responses);

          if (!_this3.disabledItems) {
            batchApi.addChangeset(_lodash2.default.map(items, function (item) {
              return batchApi.createListItem(_this3.itemsList, item);
            }));
          }

          return batchApi.run();
        }).then(function (responses) {
          if (_lodash2.default.some(responses, function (r) {
            return _lodash2.default.has(r, 'error');
          })) throw JSON.stringify(responses);
          var results = _lodash2.default.map(responses, function (r) {
            return Number(r.entry.content[0]['m:properties'][0]['d:Id'][0]._);
          });

          batchApi.addChangeset([batchApi.createListItem(_this3.auditsList, _lodash2.default.assign({}, audit, {
            'ItemsId': { results: results }
          }))]);

          return batchApi.run();
        });
      });
    }
  }]);

  return SharePoint;
}();