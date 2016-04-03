'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = OAuth2;

var _querystring = require('querystring');

var _querystring2 = _interopRequireDefault(_querystring);

var _lodash = require('lodash');

var _lodash2 = _interopRequireDefault(_lodash);

var _url = require('url');

var _url2 = _interopRequireDefault(_url);

var _requestPromise = require('request-promise');

var _requestPromise2 = _interopRequireDefault(_requestPromise);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var LOGIN_URL = 'https://login.live.com';
var AUTHORIZE_URL = LOGIN_URL + '/oauth20_authorize.srf';
var TOKEN_URL = LOGIN_URL + '/oauth20_token.srf';

function OAuth2(_ref) {
  var clientId = _ref.clientId;
  var clientSecret = _ref.clientSecret;
  var redirectUri = _ref.redirectUri;

  return {
    clientId: clientId,
    clientSecret: clientSecret,
    redirectUri: redirectUri,

    getAuthorizationUrl: function getAuthorizationUrl(_ref2) {
      var scope = _ref2.scope;
      var state = _ref2.state;

      var params = _lodash2.default.extend({
        response_type: 'code',
        client_id: this.clientId,
        redirect_uri: this.redirectUri
      }, { scope: scope, state: state });

      return this.mergeUrl(AUTHORIZE_URL, params);
    },
    mergeUrl: function mergeUrl(baseUrl, params) {
      var components = _url2.default.parse(baseUrl);
      var merged = _lodash2.default.extend(_querystring2.default.parse(components.query), params);
      components.query = merged;
      return _url2.default.format(components);
    },
    refreshToken: function refreshToken(_refreshToken) {
      return this.post({
        grant_type: 'refresh_token',
        refresh_token: _refreshToken,
        client_id: this.clientId,
        client_secret: this.clientSecret
      });
    },
    requestToken: function requestToken(code) {
      return this.post({
        grant_type: 'authorization_code',
        code: code,
        client_id: this.clientId,
        client_secret: this.clientSecret,
        redirect_uri: this.redirectUri
      });
    },
    post: function post(params) {
      return (0, _requestPromise2.default)({
        method: 'POST',
        uri: TOKEN_URL,
        body: _querystring2.default.stringify(params),
        headers: {
          'content-type': 'application/x-www-form-urlencoded'
        }
      }).then(function (response) {
        return JSON.parse(response);
      });
    }
  };
}