'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.OAuth2 = OAuth2;

var _querystring = require('querystring');

var _querystring2 = _interopRequireDefault(_querystring);

var _lodash = require('lodash');

var _url = require('url');

var _url2 = _interopRequireDefault(_url);

var _requestPromise = require('request-promise');

var _requestPromise2 = _interopRequireDefault(_requestPromise);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var LOGIN_URL = 'https://login.live.com';
var AUTHORIZE_URL = LOGIN_URL + '/oauth20_authorize.srf';
var TOKEN_URL = LOGIN_URL + '/oauth20_token.srf';

function OAuth2() {
  var _ref = arguments.length <= 0 || arguments[0] === undefined ? { authorizeUri: AUTHORIZE_URL, tokenUri: TOKEN_URL } : arguments[0];

  var clientId = _ref.clientId;
  var clientSecret = _ref.clientSecret;
  var redirectUri = _ref.redirectUri;
  var authorizeUri = _ref.authorizeUri;
  var tokenUri = _ref.tokenUri;
  var realm = _ref.realm;
  var resource = _ref.resource;

  return {
    clientId: clientId,
    clientSecret: clientSecret,
    redirectUri: redirectUri,
    authorizeUri: authorizeUri,
    tokenUri: tokenUri,
    realm: realm,
    resource: resource,

    getAuthorizationUrl: function getAuthorizationUrl(_ref2) {
      var scope = _ref2.scope;
      var state = _ref2.state;

      var params = (0, _lodash.extend)({
        response_type: 'code',
        client_id: this.clientId,
        redirect_uri: this.redirectUri
      }, { scope: scope, state: state });

      return this.mergeUrl(authorizeUri, params);
    },
    mergeUrl: function mergeUrl(baseUrl, params) {
      var components = _url2.default.parse(baseUrl);
      var merged = (0, _lodash.extend)(_querystring2.default.parse(components.query), params);
      components.query = merged;
      return _url2.default.format(components);
    },
    refreshToken: function refreshToken(_refreshToken) {
      return this.post({
        grant_type: 'refresh_token',
        refresh_token: _refreshToken,
        client_id: this.getClientId(),
        client_secret: this.clientSecret,
        resource: this.resource
      });
    },
    requestToken: function requestToken(code) {
      return this.post({
        grant_type: 'authorization_code',
        code: code,
        client_id: this.getClientId(),
        client_secret: this.clientSecret,
        redirect_uri: this.redirectUri,
        resource: this.resource
      });
    },
    getClientId: function getClientId() {
      if (this.realm) {
        return this.clientId + '@' + this.realm;
      } else {
        return this.clientId;
      }
    },
    post: function post(params) {
      return (0, _requestPromise2.default)({
        method: 'POST',
        uri: tokenUri,
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