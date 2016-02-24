'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.Authentication = Authentication;

var _fs = require('fs');

var _fs2 = _interopRequireDefault(_fs);

var _xml2json = require('xml2json');

var _xml2json2 = _interopRequireDefault(_xml2json);

var _requestPromise = require('request-promise');

var _requestPromise2 = _interopRequireDefault(_requestPromise);

var _lodash = require('lodash');

var _lodash2 = _interopRequireDefault(_lodash);

var _misc = require('./misc');

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var saml = _fs2.default.readFileSync(__dirname + '/../config/saml.xml').toString();

var getCustomerDomain = function getCustomerDomain(host) {
  var hostParts = host.split('://');
  var hostname = hostParts[1].split('/')[0].split('.')[0];
  return hostname;
};

var extractCookies = function extractCookies(headers) {
  var cookies = {};
  _lodash2.default.each(headers['set-cookie'], function (value) {
    var parsedCookies = value.split(/\=(.+)?/);
    parsedCookies[1] = parsedCookies[1].substr(0, parsedCookies[1].indexOf(';'));
    cookies[parsedCookies[0]] = parsedCookies[1];
  });

  return cookies;
};

var buildRequest = function buildRequest(username, password, host) {
  //Replace username, pwd and URL into SAML.xml
  var body = saml;
  body = body.replace('{username}', username);
  body = body.replace('{password}', password);
  body = body.replace('{url}', host);
  return body;
};

var getDigest = function getDigest(_ref) {
  var cookies = _ref.cookies;
  var domain = _ref.domain;

  var url = 'https://' + domain + '.sharepoint.com/_api/contextinfo';

  var headers = {
    'Cookie': 'FedAuth=' + cookies.FedAuth + ';' + 'rtFa=' + cookies.rtFa,
    'Content-Type': 'application/json; odata=verbose',
    'Accept': 'application/json; odata=verbose',
    'User-Agent': _misc.USER_AGENT
  };

  return _requestPromise2.default.post({ url: url, headers: headers }).then(function (resp) {
    var data = JSON.parse(resp);
    var requestDigest = data.d.GetContextWebInformation.FormDigestValue;
    var requestDigestTimeoutSeconds = data.d.GetContextWebInformation.FormDigestTimeoutSeconds;

    return {
      requestDigest: requestDigest,
      requestDigestTimeoutSeconds: requestDigestTimeoutSeconds,
      FedAuth: cookies.FedAuth,
      rtFa: cookies.rtFa
    };
  });
};

var getToken = function getToken(_ref2) {
  var username = _ref2.username;
  var password = _ref2.password;
  var host = _ref2.host;

  var request = buildRequest(username, password, host);
  var domain = getCustomerDomain(host);
  var url = 'https://login.microsoftonline.com/extSTS.srf';

  var headers = {
    'User-Agent': _misc.USER_AGENT
  };

  return _requestPromise2.default.post({ url: url, body: request, headers: headers }).then(function (resp) {
    var body = _xml2json2.default.toJson(resp, { object: true });

    var responseBody = body['S:Envelope']['S:Body'];
    // let samlError = responseBody['S:Fault'];

    var token = responseBody['wst:RequestSecurityTokenResponse']['wst:RequestedSecurityToken']['wsse:BinarySecurityToken'].$t;
    return { token: token, domain: domain };
  });
};

// Get the Cookies
var getCookies = function getCookies(_ref3) {
  var token = _ref3.token;
  var domain = _ref3.domain;

  var url = 'https://' + domain + '.sharepoint.com/_forms/default.aspx?wa=wsignin1.0';

  var headers = {
    'User-Agent': _misc.USER_AGENT
  };

  var options = { url: url,
    body: token,
    resolveWithFullResponse: true,
    followAllRedirects: true,
    headers: headers,
    jar: true };

  return _requestPromise2.default.post(options).then(function (response) {
    return { domain: domain, cookies: extractCookies(response.headers) };
  });
};

function Authentication(_ref4) {
  var username = _ref4.username;
  var password = _ref4.password;
  var host = _ref4.host;

  return {
    request: function request() {
      return getToken({ username: username, password: password, host: host }).then(getCookies).then(getDigest);
    }
  };
}