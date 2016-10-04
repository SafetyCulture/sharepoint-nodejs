import fs from 'fs';
import parser from 'xml2json';
import rp from 'request-promise';
import { each } from 'lodash';

import { USER_AGENT } from './misc';

const saml = fs.readFileSync(__dirname + '/../config/saml.xml').toString();

const getCustomerDomain = (host) => {
  let hostParts = host.split('://');
  let hostname = hostParts[1].split('/')[0].split('.')[0];
  return hostname;
};

const extractCookies = (headers) => {
  let cookies = {};
  each(headers['set-cookie'], (value) => {
    let parsedCookies = value.split(/\=(.+)?/);
    parsedCookies[1] = parsedCookies[1].substr(0, parsedCookies[1].indexOf(';'));
    cookies[parsedCookies[0]] = parsedCookies[1];
  });

  return cookies;
};

const buildRequest = (username, password, host) => {
  //Replace username, pwd and URL into SAML.xml
  let body = saml;
  body = body.replace('{username}', username);
  body = body.replace('{password}', password);
  body = body.replace('{url}', host);
  return body;
};

const getDigest = ({cookies, domain}) => {
  let url = 'https://' + domain + '.sharepoint.com/_api/contextinfo';

  let headers = {
    'Cookie': 'FedAuth=' + cookies.FedAuth + ';' + 'rtFa=' + cookies.rtFa,
    'Content-Type': 'application/json; odata=verbose',
    'Accept': 'application/json; odata=verbose',
    'User-Agent': USER_AGENT
  };

  return rp.post({url: url, headers: headers }).then((resp) => {
    let data = JSON.parse(resp);
    let requestDigest = data.d.GetContextWebInformation.FormDigestValue;
    let requestDigestTimeoutSeconds = data.d.GetContextWebInformation.FormDigestTimeoutSeconds;

    return {
      requestDigest: requestDigest,
      requestDigestTimeoutSeconds: requestDigestTimeoutSeconds,
      FedAuth: cookies.FedAuth,
      rtFa: cookies.rtFa
    };
  });
};

const getToken = ({ username, password, host }) => {
  let request = buildRequest(username, password, host);
  let domain = getCustomerDomain(host);
  let url = 'https://login.microsoftonline.com/extSTS.srf';

  const headers = {
    'User-Agent': USER_AGENT
  };

  return rp.post({url: url, body: request, headers: headers}).then((resp) => {
    console.info(`Token Response ${resp}`);
    let body = parser.toJson(resp, {object: true});

    let responseBody = body['S:Envelope']['S:Body'];
    let samlError = responseBody['S:Fault'];

    if (typeof responseBody === 'undefined') {
      return Promise.reject(`no response ${resp}`);
    }

    if (samlError) {
      return Promise.reject(samlError);
    }

    let token = responseBody['wst:RequestSecurityTokenResponse']['wst:RequestedSecurityToken']['wsse:BinarySecurityToken'].$t;
    return {token, domain};
  });
};

// Get the Cookies
const getCookies = ({token, domain}) => {
  let url = 'https://' + domain + '.sharepoint.com/_forms/default.aspx?wa=wsignin1.0';

  const headers = {
    'User-Agent': USER_AGENT
  };

  let options = { url: url,
                  body: token,
                  resolveWithFullResponse: true,
                  followAllRedirects: true,
                  headers: headers,
                  jar: true };

  return rp.post(options).then((response) => {
    return {domain: domain, cookies: extractCookies(response.headers)};
  });
};

export function Authentication({ username, password, host }) {
  return {
    request: () => {
      return getToken({ username, password, host })
                  .then(getCookies)
                  .then(getDigest);
    }
  };
}
