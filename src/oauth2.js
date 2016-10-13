import querystring from 'querystring';
import {extend, contains, forOwn} from 'lodash';
import url from 'url';
import rp from 'request-promise';
import {log as log_} from './logger';

export const asStringWithoutSensitiveFields = (queryStringObject) => {
  const overwriteWith = 'omitted';
  const sensitiveFields = ['refresh_token', 'client_secret', 'code'];

  let sanitisedQueryString = {};

  forOwn(queryStringObject, (value, key) => {
    sanitisedQueryString[key] = contains(sensitiveFields, key) ? overwriteWith : value;
  });

  return querystring.stringify(sanitisedQueryString);
};

export function OAuth2({clientId, clientSecret, redirectUri, authorizeUri, tokenUri, realm, resource, log = log_}) {
  return {
    clientId,
    clientSecret,
    redirectUri,
    authorizeUri,
    tokenUri,
    realm,
    resource,

    getAuthorizationUrl({scope, state}) {
      let params = extend({
        response_type: 'code',
        client_id: this.clientId,
        redirect_uri: this.redirectUri
      }, {scope, state});

      return this.mergeUrl(authorizeUri, params);
    },

    mergeUrl(baseUrl, params) {
      let components = url.parse(baseUrl);
      let merged = extend(querystring.parse(components.query),
        params);
      components.query = merged;
      return url.format(components);
    },

    refreshToken(refreshToken) {
      return this.post({
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        client_id: this.getClientId(),
        client_secret: this.clientSecret,
        resource: this.resource
      });
    },

    requestToken(code) {
      return this.post({
        grant_type: 'authorization_code',
        code: code,
        client_id: this.getClientId(),
        client_secret: this.clientSecret,
        redirect_uri: this.redirectUri,
        resource: this.resource
      });
    },

    getClientId() {
      let id = this.realm ?
        `${this.clientId}@${this.realm}` : this.clientId;
      log.info(`Add-in full client ID = ${id} Realm = ${this.realm} Short client ID = ${this.clientId}`);
      return id;
    },

    post(params) {
      log.info(`POST to ${tokenUri} with ${asStringWithoutSensitiveFields(params)}`);

      return rp({
        method: 'POST',
        uri: tokenUri,
        body: querystring.stringify(params),
        headers: {
          'content-type': 'application/x-www-form-urlencoded'
        }
      }).catch((err) => {
        log.error(`POST failed with ${err}`);
        throw err;
      }).then((response) => {
        try {
          return JSON.parse(response);
        } catch (err) {
          // In this case we do not expect sensitive info to be part of the response so ok to log it.
          log.error(`Failed to deserialise POST response: ${response}`);
          throw err;
        }
      });
    }
  };
}
