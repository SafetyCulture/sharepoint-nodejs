import querystring from 'querystring';
import _ from 'lodash';
import url from 'url';
import rp from 'request-promise';

const LOGIN_URL = 'https://login.live.com';
const AUTHORIZE_URL = `${LOGIN_URL}/oauth20_authorize.srf`;
const TOKEN_URL = `${LOGIN_URL}/oauth20_token.srf`;

export function OAuth2({ clientId, clientSecret, redirectUri, authorizeUri, tokenUri, realm, resource } = {authorizeUri: AUTHORIZE_URL, tokenUri: TOKEN_URL}) {
  return {
    clientId,
    clientSecret,
    redirectUri,
    authorizeUri,
    tokenUri,
    realm,
    resource,

    getAuthorizationUrl({ scope, state }) {
      let params = _.extend({
        response_type: 'code',
        client_id: this.clientId,
        redirect_uri: this.redirectUri
      }, { scope, state });

      return this.mergeUrl(authorizeUri, params);
    },

    mergeUrl(baseUrl, params) {
      let components = url.parse(baseUrl);
      let merged = _.extend(querystring.parse(components.query),
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
      if (this.realm) {
        return `${this.clientId}@${this.realm}`;
      }
      else {
        return this.clientId;
      }
    },

    post(params) {
      return rp({
        method: 'POST',
        uri: tokenUri,
        body: querystring.stringify(params),
        headers: {
          'content-type': 'application/x-www-form-urlencoded'
        }
      }).then(function(response) {
        return JSON.parse(response);
      });
    }
  };
}
