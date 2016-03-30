import querystring from 'querystring';
import _ from 'lodash';
import url from 'url';
import rp from 'request-promise';

const LOGIN_URL = 'https://login.live.com';
const AUTHORIZE_URL = `${LOGIN_URL}/oauth20_authorize.srf`;
const TOKEN_URL = `${LOGIN_URL}/oauth20_token.srf`;

export default function OAuth2({ clientId, clientSecret, redirectUri }) {
  return {
    clientId,
    clientSecret,
    redirectUri,

    getAuthorizationUrl({ scope, state }) {
      let params = _.extend({
        response_type: 'code',
        client_id: this.clientId,
        redirect_uri: this.redirectUri
      }, { scope, state });

      return this.mergeUrl(AUTHORIZE_URL, params);
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
        client_id: this.clientId,
        client_secret: this.clientSecret
      });
    },

    requestToken(code) {
      return this.post({
        grant_type: 'authorization_code',
        code: code,
        client_id: this.clientId,
        client_secret: this.clientSecret,
        redirect_uri: this.redirectUri
      });
    },

    post(params) {
      return rp({
        method: 'POST',
        uri: TOKEN_URL,
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
