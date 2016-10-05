import express from 'express';

import * as SharePoint from '../../src/index';

export const app = express();

const oauth = SharePoint.OAuth2({
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  redirectUri: process.env.CALLBACK_URL,
  authorizeUri: process.env.CALLBACK_URL
});

app.get(`/`, (req, res) => {
  const code = req.param.code;

  oauth.requestToken(code).then((response) => {
    console.log(response.access_token);
    console.log(response.refresh_token);
    res.send(`Authorized!`);
  });
});

app.get(`/request`, (req, res) => {
  let url = oauth.getAuthorizationUrl({ scope: `wl.offline_access` });
  res.redirect(url);
});

app.listen(3000);
