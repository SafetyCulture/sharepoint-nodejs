import express from 'express';

import * as SharePoint from '../../src/index';

export const app = express();

const sharePointSiteUrl = 'https://your.sharepoint.com/tenantsite';

// These can be established automatically by requesting e.g.`https://your.sharepoint.com/_vti_bin/client.svc
// and parsing the headers of the response.
const sharePointRealm = 'tenant site realm';
const sharePointResource = 'tenant site resource';

// clientId, clientSecret and redirectUri are the values registered with
// the Office 365 tenant security token service (ACS) using e.g. appregnew.aspx
const oauth = SharePoint.OAuth2({
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  redirectUri: process.env.CALLBACK_URL,
  authorizeUri: `${sharePointSiteUrl}/_layouts/oauthauthorize.aspx`,
  tokenUri: `https://accounts.accesscontrol.windows.net/${sharePointRealm}/tokens/OAuth/2`,
  realm: sharePointRealm,
  resource: `${sharePointResource}/${sharePointSiteHostName}@${sharePointRealm}`
});

// Redirect URI handler
app.get(`/`, (req, res) => {
  const code = req.param.code;

  oauth.requestToken(code).then((response) => {
    console.log(response.access_token);
    console.log(response.refresh_token);
    res.send(`Authorized!`);
  });
});

// SharePoint add-in trust establishment handler
app.get(`/request`, (req, res) => {
  let url = oauth.getAuthorizationUrl({ scope: `wl.offline_access` });
  res.redirect(url);
});

app.listen(3000);
