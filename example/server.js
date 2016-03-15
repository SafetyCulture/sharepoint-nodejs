import express from 'express';
import OAuth2 from '../src/oauth2.js';

export const app = express();

const oauth2 = OAuth2({
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  redirectUri: process.env.CALLBACK_URL
});

app.get(`/`, (req, res) => {
  const code = req.param(`code`);

  oauth2.requestToken(code).then((response) => {
    console.log(response.access_token);
    console.log(response.refresh_token);
    res.send(`Authorized!`);
  });

});

app.get(`/request`, (req, res) => {
  let url = oauth2.getAuthorizationUrl({ scope: `wl.offline_access` });
  res.redirect(url);
});


app.listen(3000);
