import { SharePoint, Authentication } from '../src/index';

const username = process.env.SHAREPOINT_USERNAME;
const password = process.env.SHAREPOINT_PASSWORD;
const host = process.env.SHAREPOINT_URL;

const authentication = Authentication({ username, password, host });

authentication.request().then(auth => {
  console.log(auth);
});
