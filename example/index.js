import { SharePoint, Authentication } from '../src/index';
console.log(Authentication);

const username = "";
const password = "";
const host = "";

const authentication = Authentication({ username, password, host });

authentication.request().then(auth => {
  console.log(auth);
});
