import { Batch, Authentication } from '../../src/index';

const username = process.env.SHAREPOINT_USERNAME;
const password = process.env.SHAREPOINT_PASSWORD;
const host = process.env.SHAREPOINT_URL;

const authentication = Authentication({ username, password, host });

authentication.request().then(auth => {
  const batch = new Batch(host, auth);
  batch.getList('SafetyCulture Audits');

  batch.run().then((resp) => {
    console.log(resp);
  });
});
