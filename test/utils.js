export function generateSchemaErrorMsg(errors) {
  return errors.reduce((errStr, error) => {
    return errStr + 'Invalid ' + error.keyword + ' on ' + error.path + '. ';
  }, '');
}

// Formats SharePoints URLs so nock can catch them
export const format = (endpoint) =>
  encodeURI(`/_api/web${endpoint}`).replace(/'/g, '%27').replace(/\$/g, '%24');
