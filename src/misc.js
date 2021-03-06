import { isArray, isObject, isString, map, transform } from 'lodash';

export const USER_AGENT = 'SafetyCulture SharePoint';

/**
* Formats a response to replace '_x0020_' with spaces.
* @param {object} res Response to deep replace
* @returns {object} Formatted response
*/
export function formatResponse(res) {
  return transform(res, (result, val, key) => {
    let newVal = val;
    const newKey = isString(key) ? key.replace(/_x0020_/g, ' ') : key;

    if (isArray(val)) newVal = map(val, formatResponse);
    if (isObject(val)) newVal = formatResponse(val);

    result[newKey] = newVal;
  });
}

/**
 * Support either token (oauth2) or cookie based authentication
 * to sharepoint API
 */
export function getAuthHeaders(auth) {
  if (auth.token !== undefined) {
    return {'Authorization': `Bearer ${auth.token}`};
  }

  return {'Cookie': `FedAuth=${auth.FedAuth};rtFa=${auth.rtFa};`,
          'X-RequestDigest': auth.requestDigest};
}
