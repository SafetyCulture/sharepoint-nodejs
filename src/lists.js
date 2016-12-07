import { mapKeys } from 'lodash';

export const STD_NEWLINE_REGEX = /(\n|\r|\\r|\\n)/gm;
export const AFFIXED_NEWLINE_REGEX = /^(\n|\r|\\r|\\n)$/gm;
export const LIST_TEMPLATES = {
  STANDARD: 100,
  LIBRARY: 101
};

// String builders for list URIS
// List title must be URI Encoded and Apostrophes must be duplicated to avoid errors
export const listURI = title => `/lists/GetByTitle('${encodeURI(title).replace(/\'/g, "''")}')`;
export const listGuidUri = guid => `/Lists(guid'${guid}')`;

// This function converts a string to the encoding style Sharepoint uses.
// 1. Handles newlines and newline strings.
// 2. URI/Percent encode the string.
// 3. Using regex, replace all encoded symbol codes with the Sharepoint equivalent.
// 4. Finish off by explicitly replacing safe URI symbols with Sharepoint codes.
// References:
// http://www.blooberry.com/indexdot/html/topics/urlencoding.htm
// https://abstractspaces.wordpress.com/2008/05/07/sharepoint-column-names-internal-name-mappings-for-non-alphabet/
export const sharepointEscapeChars = (str) => {
  // Strip leading/trailing newlines and replace others with whitespace
  let result = str.replace(AFFIXED_NEWLINE_REGEX, '').replace(STD_NEWLINE_REGEX, ' ');
  return encodeURI(result).replace(/(\%)([a-zA-Z0-9]{2})/g, '_x00$2_') // convert all unreserved
                          .replace(/\$/g, '_x0024_')  // $
                          .replace(/\-/g, '_x002d_')  // -
                          .replace(/\./g, '_x002e_')  // .
                          .replace(/\+/g, '_x002b_')  // +
                          .replace(/\!/g, '_x0021_')  // !
                          .replace(/\&/g, '_x0026_')  // &
                          .replace(/\)/g, '_x0029_')  // (
                          .replace(/\(/g, '_x0028_')  // )
                          .replace(/\?/g, '_x003f_')  // ?
                          .replace(/\=/g, '_x003d_')  // =
                          .replace(/\*/g, '_x002a_')  // *
                          .replace(/\,/g, '_x002c_')  // ,
                          .replace(/\//g, '_x002f_')  // /
                          .replace(/\'/g, '_x0027_')  // '
                          .replace(/\@/g, '_x0040_')  // @
                          .replace(/\:/g, '_x003a_')  // :
                          .replace(/\;/g, '_x003b_')  // ;
                          .replace(/\#/g, '_x0023_'); // #
};

export const listType = name => `SP.Data.${sharepointEscapeChars(name.charAt(0).toUpperCase() + name.slice(1))}ListItem`;
export const libraryType = name => `SP.Data.${sharepointEscapeChars(name.charAt(0).toUpperCase() + name.slice(1))}Item`;
// Small helper to replace spaces in keys with '_x0020_' within an object
export const fillSpaces = data => mapKeys(data, (val, key) => sharepointEscapeChars(key));
