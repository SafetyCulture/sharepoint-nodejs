import _ from 'lodash';

export const LIST_TEMPLATES = {
  STANDARD: 100,
  LIBRARY: 101
};

// String builder for list URIS
export const listURI = title => `/lists/GetByTitle('${title}')`;

// Escapes non-alphanumerical chars to appropriate SharePoint internal code
// https://abstractspaces.wordpress.com/2008/05/07/sharepoint-column-names-internal-name-mappings-for-non-alphabet/
export const sharepointEscapeChars = (str) => {
  return str.replace(/ /g, '_x0020_')      // whitespace
            .replace(/\`/g, '_x0060_')     // backtick
            .replace(/\//g, '_x002f_')     // forwardslash
            .replace(/\./g, '_x002e_')     // period
            .replace(/\,/g, '_x002c_')     // comma
            .replace(/\?/g, '_x003f_')     // questionmark
            .replace(/\>/g, '_x003e_')     // right angle bracket
            .replace(/\</g, '_x003c_')     // left angle bracket
            .replace(/\\/g, '_x005c_')     // backslash
            .replace(/\'/g, '_x0027_')     // apostrophe
            .replace(/\;/g, '_x003b_')     // semicolon
            .replace(/\|/g, '_x007c_')     // pipe
            .replace(/\"/g, '_x0022_')     // quotation
            .replace(/\:/g, '_x003a_')     // colon
            .replace(/\}/g, '_x007d_')     // right curly brace
            .replace(/\{/g, '_x007b_')     // left curly brace
            .replace(/\=/g, '_x003d_')     // equals sign
            .replace(/\-/g, '_x002d_')     // minus sign
            .replace(/\+/g, '_x002b_')     // plus sign
            .replace(/\)/g, '_x0029_')     // right paranthesis
            .replace(/\(/g, '_x0028_')     // left paranthesis
            .replace(/\*/g, '_x002a_')     // asterisk
            .replace(/\&/g, '_x0026_')     // ampersand
            .replace(/\^/g, '_x005e_')     // caret
            .replace(/\%/g, '_x0025_')     // percent
            .replace(/\$/g, '_x0024_')     // dollar
            .replace(/\#/g, '_x0023_')     // hash
            .replace(/\@/g, '_x0040_')     // at symbol
            .replace(/\!/g, '_x0021_')     // exclamation
            .replace(/\~/g, '_x007e_');    // tilde
};
export const listType = name => `SP.Data.${sharepointEscapeChars(name.charAt(0).toUpperCase() + name.slice(1))}ListItem`;
export const libraryType = name => `SP.Data.${sharepointEscapeChars(name.charAt(0).toUpperCase() + name.slice(1))}Item`;
// Small helper to replace spaces in keys with '_x0020_' within an object
export const fillSpaces = data => _.mapKeys(data, (val, key) => sharepointEscapeChars(key));
