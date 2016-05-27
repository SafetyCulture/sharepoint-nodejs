import _ from 'lodash';

export const LIST_TEMPLATES = {
  STANDARD: 100,
  LIBRARY: 101
};

// String builder for list URIS
export const listURI = title => `/lists/GetByTitle('${title}')`;
export const listType = name => `SP.Data.${name.replace(/\/|\-/g, '').replace(/ /g, '_x0020_')}ListItem`;

// Small helper to replace spaces in keys with '_x0020_' within an object
export const fillSpaces = data =>
  _.mapKeys(data, (val, key) => key.replace(/ /g, '_x0020_'));
