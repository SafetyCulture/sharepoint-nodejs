export const LIST_TEMPLATES = {
  STANDARD: 100,
  LIBRARY: 101
};

// String builder for list URIS
export const listURI = title => `/lists/GetByTitle('${title}')`;
export const listType = name => `SP.Data.${name.replace(/ /g, '_x0020_')}Item`;
