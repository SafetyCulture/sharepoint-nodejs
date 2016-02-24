'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.Files = Files;

var _fs = require('fs');

var _fs2 = _interopRequireDefault(_fs);

var _bluebird = require('bluebird');

var _bluebird2 = _interopRequireDefault(_bluebird);

var _path = require('path');

var _path2 = _interopRequireDefault(_path);

var _lists = require('./lists');

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var readFile = _bluebird2.default.promisify(_fs2.default.readFile);

function Files(api) {
  // String builder for list URIS
  // const folderUrl = title => `GetFolderByServerRelativeUrl('${title}')`;

  return {
    /**
     * Upload a file to Sharepoint
     *
     * @param {string} list The name of the list to upload the file too
     * @param {string} filePath The location of the file to upload
     * @param {string} folderName The name of the folder in the lsit to upload too
     * @returns {Promise} Resolves on success, rejects with error from SP
     */
    uploadFile: function uploadFile(_ref) {
      var list = _ref.list;
      var fileLocation = _ref.fileLocation;
      var _ref$folderName = _ref.folderName;
      var folderName = _ref$folderName === undefined ? null : _ref$folderName;
      var _ref$overwrite = _ref.overwrite;
      var overwrite = _ref$overwrite === undefined ? false : _ref$overwrite;

      var fileName = _path2.default.basename(fileLocation);
      var folder = folderName ? folderUrl(folderName) : 'RootFolder';
      var options = '';

      if (overwrite === true) {
        options = ', overwrite=true';
      }

      return readFile(fileLocation).then(function (data) {
        var headers = { 'content-length': data.length };
        return api._axios.post((0, _lists.listURI)(list) + '/' + folder + '/Files/Add(url=\'' + fileName + '\'' + options + ')?$expand=ListItemAllFields', data, { headers: headers }).then(function (response) {
          return response.data;
        });
      });
    },

    /**
     * Add file to a list item
     *
     * @param {string} list The name of the list to upload the file too
     * @param {string} itemId The id of the item to attach the file too
     * @param {string} filePath The location of the file to upload
     * @returns {Promise} Resolves on success, rejects with error from SP
     */
    attachFileToItem: function attachFileToItem(list, itemId, fileLocation) {
      var fileName = _path2.default.basename(fileLocation);

      return readFile(fileLocation).then(function (data) {
        var headers = { 'content-length': data.length };
        return api._axios.post((0, _lists.listURI)(list) + '/items({' + itemId + ')/Files/Add(FileName=\'' + fileName + '\', overwrite=true)', data, { headers: headers });
      });
    }
  };
}