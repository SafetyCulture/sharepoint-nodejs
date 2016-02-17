import fs from 'fs';
import Promise from 'bluebird';
import path from 'path';

const readFile = Promise.promisify(fs.readFile);

export function Files(api) {
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
    uploadFile: (list, fileLocation, folderName = null) => {
      let fileName = path.basename(fileLocation);
      let folder = folderName ? folderUrl(folderName) : 'RootFolder';

      return readFile(fileLocation).then((data) => {
        let headers = {'content-length': data.length};
        return api._axios.post(`${listURI(list)}/${folder}/Files/Add(url='${fileName}', overwrite=true)?$expand=ListItemAllFields`,
                                data, {headers: headers})
                          .then((response) => {
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
    attachFileToItem: (list, itemId, fileLocation) => {
      let fileName = path.basename(fileLocation);

      return readFile(fileLocation).then((data) => {
        let headers = {'content-length': data.length};
        return api._axios.post(`${listURI(list)}/items({${itemId})/Files/Add(FileName='${fileName}', overwrite=true)`,
                                data, {headers: headers});
      });
    }
  };
}
