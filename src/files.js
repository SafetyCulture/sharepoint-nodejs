import fs from 'fs';
import Promise from 'bluebird';
import path from 'path';

import { listURI } from './lists';

const readFile = Promise.promisify(fs.readFile);

export function Files(api) {
  return {
    /**
     * Upload a file to a Document Library or Folder in Sharepoint
     *
     * @param {string} libraryTitle The title of the library to upload the file too
     * @param {string} folderRelativeUrl The Server Relative URL of the target folder
     * @param {string} filePath The location of the file to upload
     * @returns {Promise} Resolves on success, rejects with error from SP
     */
    uploadFileToLibraryOrFolder: ({ libraryTitle, folderRelativeUrl, fileLocation, overwrite = false }) => {
      let fileName = path.basename(fileLocation);
      let options = overwrite ? ', overwrite=true' : '';
      let targetUri = '';
      if (libraryTitle) {
        targetUri = `${listURI(libraryTitle)}/RootFolder/Files/Add(url='${fileName}'${options})?$expand=ListItemAllFields`;
      } else if (folderRelativeUrl) {
        targetUri = `/GetFolderByServerRelativeUrl('${folderRelativeUrl}')/Files/Add(url='${fileName}'${options})?$expand=ListItemAllFields`;
      } else {
        return Promise.reject(`Library Title or Folder's Server Relative URL must be provided.`);
      }
      return readFile(fileLocation).then((data) => {
        let headers = {'content-length': data.length};
        return api._axios.post(targetUri, data, {headers: headers})
        .then(response => response.data);
      });
    },

    /**
     * Attach a file to a list item
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
        return api._axios.post(`${listURI(list)}/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`,
                                data, {headers: headers});
      });
    }
  };
}
