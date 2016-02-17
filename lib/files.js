import url from 'url';
import _ from 'lodash';
import axios from 'axios';
import uuid from 'node-uuid';
import { parseString } from 'xml2js';


const folderUrl = title => `/lists/GetFolderByServerRelativeUrl('${title}')`;

url: http://site url/_api/web/GetFolderByServerRelativeUrl('/Folder Name')/Files/Add(url='file name', overwrite=true)


  

method: POST
body: contents of binary file
headers:
    Authorization: "Bearer " + accessToken
    X-RequestDigest: form digest value
    content-type: "application/json;odata=verbose"
    content-length:length of post body
