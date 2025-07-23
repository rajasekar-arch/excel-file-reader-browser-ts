import {readFileAsArrayBuffer, getWorkbookFromData} from '../src/file-reader-functions/excel-file-functions';



// file path
const filePath = './Untitled spreadsheet.xlsx';
const file = new File([new Blob()], filePath); // Simulating a File object for testing

readFileAsArrayBuffer(file).then(arrayBuffer => {
  console.log('ArrayBuffer:', arrayBuffer);
}).catch(error => {
  console.error('Error reading file as ArrayBuffer:', error);
});