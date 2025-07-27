# excel-file-reader-browser-ts

[![Hits](https://hits.sh/github.com/rajasekar-arch/excel-file-reader-browser-ts.svg?style=flat-square)](https://hits.sh/github.com/rajasekar-arch/excel-file-reader-browser-ts/)

A robust and efficient TypeScript utility
designed specifically for browser environments to read Excel files
(.xlsx, .xls) from user-uploaded File objects. This package provides
convenient functions to easily get the total number of rows, extract
sheet headers, and check if specific columns contain data, all
client-side.

# Table of Contents Description

    Features

    Installation

    Usage

    Getting the File Object from User Input

    Get Total Row Count

    Get Headers

    Check if Column is Populated

    API Reference

    Error Handling

    Contributing

    Demo Link

# Description

excel-file-reader-browser-ts addresses the need for
client-side Excel file processing in web applications. Unlike
server-side solutions that rely on file paths, this package directly
works with File objects obtained from HTML \<input type=\"file\"\>
elements. It\'s built with TypeScript, offering strong type safety and
clear interfaces, making it ideal for Angular, React, Vue, and other
front-end frameworks.

It internally utilizes the browser-compatible features of the xlsx
library, abstracting away the complexities of FileReader and binary data
parsing, so you can focus on your application logic.

# Features

Browser-Compatible: Works directly with File objects, no
Node.js fs module required.

Get Total Row Count: Quickly determine the total number of rows in an
Excel sheet, including headers.

Extract Headers: Retrieve the names of all columns from the first row of
a specified sheet.

Check Column Population: Verify if a column (identified by its header
name) contains any non-empty values in its data rows.

TypeScript Support: Fully typed for a better developer experience and
compile-time safety.

Promise-based API: All functions return Promises for easy asynchronous
handling.

Robust Error Handling: Provides clear error messages for issues like no
file selected, sheet not found, or header not found.

Installation To install excel-file-reader-browser-ts in your front-end
project, use npm or yarn:

# npm install excel-file-reader-browser-ts xlsx

or

# yarn add excel-file-reader-browser-ts xlsx

Note: While excel-file-reader-browser-ts handles xlsx internally, you
might still need to install xlsx directly if your project uses other
xlsx utilities.

# Usage

```javascript
import {
  readFileAsArrayBuffer,
  getExcelRowCount,
  getExcelHeaders,
  isExcelColumnPopulated,
  findSpecialCharacterCells,
} from "excel-file-reader-browser-ts";

// Inside class or function

// Get File Name Extension

const fileType = getFileType(fileName);
// console.log(fileType)


// To Read the excel content as a array buffer
const buffer = await readFileAsArrayBuffer(file);

// Total Row count in excel sheet
const rowCount = await getExcelRowCount(buffer);

// Get Column Count
const columnCount = await getExcelColumnCount(buffer);

// Headers only
const allHeaders = await getExcelHeaders(buffer);

// Is Any special chars are avaialble in the cells
const specialChars = await findSpecialCharacterCells(buffer);

// Get Number of Columns
const columnsCount = await getExcelColumnCount(buffer);

// Get Raw Data in a 2D array Format

const getRawData = await getExcelRawData(buffer);

// Get Excel Meta data

const meta = await getExcelMetadata(file);
// console.log(meta);

// Download the uploaded file as CSV

const buffer = await readFileAsArrayBuffer(file);
await downloadExcelAsCSV(buffer);

// Download the uploaded file as JSON

const buffer = await readFileAsArrayBuffer(file);
await downloadExcelAsJSON(buffer);
```

# Contributing

Contributions are welcome! If you find a bug or have a
feature request, please open an issue on the GitHub repository.

# Demo Link

[Try This Link to See the Real Time Usage](https://rajasekar-arch.github.io/excel-file-reader-ui/)

-It helps you to understand, how this package will work in real time

❤️ Support & Share

If you find `excel-file-reader-browser-ts` helpful, please consider sharing your experience!
Post about it on LinkedIn, Twitter, or any social media platform.
Tag me or the project so we can feature your post!!
Or simply drop me an email at: `rajasekar_e_c@outlook.com` – I’d love to hear how you are using this package!

Your support helps others discover this library and keeps the project growing!
