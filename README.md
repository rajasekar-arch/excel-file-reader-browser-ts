# excel-file-reader-browser-ts 
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

    License

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

npm install excel-file-reader-browser-ts xlsx \# or yarn add
excel-file-reader-browser-ts xlsx

Note: While excel-file-reader-browser-ts handles xlsx internally, you
might still need to install xlsx directly if your project uses other
xlsx utilities.

# Usage 
Here\'s how to use the functions provided by
excel-file-reader-browser-ts in a typical web application scenario.

First, import the necessary functions:

// For TypeScript import { getExcelRowCount, getExcelHeaders,
isExcelColumnPopulated } from \'excel-file-reader-browser-ts\'; // For
JavaScript (CommonJS) // const { getExcelRowCount, getExcelHeaders,
isExcelColumnPopulated } = require(\'excel-file-reader-browser-ts\');

Let\'s assume you have an HTML input for file selection:

\<!\-- In your HTML template (e.g., Angular component template, React
JSX, Vue template) \--\> \<input type=\"file\" id=\"excelFileInput\"
accept=\".xlsx, .xls\"\> \<button id=\"processButton\"\>Process
Excel\</button\> \<div id=\"results\"\>\</div\> \<div
id=\"error\"\>\</div\>

Getting the File Object from User Input The first step is always to get
the File object from the user\'s input.

// Example in plain JavaScript/TypeScript (adapt for your framework)

const excelFileInput = document.getElementById(\'excelFileInput\') as
HTMLInputElement; let selectedFile: File \| null = null;

excelFileInput.addEventListener(\'change\', (event) =\> { const input =
event.target as HTMLInputElement; if (input.files && input.files.length
\> 0) { selectedFile = input.files\[0\]; console.log(\'File selected:\',
selectedFile.name); } else { selectedFile = null; console.log(\'No file
selected.\'); } });

document.getElementById(\'processButton\')?.addEventListener(\'click\',
async () =\> { const resultsDiv = document.getElementById(\'results\');
const errorDiv = document.getElementById(\'error\'); if (resultsDiv)
resultsDiv.innerHTML = \'\'; if (errorDiv) errorDiv.innerHTML = \'\';

if (!selectedFile) { if (errorDiv) errorDiv.textContent = \'Please
select an Excel file first.\'; return; }

// Now you can call the package functions with \`selectedFile\` //
Example calls are shown below for each function. // Remember to handle
errors for each call! });

Get Total Row Count This function returns the total number of rows in a
specified Excel sheet, including header rows.

// Inside your processButton click handler or similar async function //
assuming \`selectedFile\` is available

try { const totalRows = await getExcelRowCount(selectedFile,
\'Sheet1\'); // Or omit \'Sheet1\' for the first sheet if (resultsDiv)
resultsDiv.innerHTML += \`\<p\>Total rows in Sheet1:
\<strong\>\${totalRows}\</strong\>\</p\>\`; } catch (error: any) { if
(errorDiv) errorDiv.textContent = \`Error getting row count:
\${error.message}\`; }

// Example for an empty sheet (if your Excel file has one named
\'EmptySheet\') try { const emptyRows = await
getExcelRowCount(selectedFile, \'EmptySheet\'); if (resultsDiv)
resultsDiv.innerHTML += \`\<p\>Total rows in EmptySheet:
\<strong\>\${emptyRows}\</strong\>\</p\>\`; // Expected: 0 } catch
(error: any) { if (errorDiv) errorDiv.textContent = \`Error getting
empty sheet row count: \${error.message}\`; }

Get Headers This function retrieves an array of header names from the
first row of an Excel sheet.

// Inside your processButton click handler or similar async function //
assuming \`selectedFile\` is available

try { const headers = await getExcelHeaders(selectedFile, \'Sheet1\');
// Or omit \'Sheet1\' for the first sheet if (resultsDiv)
resultsDiv.innerHTML += \`\<p\>Headers: \<strong\>\${headers.join(\',
\')}\</strong\>\</p\>\`; } catch (error: any) { if (errorDiv)
errorDiv.textContent = \`Error getting headers: \${error.message}\`; }

Check if Column is Populated This function checks if a column,
identified by its header name, contains any non-empty values in its data
rows. It returns true if at least one cell in that column (excluding the
header) has a value, false otherwise.

// Inside your processButton click handler or similar async function //
assuming \`selectedFile\` is available

try { // Check a populated column const isProductNamePopulated = await
isExcelColumnPopulated(selectedFile, \'Product Name\', \'Sheet1\'); if
(resultsDiv) resultsDiv.innerHTML += \`\<p\>Is \'Product Name\' column
populated? \<strong\>\${isProductNamePopulated}\</strong\>\</p\>\`; //
Expected: true

// Check a potentially empty column (e.g., \'Description\' from previous
example) const isDescriptionPopulated = await
isExcelColumnPopulated(selectedFile, \'Description\', \'Sheet1\'); if
(resultsDiv) resultsDiv.innerHTML += \`\<p\>Is \'Description\' column
populated? \<strong\>\${isDescriptionPopulated}\</strong\>\</p\>\`; //
Expected: true (if it has any data)

// Check a non-existent header (this will throw an error) try { await
isExcelColumnPopulated(selectedFile, \'NonExistentHeader\', \'Sheet1\');
} catch (error: any) { if (errorDiv) errorDiv.textContent = \`Error for
non-existent header (expected): \${error.message}\`; }

} catch (error: any) { // Catch any unexpected errors from the above
calls if (errorDiv) errorDiv.textContent = \`An unexpected error
occurred during column check: \${error.message}\`; }

API Reference All functions are asynchronous and return Promises. They
accept a File object (or null) as the primary input.

getExcelRowCount(file: File \| null, sheetName?: string):
Promise\<number\> file (File \| null): The File object obtained from a
user\'s file input. If null, an error will be thrown.

sheetName (string, optional): The name of the specific sheet to read. If
omitted, the first sheet in the workbook will be used.

Returns: Promise\<number\> - The total number of rows in the specified
sheet. This count includes header rows.

getExcelHeaders(file: File \| null, sheetName?: string):
Promise\<string\[\]\> file (File \| null): The File object obtained from
a user\'s file input. If null, an error will be thrown.

sheetName (string, optional): The name of the specific sheet to read. If
omitted, the first sheet in the workbook will be used.

Returns: Promise\<string\[\]\> - An array of strings representing the
header names from the first row of the specified sheet. Returns an empty
array if no headers are found or the sheet is empty.

isExcelColumnPopulated(file: File \| null, headerName: string,
sheetName?: string): Promise\<boolean\> file (File \| null): The File
object obtained from a user\'s file input. If null, an error will be
thrown.

headerName (string): The exact name of the header column to check for
values. This is case-sensitive.

sheetName (string, optional): The name of the specific sheet to read. If
omitted, the first sheet in the workbook will be used.

Returns: Promise\<boolean\> - true if the specified column contains at
least one non-empty (not null, undefined, or whitespace-only string)
value in its data rows; false otherwise.

Error Handling All functions will throw an Error (and reject their
Promises) in the following scenarios:

The file argument is null.

The sheetName provided does not exist in the Excel workbook.

For isExcelColumnPopulated, if the headerName does not exist in the
specified sheet.

If the Excel file is corrupted or cannot be parsed by the xlsx library.

Any issues during the internal FileReader process (e.g., file access
errors).

It is highly recommended to wrap calls to these functions in
try\...catch blocks or use .catch() with Promises to handle potential
errors gracefully in your UI.

# Contributing 
Contributions are welcome! If you find a bug or have a
feature request, please open an issue on the GitHub repository.

# License 
This project is licensed under the MIT License - see the LICENSE
file for details.
