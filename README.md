# Printful Syncing Automation App

A Node.js application that automates the process of finding un-synced listing variations from a Printful store and syncing them with the correct Printful product/image.

## Features

- Fetches listings with zero synced variations from Printful and generates a template spreadsheet.
- Allows the user to upload a completed template spreadsheet.
- Automatically syncs the variations on Printful based on the uploaded spreadsheet.
- Provides a simple user interface in the browser using Handlebars.
- Packaged into a portable executable file using PKG, allowing the client to run the app on their PC as needed.

## How it Works

The application provides two main functionalities:

1. **Creating a Template Spreadsheet**: Initiated by a GET request to the `/createTemplate` endpoint. This fetches listings with zero synced variations from Printful and generates a template spreadsheet. The spreadsheet is saved in the root directory of the application.

2. **Syncing Product Variations**: Initiated by a PUT request to the `/syncPrintful` endpoint. The user can upload a spreadsheet with product variations to be synced to Printful. The application reads the spreadsheet, and for each row, it sends a request to Printful to sync the product variation.

The user interface is rendered using Handlebars. After each operation, a success or failure message is displayed on the page.

## Running the App

The application is packaged into a portable executable file using PKG. This allows the client to run the app on their PC as needed.

## Dependencies

The application uses the following dependencies:

- body-parser
- dotenv
- exceljs
- express
- hbs
- multer
- node-fetch

## Note

The Printful API key is stored in a `.env` file which is ignored by Git.