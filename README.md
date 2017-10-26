# GoogleSheetCryptoPortfolio
A dotnet core application which pushes crypto currency value changes to a google sheet

This application will update a basic google spreadsheet template which will keep track of your crypto currency portfolio for Bitcoin, Ethereum, Ripple, Litecoin and OMG.

The sheet will be updated every time the process runs, using APIs from Crypto.watch

## Setup
To use this program you will need to configure the following
 - 1. Template Google Spreadsheet
 - 2. Google Sheets service account (For client_secret.json)
 - 3. AppConfig.json

Then build and run the program, schedule every 5/10 minutes.

### 1. Template Google Spreadsheet
Make a **copy** of this Google spreadsheet in your own Google Drive: 
#### USD Base Currency
https://docs.google.com/spreadsheets/d/1AvrVp56bQA5Fa04IxKn3ScwdM59QkReHWjO6qnVtXSA/
#### EUR Base Currency
https://docs.google.com/spreadsheets/d/11WWKEgkNgIoxfv8NtBu1R_NRWrbEqPzv-2D3qeG9N1E/

### 2. Set up the Google Sheets service account
Go to https://console.developers.google.com to set up a service account. 

To do this you need to perform multiple steps:
 - You will need to set up a Project, which you can name however you want
 - Enable the Google Drive and Google Sheets APIs
 - Visit the "Credentials" area and "Manage Service Accounts"
 - Add a new service account, I chose the 'Project Owner' role, tick "Furnish Private Key" and choose the JSON method
 - Rename the file that is downloaded to 'client_secret.json' and deploy this in the root of your application install
 
### 3. AppConfig.json
The AppConfig.json file will store the details of which spreadsheet to connect to and which tab to update. 

Open your copy of the template and copy the Spreadsheet ID from the URL.  The ID is the long string in the URL, e.g. the template spreadsheet linked above has the ID **1AvrVp56bQA5Fa04IxKn3ScwdM59QkReHWjO6qnVtXSA**

In the AppConfig.json file create a JSON structure for the Spreadsheet ID, tab name and base currency (only USD and EUR supported base currencies) in the following format:

```json
{
  "sheets": [
    {
      "id": "your-sheet-id",
      "tabname": "My Portfolio",
	    "basecurrency": "USD"
    }
  ]
}
```

Note, you can update more than one sheet by adding multiple items to the array.

## Build & Deploy
Build the application to your preferred target, place your client_secret.json and AppConfig.json in your release path and run the program.

NOTE: Has not been tested in any other platform other than Windows 10 (winx64)
