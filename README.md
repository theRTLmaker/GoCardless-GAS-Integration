# GoCardless-GAS-Integration

Google Apps Script integration with GoCardless API to create a personal finance tracker

Google Sheets Template: [Link](https://docs.google.com/spreadsheets/d/1nrxX-XEwtZW7U2Wft8X2VwDHoo80yLWcaF9ESQgHve0/copy?usp=sharing)

## Description

This project integrates Google Apps Script with the GoCardless API to create a powerful personal finance tracker. It allows users to connect their bank accounts, fetch transaction data, and manage their finances directly within Google Sheets.

## Features

- Connect multiple bank accounts using GoCardless API
- Fetch and update transaction data automatically
- Categorize transactions
- Generate financial reports and insights
- User-friendly bank selection interface with search functionality
- Customizable column mapping for transaction data
- Error handling and rate limit management
- Support for both booked and pending transactions

## Setup

1. Create a new Google Sheet
2. Open the Script Editor (Tools > Script editor)
3. Copy the contents of the `src` folder into the Script Editor
4. Set up GoCardless API credentials (instructions below)

## Usage

1. Open your Google Sheet
2. Use the "GoCardless" menu to:
   - Initialize the integration
   - Link bank accounts
   - Fetch transaction data
   - Customize column mappings

## Bank Account Linking

1. Select "Link Account" from the GoCardless menu
2. Enter the country code for your bank
3. Use the search function to find your bank in the list
4. Click on your bank to start the linking process
5. Follow the on-screen instructions to complete the authentication

## Customizing Column Mappings

1. Select "Configure Column Mappings" from the GoCardless menu
2. Use the dialog to select which transaction fields to include
3. Specify the column letter for each selected field
4. Save your customized mapping

## GoCardless API Setup

1. Sign up for a GoCardless account
2. Obtain your API credentials (Secret ID and Secret Key)
3. Use these credentials when initializing the integration in Google Sheets

## Error Handling

- The script includes robust error handling for API requests
- Rate limit errors are managed with appropriate waiting periods
- User-friendly error messages are displayed for common issues

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License. See the [LICENSE](#license-1) section at the end of this file for details.

## License

MIT License

Copyright (c) 2024 Francisco Soares Mendes

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
