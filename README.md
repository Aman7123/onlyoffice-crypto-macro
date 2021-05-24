# onlyoffice-crypto-macro

## Installation

Inside of the OnlyOffice Spreadsheet application select `Plugins` from the menu bar then select `Macros` from that submenu. Paste the macros.js file inside a new macro.

## Documentation
For this to work correctly you should setup the config options at the top of the function in the `macro.js` file. Layout the table in OnlyOffice Spreadsheet like show below:

| Coin | Holding | USD |   |   |
|------|---------|-----|---|---|
| BTC  | 0.150   |     |   |   |
| DOGE | 301     |     |   |   |
| ADA  | 120     |     |   |   |

The function expects to find the coins market abbriviation within the `A` column on a spreadsheet. The macro reaches out to coinapi.io for the current price of the coin in USD, this is an HTTP GET request and the response is in JSON. Our macro then parses this JSON to find the `rate` key/value pair and places that in the USD row for each coin.

When making too many requests in one day this could lead to maxing out the rate limit of the coinapi.io account. To help solve the issue of blank prices being loaded into the spreadsheet I inluded an IF statement to ensure the response code came back as a 200 before populating.

The debug function is still a work in progress!

## Environment
* https://github.com/ONLYOFFICE/docker-onlyoffice-owncloud
