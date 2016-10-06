# SharePoint Node.js Module


## General

The 'core' SharePoint integration library. It:

a) Does SharePoint authentication with:

* OAuth 2.0
* Digest auth

b) Batch-uploads data to SharePoint

Only works with SharePoint Online so far.

## Dependencies

Node 4 LTS, ES6

## Tests

The tests point to an internal test site created specifically for this reason:

https://safetyculture.sharepoint.com/IntegrationAutomatedTestSite

## Examples
 
Using username/password auth: `example/digest/index.js`
Using OAuth  2.0: `example/oauth/server.js`
