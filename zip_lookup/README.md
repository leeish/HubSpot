# Zip Lookup Across HubDB Tables

HubDB currently has a limitation of 10K (soon to be 20k) rows. There are something upwards of 44k zip codes in the US. If a customer wants to do some sort of lookup against zip codes they'd need to use a third party API or do mulitple API calls to HubDB until they find the right data.

This serverless function is an example of using multiple HubDB tables to store data and then using GraphQL to search across those tables at the same time. The GraphQL endpoint could be hit directly, but this encapsulates some logic as well as provides a public endpoint to serve this purpose.

## Requirements

The serverless functions expects requires the following: 

1. A secret named `TOKEN` with the following scopes:
    a. `hubdb`
    b. `collector.graphql_schema.read`
    c. `collector.graphql_query.execute`
2. Four HubDB tables named:
    a. `zip_codes_collection`
    b. `zip_codes_collection_2`
    c. `zip_codes_collection_3`
    d. `zip_codes_collection_4`

## Usage

Once deployed, run a GET request against the URL: `https://[portal_id].hs-sites.com/_hcms/api/zip-lookup?zip=[zip_code]`

[Example Link](https://21975837.hs-sites.com/_hcms/api/zip-lookup?zip=75081)

## TODO

1. Handle missing / invalid zip codes
2. Refactor so the tables / parameters can be configurable and a single endpoint could be reused