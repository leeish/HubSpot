const axios = require('axios').default;

exports.main = (context, sendResponse) => {

    try {

        if(context.params.zip.length == 0 || context.params.zip[0] == ''){
            console.log(`Error with params.zip ${params.zip}`)
            throw new Error('Argument Error: Zip code is required');
        }

        const data = JSON.stringify({
            "operationName": "zips",
            "query": `query zips($zip: String!) { 
                HUBDB { 
                    zip_codes_collection(  filter: {zip__eq: $zip}) { items { city state } } 
                    zip_codes_2_collection(filter: {zip__eq: $zip}) { items { city state } } 
                    zip_codes_3_collection(filter: {zip__eq: $zip}) { items { city state } } 
                    zip_codes_4_collection(filter: {zip__eq: $zip}) { items { city state } }
                }
            }`,
            "variables": { "zip": context.params.zip[0] }
        });

        var config = {
            method: 'post',
            url: 'https://api.hubapi.com/collector/graphql',
            headers: { 
            'Authorization': `Bearer ${context.secrets.TOKEN}`,
            'Content-Type': 'application/json'
            },
            data : data
        };

        axios(config)
            .then(function (response) {
                const result = extractZipData(response.data.data.HUBDB);
                sendResponse({
                    body: result,
                    statusCode: (Object.keys(result).length == 0) ? 404 : 200
                });
            })
            .catch(function (error) {
                throw new Error(error);
            })
    } catch(e){
        sendResponse({
            body: `${e}`,
            statusCode: 500,
        });
    }
    
};

function extractZipData(data){
    for(let collection of Object.keys(data)){
        if(data[collection].items.length){
            return data[collection].items[0];
        }
    }
    return {};
}