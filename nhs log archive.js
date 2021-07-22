const { DefaultAzureCredential } = require("@azure/identity");
const { SecretClient } = require("@azure/keyvault-secrets");
const { BlobServiceClient } = require("@azure/storage-blob");
const axios = require("axios");
let globalContext;

const storageAccountName = "storageaccountcimonafc1";
const storageContainerName = "logarchive";
const keyVaultName = "ci-monitoringv4-vault";

async function getSecret(secretName) {
    const credential = new DefaultAzureCredential();

    const url = `https://${keyVaultName}.vault.azure.net`;

    const client = new SecretClient(url, credential);
    const secret = await client.getSecret(secretName);

    return secret.value;
}

async function generateOAuthToken(clientId, clientSecret, tenantId, scope) {
    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const data = `scope=https%3A%2F%2F${scope}/.default&client_id=${clientId}&grant_type=client_credentials&client_secret=${clientSecret}`;
    const headers = { "Content-Type": "application/x-www-form-urlencoded" };
    const method = "POST";
    const options = { method, headers, data, url };

    globalContext.log("Generating OAuth token...");

    try {
        const res = await axios(options);
        return res.data.access_token;
    } catch(err) {
        globalContext.log(err.message);
        return null;
    }   
}

// Format the data returned by the log analytics api
function restructureQueryResult(queryResult) {
    let { rows, columns } = queryResult.tables[0];

    let result = [];
    rows.forEach(row => {
        let resultObject = {};
        row.forEach((cell, index) => {
            if (cell != null) {
                resultObject[columns[index].name] = cell;
            }
        });
        result.push(resultObject);
    });

    return result;
}

async function runLogAnalyticsQuery(workSpaceId, oAuthToken, query) {
    const url = `https://api.loganalytics.io/v1/workspaces/${workSpaceId}/query/?query=${encodeURI(query)}`;
    const headers = { "Authorization": "Bearer "+ oAuthToken, "Content-Type": "application/json" };
    const method = "GET";
    const options = { url, headers, method };

    try {
        const res = await axios(options);
        return res.data;
    } catch(err) {
        globalContext.log(err.message);
        return null;
    }
}

// This will get the list of tables in a log analytics workspace - not sure if it's required
async function getLogAnalyticsTables(workspaceId, oAuthToken) {
    const query = "search * | distinct Type";

    const queryResult = await runLogAnalyticsQuery(workspaceId, oAuthToken, query);

    return queryResult.tables[0].rows.map(row => row[0]);
}

async function getDataFromTable(tableName, startDateTime, workspaceId, oAuthToken) {
    const endDateTime = startDateTime + 24*60*60*1000 - 1;

    let query = `${tableName} | where TimeGenerated between(datetime(${new Date(startDateTime).toISOString()}) .. datetime(${new Date(endDateTime).toISOString()}))`;
    
    let queryResult;
    try {
        queryResult = await runLogAnalyticsQuery(workspaceId, oAuthToken, query);
    } catch (err) {
        globalContext.log(err.message);
        return null;
    }
    
    if (!queryResult) {
        return null;
    }
    return restructureQueryResult(queryResult);
}

// This will currrently save each piece of data passed as a new blob in the specified storage account
async function storeTableAsBlob(tableName, tableData, startDateTime) {
    const credential = new DefaultAzureCredential();

    const url = `https://${storageAccountName}.blob.core.windows.net`;

    const blobServiceClient = new BlobServiceClient(url, credential);

    const containerClient = blobServiceClient.getContainerClient(storageContainerName);

    const content = JSON.stringify(tableData);
    const blobName = tableName + new Date(startDateTime).toISOString() + ".json";
    const blockBlobClient = containerClient.getBlockBlobClient(blobName);
    const uploadBlobResponse = await blockBlobClient.upload(content, content.length);

    globalContext.log(`Upload block blob ${blobName} successfully`, uploadBlobResponse.requestId);
}

async function getLogAnalyticsLogs(tableName, startDateTime, workspaceId, oAuthToken) {
    let tableData = await getDataFromTable(tableName, startDateTime, workspaceId, oAuthToken);
    if (!tableData) {
        return
    }
    globalContext.log("Saving data to storage blob");
    await storeTableAsBlob(tableName, tableData, startDateTime);
}

async function getO365AdminLogs(clientId, clientSecret, tenantId, startDateTime) {
    // For this to work the app must have the api permission Office 365 Management APIs ActivityFeed.Read
    const oAuthToken = await generateOAuthToken(clientId, clientSecret, tenantId, "manage.office.com");

    // Trim the milliseconds out of the datetime
    const startTime = new Date(startDateTime).toISOString().substring(0, 19);
    const endTime = new Date(startDateTime + 24*60*60*1000 - 1).toISOString().substring(0, 19);
    
    // For each contentType, get the list of contents for a given time period
    // These need to be registed using:
    // https://manage.office.com/api/v1.0/{tenantId}/activity/feed/subscriptions/start?contentType={contentType}
    let result = {};
    const contentTypes = ["Audit.AzureActiveDirectory", "Audit.Exchange", "Audit.General", "Audit.SharePoint"]
    contentTypes.forEach(async (contentType) => {
        const url = `https://manage.office.com/api/v1.0/${tenantId}/activity/feed/subscriptions/content?contentType=${contentType}&startTime=${startTime}&endTime=${endTime}`;

        const headers = { "Authorization": "Bearer "+ oAuthToken, "Content-Type": "application/json" };
        const method = "GET";
        const options = { url, headers, method }

        try {
            const res = await axios(options);

            // For each bit of content returned by the request, make a further request to get the actual content
            res.data.forEach(async (item) => {
                const { contentUri } = item;
                const contentOptions = { url: contentUri, headers, method }
                
                try {
                    const res = await axios(contentOptions);
                    
                    // Add the data to a single object to store in a blob
                    if (contentType in result) {
                        result[contentType] = result[contentType].concat(res.data);
                    } else {
                        result[contentType] = res.data;
                    }
                } catch(err) {
                    globalContext.log(err.message);
                }
            });
        } catch(err) {
            globalContext.log(err.message);
        }

        storeTableAsBlob("Office365Admin", result, startDateTime);
    });
}

module.exports = async function (context, req) {
    // So you can write to the log when debugging
    globalContext = context;

    let startDateTime = 0;
    try {
        startDateTime = parseInt(req.query.startDateTime);
    }
    catch (err) {
        globalContext.log("No start date provided. Defaulting to yesterday.")
    }
    if (!startDateTime) {
        startDateTime = new Date().setHours(0,0,0,0) - 24*60*60*1000
    }

    const clientId = await getSecret("azureclientid");
    const clientSecret = await getSecret("azuresecret");
    // const subscriptionId = await getSecret("azuresubscriptionid");
    const tenantId = await getSecret("azuretenant");
    const workspaceId = await getSecret("loganalyticsworkspaceid");

    const oAuthToken = await generateOAuthToken(clientId, clientSecret, tenantId, "api.loganalytics.io");

    await getLogAnalyticsLogs("ADFSAudit", startDateTime, workspaceId, oAuthToken);
    await getLogAnalyticsLogs("AzureAdmin", startDateTime, workspaceId, oAuthToken);

    // This has not yet been tested!!
    await getO365AdminLogs(clientId, clientSecret, tenantId, startDateTime);

    const responseMessage = "Successfully got and saved data.";

    globalContext.log("JavaScript HTTP trigger function processed a request.");

    globalContext.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };
}
