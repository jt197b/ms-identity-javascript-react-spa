/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { LogLevel } from "@azure/msal-browser";

/**
 * Configuration object to be passed to MSAL instance on creation. 
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md 
 */
export const msalConfig = {
    auth: {
        clientId: "d5036aa2-0da6-4f55-b76d-83c7ae38de9b",
        authority: "https://login.microsoftonline.com/e741d71c-c6b6-47b0-803c-0f3b32b07556",
        redirectUri: "http://localhost:3000/"
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {	
        loggerOptions: {	
            loggerCallback: (level, message, containsPii) => {	
                if (containsPii) {		
                    return;		
                }		
                switch (level) {
                    case LogLevel.Error:
                        console.error(message);
                        return;
                    case LogLevel.Info:
                        console.info(message);
                        return;
                    case LogLevel.Verbose:
                        console.debug(message);
                        return;
                    case LogLevel.Warning:
                        console.warn(message);
                        return;
                    default:
                        return;
                }	
            }	
        }	
    }
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit: 
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
export const scopeBase = ["https://analysis.windows.net/powerbi/api/Report.Read.All"];

/**
 * Add here the scopes to request when obtaining an access token for MS Graph API. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
export const graphConfig = {
    graphMeEndpoint: "Enter_the_Graph_Endpoint_Herev1.0/me" //e.g. https://graph.microsoft.com/v1.0/me
};

// End point URL for Power BI API
export const powerBiApiUrl = "https://api.powerbi.com/";

// Client Id (Application Id) of the AAD app.
export const clientId = "d5036aa2-0da6-4f55-b76d-83c7ae38de9b";

// Id of the workspace where the report is hosted
export const workspaceId = "a81e25a2-2ee6-40bb-a4ac-9a2cb4538f01";

// Id of the report to be embedded
export const reportId = "10cb3f4c-3c9c-4d45-a18f-b40a21da4bc0";

// Id of the dataset to be embedded
export const datasetId = "def547c7-5fcb-4e75-9e23-99c5168906e0";
