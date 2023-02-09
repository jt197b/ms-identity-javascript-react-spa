import React, { useState, useRef } from 'react';
import './styles/App.css';
import { PageLayout } from './components/PageLayout';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import Button from 'react-bootstrap/Button';
import { scopeBase, workspaceId, reportId, powerBiApiUrl, datasetId } from './authConfig';
import { callMsGraph } from './graph';
import { ProfileData } from './components/ProfileData';
import { MsalContext } from "@azure/msal-react";
import { useIsAuthenticated } from "@azure/msal-react";
import { useEffect } from 'react';
import ReportEmbed from './components/ReportEmbed';

/**
 * Renders information about the signed-in user or a button to retrieve data about the user
 */
const ProfileContent = () => {
    const { instance, accounts } = useMsal();
    const [graphData, setGraphData] = useState(null);
    const reportRef = useRef(null);

    const [error, setError] = useState("");

    // msal
    const [accessToken, setAccessToken] = useState(null);
    const [username, setUsername] = useState("");
    const isAuthenticated = useIsAuthenticated(false);

    // pbi
    const [embedUrl, setEmbedUrl] = useState("https://app.powerbi.com/reportEmbed?reportId=10cb3f4c-3c9c-4d45-a18f-b40a21da4bc0&groupId=a81e25a2-2ee6-40bb-a4ac-9a2cb4538f01&w=2");
    const [embedToken, setEmbedToken] = useState("");

    const loginRequest = {
        scopes: scopeBase,
        account: accounts[0]
    };


    function authenticate() {

        console.log(accounts);
    
        // Silently acquires an access token which is then attached to a request for MS Graph data
        instance
            .acquireTokenSilent(loginRequest)
            .then((resp) => {
                console.log(resp.accessToken);
                setAccessToken(resp.accessToken);
                return resp.accessToken;
                //let reportData = getEmbedUrl();
                //console.log(reportData);
            })
    }

    const getEmbedToken  = () => {
        const reportInGroupApi = powerBiApiUrl + "v1.0/myorg/groups/" + workspaceId + "/reports/" + reportId;

        
        // Get report info by calling the PowerBI REST API
        fetch(reportInGroupApi, { 
            method: 'GET', 
            headers: new Headers(
                {
                'Authorization': 'Bearer ' + accessToken, 
                "datasets": [
                    {
                    "id": datasetId
                    }
                ],
                "reports": [
                    {
                    "allowEdit": true,
                    "id": reportId
                    }
                ]
            }),
        }).then(resp => 
            {
                console.log(resp);
                return resp.json();
            }).then(function(data) {
              // `data` is the parsed version of the JSON returned from the above endpoint.
              console.log(data);
            });
    }

    const getEmbedUrl = () => {

        // Get report info by calling the PowerBI REST API
        fetch("https://api.powerbi.com/v1.0/myorg/groups/a81e25a2-2ee6-40bb-a4ac-9a2cb4538f01/reports/10cb3f4c-3c9c-4d45-a18f-b40a21da4bc0", { 
            method: 'GET', 
            headers: new Headers(
                {
                'Authorization': 'Bearer ' + accessToken
            }),
        }).then(resp => 
            {
                console.log(resp);
                return resp.json();
            }).then(function(data) {
              // `data` is the parsed version of the JSON returned from the above endpoint.
              console.log(data);
            });
    }

    useEffect(() => {
        console.log(accessToken)
    },[accessToken]);

    return (
        <>
            <h5 className="card-title">Welcome {accounts[0].name}</h5>
            {!accessToken || !embedUrl ? (
                <div>
                    <Button variant="secondary" onClick={authenticate}>
                        Get Access and Embed Token
                    </Button>
                   
                </div>
            ) : (
                <div>
                    <p>Access Token: </p>
                    {accessToken}

                    <p>Embed URL: </p>
                    {embedUrl}
                    <ReportEmbed
                        accessToken={accessToken}
                        embedUrl={embedUrl}
                    />
                </div>
            )}
        </>
    );
};

/**
 * If a user is authenticated the ProfileContent component above is rendered. Otherwise a message indicating a user is not authenticated is rendered.
 */
const MainContent = () => {
    return (
        <div className="App">
            <AuthenticatedTemplate>
                <ProfileContent />
            </AuthenticatedTemplate>

            <UnauthenticatedTemplate>
                <h5 className="card-title">Please sign-in to see your profile information.</h5>
            </UnauthenticatedTemplate>
        </div>
    );
};

export default function App() {

    useEffect(() => {
        console.log("loading")
    },[]);

    return (
        <PageLayout>
            <MainContent />
        </PageLayout>
    );
}
