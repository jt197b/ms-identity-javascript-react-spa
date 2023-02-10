import React, { useState, useRef } from 'react';
import './styles/App.css';
import { PageLayout } from './components/PageLayout';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import Button from 'react-bootstrap/Button';
import { scopeBase, workspaceId, reportId, powerBiApiUrl, datasetId } from './authConfig';
import { useIsAuthenticated } from "@azure/msal-react";
import { useEffect } from 'react';
import ReportEmbed from './components/ReportEmbed';

const PbiContent = () => {

    const { instance, accounts } = useMsal();

    // msal
    const [accessToken, setAccessToken] = useState(localStorage.getItem("msalToken") ? localStorage.getItem("msalToken") : null);
    const [authFailed, setAuthFailed] = useState(false);

    // pbi
    const [embedUrl, setEmbedUrl] = useState("https://app.powerbi.com/reportEmbed?reportId=10cb3f4c-3c9c-4d45-a18f-b40a21da4bc0&groupId=a81e25a2-2ee6-40bb-a4ac-9a2cb4538f01&w=2");

    const loginRequest = {
        scopes: scopeBase,
        account: accounts[0]
    };

    function authenticate() {

        // silently acquires an access token
        instance
            .acquireTokenSilent(loginRequest)
            .then((resp) => {
                localStorage.setItem("msalToken", resp.accessToken);
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
                return resp.json();
            }).then(function(data) {
                console.log(data);
            });
    }

    useEffect(() => {
        try {
            authenticate();
            //console.log(embedUrl);
            //console.log(accessToken);
            
        } catch (err) {
            console.log(err);
            console.log("Auth failed. Rendering backup embed... ");
            setAuthFailed(true);
        }
    }, [accessToken, authFailed]);

    return (
        <>
            <h5 className="card-title">Welcome {accounts[0].name}</h5>
            {
                (accessToken && !authFailed) ?
                <div>
                    <ReportEmbed
                        accessToken={accessToken}
                        embedUrl={embedUrl}
                    />
                </div>
                :
                <div>
                    <Button variant="secondary" onClick={authenticate}>
                        Retry Connection
                    </Button>
                    <p> Report loaded successfully but currently cannot connect to MyDashboard. </p>
                    <iframe
                        title="CE Executive Dashboard" width="1140" height="700"
                        src="https://app.powerbi.com/reportEmbed?reportId=10cb3f4c-3c9c-4d45-a18f-b40a21da4bc0&autoAuth=true&ctid=e741d71c-c6b6-47b0-803c-0f3b32b07556" frameborder="0" allowFullScreen="true">
                    </iframe>
                </div>
            }
        </>
    );
};


//if a user is authenticated the pbiContent component above is rendered. Otherwise a message indicating a user is not authenticated is rendered.
export default function App() {

    const { instance, inProgress  } = useMsal();
    const isAuthenticated = useIsAuthenticated();

    const handleLogin = () => {
        instance.loginRedirect(scopeBase).catch(e => {
            console.log(e);
        });
    }

    useEffect(() => {

        console.log(isAuthenticated);

        if (!isAuthenticated) {
            console.log("Logging you in");
            handleLogin();
        }
        
    }, [isAuthenticated]);

    return (
        <PageLayout>

            <div className="App">
                <AuthenticatedTemplate>
                    <PbiContent />
                </AuthenticatedTemplate>

                <UnauthenticatedTemplate>
                    <h5 className="card-title">Signing you in... </h5>
                    {/*
                    <iframe
                        title="CE Executive Dashboard" width="1140" height="700"
                        src="https://app.powerbi.com/reportEmbed?reportId=10cb3f4c-3c9c-4d45-a18f-b40a21da4bc0&autoAuth=true&ctid=e741d71c-c6b6-47b0-803c-0f3b32b07556" frameborder="0" allowFullScreen="true">
                    </iframe>
                    */}
                </UnauthenticatedTemplate>
            </div>

        </PageLayout>
    );
}
