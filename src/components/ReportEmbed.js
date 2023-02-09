import React, {useEffect, useState } from "react";
import { service, factories, models, IEmbedConfiguration } from 'powerbi-client';
import { scopeBase, workspaceId, reportId, powerBiApiUrl, datasetId } from '../authConfig';
import { PowerBIEmbed } from 'powerbi-client-react';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';

const powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);

export default function ReportEmbed({accessToken, embedUrl}) {

    let reportContainer;
    let reportRef;

    const { instance, accounts } = useMsal();
    const [loading, setLoading] = useState(true);

    const loginRequest = {
        scopes: scopeBase,
        account: accounts[0]
    };

    const getNewAccessToken = () => {
        // Silently acquires an access token which is then attached to a request for MS Graph data
        instance
            .acquireTokenSilent(loginRequest)
            .then((resp) => {
                console.log(resp.accessToken);
                return resp.accessToken;
            })
    }
    
    const [embedConfig, setEmbedConfig] = useState({
		type: 'report',   // Supported types: report, dashboard, tile, visual and qna
		id: reportId,
		embedUrl: embedUrl,
		accessToken: accessToken,
		tokenType: models.TokenType.Aad,
        eventHooks: {
            accessTokenProvider: getNewAccessToken
        },
		settings: {
			panes: {
				filters: {
					expanded: false,
					visible: false
				}
			},
			background: models.BackgroundType.Transparent,
		}
	});

    reportRef = React.createRef();


    useEffect(() => {
        console.log("Rendering ...");

        if (reportRef !== null) {
            reportContainer = reportRef['current'];
        }
    }, []);

    return (
        <div>
            {!loading ?
                <PowerBIEmbed
                embedConfig = {embedConfig}
            
                eventHandlers = { 
                    new Map([
                        ['loaded', function () {console.log('Report loaded');}],
                        ['rendered', function () {console.log('Report rendered');}],
                        ['error', function (event) {console.log(event.detail);}]
                    ])
                }
                    
                cssClassName = { "report-style-class" }
            
                getEmbeddedComponent = { (embeddedReport) => {
                    this.report = embeddedReport;
                }}
            />
                :
                <PowerBIEmbed
                embedConfig = {embedConfig}
            
                eventHandlers = { 
                    new Map([
                        ['loaded', function () {console.log('Report loaded');}],
                        ['rendered', function () {console.log('Report rendered');}],
                        ['error', function (event) {console.log(event.detail);}]
                    ])
                }
                    
                cssClassName = { "report-style-class" }
            
                getEmbeddedComponent = { (embeddedReport) => {
                    window.report = embeddedReport;
                }}
            />
                
            }
        </div>
    )
}
