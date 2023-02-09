import React, {useEffect, useState } from "react";
import { service, factories, models, IEmbedConfiguration } from 'powerbi-client';
import { scopeBase, workspaceId, reportId, powerBiApiUrl, datasetId } from '../authConfig';
import { PowerBIEmbed } from 'powerbi-client-react';
import { useMsal } from '@azure/msal-react';
import { styles } from "../styles/pbi.css";
import { render } from "react-dom";

const powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);

export default function ReportEmbed({accessToken, embedUrl}) {

    let reportRef;

    const { instance, accounts } = useMsal();

    const loginRequest = {
        scopes: scopeBase,
        account: accounts[0]
    };

    function resizeIFrameToFitContent(iFrame) {

        var reportContainer = document.getElementById('reportContainer');
        console.log(window.innerWidth);
        console.log(reportContainer.clientWidth);

        iFrame.width = window.innerWidth;
        iFrame.height = reportContainer.clientHeight;
    }

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
        pageView : "fitToWidth",
		settings: {
			panes: {
				filters: {
					expanded: false,
					visible: false
				}
			},
            layoutType: models.LayoutType.Custom,
            customLayout: {
               displayOption: models.DisplayOption.FitToWidth
            }, 
			background: models.BackgroundType.Transparent,
		}
	});

    reportRef = React.createRef();

    const loading = (
        <div
            id="reportContainer"
            ref={reportRef} >
            Loading the report...
        </div>
    );


    useEffect(() => {
        console.log("Rendering ...");
        renderReport();
        // resize PBI iframe
        var iframes = document.querySelectorAll("iframe");


        try {
            for (var i = 0; i < iframes.length; i++) {
                console.log(iframes[i]);
                resizeIFrameToFitContent(iframes[i]);
                iframes[i].attributes.removeNamedItem("style");
            }
        } catch(err) {
            console.log(err);
        }

    }, []);

    const renderReport = () => {

        let reportContainer = document.getElementById("reportContainer");
        console.log(reportContainer);
        const report = powerbi.embed(reportContainer, embedConfig);

        // Clear any other loaded handler events
        report.off("loaded");

        // Triggers when a content schema is successfully loaded
        report.on("loaded", function () {
            console.log("Report load successful");
        });

        // Clear any other rendered handler events
        report.off("rendered");

        // Triggers when a content is successfully embedded in UI
        report.on("rendered", function () {
            console.log("Report render successful");
        });

        report.off("dataSelected");

        // Triggers when a content is successfully embedded in UI
        report.on("dataSelected", function (e) {
            console.log(e);

            let metricSelected = e.detail.dataPoints[0].values[1].value;
            console.log(metricSelected);
        });
        

        // Clear any other error handler event
        report.off("error");

        // Below patch of code is for handling errors that occur during embedding
        report.on("error", function (event) {
            const errorMsg = event.detail;

            // Use errorMsg variable to log error in any destination of choice
            console.error(errorMsg);
        });

        return loading;
    }

    return (
        loading
    )
}

/*
<PowerBIEmbed
                    embedConfig = {embedConfig}
                
                    eventHandlers = { 
                        new Map([
                            ['loaded', function () {console.log('Report loaded');}],
                            ['rendered', function () {console.log('Report rendered');}],
                            ['error', function (event) {console.log(event.detail);}],
                            ['dataSelected', function (event) {
                                let data = event.detail;
                                console.log("Event - dataSelected:\n", data);
                            }]
                        ])
                    }
                        
                    cssClassName = { "report-style-class" }
                
                    getEmbeddedComponent = { (embeddedReport) => {
                        window.report = embeddedReport;
                    }}
                />
*/
