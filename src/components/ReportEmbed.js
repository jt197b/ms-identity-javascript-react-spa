import React, {useEffect, useState } from "react";
import { service, factories, models } from 'powerbi-client';
import { scopeBase, reportId, powerBiApiUrl } from '../authConfig';
import { useMsal } from '@azure/msal-react';
import { styles } from "../styles/pbi.css";

const powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);

export default function ReportEmbed({accessToken, embedUrl}) {

    let reportRef;
    const { instance, accounts } = useMsal();
    const [isRendered, setIsRendered] = useState(true);

    const loginRequest = {
        scopes: scopeBase,
        account: accounts[0]
    };

    // silently acquires an access token on expiry
    const getNewAccessToken = () => {
        instance
            .acquireTokenSilent(loginRequest)
            .then((resp) => {
                console.log(resp.accessToken);
                return resp.accessToken;
            })
    }
    
    function resizeIFrameToFitContent(iFrame) {
        var reportContainer = document.getElementById('reportContainer');
        iFrame.width = reportContainer.clientWidth;
        iFrame.height = reportContainer.clientHeight;
    }

    const [embedConfig, setEmbedConfig] = useState({
		type: 'report',   // supported types: report, dashboard, tile, visual and qna
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
            layoutType: models.LayoutType.Custom,
            customLayout: {
               displayOption: models.DisplayOption.FitToWidth
            }, 
			background: models.BackgroundType.Transparent,
		}
	});

    reportRef = React.createRef();

    const loading = (
        <div>
            {isRendered && <p>Loading ...</p>}
            <div
                id="reportContainer"
                ref={reportRef} >
                Loading ...
            </div>
        </div>
    );

    const renderReport = () => {

        let reportContainer = document.getElementById("reportContainer");
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
            setIsRendered(false);
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

    const bootstrap = () => {
        console.log("Bootstrapping... ")
        let reportContainer = document.getElementById("reportContainer");
        powerbi.bootstrap(
            reportContainer,
            {
                type: 'report',
            }
        )
    }

    useEffect(() => {

        let reportContainer = document.getElementById("reportContainer");
        renderReport();

        if (reportContainer.textContent === "Loading ...") {
            bootstrap();
        } else { 
            renderReport();
        }

        // resize PBI iframe
        var iframes = document.querySelectorAll("iframe");

        try {
            for (var i = 0; i < iframes.length; i++) {
                resizeIFrameToFitContent(iframes[i]);
                iframes[i].attributes.removeNamedItem("style");
            }
        } catch(err) {
            console.log(err);
        }
    }, []);

    return (loading)
}
