using ReportExecutionService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel;
using System.Threading.Tasks;

namespace SSRSREPORTS
{
    public static class SSRSDownloader
    {
        const string SSRSUsername = "sa";
        const string SSRSDomain = "";
        const string SSRSPassword = "Temp123";
        const string SSRSReportExecutionUrl = "http://desktop-ffaakts/ReportServer/ReportExecution2005.asmx";
        const string SSRSFolderPath = "DemoReports";
        const string ReportName = "Users";

        public async static Task<byte[]> GenerateSSRSReport(string County)
        {
            var binding = new BasicHttpBinding(BasicHttpSecurityMode.None);
            //binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Windows;
            //binding.MaxReceivedMessageSize = 10485760; //10MB limit

            //Create the execution service SOAP Client
            var rsExec = new ReportExecutionServiceSoapClient(binding, new EndpointAddress(SSRSReportExecutionUrl));

            //Setup access credentials.
            var clientCredentials = new NetworkCredential(SSRSUsername, SSRSPassword, SSRSDomain);
            if (rsExec.ClientCredentials != null)
            {
                rsExec.ClientCredentials.Windows.AllowedImpersonationLevel =
                    System.Security.Principal.TokenImpersonationLevel.Impersonation;
                rsExec.ClientCredentials.Windows.ClientCredential = clientCredentials;
            }

            //This handles the problem of "Missing session identifier"
            rsExec.Endpoint.EndpointBehaviors.Add(new ReportingServicesEndpointBehavior());

            await rsExec.LoadReportAsync(null, "/" + SSRSFolderPath + "/" + ReportName, null);

            //TODO: determine parameters
            //Set the parameters asked for by the report
            ReportExecutionService.ParameterValue[] reportParam = new ReportExecutionService.ParameterValue[1];

            reportParam[0] = new ReportExecutionService.ParameterValue();
            reportParam[0].Name = "County";
            reportParam[0].Value = County.ToString();

            await rsExec.SetExecutionParametersAsync(null, null, reportParam, "en-us");

            //run the report
            const string deviceInfo = @"<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";
            var response = await rsExec.RenderAsync(new RenderRequest(null, null, "PDF", deviceInfo));

            //spit out the result
            var byteResults = response.Result;

            return byteResults;
        }
    }
}
