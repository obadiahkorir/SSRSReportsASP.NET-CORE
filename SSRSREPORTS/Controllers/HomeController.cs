using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using ReportExecutionService;
using SSRSREPORTS.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.ServiceModel.Dispatcher;
using System.Threading.Tasks;

namespace SSRSREPORTS.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {

            
            return View();
        }
        public async Task GetSSRSFileBytes()
        {
            var ssrsFileBytes = await SSRSDownloader.GenerateSSRSReport("Male");
        }
        public IActionResult Privacy()
        {
            return View();
        }
        [HttpGet("GetPDFReport")]
        public async Task<IActionResult> GetPDFReport( CountiesModel countiesModel)
        {
            string CountyName = countiesModel.County;
            string reportName = "DemoUsers2";
            IDictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("County", CountyName);
            string languageCode = "en-us";

            byte[] reportContent = await this.RenderReport(reportName, parameters, languageCode, "PDF");

            Stream stream = new MemoryStream(reportContent);

            return new FileStreamResult(stream, "application/pdf");

        }
        [HttpGet("GetExcelReport")]
        public async Task<IActionResult> GetExcelReport(CountiesModel countiesModel)
        {
            string CountyName = countiesModel.County;
            string reportName = "DemoUsers2";
            IDictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("County", CountyName);
            string languageCode = "en-us";

            byte[] reportContent = await this.RenderReport(reportName, parameters, languageCode, "EXCEL");

            Stream stream = new MemoryStream(reportContent);

            return new FileStreamResult(stream, "application/vnd.ms-excel");

        }
        [HttpGet("GetWordReport")]
        public async Task<IActionResult> GetWordReport()
        {
            string reportName = "DemoUsers2";
            IDictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("County", "Nairobi12");
            string languageCode = "en-us";

            byte[] reportContent = await this.RenderReport(reportName, parameters, languageCode, "WORD");

            Stream stream = new MemoryStream(reportContent);

            return new FileStreamResult(stream, "application/msword");

        }
        [HttpGet("GetPowerPointReport")]
        public async Task<IActionResult> GetPowerPointReport()
        {
            string reportName = "DemoUsers3";
            IDictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("County", "Nairobi12");
            string languageCode = "en-us";

            byte[] reportContent = await this.RenderReport(reportName, parameters, languageCode, "PPTX");

            Stream stream = new MemoryStream(reportContent);

            return new FileStreamResult(stream, "application/vnd.ms-powerpoint");

        }
        private async Task<byte[]> RenderReport(string reportName, IDictionary<string, object> parameters, string languageCode, string exportFormat)
        {
            const string SSRSUsername = "Obadiah Korir";
            const string SSRSDomain = "DESKTOP-FFAAKTS";
            const string SSRSPassword = "TreeFresh5$";
            const string SSRSReportExecutionUrl = "http://desktop-ffaakts/ReportServer/ReportExecution2005.asmx";
            const string SSRSFolderPath = "DemoReports";
            string reportPath = string.Format("{0}{1}", SSRSFolderPath, reportName);

            //
            // Binding setup, since ASP.NET Core apps don't use a web.config file
            var binding = new BasicHttpBinding(BasicHttpSecurityMode.None);
            //binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Ntlm;
            binding.MaxReceivedMessageSize = 999999999; //10MB limit
            binding.MaxBufferSize = 999999999;
            binding.OpenTimeout = new TimeSpan(0, 90, 0);
            binding.CloseTimeout = new TimeSpan(0, 90, 0);
            binding.SendTimeout = new TimeSpan(0, 90, 0);
            binding.ReceiveTimeout = new TimeSpan(0, 90, 0);

            //Create the execution service SOAP Client
            ReportExecutionServiceSoapClient reportClient = new ReportExecutionServiceSoapClient(binding, new EndpointAddress(SSRSReportExecutionUrl));

            //Setup access credentials. Here use windows credentials.
            var clientCredentials = new NetworkCredential(SSRSUsername, SSRSPassword, SSRSDomain);
            reportClient.ClientCredentials.Windows.AllowedImpersonationLevel = System.Security.Principal.TokenImpersonationLevel.Impersonation;
            reportClient.ClientCredentials.Windows.ClientCredential = clientCredentials;

            //This handles the problem of "Missing session identifier"
            reportClient.Endpoint.EndpointBehaviors.Add(new ReportingServiceEndPointBehavior());

            //string historyID = null;
            //TrustedUserHeader trustedUserHeader = new TrustedUserHeader();
            //ExecutionHeader execHeader = new ExecutionHeader();

            //trustedUserHeader.UserName = clientCredentials.UserName;

            //
            // Load the report
            //
            var taskLoadReport = await reportClient.LoadReportAsync(null, "/" + SSRSFolderPath + "/" + reportName, null);
            // Fixed the exception of "session identifier is missing".
            reportClient.Endpoint.EndpointBehaviors.Add(new ReportingServicesEndpointBehavior());

            //
            //Set the parameteres asked for by the report
            //
            ParameterValue[] reportParameters = null;
            if (parameters != null && parameters.Count > 0)
            {
                reportParameters = taskLoadReport.executionInfo.Parameters.Where(x => parameters.ContainsKey(x.Name)).Select(x => new ParameterValue() { Name = x.Name, Value = parameters[x.Name].ToString() }).ToArray();
            }

            await reportClient.SetExecutionParametersAsync(null, null, reportParameters, languageCode);
            // run the report
            const string deviceInfo = @"<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";

            var response = await reportClient.RenderAsync(new RenderRequest(null, null, exportFormat ?? exportFormat, deviceInfo));

            //spit out the result
            return response.Result;
        }
        public class ReportingServiceEndPointBehavior : IEndpointBehavior
        {
            public void AddBindingParameters(ServiceEndpoint endpoint, BindingParameterCollection bindingParameters) { }

            public void ApplyClientBehavior(ServiceEndpoint endpoint, ClientRuntime clientRuntime)
            {
                clientRuntime.ClientMessageInspectors.Add(new ReportingServiceExecutionInspector());
            }

            public void ApplyDispatchBehavior(ServiceEndpoint endpoint, EndpointDispatcher endpointDispatcher) { }

            public void Validate(ServiceEndpoint endpoint) { }
        }

        public class ReportingServiceExecutionInspector : IClientMessageInspector
        {
            private MessageHeaders headers;

            public void AfterReceiveReply(ref Message reply, object correlationState)
            {
                var index = reply.Headers.FindHeader("ExecutionHeader", "http://schemas.microsoft.com/sqlserver/2005/06/30/reporting/reportingservices");
                if (index >= 0 && headers == null)
                {
                    headers = new MessageHeaders(MessageVersion.Soap11);
                    headers.CopyHeaderFrom(reply, reply.Headers.FindHeader("ExecutionHeader", "http://schemas.microsoft.com/sqlserver/2005/06/30/reporting/reportingservices"));
                }
            }

            public object BeforeSendRequest(ref Message request, IClientChannel channel)
            {
                if (headers != null)
                    request.Headers.CopyHeadersFrom(headers);

                return Guid.NewGuid(); //https://msdn.microsoft.com/en-us/library/system.servicemodel.dispatcher.iclientmessageinspector.beforesendrequest(v=vs.110).aspx#Anchor_0
            }
        }
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
