using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.ServiceModel.Dispatcher;
using System.Threading.Tasks;

namespace SSRSREPORTS
{
    internal class ReportingServicesEndpointBehavior : IEndpointBehavior
    {
        public void AddBindingParameters(ServiceEndpoint endpoint, BindingParameterCollection bindingParameters) { }

        public void ApplyClientBehavior(ServiceEndpoint endpoint, ClientRuntime clientRuntime)
        {
            clientRuntime.ClientMessageInspectors.Add(new ReportingServicesExecutionInspector());
        }

        public void ApplyDispatchBehavior(ServiceEndpoint endpoint, EndpointDispatcher endpointDispatcher) { }

        public void Validate(ServiceEndpoint endpoint) { }
    }

    internal class ReportingServicesExecutionInspector : IClientMessageInspector
    {
        private MessageHeaders headers;

        public void AfterReceiveReply(ref Message reply, object correlationState)
        {
            var index = reply.Headers.FindHeader("ExecutionHeader", "http://schemas.microsoft.com/sqlserver/2005/06/30/reporting/reportingservices");
            if (index >= 0 & headers == null)
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
}
