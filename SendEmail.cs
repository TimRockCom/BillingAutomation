using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Rimrock.Fieldboss.Common;
using System;
using System.Activities;

namespace Rimrock.Fieldboss.Workflows
{
    public class SendEmail : WorkFlowActivityBase
    {
        [RequiredArgument]
        [Input("Email to send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> SourceEmail
        { get; set; }

        [Output("Email Subject")]
        public OutArgument<string> Subject { get; set; }

        public override void ExecuteCRMWorkFlowActivity(CodeActivityContext ec, LocalWorkflowContext crmWorkflowContext)
        {
            if (crmWorkflowContext == null)
                throw new ArgumentNullException("crmWorkflowContext");

            ContextWrapper contextWrapper = new ContextWrapper
            (
                crmWorkflowContext.TracingService,
                crmWorkflowContext.OrganizationService
            );

            EntityReference email = SourceEmail.Get(ec);
            SendEmailResponse ser = contextWrapper.service.Execute(
                new SendEmailRequest()
                {
                    EmailId = email.Id,
                    IssueSend = true
                }
              ) as SendEmailResponse;

            if (ser != null)
            {
                Subject.Set(ec, ser.Subject);
            }
        }
    }
}
