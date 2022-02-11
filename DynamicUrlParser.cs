using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Rimrock.Fieldboss.Common;
using System;
using System.Activities;
using System.ServiceModel;

namespace Rimrock.Fieldboss.Workflows
{
    /// <summary>
    /// Used to parse the Dynamics CRM 'Record Url (Dynamic)' that can be created by workflows and dialogs
    /// </summary>
    public class DynamicUrlParser : WorkFlowActivityBase
    {
        public string Url { get; set; }
        public int EntityTypeCode { get; set; }
        public Guid EntityId { get; set; }
        public string EntityName { get; set; }


        [Input("Record Dynamic Url")]
        [RequiredArgument]
        public InArgument<string> RecordUrl { get; set; }

        [Output("Record Guid")]
        public OutArgument<string> RecordGuid { get; set; }

        [Output("Record Entity Logical Name")]
        public OutArgument<string> EntityLogicalName { get; set; }

        [Output("Invoice Entity Reference")]
        [ReferenceTarget("invoice")]
        [ArgumentDescription("If dynamic url points to an invoice, this will contain data")]
        public OutArgument<EntityReference> Invoice { get; set; }

        [Output("Email Entity Reference")]
        [ReferenceTarget("email")]
        [ArgumentDescription("If dynamic url points to an email, this will contain data")]
        public OutArgument<EntityReference> Email { get; set; }

        [Output("Service Activity Reference")]
        [ReferenceTarget("serviceappointment")]
        [ArgumentDescription("If dynamic url points to a service activity, this will contain data")]
        public OutArgument<EntityReference> ServiceActivity { get; set; }

        [Output("Quote Entity Reference")]
        [ReferenceTarget("quote")]
        [ArgumentDescription("If dynamic url points to a quote, this will contain data")]
        public OutArgument<EntityReference> Quote { get; set; }

        [Output("Account Entity Reference")]
        [ReferenceTarget("account")]
        [ArgumentDescription("If dynamic url points to an account, this will contain data")]
        public OutArgument<EntityReference> Account { get; set; }

        [Output("Building Location Entity Reference")]
        [ReferenceTarget("fsip_buildinglocation")]
        [ArgumentDescription("If dynamic url points to a building location, this will contain data")]
        public OutArgument<EntityReference> BuildingLocaton { get; set; }

        [Output("Work Order Entity Reference")]
        [ReferenceTarget("salesorder")]
        [ArgumentDescription("If dynamic url points to a work order, this will contain data")]
        public OutArgument<EntityReference> WorkOrder { get; set; }

        [Output("Maintenance Contract Entity Reference")]
        [ReferenceTarget("fsip_maintenancecontract")]
        [ArgumentDescription("If dynamic url points to a maintenance contract, this will contain data")]
        public OutArgument<EntityReference> MaintenanceContract { get; set; }

        [Output("Project Entity Reference")]
        [ReferenceTarget("new_project")]
        [ArgumentDescription("If dynamic url points to a project, this will contain data")]
        public OutArgument<EntityReference> Project { get; set; }

        [Output("Device Entity Reference")]
        [ReferenceTarget("new_servloc")]
        [ArgumentDescription("If dynamic url points to a device, this will contain data")]
        public OutArgument<EntityReference> Device { get; set; }


        public override void ExecuteCRMWorkFlowActivity(CodeActivityContext executionContext, LocalWorkflowContext crmWorkflowContext)
        {
            if (crmWorkflowContext == null)
                throw new ArgumentNullException("crmWorkflowContext");

            ContextWrapper contextWrapper = new ContextWrapper
            (
                crmWorkflowContext.TracingService,
                crmWorkflowContext.OrganizationService
            );
            try
            {
                InternalDynamicUrlParser(RecordUrl.Get<string>(executionContext));
                RecordGuid.Set(executionContext, EntityId.ToString());
                EntityName = GetEntityLogicalName(contextWrapper.service);
                EntityLogicalName.Set(executionContext, EntityName);
                switch (EntityName.ToLower())
                {
                    case "quote": Quote.Set(executionContext, new EntityReference(EntityName, EntityId)); break;
                    case "account": Account.Set(executionContext, new EntityReference(EntityName, EntityId)); break;
                    case "fsip_buildinglocation": BuildingLocaton.Set(executionContext, new EntityReference(EntityName, EntityId)); break;
                    case "salesorder": WorkOrder.Set(executionContext, new EntityReference(EntityName, EntityId)); break;
                    case "fsip_maintenancecontract": MaintenanceContract.Set(executionContext, new EntityReference(EntityName, EntityId)); break;
                    case "new_project": Project.Set(executionContext, new EntityReference(EntityName, EntityId)); break;
                    case "new_servloc": MaintenanceContract.Set(executionContext, new EntityReference(EntityName, EntityId)); break;
                    case "serviceappointment": ServiceActivity.Set(executionContext, new EntityReference(EntityName, EntityId)); break;
                    case "invoice": Invoice.Set(executionContext, new EntityReference(EntityName, EntityId)); break;
                    case "email":
                        Email.Set(executionContext, new EntityReference(EntityName, EntityId));
                        var regInvoice = getEmailProperties(contextWrapper);
                        Invoice.Set(executionContext, regInvoice);
                        break;
                }
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                contextWrapper.trace.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }
        }

        private EntityReference getEmailProperties(ContextWrapper contextWrapper)
        {
            ColumnSet emailColumns = new ColumnSet("regardingobjectid");
            Entity email = contextWrapper.service.Retrieve(EntityName, EntityId, emailColumns);
            EntityReference invRef = RRUtils.getEntityReference(email, "regardingobjectid");
            if (invRef == null || invRef.LogicalName != "invoice")
            {
                String msg = String.Format("Email's regarding field is not set to invoice.");
                throw new InvalidPluginExecutionException(msg);
            }

            return invRef;
        }

        /// <summary>
        /// Parse the dynamic url in constructor
        /// </summary>
        /// <param name="url"></param>
        private void InternalDynamicUrlParser(string url)
        { 
            try
            {
                Url = url;
                var uri = new Uri(url);
                int found = 0;

                string[] parameters = uri.Query.TrimStart('?').Split('&');
                foreach (string param in parameters)
                {
                    var nameValue = param.Split('=');
                    switch (nameValue[0])
                    {
                        case "etc":
                            EntityTypeCode = int.Parse(nameValue[1]);
                            found++;
                            break;
                        case "id":
                            EntityId = new Guid(nameValue[1]);
                            found++;
                            break;
                    }
                    if (found > 1) break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(String.Format("Url '{0}' is incorrectly formated for a Dynamics CRM Dynamics Url", url), ex);
            }
        }

        /// <summary>
        /// Find the Logical Name from the entity type code - this needs a reference to the Organization Service to look up metadata
        /// </summary>
        /// <param name="service"></param>
        /// <returns></returns>
        public string GetEntityLogicalName(IOrganizationService service)
        {
            var entityFilter = new MetadataFilterExpression(LogicalOperator.And);
            entityFilter.Conditions.Add(new MetadataConditionExpression("ObjectTypeCode ", MetadataConditionOperator.Equals, this.EntityTypeCode));
            var propertyExpression = new MetadataPropertiesExpression { AllProperties = false };
            propertyExpression.PropertyNames.Add("LogicalName");
            var entityQueryExpression = new EntityQueryExpression()
            {
                Criteria = entityFilter,
                Properties = propertyExpression
            };

            var retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest()
            {
                Query = entityQueryExpression
            };

            var response = (RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest);

            if (response.EntityMetadata.Count == 1)
            {
                return response.EntityMetadata[0].LogicalName;
            }
            return null;
        }
    }
}