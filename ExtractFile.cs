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
    public class ExtractFile : WorkFlowActivityBase
    {
        public string Url { get; set; }
        public int EntityTypeCode { get; set; }
        public Guid EntityId { get; set; }
        public string EntityName { get; set; }


        [Input("Created Document")]
        [ReferenceTarget("ptm_mscrmaddonstemp")]
        [RequiredArgument]
        public InArgument<EntityReference> AttachmentRecord { get; set; }

        [Input("Email Entity Reference")]
        [RequiredArgument]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> Email { get; set; }

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
                contextWrapper.trace.Trace("checking input parameters");
                var emailReference = Email.Get<EntityReference>(executionContext);
                var docReference = AttachmentRecord.Get<EntityReference>(executionContext);
                contextWrapper.trace.Trace("looking for attachment");
                var docAttach = GetFileFromNote(docReference, contextWrapper.service);
                contextWrapper.trace.Trace("attaching file to email");
                AttachFileToEmail(contextWrapper.service, docAttach, emailReference);
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                contextWrapper.trace.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }
        }

        private void AttachFileToEmail(IOrganizationService service, Entity docAttach, EntityReference emailReference)
        {
            Entity _Attachment = new Entity("activitymimeattachment");
            _Attachment["subject"] = docAttach["filename"];
            _Attachment["filename"] = docAttach["filename"];
            _Attachment["body"] = docAttach["documentbody"];
            _Attachment["mimetype"] = docAttach["mimetype"];
            _Attachment["attachmentnumber"] = 1;
            _Attachment["objectid"] = emailReference;
            _Attachment["objecttypecode"] = emailReference.LogicalName;
            service.Create(_Attachment);
        }

        /// <summary>
        /// takes a reference to a record and gets the first note with a file attached
        /// </summary>
        /// <param name="docReference"></param>
        /// <param name="service"></param>
        /// <returns>entity record for the file attachment, if found. null if not found, or more than one found.</returns>
        private Entity GetFileFromNote(EntityReference docReference, IOrganizationService service)
        {
            QueryExpression qe = new QueryExpression
            {
                EntityName = "annotation",
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression
                        {
                            AttributeName = "objectid", Operator = ConditionOperator.Equal, Values = { docReference.Id }
                        },
                        new ConditionExpression
                        {
                            AttributeName="isdocument", Operator = ConditionOperator.Equal, Values = { true }
                        }
                    }
                }
            };
            Entity note = null;
            EntityCollection notes = service.RetrieveMultiple(qe);
            if (notes.Entities.Count == 1)
                note = notes.Entities[0];

            return note;
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
