using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using System.Text;
using System.Xml;

namespace MsCrmTools.AuditCenter
{
    /// <summary>
    /// Class for querying Crm Metadata
    /// </summary>
    internal class MetadataHelper
    {
        public static EntityMetadataCollection LoadEntities(IOrganizationService service)
        {
            var entityQueryExpression = new EntityQueryExpression
            {
                Properties = new MetadataPropertiesExpression("LogicalName", "DisplayName", "Attributes", "IsAuditEnabled", "ObjectTypeCode"),
                AttributeQuery = new AttributeQueryExpression
                {
                    Criteria = new MetadataFilterExpression
                    {
                        Conditions =
                        {
                            new MetadataConditionExpression("LogicalName", MetadataConditionOperator.NotIn, new[]{"traversedpath","versionnumber"})
                        }
                    },
                    Properties = new MetadataPropertiesExpression("DisplayName", "LogicalName", "AttributeType", "IsAuditEnabled", "AttributeOf", "EntityLogicalName"),
                }
            };
            var retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest
            {
                Query = entityQueryExpression,
                ClientVersionStamp = null
            };

            return ((RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest)).EntityMetadata;
        }

        public static EntityMetadataCollection LoadEntitiesBelowv9(IOrganizationService service)
        {
            var entityQueryExpression = new EntityQueryExpression
            {
                Properties = new MetadataPropertiesExpression("LogicalName", "DisplayName", "Attributes", "IsAuditEnabled", "ObjectTypeCode"),
                AttributeQuery = new AttributeQueryExpression
                {
                    Criteria = new MetadataFilterExpression
                    {
                        Conditions =
                        {
                            new MetadataConditionExpression("LogicalName", MetadataConditionOperator.NotIn, new[]{"traversedpath","versionnumber"})
                        }
                    }
                }
            };
            var retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest
            {
                Query = entityQueryExpression,
                ClientVersionStamp = null
            };

            return ((RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest)).EntityMetadata;
        }

        /// <summary>
        /// Retrieves main forms for the specified entity
        /// </summary>
        /// <param name="logicalName">Entity logical name</param>
        /// <param name="oService">Crm organization service</param>
        /// <returns>Document containing all forms definition</returns>
        public static XmlDocument RetrieveEntityForms(string logicalName, IOrganizationService oService)
        {
            QueryByAttribute qba = new QueryByAttribute("systemform");
            qba.Attributes.AddRange("objecttypecode", "type");
            qba.Values.AddRange(logicalName, 2);
            qba.ColumnSet = new ColumnSet(true);

            EntityCollection ec = oService.RetrieveMultiple(qba);

            StringBuilder allFormsXml = new StringBuilder();
            allFormsXml.Append("<root>");

            foreach (Entity form in ec.Entities)
            {
                allFormsXml.Append(form["formxml"]);
            }

            allFormsXml.Append("</root>");

            XmlDocument docAllForms = new XmlDocument();
            docAllForms.LoadXml(allFormsXml.ToString());

            return docAllForms;
        }
    }
}