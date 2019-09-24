using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ImproveGroup
{
    public class ReviseBidSheet : IPlugin
    {
        IPluginExecutionContext context;
        ITracingService tracing;
        IOrganizationServiceFactory servicefactory;
        IOrganizationService service;
        public void Execute(IServiceProvider serviceProvider)
        {
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = servicefactory.CreateOrganizationService(context.UserId);
            try
            {

                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    EntityReference entity = (EntityReference)context.InputParameters["Target"];
                    if (entity.LogicalName != "ig1_bidsheet")
                    {
                        return;
                    }

                    var bidSheetId = string.Empty;
                    bidSheetId = entity.Id.ToString().Replace("{", "").Replace("}", "");

                    Entity existingBidSheet = service.Retrieve("ig1_bidsheet", new Guid(bidSheetId), new ColumnSet(true));
                    Entity bidSheet = new Entity("ig1_bidsheet");
                    foreach (KeyValuePair<string, object> attr in existingBidSheet.Attributes)
                    {
                        if (attr.Key == "statecode" || attr.Key == "statuscode" || attr.Key == "ig1_bidsheetid")
                            continue;
                        if (attr.Key == "ig1_status")
                        {
                            bidSheet[attr.Key] = new OptionSetValue(Convert.ToInt32(286150002));
                        }
                        else
                        {
                            bidSheet[attr.Key] = attr.Value;
                        }
                    }
                    Guid newBidSheetId = service.Create(bidSheet);

                    //Nazish - 10-07-2019 - Cloning the child records from exiating Bid Sheet.
                    CloneBidSheetCategoryVendors(newBidSheetId, existingBidSheet);
                    CloneBidSheetProducts(newBidSheetId, existingBidSheet);
                    CloneBidSheetLineItems(newBidSheetId, existingBidSheet);
                    CloneAssociatedCost(newBidSheetId, existingBidSheet);

                    //Nazish - 10-07-2019 - Updating the Revision Id
                    UpdateRevisionId(existingBidSheet);

                }
            }
            catch (Exception ex)
            {
                Entity errorLog = new Entity("ig1_pluginserrorlogs");
                errorLog["ig1_name"] = "Error";
                errorLog["ig1_errormessage"] = ex.Message;
                errorLog["ig1_errordescription"] = ex.InnerException;
                service.Create(errorLog);
                throw;
            }
        }
        protected void CloneBidSheetCategoryVendors(Guid newBidSheetId, Entity existingBidSheet)
        {
            try
            {
                QueryByAttribute query = new QueryByAttribute("ig1_bscategoryvendor");
                query.ColumnSet = new ColumnSet(true);
                query.Attributes.AddRange("ig1_bidsheet");
                query.Values.AddRange(existingBidSheet.Id);
                List<Entity> existingBidSheetCategoryVendor = service.RetrieveMultiple(query).Entities.ToList();

                foreach (Entity bidSheetCategoryVendor in existingBidSheetCategoryVendor)
                {
                    Entity newBidSheetCategoryVendor = new Entity("ig1_bscategoryvendor");
                    foreach (KeyValuePair<string, object> attr in bidSheetCategoryVendor.Attributes)
                    {
                        if (attr.Key == "ig1_bscategoryvendorid")
                            continue;

                        newBidSheetCategoryVendor[attr.Key] = attr.Value;
                    }

                    newBidSheetCategoryVendor["ig1_bidsheet"] = new EntityReference("ig1_bidsheet", newBidSheetId);

                    service.Create(newBidSheetCategoryVendor);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        protected void CloneBidSheetProducts(Guid newBidSheetId, Entity existingBidSheet)
        {
            try
            {
                QueryByAttribute query = new QueryByAttribute("ig1_bidsheetproduct");
                query.ColumnSet = new ColumnSet(true);
                query.Attributes.AddRange("ig1_bidsheet");
                query.Values.AddRange(existingBidSheet.Id);
                List<Entity> existingBidSheetProducts = service.RetrieveMultiple(query).Entities.ToList();

                foreach (Entity bidSheetProducts in existingBidSheetProducts)
                {
                    Entity newBidSheetProducts = new Entity("ig1_bidsheetproduct");
                    foreach (KeyValuePair<string, object> attr in bidSheetProducts.Attributes)
                    {
                        if (attr.Key == "ig1_bidsheetproductid")
                            continue;

                        newBidSheetProducts[attr.Key] = attr.Value;
                    }

                    newBidSheetProducts["ig1_bidsheet"] = new EntityReference("ig1_bidsheet", newBidSheetId);

                    service.Create(newBidSheetProducts);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        protected void CloneBidSheetLineItems(Guid newBidSheetId, Entity existingBidSheet)
        {
            try
            {
                QueryByAttribute query = new QueryByAttribute("ig1_bidsheetpricelistitem");
                query.ColumnSet = new ColumnSet(true);
                query.Attributes.AddRange("ig1_bidsheet");
                query.Values.AddRange(existingBidSheet.Id);
                List<Entity> existingBidSheetLineItems = service.RetrieveMultiple(query).Entities.ToList();

                foreach (Entity bidSheetLineItem in existingBidSheetLineItems)
                {
                    Entity newBidSheetLineItems = new Entity("ig1_bidsheetpricelistitem");
                    foreach (KeyValuePair<string, object> attr in bidSheetLineItem.Attributes)
                    {
                        if (attr.Key == "ig1_bidsheetpricelistitemid")
                            continue;

                        newBidSheetLineItems[attr.Key] = attr.Value;
                    }

                    newBidSheetLineItems["ig1_bidsheet"] = new EntityReference("ig1_bidsheet", newBidSheetId);

                    service.Create(newBidSheetLineItems);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        protected void CloneAssociatedCost(Guid newBidSheetId, Entity existingBidSheet)
        {
            try
            {
                QueryByAttribute query = new QueryByAttribute("ig1_associatedcost");
                query.ColumnSet = new ColumnSet(true);
                query.Attributes.AddRange("ig1_bidsheet");
                query.Values.AddRange(existingBidSheet.Id);
                List<Entity> existingAssociatedCost = service.RetrieveMultiple(query).Entities.ToList();

                foreach (Entity associatedCost in existingAssociatedCost)
                {
                    Entity newAssociatedCost = new Entity("ig1_associatedcost");
                    foreach (KeyValuePair<string, object> attr in associatedCost.Attributes)
                    {
                        if (attr.Key == "ig1_associatedcostid")
                            continue;

                        newAssociatedCost[attr.Key] = attr.Value;
                    }

                    newAssociatedCost["ig1_bidsheet"] = new EntityReference("ig1_bidsheet", newBidSheetId);

                    service.Create(newAssociatedCost);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        protected void UpdateRevisionId(Entity existingBidSheet)
        {
            Entity entity = service.Retrieve("ig1_bidsheet", existingBidSheet.Id, new ColumnSet("ig1_revisionid", "ig1_status"));
            try
            {
                var revisionID = entity.GetAttributeValue<int>("ig1_revisionid");
                var status = entity.GetAttributeValue<OptionSetValue>("ig1_status");
                if (!revisionID.Equals(null) && !revisionID.Equals(""))
                {
                    revisionID++;
                    entity["ig1_revisionid"] = revisionID;
                }
                else
                {
                    entity["ig1_revisionid"] = 0;
                }
                entity["ig1_status"]= new OptionSetValue(Convert.ToInt32(286150001));
                service.Update(entity);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}