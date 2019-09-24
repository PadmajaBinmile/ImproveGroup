/*Nazish - 06-05-2019 -  This plugin is created to create the record within the "Bid Sheet Price List Item" entity and to create the opportunity line
 */
using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ImproveGroup
{
    public class AddProductGroupsToBillOfMaterials : IPlugin
    {
        ITracingService tracingService;
        IPluginExecutionContext context;
        IOrganizationServiceFactory serviceFactory;
        IOrganizationService service;

        public void Execute(IServiceProvider serviceProvider)
        {
            #region Setup
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            if (context.InputParameters.Equals(null))
            {
                return;
            }
            try
            {
                string bidSheetId = context.InputParameters["bidSheetId"].ToString();
                var opportunityId = context.InputParameters["opportunityId"].ToString();
                if (opportunityId.Equals("") || opportunityId.Equals(null) || bidSheetId.Equals("") || bidSheetId.Equals(null))
                {
                    return;
                }
                //fetching all the record from the bid sheet product entity based on product group and bid sheet value.
                var fetchData = new
                {
                    ig1_bidsheet = bidSheetId,
                    ig1_isgrouped = "1"
                };
                var fetchXml = $@"
                                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='ig1_bidsheetproduct'>
                                    <attribute name='ig1_productgroup' />
                                    <attribute name='ig1_totalamount' />
                                    <filter type='and'>
                                        <condition attribute='ig1_isgrouped' operator='eq' value='{fetchData.ig1_isgrouped}'/>
                                        <condition attribute='ig1_productgroup' operator='not-null' />
                                        <condition attribute='ig1_bidsheet' operator='eq' value='{fetchData.ig1_bidsheet}'/>
                                    </filter>
                                    <link-entity name='product' from='productid' to='ig1_productgroup'>
                                        <attribute name='ig1_bidsheetcategory' alias='category' />
                                        <attribute name='defaultuomid' alias='unitId' />
                                        <attribute name='ig1_freight' alias='freight' />
                                        <attribute name='ig1_projectlu' alias='projectlu' />
                                        <link-entity name='ig1_bidsheetcategory' from='ig1_bidsheetcategoryid' to='ig1_bidsheetcategory' >
                                            <attribute name='ig1_rate' alias='rate' />
                                            <attribute name='ig1_projecthours' alias='projectHour' />
                                            <attribute name='ig1_defaultmatcostmargin' alias='margin' />
                                            <attribute name='ig1_laborunit' alias='laborUnit' />
                                      </link-entity>
                                    </link-entity>
                                    </entity>
                                </fetch>";

                EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                foreach (var groupId in result.Entities)
                {
                    //Variable are declared to create the record within the bid sheet proce list item
                    var productGroup = (EntityReference)(groupId.Attributes["ig1_productgroup"]);
                    var id = (Guid)productGroup.Id;
                    var categoryId = (Guid)((EntityReference)((AliasedValue)groupId["category"]).Value).Id;
                    var unitId = (Guid)((EntityReference)((AliasedValue)groupId["unitId"]).Value).Id;
                    var freight=new Money(0);
                    if (groupId.Attributes.Contains("freight"))
                    {
                        freight = ((Money)((AliasedValue)(groupId["freight"])).Value);
                    }
                    var rate = new decimal(0);
                    if(groupId.Attributes.Contains("rate"))
                    {
                        rate = (decimal)((AliasedValue)groupId["rate"]).Value;
                    }
                    var projectHour = new decimal(0);
                    if (groupId.Attributes.Contains("projectHour"))
                    {
                        projectHour = (decimal)((AliasedValue)groupId["projectHour"]).Value;
                    }
                    var productLu = new decimal(0);
                    if (groupId.Attributes.Contains("laborUnit"))
                    {
                        productLu= (decimal)((AliasedValue)groupId["laborUnit"]).Value;
                    }
                    var margin = new decimal();
                    if (groupId.Attributes.Contains("margin"))
                    {
                        margin = (decimal)((AliasedValue)groupId["margin"]).Value;
                    }
                    //Calculating the project cost based on the value of prjectHour and rate from the category
                    var cost = projectHour * rate;
                    var ProjectCost = new Money(cost);
                    // Calculating the freight total amount.
                    var freightTotal = new Money(freight.Value);
                    if (margin < 100 && margin != 0)
                    {
                        freightTotal = new Money(freight.Value / (1 - margin / 100));
                    }
                    //below code is used to calculate the total amount of bid sheet products at group level
                    var materialCost = new decimal();

                    foreach (var products in result.Entities)
                    {
                        var pid = (Guid)((EntityReference)(products.Attributes["ig1_productgroup"])).Id;
                        if (id.Equals(pid))
                        {
                            if (products.Attributes.Contains("ig1_totalamount"))
                            {
                                var amount = (Money)(products.Attributes["ig1_totalamount"]);
                                materialCost += (decimal)amount.Value; // adding all the bid sheet products amount to calculate the total amount of group.
                            }
                            else
                            {
                                var amount = new Money(0);
                                materialCost += (decimal)amount.Value; // adding all the bid sheet products amount to calculate the total amount of group.
                            }
                        }
                    }
                    //Call CreateBillOfMaterial Method
                    CreateBillOfMaterial(bidSheetId, productGroup, id, categoryId, projectHour, ProjectCost, freight, productLu, margin, freightTotal, materialCost, opportunityId);
                }

            }
            catch (Exception ex)
            {
                Entity errorLog = new Entity("ig1_pluginserrorlogs");
                errorLog["ig1_name"] = "Error";
                errorLog["ig1_errormessage"] = ex.Message;
                errorLog["ig1_errordescription"] = ex.StackTrace.ToString(); ;
                service.Create(errorLog);
                throw;
            }
        }


        protected void CreateBillOfMaterial(string bidSheetId, EntityReference productGroup, Guid id, Guid categoryId, decimal projectHour, Money ProjectCost, Money freight, decimal productLu, decimal margin, Money freightTotal, decimal materialCost, string opportunityId)
        {
            var flag = true;
            var fetchDataPriceList = new
            {
                ig1_bidsheet = bidSheetId.ToLower(),
                ig1_productId=id.ToString()
            };

            // fetching the records from the "bid sheet pricelist item" entity based on the bidsheet and product group.
            var fetchXmlPriceList = $@"
                                        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='ig1_bidsheetpricelistitem'>
                                            <attribute name='ig1_bidsheetpricelistitemid' />
                                            <attribute name='ig1_product' />
                                            <attribute name='ig1_bidsheet' />
                                            <filter type='and'>
                                              <condition attribute='ig1_bidsheet' operator='eq' value='{fetchDataPriceList.ig1_bidsheet}'/>
                                              <condition attribute='ig1_product' operator='eq' value='{fetchDataPriceList.ig1_productId}' />
                                            </filter>
                                          </entity>
                                        </fetch>";
            EntityCollection priceListResult = service.RetrieveMultiple(new FetchExpression(fetchXmlPriceList));
            var recordCount = priceListResult.Entities.Count;
            var bidSheetPriceListItemId = new Guid();
            //below code is used to make sure that the one group sould be added for the single bid sheet 
            if(recordCount>0)
            {
                var PriceListproductGroup = (EntityReference)(priceListResult.Entities[0].Attributes["ig1_product"]);
                if (PriceListproductGroup.Id.Equals(id))
                {
                    bidSheetPriceListItemId = (Guid)priceListResult.Entities[0].Attributes["ig1_bidsheetpricelistitemid"];
                    flag = false;
                }
            }
            // if product group does not exist then the new group will be added into the "bid sheet price list item" entity
            if (flag)
            {

                Entity entity = new Entity("ig1_bidsheetpricelistitem");
                entity["ig1_name"] = productGroup.Name.ToString();
                entity["ig1_product"] = new EntityReference("product", id);
                entity["ig1_bidsheet"] = new EntityReference("ig1_bidsheet", new Guid(bidSheetId));
                entity["ig1_category"] = new EntityReference("ig1_bidsheetcategory", categoryId);
                entity["ig1_projecthours"] = projectHour;
                entity["ig1_projectcost"] = ProjectCost;
                entity["ig1_freightamount"] = freight;
                entity["ig1_projectlu"] = productLu;
                entity["ig1_luextend"] = materialCost* productLu;
                entity["ig1_markup"] = margin;
                entity["ig1_freighttotal"] = freightTotal;
                entity["ig1_materialcost"] = new Money(materialCost);
                entity["ig1_unitprice"] = materialCost;
                entity["ig1_quantity"] = Convert.ToInt32(1);
                entity["ig1_opportunity"] = new EntityReference("opportunity", new Guid(opportunityId));
                service.Create(entity);
            }
            else
            {
                Entity entity = service.Retrieve("ig1_bidsheetpricelistitem", bidSheetPriceListItemId, new ColumnSet(true));
                entity["ig1_name"] = productGroup.Name.ToString();
                entity["ig1_product"] = new EntityReference("product", id);
                entity["ig1_bidsheet"] = new EntityReference("ig1_bidsheet", new Guid(bidSheetId));
                entity["ig1_category"] = new EntityReference("ig1_bidsheetcategory", categoryId);
                //entity["ig1_projecthours"] = projectHour;
                //entity["ig1_projectcost"] = ProjectCost;
                //entity["ig1_freightamount"] = freight;
                //entity["ig1_projectlu"] = projectLu;
                //entity["ig1_markup"] = markup;
                //entity["ig1_freighttotal"] = freightTotal;
                //entity["ig1_materialcost"] = new Money(materialCost);
                entity["ig1_opportunity"] = new EntityReference("opportunity", new Guid(opportunityId));
                service.Update(entity);
            }
        }
        
    }
}
