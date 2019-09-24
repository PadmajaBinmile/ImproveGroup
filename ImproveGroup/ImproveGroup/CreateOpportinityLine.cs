using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
namespace ImproveGroup
{
    public class CreateOpportinityLine : IPlugin
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

                //Nazish - Fetching all products which are associated to Bill of Material
                var fetchData = new
                {
                    ig1_bidsheet = bidSheetId
                };
                var fetchXmlBOMProducts = $@"
                                <fetch>
                                  <entity name='product'>
                                    <attribute name='defaultuomid' />
                                    <attribute name='productid' />
                                    <attribute name='name' />
                                    <link-entity name='ig1_bidsheetpricelistitem' from='ig1_product' to='productid'>
                                      <attribute name='ig1_materialcost' alias='materialCost'/>
                                      <attribute name='ig1_sdt' alias='sdt'/>
                                      <filter type='and'>
                                        <condition attribute='ig1_bidsheet' operator='eq' value='{fetchData.ig1_bidsheet}'/>
                                      </filter>
                                      <link-entity name='ig1_associatedcost' from='ig1_bidsheetcategory' to='ig1_category' link-type='outer'>
                                        <attribute name='ig1_margin' alias='margin'/>
                                        <filter type='and'>
                                          <condition attribute='ig1_bidsheet' operator='eq' value='{fetchData.ig1_bidsheet}'/>
                                        </filter>
                                      </link-entity>
                                    </link-entity>
                                  </entity>
                                </fetch>";
                EntityCollection BOMProductsData = service.RetrieveMultiple(new FetchExpression(fetchXmlBOMProducts));
                var unitId = new Guid();
                if (BOMProductsData.Entities.Count > 0)
                {
                    foreach (var item in BOMProductsData.Entities)
                    {
                        var productId = new Guid();
                        var materialCost = new Money(0);
                        var totalMaterialCost = new Money(0);
                        var margin = new decimal(0);

                        if (item.Attributes.Contains("productid"))
                        {
                            productId = (Guid)item.Attributes["productid"];
                        }
                        if (item.Attributes.Contains("defaultuomid"))
                        {
                            var unit = (EntityReference)item.Attributes["defaultuomid"];
                            unitId = (Guid)unit.Id;
                        }
                        if (item.Attributes.Contains("materialCost"))
                        {
                            materialCost = (Money)((AliasedValue)item.Attributes["materialCost"]).Value;
                        }
                        if (item.Attributes.Contains("margin"))
                        {
                            margin = (decimal)((AliasedValue)item.Attributes["margin"]).Value;
                        }
                        if (margin > 0)
                        {
                            totalMaterialCost = new Money(materialCost.Value /(1-margin/100));
                        }
                        else
                        {
                            totalMaterialCost = new Money(materialCost.Value);
                        }
                        if (item.KeyAttributes.Contains("sdt"))
                        {
                            totalMaterialCost = new Money((decimal)totalMaterialCost.Value + (decimal)((AliasedValue)item.Attributes["sdt"]).Value); 
                        }
                        CreatePriceListItem(opportunityId, productId, unitId, totalMaterialCost);
                        CreateOpportunityLine(productId, opportunityId, unitId);
                    }
                }
                //CreateProjectCostAllowancesPriceListItem(opportunityId, unitId, bidSheetId);
                //CreateProjectCostAllowancesOpportunityLine(opportunityId, unitId);
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
        protected void CreatePriceListItem(string opportunityId, Guid productId, Guid unitId, Money materialCost)
        {
            var fetchDataOpportunity = new
            {
                PriceLevelopportunityid = opportunityId
            };
            var PriceLevelFetch = $@"
                                    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='pricelevel'>
                                        <attribute name='transactioncurrencyid' />
                                        <attribute name='name' />
                                        <attribute name='pricelevelid' />
                                        <link-entity name='opportunity' from='pricelevelid' to='pricelevelid'>
                                            <filter type='and'>
                                            <condition attribute='opportunityid' operator='eq' value='{fetchDataOpportunity.PriceLevelopportunityid}'/>
                                            </filter>
                                        </link-entity>
                                        </entity>
                                    </fetch>";
            EntityCollection priceLevelData = service.RetrieveMultiple(new FetchExpression(PriceLevelFetch));
            foreach (var priceLevel in priceLevelData.Entities)
            {
                var productPriceLevelData = new
                {
                    productid = productId
                };
                var productPriceLevelfetchXml = $@"
                                                    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                        <entity name='productpricelevel'>
                                                        <attribute name='productid' />
                                                        <attribute name='pricelevelid' />
                                                        <attribute name='productpricelevelid' />
                                                        <attribute name='productidname' />
                                                        <filter type='and'>
                                                            <condition attribute='productid' operator='eq' value='{productPriceLevelData.productid}'/>
                                                            </filter>
                                                        </entity>
                                                    </fetch>";
                var check = true;
                var productPriceLevelId = new Guid();
                EntityCollection productPriceLevel = service.RetrieveMultiple(new FetchExpression(productPriceLevelfetchXml));
                foreach (var product in productPriceLevel.Entities)
                {
                    var proId = (EntityReference)(product.Attributes["productid"]);
                    if (proId.Id.Equals(productId) && !productPriceLevel.Equals(null))
                    {
                        check = false;
                        productPriceLevelId = (Guid)product.Attributes["productpricelevelid"];
                    }
                }
                var priceList = priceLevel.Attributes["pricelevelid"];
                if (check)
                {
                    Entity priceListItem = new Entity("productpricelevel");
                    priceListItem["pricelevelid"] = new EntityReference("pricelevel", (Guid)priceList);
                    priceListItem["productid"] = new EntityReference("product", productId);
                    priceListItem["uomid"] = new EntityReference("uom", unitId);
                    priceListItem["amount"] = materialCost;
                    service.Create(priceListItem);
                }
                else
                {

                    Entity priceListItem = service.Retrieve("productpricelevel", productPriceLevelId, new ColumnSet("pricelevelid", "productid", "uomid", "amount"));
                    priceListItem["pricelevelid"] = new EntityReference("pricelevel", (Guid)priceList);
                    priceListItem["productid"] = new EntityReference("product", productId);
                    priceListItem["uomid"] = new EntityReference("uom", unitId);
                    priceListItem["amount"] = materialCost;
                    service.Update(priceListItem);
                }
            }
        }
        protected void CreateOpportunityLine(Guid productid, string opportunityId, Guid unitId)
        {
            var isProductNotExist = true;
            var fetchData = new
            {
                opportunityid = opportunityId,
                productid = productid
            };
            var fetchXml = $@"
                            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='opportunityproduct'>
                                <attribute name='opportunityproductid' />
                                <attribute name='opportunityid' />
                                <attribute name='uomid' />
                                <attribute name='productid' />
                                <attribute name='quantity' />
                                <filter type='and'>
                                  <condition attribute='opportunityid' operator='eq' value='{fetchData.opportunityid}'/>
                                  <condition attribute='productid' operator='eq' value='{fetchData.productid}'/>
                                </filter>
                              </entity>
                            </fetch>";

            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            var recordCount = result.Entities.Count;
            var productId = new Guid();
            var opportunityProductId = new Guid();
            if (recordCount > 0)
            {
                productId = (Guid)((EntityReference)result.Entities[0].Attributes["productid"]).Id;
                opportunityProductId = (Guid)(result.Entities[0].Attributes["opportunityproductid"]);
                isProductNotExist = false;
            }

            if (isProductNotExist)
            {
                Entity opportunityLine = new Entity("opportunityproduct");
                opportunityLine["productid"] = new EntityReference("product", productid);
                opportunityLine["opportunityid"] = new EntityReference("opportunity", new Guid(opportunityId));
                opportunityLine["uomid"] = new EntityReference("uom", unitId);
                opportunityLine["quantity"] = new decimal(1);
                service.Create(opportunityLine);
            }
            else
            {
                Entity opportunityLine = service.Retrieve("opportunityproduct", opportunityProductId, new ColumnSet("productid", "opportunityid", "uomid", "quantity"));
                opportunityLine["productid"] = new EntityReference("product", productId);
                opportunityLine["opportunityid"] = new EntityReference("opportunity", new Guid(opportunityId));
                opportunityLine["uomid"] = new EntityReference("uom", unitId);
                opportunityLine["quantity"] = new decimal(1);
                service.Update(opportunityLine);
            }
        }
        /*protected void CreateProjectCostAllowancesPriceListItem(string opportunityId, Guid unitId, string bidSheetId)
        {
            var productId = "853c0615-f893-e911-a95d-000d3a1d5d22";
            var indirectCost = new Money(0);
            var fetchBidSheetData = new
            {
                ig1_bidsheetid = bidSheetId.Replace("{", "").Replace("}", "")
            };
            var BidSheetFetchXml = $@"
                            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='ig1_bidsheet'>
                                <attribute name='ig1_freightamount' />
                                <attribute name='ig1_indirectcost' />
                                <filter type='and'>
                                  <condition attribute='ig1_bidsheetid' operator='eq' value='{fetchBidSheetData.ig1_bidsheetid}'/>
                                </filter>
                              </entity>
                            </fetch>";
            EntityCollection BidSheetData = service.RetrieveMultiple(new FetchExpression(BidSheetFetchXml));
            if (BidSheetData.Entities[0].Attributes.Contains("ig1_indirectcost"))
            {
                decimal Cost = (decimal)BidSheetData.Entities[0].Attributes["ig1_indirectcost"];
                indirectCost = new Money(Cost);
            }


            var fetchDataOpportunity = new
            {
                PriceLevelopportunityid = opportunityId
            };
            var PriceLevelFetch = $@"
                                    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='pricelevel'>
                                        <attribute name='transactioncurrencyid' />
                                        <attribute name='name' />
                                        <attribute name='pricelevelid' />
                                        <link-entity name='opportunity' from='pricelevelid' to='pricelevelid'>
                                            <filter type='and'>
                                            <condition attribute='opportunityid' operator='eq' value='{fetchDataOpportunity.PriceLevelopportunityid}'/>
                                            </filter>
                                        </link-entity>
                                        </entity>
                                    </fetch>";
            EntityCollection priceLevelData = service.RetrieveMultiple(new FetchExpression(PriceLevelFetch));
            foreach (var priceLevel in priceLevelData.Entities)
            {
                var productPriceLevelData = new
                {
                    productid = productId
                };
                var productPriceLevelfetchXml = $@"
                                                    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                        <entity name='productpricelevel'>
                                                        <attribute name='productid' />
                                                        <attribute name='pricelevelid' />
                                                        <attribute name='productpricelevelid' />
                                                        <attribute name='productidname' />
                                                        <filter type='and'>
                                                            <condition attribute='productid' operator='eq' value='{productPriceLevelData.productid}'/>
                                                            </filter>
                                                        </entity>
                                                    </fetch>";
                var check = true;
                var productPriceLevelId = new Guid();
                EntityCollection productPriceLevel = service.RetrieveMultiple(new FetchExpression(productPriceLevelfetchXml));
                foreach (var product in productPriceLevel.Entities)
                {
                    var pid = productId.ToLower();
                    var proId = (EntityReference)(product.Attributes["productid"]);
                    var id = proId.Id.ToString().Replace("{", "").Replace("}", "");
                    if (id.Equals(pid) && !productPriceLevel.Equals(null))
                    {
                        check = false;
                        productPriceLevelId = (Guid)product.Attributes["productpricelevelid"];
                    }
                }
                var priceList = priceLevel.Attributes["pricelevelid"];
                if (check)
                {
                    Entity priceListItem = new Entity("productpricelevel");
                    priceListItem["pricelevelid"] = new EntityReference("pricelevel", (Guid)priceList);
                    priceListItem["productid"] = new EntityReference("product", new Guid(productId));
                    priceListItem["uomid"] = new EntityReference("uom", unitId);
                    priceListItem["amount"] = indirectCost;
                    service.Create(priceListItem);
                }
                else
                {

                    Entity priceListItem = service.Retrieve("productpricelevel", productPriceLevelId, new ColumnSet("pricelevelid", "productid", "uomid", "amount"));
                    priceListItem["pricelevelid"] = new EntityReference("pricelevel", (Guid)priceList);
                    priceListItem["productid"] = new EntityReference("product", new Guid(productId));
                    priceListItem["uomid"] = new EntityReference("uom", unitId);
                    priceListItem["amount"] = indirectCost;
                    service.Update(priceListItem);
                }
            }

        }*/
        /*protected void CreateProjectCostAllowancesOpportunityLine(string opportunityId, Guid unitId)
        {
            var projetcCostAllowancesId =  "853C0615-F893-E911-A95D-000D3A1D5D22";
            var isProductNotExist = true;
            var fetchData = new
            {
                opportunityid = opportunityId,
                productid = projetcCostAllowancesId
            };
            var fetchXml = $@"
                            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='opportunityproduct'>
                                <attribute name='opportunityproductid' />
                                <attribute name='opportunityid' />
                                <attribute name='uomid' />
                                <attribute name='productid' />
                                <attribute name='quantity' />
                                <filter type='and'>
                                  <condition attribute='opportunityid' operator='eq' value='{fetchData.opportunityid}'/>
                                  <condition attribute='productid' operator='eq' value='{fetchData.productid}'/>
                                </filter>
                              </entity>
                            </fetch>";

            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            var recordCount = result.Entities.Count;
            var productId = new Guid();
            var opportunityProductId = new Guid();
            if (recordCount > 0)
            {
                productId = (Guid)((EntityReference)result.Entities[0].Attributes["productid"]).Id;
                opportunityProductId = (Guid)(result.Entities[0].Attributes["opportunityproductid"]);
                isProductNotExist = false;
            }

            if (isProductNotExist)
            {
                Entity opportunityLine = new Entity("opportunityproduct");
                opportunityLine["productid"] = new EntityReference("product", new Guid(projetcCostAllowancesId));
                opportunityLine["opportunityid"] = new EntityReference("opportunity", new Guid(opportunityId));
                opportunityLine["uomid"] = new EntityReference("uom", unitId);
                opportunityLine["quantity"] = new decimal(1);
                service.Create(opportunityLine);
            }
            else
            {
                Entity opportunityLine = service.Retrieve("opportunityproduct", opportunityProductId, new ColumnSet("productid", "opportunityid", "uomid", "quantity"));
                opportunityLine["productid"] = new EntityReference("product", productId);
                opportunityLine["opportunityid"] = new EntityReference("opportunity", new Guid(opportunityId));
                opportunityLine["uomid"] = new EntityReference("uom", unitId);
                opportunityLine["quantity"] = new decimal(1);
                service.Update(opportunityLine);
            }
        }*/
    }
}