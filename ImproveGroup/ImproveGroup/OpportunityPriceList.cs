using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ImproveGroup
{
    public class OpportunityPriceList : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Setup
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.LogicalName != "opportunity")
                {
                    return;
                }
                try
                {
                    if (entity.Attributes.Contains("transactioncurrencyid"))
                    {
                        var transactionCurrency = (EntityReference)entity.Attributes["transactioncurrencyid"];
                        var currencyId = (Guid)transactionCurrency.Id;

                        Entity priceList = new Entity("pricelevel");
                        priceList["name"] = entity.Attributes["name"]+entity.Id.ToString();
                        priceList["transactioncurrencyid"] = new EntityReference(transactionCurrency.LogicalName, currencyId);
                        service.Create(priceList);
                    }
                    if (entity.Attributes.Contains("pricelevelid"))
                    {
                        var priceListId = new Guid();
                        var fetchData = new
                        {
                            name = entity.Attributes["name"] + entity.Id.ToString()
                        };
                        var fetchXml = $@"
                                        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='pricelevel'>
                                            <attribute name='name' />
                                            <attribute name='pricelevelid' />
                                            <filter type='and'>
                                                <condition attribute='name' operator='eq' value='{fetchData.name}'/>
                                            </filter>
                                            </entity>
                                        </fetch>";
                        EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        foreach (var priceList in result.Entities)
                        {
                            priceListId = (Guid)priceList.Attributes["pricelevelid"];
                            entity["pricelevelid"] = new EntityReference(priceList.LogicalName, priceList.Id);
                            service.Update(entity);

                        }

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
        }
    }
}
