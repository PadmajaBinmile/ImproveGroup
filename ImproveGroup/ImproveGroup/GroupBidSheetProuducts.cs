using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace ImproveGroup
{
    public class GroupBidSheetProuducts : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Setup
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            if (context.InputParameters.Equals(null))
            {
                return;
            }
            try
            {
                string productGroup = context.InputParameters["selectedGroupID"].ToString();
                string getRecords = context.InputParameters["selectedRecordsId"].ToString();
                string[] records = getRecords.Split(new string[] { "," }, StringSplitOptions.None);
                foreach (string recordId in records)
                {
                    Entity entity = service.Retrieve("ig1_bidsheetproduct", new Guid(recordId), new ColumnSet("ig1_productgroup", "ig1_isgrouped"));
                    entity["ig1_productgroup"] = new EntityReference("product", new Guid(productGroup));
                    entity["ig1_isgrouped"] = true;
                    service.Update(entity);
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
