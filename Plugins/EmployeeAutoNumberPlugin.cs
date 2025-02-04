using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace EMS.Plugins
{
    public class EmployeeAutoNumberPlugin : IPlugin
    {
        /// <summary>
        /// Plugin's Execute with input parameter as ServiceProvider object which can be used to get the plugin context information, Input Parameters. 
        /// This method is called by the Plugin Event Framework when the event is fired
        /// </summary>
        /// <param name="serviceProvider"></param>
        public void Execute(IServiceProvider serviceProvider)
        {
            // Get the plugin run time execution context.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Get the tracing object to log the debug information.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Get the factory object to get the Org Service.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            // Create the Org Service Object.
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                {
                    //Skip the execution if the plugin is registered on a different entity.
                    if (target.LogicalName != "dev_employee")
                        return;

                    // Query to get the last generated Auto Number from the config table.
                    QueryExpression query = new QueryExpression("dev_autonumber");
                    FilterExpression filter = new FilterExpression();
                    ColumnSet columns = new ColumnSet("dev_lastcreatednumber", "dev_entity");
                    query.ColumnSet = columns;
                    ConditionExpression condition = new ConditionExpression("dev_entity", ConditionOperator.Equal, target.LogicalName);
                    query.Criteria.AddCondition(condition);
                    var results = service.RetrieveMultiple(query);
                    if (results != null && results.Entities != null)
                    {
                        tracingService.Trace($"Auto number record found with last generated number {results.Entities[0].Attributes["dev_lastcreatednumber"]}");

                        //Obtain a write-lock on the auto number entity.
                        Entity updateLock = new Entity("dev_autonumber", results.Entities[0].Id);
                        updateLock.Attributes.Add("dev_updateinprogress", true);
                        service.Update(updateLock);

                        int currentAutoNumber = (int)results.Entities[0].Attributes["dev_lastcreatednumber"];
                        currentAutoNumber++;

                        target.Attributes["dev_employeenumber"] = currentAutoNumber.ToString();

                        Entity updateAutoNumberConfig = new Entity("dev_autonumber", results.Entities[0].Id);
                        updateAutoNumberConfig.Attributes.Add("dev_updateinprogress", false);
                        updateAutoNumberConfig.Attributes.Add("dev_lastcreatednumber", currentAutoNumber);
                        service.Update(updateAutoNumberConfig);
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException($"No auto number record found for the entity {target.LogicalName}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"An exception occured in Autonumber generation for the Employee  record {ex.Message}");
            }
        }
    }
}

