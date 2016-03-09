using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lally.Andrew.Crm.Plugins
{
    /// <summary>
    /// Plugin to validate that an order is being created by converting a quote to an order. The plugin 
    /// should only be used on an sales order entity and on the create message. We are going to fail fast 
    /// if it is on anything else. 
    /// </summary>
    public class ValidateOrderCreatedFromQuote : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //Retreive the Plugin Execution Context
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Verify that we were able to retrieve the plugin context succesfully 
            if (context == null)
                throw new InvalidPluginExecutionException(OperationStatus.Failed, "Could not retrieve an IPluginExecutionContext object");

            if (tracer == null)
                throw new InvalidPluginExecutionException(OperationStatus.Failed, "Could not retrieve an ITracingService object");
 
            /**
            * I prefer to use the fail fast and loud approach as I believe it makes it easier later on for system administrators and myself in the event that someone 
            * accidentally attaches this to the wrong entity or message. We could let it continue and a different error could be thrown if we try to access 
            * something that isn't available. Alternatively we could just return without doing anything, both would make it a little more difficult for the administrator to debug
            * and would probably result in a phone call because the errors wouldn't make much sense without looking at the code.
            */
            if (!string.Equals(context.PrimaryEntityName, "salesorder", StringComparison.InvariantCultureIgnoreCase))
                throw new InvalidPluginExecutionException(OperationStatus.Canceled,
                    "The ValidateOrderCreatedFromQuote plugin should only be attached to the salesorder entity but is attached to the " + context.PrimaryEntityName + " entity");

            //Do the same checking to make sure we are on the create message 
            if (!string.Equals(context.MessageName, "create", StringComparison.InvariantCultureIgnoreCase))
                throw new InvalidPluginExecutionException(OperationStatus.Canceled,
                    "The ValidateOrderCreatedFromQuote plugin should only be attached to the create message but is attached to the " + context.MessageName + " message");

            //This plugin is useless if it is not running synchronously so lets give an error if we are not in that state. 
            if (context.Mode != (int)SdkMessageProcessingStepMode.Synchronous)
                throw new InvalidPluginExecutionException(OperationStatus.Canceled,
                    "The ValidateOrderCreatedFromQuote plugin only works when executed synchronously but it is being executed asynchrously");

            /**
            * This check is not really 100% necessary but is more of a best practice. We could let the plugin run in any stage but it would probably be best to atleast run it in 
            * one that comes before anything is written to the database so that things arn't written and then rolled back. Because the plugin can still function and work correctly 
            * we are just going to log the issue for the admin to see if they have tracing turned on
            */
            if (context.Stage > (int)SdkMessageProcessingStepStage.Preoperation)
                tracer.Trace("The ValidateOrderCreatedFromQuote is registered after the database write, it is best practice to register in the PreOperation stage");

            
            /**
             * Once everything is validated we can check the parent context to make sure the salesorder create process was initiated beacuse of a 
             * convertquotetosalesorder message. The parent context can be thought of as the step that caused this step to run. So for example 
             * if you create a plugin on the quote post create and then in that create quote details. If you would create another plugin on the quote 
             * detail create, the parentContext would be the context of your quote plugin for the quote details your plugin created.
             * (This may be confusing so see my blog post for a nice diagram) 
             *
             * How this applies to this situation is that when you want to convert a quote you use the convertquotetosalesordermessage when that is being 
             * executed it creates the order entity similar the above scenario so the parent context will contain that message. 
            **/
            for (var parentContext = context.ParentContext; parentContext != null; parentContext = parentContext.ParentContext)
            {
                //We are looking for the convertquotetosalesorder message
                if (string.Equals(parentContext.MessageName, "convertquotetosalesorder", StringComparison.InvariantCultureIgnoreCase))
                    return; //Because we allow this type of action we will simply return
            }

            //If you have certain users / groups that should be allowed to directly create sales order you could check for that 
            var userId = context.InitiatingUserId;
            //Once you have the user Id you can query for the groups they are in. 

            //If we get to this point none of the parent context messages were correct so throw an exception explaining to the user that you must create a quote first 
            //This exception and the string will be seen by the user when they try to save, it will also be returned to anything that tries to use the API to create a sales order. 
            throw new InvalidPluginExecutionException(OperationStatus.Canceled,
                "You can only create a sales order through the quote to sales order process. Please create a quote and then convert it to a sales order");



        }
    }
}
