using System;
using System.Activities;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Workflow;

namespace ActivateMailBoxes
{

    #region Public Methods
    public sealed class MailBoxCheck : CodeActivity
    {
        [ComVisible(false)]
        [MTAThread] 
        protected override void Execute(CodeActivityContext executionContext)
        {
            //Create the context
            var tracingService = executionContext.GetExtension<ITracingService>();
            var context = executionContext.GetExtension<IWorkflowContext>();
            var serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var orgContex = new OrganizationServiceContext(service);
            

            try
            {
                //get inactive mailboxes in the system
                var inactiveRecords = orgContex.CreateQuery<Mailbox>()
                    .Where(_ => (_.StateCode == MailboxState.Inactive))
                    .Where(_ => _.EmailAddress != null && _.EmailAddress != string.Empty)
                    .Select(_ => new
                    {
                        ID = _.MailboxId,
                        UserName = _.Username,
                        Email = _.EmailAddress,
                    }).ToList();
                
                var mailboxes = inactiveRecords.Select(inactiveRecord => new InactiveMailbox
                {
                    MailboxId = inactiveRecord.ID == Guid.Empty ? Guid.Empty : inactiveRecord.ID,
                    EmailAddress = string.IsNullOrEmpty(inactiveRecord.Email) ? string.Empty : inactiveRecord.Email,
                    UserName = string.IsNullOrEmpty(inactiveRecord.UserName) ? string.Empty : inactiveRecord.UserName
                }).ToList();

                //activate mailboxes
                foreach (var record in from mailbox in mailboxes
                    let mailboxId = mailbox.MailboxId
                    where mailboxId != null
                    where mailboxId != Guid.Empty  let setStateResponse = (SetStateResponse) service.Execute(new SetStateRequest
                {
                    EntityMoniker = new EntityReference
                    {
                        Id = mailboxId.Value,
                        LogicalName = Mailbox.EntityLogicalName
                    },
                    State = new OptionSetValue(0),
                    Status = new OptionSetValue(1)
                }) select new Mailbox
                {
                    MailboxId =  mailboxId,
                    IncomingEmailDeliveryMethod = new OptionSetValue(2),
                    OutgoingEmailDeliveryMethod = new OptionSetValue(2),
                    EmailRouterAccessApproval = new OptionSetValue(1),
                    TestEmailConfigurationScheduled = true
                })
                {
                    service.Update(record);
                }


            }
            catch (InvalidWorkflowException ex)
            {
               tracingService.Trace("Exception Occurred" + ex.Message);
               throw new InvalidWorkflowException("No emailboxes got activated " + ex.Message);
            }
        }
    #endregion
         

        #region Input Variables
        
        #endregion


        #region Output Variables

        #endregion

    }
}
