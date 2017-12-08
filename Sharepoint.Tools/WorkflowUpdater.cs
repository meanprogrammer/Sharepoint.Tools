using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using System.Security;

namespace Sharepoint.Tools
{
    public class WorkflowUpdater
    {
        public void ExecuteUpdate()
        {
            string siteCollectionUrl = "https://adbdev.sharepoint.com/teams/foo26";
            string userName = "vdudan@adbdev.onmicrosoft.com";
            string password = "Verbinden1";

            // Namespace: Microsoft.SharePoint.Client  
            ClientContext ctx = new ClientContext(siteCollectionUrl);

            // Namespace: System.Security
            SecureString secureString = new SecureString();
            password.ToList().ForEach(secureString.AppendChar);

            // Namespace: Microsoft.SharePoint.Client  
            ctx.Credentials = new SharePointOnlineCredentials(userName, secureString);

            // Namespace: Microsoft.SharePoint.Client  
            Site site = ctx.Site;

            ctx.Load(site);
            ctx.ExecuteQuery();

            Web web = ctx.Web;
            ctx.Load(web);

            ctx.ExecuteQuery();

            ListCollection lists = web.Lists;
            ctx.Load(lists);
            ctx.ExecuteQuery();

            List docs = lists.GetByTitle("Documents");

            ctx.Load(docs);
            ctx.ExecuteQuery();

            ctx.Load(docs.WorkflowAssociations);
            ctx.ExecuteQuery();

            var servicesManager = new WorkflowServicesManager(ctx, web);
            var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
            var subscriptions = subscriptionService.EnumerateSubscriptionsByList(docs.Id);

            ctx.Load(subscriptions);
            ctx.ExecuteQuery();

            var wfh = lists.GetByTitle("Workflow History");
            var wft = lists.GetByTitle("Update Document Type Workflow Tasks");
            var dwt = lists.GetByTitle("Workflow Tasks");
            ctx.Load(wfh);
            ctx.Load(wft);
            ctx.Load(dwt);
            ctx.ExecuteQuery();


            foreach (var s in subscriptions)
            {
                Console.WriteLine(s.Name);

                if (
                    s.Name.Equals("Update ADB Project Document Type") ||
                    s.Name.Equals("Update ADB Country Document Type") ||
                    s.Name.Equals("Update ADB Document Type")
                    )
                {
                    s.SetProperty("HistoryListId", wfh.Id.ToString());
                    s.SetProperty("TaskListId", wft.Id.ToString());
                    s.SetProperty("FormData", string.Empty);
                    subscriptionService.PublishSubscriptionForList(s, docs.Id);
                }
                else
                {
                    s.SetProperty("HistoryListId", wfh.Id.ToString());
                    s.SetProperty("TaskListId", dwt.Id.ToString());
                    s.SetProperty("FormData", "");
                    subscriptionService.PublishSubscriptionForList(s, docs.Id);
                }
            }
            ctx.ExecuteQuery();
        }
    }
}
