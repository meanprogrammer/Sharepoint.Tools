using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Sharepoint.Tools
{
    class Program
    {
        static void Main(string[] args)
        {
            string adminTenantSiteUrl = "https://adbdev-admin.sharepoint.com";
            string templateSiteUrl = "https://adbdev.sharepoint.com/teams/template_collab";
            string userName = "vdudan@adbdev.onmicrosoft.com";
            string siteName = "foo161";



            SecureString passWord = new SecureString();
            foreach (char c in "Verbinden1".ToCharArray()) passWord.AppendChar(c);

            CreateSite cs = new CreateSite(siteName);
            string createdSiteUrl = cs.Execute();

            //string createdSiteUrl = string.Format("https://adbdev.sharepoint.com/teams/{0}", siteName);


            ProvisioningTemplate template = TemplateManager.GetProvisioningTemplate(ConsoleColor.White, templateSiteUrl, userName, passWord);

            TemplateManager.ApplyProvisioningTemplate(createdSiteUrl, userName, passWord);

            using (ClientContext targetContext = new ClientContext(createdSiteUrl))
            {
                targetContext.Credentials = new SharePointOnlineCredentials(userName, passWord);
                targetContext.RequestTimeout = Timeout.Infinite;

                var web = targetContext.Web;
                targetContext.Load(web);
                targetContext.ExecuteQueryRetry();
                web.EnsureProperties(c => c.Lists);
                List doc = web.Lists.FirstOrDefault(c => c.Title == "Documents");

                targetContext.Load(doc);
                targetContext.Load(doc.ContentTypes);
                var cts = doc.ContentTypes;
                targetContext.ExecuteQueryRetry();


                doc.EnsureProperties(c => c.Fields, c => c.ContentTypes);

                foreach (var ct in doc.ContentTypes)
                {
                    //targetContext.Load(ct);
                    //targetContext.ExecuteQueryRetry();
                    ct.EnsureProperties(c => c.FieldLinks, c => c.Fields);
                    if (ct.Name == "ADB Document" || ct.Name == "ADB Project Document" || ct.Name == "ADB Country Document")
                    {
                        foreach (var item in ct.FieldLinks)
                        {
                            if (item.Name == "ADBDocumentTypeValue" || item.Name == "ADBContentGroup")
                            {
                                item.Hidden = true;
                                Console.WriteLine("Updated Field Visibility Setting");
                            }
                        }
                        ct.Update(false);
                    }
                }


                targetContext.Load(doc.WorkflowAssociations);
                targetContext.ExecuteQuery();

                var servicesManager = new WorkflowServicesManager(targetContext, web);
                var subscriptionService = servicesManager.GetWorkflowSubscriptionService();
                var subscriptions = subscriptionService.EnumerateSubscriptionsByList(doc.Id);

                targetContext.Load(subscriptions);
                targetContext.ExecuteQuery();

                var wfh = web.Lists.GetByTitle("Workflow History");
                var wft = web.Lists.GetByTitle("Update Document Type Workflow Tasks");
                var dwt = web.Lists.GetByTitle("Workflow Tasks");
                targetContext.Load(wfh);
                targetContext.Load(wft);
                targetContext.Load(dwt);
                targetContext.ExecuteQuery();


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
                        subscriptionService.PublishSubscriptionForList(s, doc.Id);
                    }
                    else
                    {
                        s.SetProperty("HistoryListId", wfh.Id.ToString());
                        s.SetProperty("TaskListId", dwt.Id.ToString());
                        s.SetProperty("FormData", "");
                        subscriptionService.PublishSubscriptionForList(s, doc.Id);
                    }
                }
                targetContext.ExecuteQuery();


                string[] fieldsForRemoval = new string[] { "Update ADB Country Document Type", "Update ADB Document Type", "Update ADB Project Document Type", "Log Activity", "Log Activity Native" };

                foreach (string fieldName in fieldsForRemoval)
                {
                    var f = doc.Fields.GetByInternalNameOrTitle(fieldName);
                    targetContext.Load(f);
                    targetContext.ExecuteQueryRetry();

                    if (f != null)
                    {
                        f.DeleteObject();
                        targetContext.ExecuteQueryRetry();
                    }
                }




            }

            //AppInstaller apps = new AppInstaller();
            //apps.Execute();

            //WorkflowUpdater workflow = new WorkflowUpdater();
            //workflow.ExecuteUpdate();

            //Console.WriteLine(site.Url.ToString());

            Console.WriteLine("END");
            Console.ReadLine();
        }
    }
}
