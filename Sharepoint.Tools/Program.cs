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
            string siteName = "foo138";

            SecureString passWord = new SecureString();
            foreach (char c in "Verbinden1".ToCharArray()) passWord.AppendChar(c);

            CreateSite cs = new CreateSite(siteName);
           string createdSiteUrl = cs.Execute();

            //string createdSiteUrl = string.Format("https://adbdev.sharepoint.com/teams/{0}", siteName);

            /*
            using (ClientContext targetContext = new ClientContext(createdSiteUrl))
            {
                targetContext.Credentials = new SharePointOnlineCredentials(userName, passWord);
                targetContext.RequestTimeout = Timeout.Infinite;

                string fileRelativeUrl = "SitePages/Home.aspx";

                var web = targetContext.Web;
                var site = targetContext.Site;
                targetContext.Load(web);
                targetContext.Load(site);
                targetContext.ExecuteQueryRetry();
                string webUrl = web.EnsureProperty(c => c.ServerRelativeUrl);
                string serverRelativeUrl = UrlUtility.Combine(webUrl, fileRelativeUrl);

                var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);

                targetContext.Load(file, f => f.Name);
                targetContext.ExecuteQueryRetry();
                file.DeleteObject();
                targetContext.ExecuteQueryRetry();
            }
            */

            ProvisioningTemplate template = TemplateManager.GetProvisioningTemplate(ConsoleColor.White, templateSiteUrl, userName, passWord);

            /*
            ClientSidePageCollection pages = template.ClientSidePages;
            foreach (var p in pages)
            {
                p.Overwrite = true;
            }
            */

            TemplateManager.ApplyProvisioningTemplate(createdSiteUrl, userName, passWord, template);


            //TemplateManager.ApplyProvisioningTemplate(createdSiteUrl, userName, passWord);

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
                            if (item.Name == "ADBDocumentTypeValue")
                            {
                                item.Hidden = true;
                                Console.WriteLine("Updated Setting");
                            }
                        }
                        ct.Update(false);
                    }
                }



            }

            //AppInstaller apps = new AppInstaller();
            //apps.Execute();

            //WorkflowUpdater workflow = new WorkflowUpdater();
            //workflow.ExecuteUpdate();

            //Console.WriteLine(site.Url.ToString());


            Console.ReadLine();
        }
    }
}
