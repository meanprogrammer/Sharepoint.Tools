using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
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
            string siteName = "foo93";

            SecureString passWord = new SecureString();
            foreach (char c in "Verbinden1".ToCharArray()) passWord.AppendChar(c);

            CreateSite cs = new CreateSite(siteName);
            string createdSiteUrl = cs.Execute();
            //string createdSiteUrl = string.Format("https://adbdev.sharepoint.com/teams/{0}", siteName);
            ProvisioningTemplate template = TemplateManager.GetProvisioningTemplate(ConsoleColor.White, templateSiteUrl, userName, passWord);

            foreach (var ct in template.ContentTypes)
            {
                if (ct.Name == "ADB Document" || ct.Name == "ADB Country Document" || ct.Name == "ADB Project Document") {

                    var fields = ct.FieldRefs;
   

               foreach (var ff in fields) 
               {
                        
                   if (ff.Name == "ADBDocumentTypeValue") {
                            ff.Hidden = true;
                     }
                    }
                }
            }

            TemplateManager.ApplyProvisioningTemplate(createdSiteUrl, userName, passWord, template);


            //AppInstaller apps = new AppInstaller();
            //apps.Execute();
            
            //WorkflowUpdater workflow = new WorkflowUpdater();
            //workflow.ExecuteUpdate();

            //Console.WriteLine(site.Url.ToString());

            Console.ReadLine();
        }
    }
}
