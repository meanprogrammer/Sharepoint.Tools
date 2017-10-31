using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint.Tools
{
    public class AppInstaller
    {
        public void Execute()
        {
            string siteCollectionUrl = "https://adbdev.sharepoint.com/teams/foo29";
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
            Web web = ctx.Web;

            ctx.Load(site);           
            ctx.Load(web);
            ctx.ExecuteQuery();


            Stream package = null;
            try
            {
                

                package = System.IO.File.OpenRead("ADBPublishDocument.app");
                AppInstance appInstance = web.LoadAndInstallApp(package);

                ctx.Load(appInstance);
                ctx.ExecuteQuery();

                if (appInstance != null && appInstance.Status == AppInstanceStatus.Initialized)
                {
                    var instance = appInstance.Id;
                }
            }
            finally
            {
                if (package != null)
                    package.Close();
            }
        }
    }
}
