using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint.Tools
{
    public class CreateSite
    {
        public void Execute()
        {
            using (ClientContext tenantContext = new ClientContext("https://adbdev-admin.sharepoint.com/"))
            {
                //Authenticate with a Tenant Administrator
                SecureString passWord = new SecureString();
                foreach (char c in "password".ToCharArray()) passWord.AppendChar(c);
                tenantContext.Credentials = new SharePointOnlineCredentials("admin@yoursite.onmicrosoft.com", passWord);

                var tenant = new Tenant(tenantContext);

                //Properties of the New SiteCollection
                var siteCreationProperties = new SiteCreationProperties();
                //New SiteCollection Url
                siteCreationProperties.Url = "https://adbdev.sharepoint.com/teams/foo41";
                //Title of the Root Site
                siteCreationProperties.Title = "Site Created from Code";
                //Email of Owner
                siteCreationProperties.Owner = "vdudan@adbdev.onmicrosoft.com";
                //Template of the Root Site. Using Team Site for now.
                siteCreationProperties.Template = "STS#0";
                //Storage Limit in MB
                siteCreationProperties.StorageMaximumLevel = 100;
                //UserCode Resource Points Allowed
                siteCreationProperties.UserCodeMaximumLevel = 50;

                //Create the SiteCollection
                SpoOperation spo = tenant.CreateSite(siteCreationProperties);

                tenantContext.Load(tenant);

                //We will need the IsComplete property to check if the provisioning of the Site Collection is complete.
                tenantContext.Load(spo, i => i.IsComplete);

                tenantContext.ExecuteQuery();

                //Check if provisioning of the SiteCollection is complete.
                while (!spo.IsComplete)
                {
                    //Wait for 30 seconds and then try again
                    System.Threading.Thread.Sleep(30000);
                    spo.RefreshLoad();
                    tenantContext.ExecuteQuery();
                }
                Console.WriteLine("SiteCollection Created.");
            }
        }
    }
}
