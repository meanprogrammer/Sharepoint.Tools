using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Sites;
using System;
using System.Security;


namespace Sharepoint.Tools
{
    public class CreateSite
    {
        private string _siteName = string.Empty;

        public CreateSite() { }
        public CreateSite(string siteName) { _siteName = siteName; }

        public string Execute()
        {
            using (ClientContext tenantContext = new ClientContext("https://adbdev-admin.sharepoint.com/"))
            {
                //Authenticate with a Tenant Administrator
                SecureString passWord = new SecureString();
                foreach (char c in "Verbinden1".ToCharArray()) passWord.AppendChar(c);
                tenantContext.Credentials = new SharePointOnlineCredentials("vdudan@adbdev.onmicrosoft.com", passWord);

                TeamSiteCollectionCreationInformation siteInformation = new TeamSiteCollectionCreationInformation() { Alias = _siteName, DisplayName = _siteName, IsPublic = true };

                var result = tenantContext.CreateSiteAsync(siteInformation);
                var returnedContext = result.GetAwaiter().GetResult();

                var web = returnedContext.Web;
                returnedContext.Load(web);
                returnedContext.Load(web.Lists);
                returnedContext.ExecuteQuery();

                web.Lists.EnsureSiteAssetsLibrary();
                returnedContext.ExecuteQuery();

                var site = returnedContext.Site;
                returnedContext.Load(site);
                returnedContext.ExecuteQuery();

                Tenant t = new Tenant(tenantContext);
                var details = t.GetSitePropertiesByUrl(returnedContext.Url, true);
                t.Context.Load(details);
                t.Context.ExecuteQuery();

                Console.WriteLine(details.Status);

                //No Script Site
                t.SetSiteProperties(returnedContext.Url, noScriptSite: false);
                t.Context.ExecuteQuery();

                details = t.GetSitePropertiesByUrl(returnedContext.Url, true);
                t.Context.Load(details);
                t.Context.ExecuteQuery();

                Console.WriteLine("SiteCollection Created.");

                return returnedContext.Url;
            }
        }
    }
}
