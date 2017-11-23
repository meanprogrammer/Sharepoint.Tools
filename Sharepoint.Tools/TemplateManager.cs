using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Sharepoint.Tools
{
    public class TemplateManager
    {
        private static string _currentDir = System.IO.Directory.GetCurrentDirectory();

        private static string GetInput(string label, bool isPassword, ConsoleColor defaultForeground)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("{0} : ", label);
            Console.ForegroundColor = defaultForeground;

            string value = "";

            for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
            {
                if (keyInfo.Key == ConsoleKey.Backspace)
                {
                    if (value.Length > 0)
                    {
                        value = value.Remove(value.Length - 1);
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        Console.Write(" ");
                        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                    }
                }
                else if (keyInfo.Key != ConsoleKey.Enter)
                {
                    if (isPassword)
                    {
                        Console.Write("*");
                    }
                    else
                    {
                        Console.Write(keyInfo.KeyChar);
                    }
                    value += keyInfo.KeyChar;

                }

            }
            Console.WriteLine("");

            return value;
        }

        public static ProvisioningTemplate GetProvisioningTemplateFromFile(ConsoleColor defaultForeground, string webUrl, string userName, SecureString pwd)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                // ctx.Credentials = new NetworkCredentials(userName, pwd);
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = Timeout.Infinite;

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is:" + ctx.Web.Title);
                Console.ForegroundColor = defaultForeground;

                ProvisioningTemplateCreationInformation ptci
                        = new ProvisioningTemplateCreationInformation(ctx.Web);

                // Create FileSystemConnector to store a temporary copy of the template 
                ptci.FileConnector = new FileSystemConnector(_currentDir, "");
                ptci.PersistBrandingFiles = true;
                ptci.PersistPublishingFiles = true;
                ptci.HandlersToProcess = Handlers.All;
                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    // Only to output progress for console UI
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                // Execute actual extraction of the template
                //ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);

                // We can serialize this template to save and reuse it
                // Optional step 
                XMLTemplateProvider provider =
                        new XMLFileSystemTemplateProvider(_currentDir, "");
                ProvisioningTemplate template = provider.GetTemplate(string.Format(@"{0}\All.xml", _currentDir));

                return template;
            }
        }

        public static ProvisioningTemplate GetProvisioningTemplate(ConsoleColor defaultForeground, string webUrl, string userName, SecureString pwd)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                // ctx.Credentials = new NetworkCredentials(userName, pwd);
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = Timeout.Infinite;

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Your site title is:" + ctx.Web.Title);
                Console.ForegroundColor = defaultForeground;

                ProvisioningTemplateCreationInformation ptci
                        = new ProvisioningTemplateCreationInformation(ctx.Web);

                // Create FileSystemConnector to store a temporary copy of the template 
                ptci.FileConnector = new FileSystemConnector(_currentDir, "");
                ptci.PersistBrandingFiles = true;
                ptci.PersistPublishingFiles = true;
                //ptci.PersistMultiLanguageResources = true;
                ptci.IncludeNativePublishingFiles = true;
                ptci.HandlersToProcess = Handlers.AuditSettings | Handlers.ContentTypes | Handlers.CustomActions | Handlers.ExtensibilityProviders | Handlers.Features | Handlers.Fields | Handlers.Files | Handlers.ImageRenditions | Handlers.Lists | Handlers.Navigation | Handlers.PageContents | Handlers.Pages | Handlers.PropertyBagEntries | Handlers.Publishing | Handlers.RegionalSettings | Handlers.SearchSettings | Handlers.SitePolicy | Handlers.SiteSecurity | Handlers.SupportedUILanguages | Handlers.TermGroups | Handlers.WebSettings | Handlers.Workflows;


                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    // Only to output progress for console UI
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                // Execute actual extraction of the template
                ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);

                // We can serialize this template to save and reuse it
                // Optional step 

                ClientSidePageCollection pages = template.ClientSidePages;
                foreach (var p in pages)
                {
                    p.Overwrite = true;
                }

                XMLTemplateProvider provider =
                        new XMLFileSystemTemplateProvider(_currentDir, "");
                provider.SaveAs(template, "All.xml");

                return template;
            }
        }

        public static void ApplyProvisioningTemplate(string webUrl, string userName, SecureString pwd, ProvisioningTemplate template)
        {
            using (ClientContext targetContext = new ClientContext(webUrl))
            {
                targetContext.Credentials = new SharePointOnlineCredentials(userName, pwd);
                targetContext.RequestTimeout = Timeout.Infinite;

                var web = targetContext.Web;
                targetContext.Load(web);
                targetContext.ExecuteQuery();

                

                ProvisioningTemplateApplyingInformation ptai
                        = new ProvisioningTemplateApplyingInformation();
                ptai.ClearNavigation = true;
                //ptai.HandlersToProcess = Handlers.All;
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                FileSystemConnector connector = new FileSystemConnector(_currentDir, "");
                template.Connector = connector;
                
                web.ApplyProvisioningTemplate(template, ptai);

            }
        }


        public static void ApplyProvisioningTemplate(string webUrl, string userName, SecureString pwd)
        {
            using (ClientContext targetContext = new ClientContext(webUrl))
            {
                targetContext.Credentials = new SharePointOnlineCredentials(userName, pwd);
                targetContext.RequestTimeout = Timeout.Infinite;

                var web = targetContext.Web;
                targetContext.Load(web);
                targetContext.ExecuteQuery();

                // Configure the XML file system provider
                XMLTemplateProvider provider =
                new XMLFileSystemTemplateProvider(_currentDir, string.Empty);

                // Load the template from the XML stored copy
                ProvisioningTemplate template = provider.GetTemplate(
                  @"All.xml");


                FileSystemConnector connector = new FileSystemConnector(_currentDir, "");
                template.Connector = connector;


                ProvisioningTemplateApplyingInformation ptai
                        = new ProvisioningTemplateApplyingInformation();
                ptai.ClearNavigation = true;
                //ptai.HandlersToProcess = Handlers.All;
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                

                web.ApplyProvisioningTemplate(template, ptai);

            }
        }
    }
}
