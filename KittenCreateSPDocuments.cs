using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using PnP.Core.Auth;
using PnP.Core.Model.SharePoint;
using PnP.Framework;
using PnP.Framework.Modernization;
using PnP.Framework.Provisioning.Model;
using System.Security;
using Microsoft.SharePoint.Client;

namespace Kitten.DocService
{
    public static class KittenCreateSPDocuments
    {
        [FunctionName("KittenCreateSPDocuments")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("Called to Create a new Document");

            string sprawdata = await new StreamReader(req.Body).ReadToEndAsync();

            // convert to json
            log.LogInformation("Sp List Event Data");
            log.LogInformation("Content: " + sprawdata);

            KittenPageParameters kpageparam =  System.Text.Json.JsonSerializer.Deserialize<KittenPageParameters>(sprawdata);
            string pagefilename = "OfferingPages/" + kpageparam.OfferingTitle + ".aspx";


            // connect to sp and create a new document - this is very very very bad
            string spurl = "https://m365x31030766.sharepoint.com/sites/KittenSolutions";  // sharepoint URL
            string spuid = "";   // User name
            string sppwd = "";  // password - bad bad bad
            Web spweb;

            var secpwd = new SecureString();
            foreach (char c in sppwd)
            {
                secpwd.AppendChar(c);
            }

            log.LogInformation("Auth the User");
            PnP.Framework.AuthenticationManager authman = new PnP.Framework.AuthenticationManager(spuid, secpwd);

            log.LogInformation("Connect to SP with the Auth User");
            using (ClientContext ctx = await authman.GetContextAsync(spurl))
            {
                spweb = ctx.Web;
                ctx.Load(spweb);
                await ctx.ExecuteQueryRetryAsync();
            }

            // got context, now need to create the page
            // create page
            PnP.Core.Model.SharePoint.IPage page = spweb.LoadClientSidePage("Templates/ServiceOffering.aspx");

            // loop through the controls and pop in the data
            foreach (var cvsctl in page.Controls)
            {
                // PageText web part control
                if (cvsctl.ControlType == 4)
                {
                    //text control so replace the text
                    log.LogInformation("Replacing some text...");
                    IPageText clitxt = (IPageText)cvsctl;
                    // clitxt.Text = "Some Values and Some Kittens for Happy times";
                    // cvsctl.Text = "Some Values and Kittens here";

                    switch(clitxt.Text)
                    {
                        case "OfferingSubHeading":
                            log.LogInformation("OfferingSubHeading");
                            clitxt.Text = kpageparam.OfferingSubHeading;
                            break;
                        case "OfferingDescription":
                            log.LogInformation("OfferingDescription");
                            clitxt.Text = kpageparam.OfferingDescription;
                            break;
                        case "OfferingContact":
                            log.LogInformation("OfferingContact");
                            clitxt.Text = kpageparam.OfferingContact;
                            break;
                        case "OfferingAddlInfo":
                            log.LogInformation("OfferingAddlInfo");
                            clitxt.Text = kpageparam.OfferingAddlInfo;
                            break;
                        case "OfferingAssetsURL":
                            log.LogInformation("OfferingAssetsURL");
                            clitxt.Text = kpageparam.OfferingAssetsURL;
                            break;
                        case "OfferingContentURL":
                            log.LogInformation("OfferingContentURL");
                            clitxt.Text = kpageparam.OfferingContentURL;
                            break;
                        default:
                            log.LogInformation("Got something: " + clitxt.Text);
                            break;
                    }

                }
                else if (cvsctl.ControlType == 3)
                {
                    // this is a PageWebPage - generic container
                    // Cloc, Image, QuickLinks, etc


                }


                // page metadata
                page.PageTitle = kpageparam.OfferingTitle;

                page.Save(pagefilename);
                page.Publish();

            }

            string responseMessage = page.RepostSourceUrl;

            return new OkObjectResult(responseMessage);

        }
    }
}
