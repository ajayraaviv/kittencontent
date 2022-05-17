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

            KittenPageParameters kpageparam = System.Text.Json.JsonSerializer.Deserialize<KittenPageParameters>(sprawdata);
            string pagefilename = "OfferingPages/" + kpageparam.OfferingTitle + ".aspx";


            // connect to sp and create a new document - this is very very very bad
            string spurl = "";  // sharepoint URL
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
            PnP.Core.Model.SharePoint.IPage page = spweb.LoadClientSidePage("Templates/ServiceOfferingtemplate.aspx");

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

                    switch (clitxt.Text)
                    {
                        case "Solution description":
                            log.LogInformation("Solution description");
                            clitxt.Text = kpageparam.OfferingSolutionDescription;
                            break;
                        case "Client Needs":
                            log.LogInformation("Client Needs");
                            clitxt.Text = kpageparam.OfferingClientNeeds;
                            break;
                        case "Addressed Client Issue":
                            log.LogInformation("Addressed Client Issue");
                            clitxt.Text = kpageparam.OfferingAddressedClientIssue;
                            break;
                        case "Business Outcome":
                            log.LogInformation("Business Outcome");
                            clitxt.Text = kpageparam.OfferingBusinessOutcome;
                            break;
                        case "Solution Owner":
                            log.LogInformation("Solution Owner");
                            clitxt.Text = kpageparam.OfferingSolutionOwner;
                            break;
                        case "Function, Service Group, Service Line":
                            log.LogInformation("Function, Service Group, Service Line");
                            clitxt.Text = kpageparam.OfferingFunctionServiceGroupServiceLine;
                            break;
                        case "Industry, Sector, Segment":
                            log.LogInformation("Industry, Sector, Segment");
                            clitxt.Text = kpageparam.OfferingIndustrySectorSegment;
                            break;
                        case "field of play":
                            log.LogInformation("field of play");
                            clitxt.Text = kpageparam.OfferingFieldOfPlay;
                            break;
                        case "9 Levers of Value":
                            log.LogInformation("9 Levers of Value");
                            clitxt.Text = kpageparam.Offering9LeversOfValue;
                            break;
                        case "phase of delivery":
                            log.LogInformation("phase of delivery");
                            clitxt.Text = kpageparam.OfferingPhaseOfDelivery;
                            break;
                        case "ESG":
                            log.LogInformation("ESG");
                            clitxt.Text = kpageparam.OfferingESG;
                            break;
                        case "COVID-19 Relevance":
                            log.LogInformation("COVID-19 Relevance");
                            clitxt.Text = kpageparam.OfferingCovid19Relevance;
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
