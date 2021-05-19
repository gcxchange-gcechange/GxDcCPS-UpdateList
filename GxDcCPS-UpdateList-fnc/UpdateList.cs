using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using System.Net;
using Microsoft.SharePoint.Client;
using System.Net.Http;
using Microsoft.Azure.WebJobs.Host;
using System.Linq;
using System;
using System.Configuration;

namespace GxDcCPS_UpdateList_fnc
{
    public static class UpdateList
    {
        [FunctionName("UpdateList")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            string siteURL = "https://gcxgce.sharepoint.com/teams/scw";
            string appOnlyId = ConfigurationManager.AppSettings["AppOnlyID"];
            string appOnlySecret = ConfigurationManager.AppSettings["AppOnlySecret"];


            // parse query parameter  
            log.Info("C# HTTP trigger function processed a request.");

            // // parse query parameter  
            string key = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "key", true) == 0)
                .Value;
            string comments = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "comments", true) == 0)
                .Value;
            string status = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "status", true) == 0).Value;


            // // Get request body  
            dynamic data = await req.Content.ReadAsAsync<object>();

            // // Set name to query string or body data  
            key = key ?? data?.name.key;
            comments = comments ?? data?.name.comments;
            status = status ?? data?.name.status;

            int id = Int32.Parse(key);

            // SharePoint App only
            ClientContext ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(siteURL, appOnlyId, appOnlySecret);

            Web web = ctx.Web;
            List list = ctx.Web.Lists.GetByTitle("Space Requests");

            ListItem oItem = list.GetItemById(id);

            oItem["Approved_x0020_Date"] = DateTime.Now;
            oItem["Reviewer_x0020_Comments"] = comments;
            oItem["_Status"] = status;

            oItem.Update();
            // list.ListDeleted
            ctx.ExecuteQuery();
            // return web == null  
            req.CreateResponse(HttpStatusCode.BadRequest, "Error retreiveing the list");
            req.CreateResponse(HttpStatusCode.OK, "Create item successfully ");
            //  }  
            return null;
        }
    }
}



