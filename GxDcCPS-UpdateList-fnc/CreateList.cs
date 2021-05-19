using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using System.Net;
using Microsoft.SharePoint.Client;
using System.Net.Http;
using Microsoft.Azure.WebJobs.Host;
using System.Linq;
using System.Configuration;

namespace GxDcCPS_UpdateList_fnc
{
    public static class CreateList
    {
        [FunctionName("CreateList")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            string siteURL = "https://gcxgce.sharepoint.com/teams/scw";
            string appOnlyId = ConfigurationManager.AppSettings["AppOnlyID"];
            string appOnlySecret = ConfigurationManager.AppSettings["AppOnlySecret"];

            // parse query parameter  
            log.Info("C# HTTP trigger function processed a request.");

            // // parse query parameter  
            string title = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "title", true) == 0)
                .Value;
            string nameFR = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "spacenamefr", true) == 0)
                .Value;
            string owner1 = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "owner1", true) == 0)
                .Value;
            string owner2 = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "owner2", true) == 0)
                .Value;
            string owner3 = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "owner3", true) == 0)
                .Value;
            string description = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "description", true) == 0)
                .Value;
            string template = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "template", true) == 0)
                .Value;
            string descriptionFr = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "descriptionFr", true) == 0)
                .Value;
            string business = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "business", true) == 0)
                .Value;
            string requester_name = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "requester_name", true) == 0)
                .Value;
            string requester_email = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "requester_email", true) == 0)
                .Value;

            // // Get request body  
            dynamic data = await req.Content.ReadAsAsync<object>();

            // // Set name to query string or body data  
            title = title ?? data?.name.title;
            nameFR = nameFR ?? data?.name.nameFR;
            owner1 = owner1 ?? data?.name.owner1;
            owner2 = owner2 ?? data?.name.owner2;
            owner3 = owner3 ?? data?.name.owner3;
            description = description ?? data?.name.description;
            template = template ?? data?.name.template;
            descriptionFr = descriptionFr ?? data?.name.descriptionFr;
            business = business ?? data?.name.business;
            requester_name = requester_name ?? data?.name.requester_name;
            requester_email = requester_email ?? data?.name.requester_email;

            log.Info("get info" + title);

            // SharePoint App only
            ClientContext ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(siteURL, appOnlyId, appOnlySecret);
            log.Info("get context");

            Web web = ctx.Web;
            List list = ctx.Web.Lists.GetByTitle("Space Requests");            
            log.Info("get list");

            ListItemCreationInformation oListItemCreationInformation = new ListItemCreationInformation();
            ListItem oItem = list.AddItem(oListItemCreationInformation);

            User userTest = web.EnsureUser(owner1);
            User userTest2 = web.EnsureUser(owner2);

            ctx.Load(userTest);
            ctx.ExecuteQuery();

            owner1 = userTest.Id.ToString() + ";#" + userTest.LoginName.ToString();
            ctx.Load(userTest2);
            ctx.ExecuteQuery();
            owner2 = userTest2.Id.ToString() + ";#" + userTest2.LoginName.ToString();
            if (owner3 != "")
            {
                User userTest3 = web.EnsureUser(owner3);
                ctx.Load(userTest3);
                ctx.ExecuteQuery();
                owner3 = userTest3.Id.ToString() + ";#" + userTest3.LoginName.ToString();
            }

            oItem["Space_x0020_Name"] = title;
            oItem["Space_x0020_Name_x0020_FR"] = nameFR;
            oItem["Owner1"] = owner1 + ";#" + owner2 + ";#" + owner3;
            oItem["Space_x0020_Description_x0020__x"] = description;
            oItem["Template_x0020_Title"] = template;
            oItem["Space_x0020_Description_x0020__x0"] = descriptionFr;
            oItem["Team_x0020_Purpose_x0020_and_x00"] = business;
            oItem["Business_x0020_Justification"] = business;
            oItem["Requester_x0020_Name"] = requester_name;
            oItem["Requester_x0020_email"] = requester_email;
            oItem["_Status"] = "Submitted";
            oItem.Update();
            ctx.ExecuteQuery();

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = string.Format(@"
            <View>
                <Query>
                    <Where>
                        <Eq>
                            <FieldRef Name='Space_x0020_Name' />
                            <Value Type='Text'>{0}</Value>
                        </Eq>
                    </Where>
                </Query>
                <ViewFields>
                    <FieldRef Name='ID'/>
                    <FieldRef Name='Space_x0020_Name'/>
                </ViewFields>
                <RowLimit>1</RowLimit>
            </View>", title);

            ListItemCollection collListItemID = list.GetItems(camlQuery);
            ctx.Load(collListItemID);
            ctx.ExecuteQuery();

            int requestID = 0;

            foreach (ListItem oListItem in collListItemID)
            {
                log.Info(oListItem["Space_x0020_Name"].ToString());
                requestID = oListItem.Id;

            }
            ListItem collListItem = list.GetItemById(requestID);
            // changes some fields 	
            collListItem["SharePoint_x0020_Site_x0020_URL"] = "https://gcxgce.sharepoint.com/teams/1000" + requestID;
            collListItem.Update();
            // executes the update of the list item on SharePoint	
            ctx.ExecuteQuery();
 
            req.CreateResponse(HttpStatusCode.InternalServerError, "Create item successfully ");
            //  }  
            return null;
        }
    }
}



