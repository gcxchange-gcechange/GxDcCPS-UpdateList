using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Azure.KeyVault;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.SharePoint.Client;
using System.Configuration;
using Microsoft.Extensions.Configuration;

namespace GxDcCPS_UpdateList2_fnc
{
    public static class UpdateList
    {
        [FunctionName("UpdateList")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,

      ExecutionContext context,
      ILogger log)
        {
            var config = new ConfigurationBuilder()
              .SetBasePath(context.FunctionAppDirectory)
              // This gives you access to your application settings in your local development environment
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              // This is what actually gets you the application settings in Azure
              .AddEnvironmentVariables()
              .Build();

            log.LogInformation("C# HTTP trigger function processed a request.");

            string KeyVault_Name = config["KeyVault_Name"];
            string Cert_Name = config["Cert_Name"];
            string appOnlyId = config["AppOnlyID"];
            string tenant_URL = config["Tenant_URL"];
            string siteURL = "https://" + tenant_URL + ".sharepoint.com/teams/scw";

            // // parse query parameter  
            string key = req.Query["key"];
            string comments = req.Query["comments"];
            string status = req.Query["status"];


            // // Get request body  
            //dynamic data = await req.Content.ReadAsAsync<object>();
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            dynamic data = JsonConvert.DeserializeObject(requestBody);

            // // Set name to query string or body data  
            key = key ?? data?.name.key;
            comments = comments ?? data?.name.comments;
            status = status ?? data?.name.status;

            log.LogInformation("get info" + key);
            int id = Int32.Parse(key);

            using (var cc = new OfficeDevPnP.Core.AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(siteURL, appOnlyId, tenant_URL+".onmicrosoft.com", KeyVaultAccess.GetKeyVaultCertificate(KeyVault_Name, Cert_Name)))

            {
                Web web = cc.Web;
            List list = cc.Web.Lists.GetByTitle("Space Requests");

            ListItem oItem = list.GetItemById(id);

            oItem["Approved_x0020_Date"] = DateTime.Now;
            oItem["Reviewer_x0020_Comments"] = comments;
            oItem["_Status"] = status;

            oItem.Update();
            cc.ExecuteQuery();
                string responseMessage = "Create item successfully ";

                return new OkObjectResult(responseMessage);
            }
        }

        class KeyVaultAccess
        {

            internal static X509Certificate2 GetKeyVaultCertificate(string keyvaultName, string name)
            {
                var serviceTokenProvider = new AzureServiceTokenProvider();
                var keyVaultClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(serviceTokenProvider.KeyVaultTokenCallback));

                // Getting the certificate
                var secret = keyVaultClient.GetSecretAsync("https://" + keyvaultName + ".vault.azure.net/", name);

                // Returning the certificate
                return new X509Certificate2(Convert.FromBase64String(secret.Result.Value));
            }
        }
    }
}
