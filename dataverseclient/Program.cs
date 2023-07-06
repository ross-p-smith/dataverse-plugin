using dotenv.net;
using dotenv.net.Utilities;
using Microsoft.Identity.Client;  // Microsoft Authentication Library (MSAL)
using Microsoft.Identity.Web;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;

namespace PowerApps.Samples
{
   /// <summary>
   /// Demonstrates Azure authentication and execution of a Dataverse WI've decided eb API function.
   /// </summary>
   class Program
   {
      static async Task Main()
      {
            // Load .env file but traverse up 4 directories from the executing directory.
            DotEnv.Load(options: new DotEnvOptions(probeForEnv: true, probeLevelsToSearch: 4));
            
            string resource = EnvReader.GetStringValue("CRM_URL");

            // Azure Active Directory app registration shared by all Power App samples.
            var clientId = EnvReader.GetStringValue("CLIENT_ID");
            var redirectUri = "http://localhost"; // Loopback for the interactive login.

            // For your custom apps, you will need to register them with Azure AD yourself.
            // See https://docs.microsoft.com/powerapps/developer/data-platform/walkthrough-register-app-azure-active-directory

            #region Authentication
            AuthenticationResult authResult;
            var authBuilder = PublicClientApplicationBuilder.Create(clientId)
                            .WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs)
                            .WithRedirectUri(redirectUri)
                            .Build();

            var scope = resource + "/.default";
            //var scope = "https://graph.microsoft.com/User.Read";
            string[] scopes = { scope };

            try
            {
                var accounts = await authBuilder.GetAccountsAsync();
                // Try to acquire an access token from the cache. If device code is required, Exception will be thrown.
                authResult = await authBuilder.AcquireTokenSilent(scopes, accounts.FirstOrDefault()).ExecuteAsync();
            }
            catch (MsalUiRequiredException)
            {
               authResult = await authBuilder.AcquireTokenInteractive(scopes).ExecuteAsync();
            }

            #endregion Authentication

            #region Client configuration
            var client = new HttpClient
            {
               // See https://docs.microsoft.com/powerapps/developer/data-platform/webapi/compose-http-requests-handle-errors#web-api-url-and-versions
               BaseAddress = new Uri(resource + "/api/data/v9.2/"),
               Timeout = new TimeSpan(0, 2, 0)    // Standard two minute timeout on web service calls.
            };

            // Default headers for each Web API call.
            // See https://docs.microsoft.com/powerapps/developer/data-platform/webapi/compose-http-requests-handle-errors#http-headers
            HttpRequestHeaders headers = client.DefaultRequestHeaders;
            headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            headers.Add("OData-MaxVersion", "4.0");
            headers.Add("OData-Version", "4.0");
            headers.Accept.Add(
               new MediaTypeWithQualityHeaderValue("application/json"));
            #endregion Client configuration

            #region Web API call

            // Invoke the Web API 'WhoAmI' unbound function.
            // See https://docs.microsoft.com/powerapps/developer/data-platform/webapi/compose-http-requests-handle-errors
            // See https://docs.microsoft.com/powerapps/developer/data-platform/webapi/use-web-api-functions#unbound-functions
            var response = await client.GetAsync("WhoAmI");

            if (response.IsSuccessStatusCode)
            {
               // Parse the JSON formatted service response (WhoAmIResponse) to obtain the user ID value.
               // See https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/whoamiresponse
               Guid userId = new Guid();

               string jsonContent = await response.Content.ReadAsStringAsync();

               // Using System.Text.Json
               using (JsonDocument doc = JsonDocument.Parse(jsonContent))
               {
                  JsonElement root = doc.RootElement;
                  JsonElement userIdElement = root.GetProperty("UserId");
                  userId = userIdElement.GetGuid();
               }

               // Alternate code, but requires that the WhoAmIResponse class be defined (see below).
               // WhoAmIResponse whoAmIresponse = JsonSerializer.Deserialize<WhoAmIResponse>(jsonContent);
               // userId = whoAmIresponse.UserId;

               Console.WriteLine($"Your user ID is {userId}");
            }
            else
            {
               Console.WriteLine("Web API call failed");
               Console.WriteLine("Reason: " + response.ReasonPhrase);
            }
            #endregion Web API call
      }
   }

   /// <summary>
   /// WhoAmIResponse class definition 
   /// </summary>
   /// <remarks>To be used for JSON deserialization.</remarks>
   /// <see cref="https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/whoamiresponse"/>
   public class WhoAmIResponse
   {
      public Guid BusinessUnitId { get; set; }
      public Guid UserId { get; set; }
      public Guid OrganizationId { get; set; }
   }
}