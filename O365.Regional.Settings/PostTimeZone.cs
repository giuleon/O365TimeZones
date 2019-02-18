using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;

namespace O365.Regional.Settings
{
    public static class PostTimeZone
    {
        [FunctionName("PostTimeZone")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // parse query parameter
            string siteURL = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "siteURL", true) == 0)
                .Value;

            string timeZoneToSet = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "timeZone", true) == 0)
                .Value;

            // Get request body
            dynamic data = await req.Content.ReadAsAsync<object>();
            if (siteURL == null)
            {
                siteURL = data?.siteURL;
                if (siteURL == null)
                {
                    return req.CreateResponse(HttpStatusCode.BadRequest, "Please pass the siteURL in the request body");
                }
            }
            if (timeZoneToSet == null)
            {
                timeZoneToSet = data?.timeZone;
                if (timeZoneToSet == null)
                {
                    return req.CreateResponse(HttpStatusCode.BadRequest, "Please pass the timeZone in the request body");
                }
            }

            // Get access to source site
            using (var ctx = new ClientContext(siteURL))
            {
                var User = Environment.GetEnvironmentVariable("spAdminUser");
                var Psw = Environment.GetEnvironmentVariable("password");
                //Provide count and pwd for connecting to the source
                var passWord = new SecureString();
                foreach (char c in Psw.ToCharArray()) passWord.AppendChar(c);
                ctx.Credentials = new SharePointOnlineCredentials(User, passWord);

                // Actual code for operations
                Web web = ctx.Web;
                Microsoft.SharePoint.Client.TimeZone tz = ctx.Web.RegionalSettings.TimeZone;
                TimeZoneCollection tzc = ctx.Web.RegionalSettings.TimeZones;
                ctx.Load(tz);
                ctx.Load(tzc);
                //ctx.Load(web);
                ctx.ExecuteQuery();

                var timeZone = tz;
                var timeZones = tzc.Where(x => x.Description == timeZoneToSet).FirstOrDefault();

                ctx.Web.RegionalSettings.TimeZone = timeZones;
                ctx.Web.Update();
                ctx.ExecuteQuery();
                log.Info("New regional settings set {0}", ctx.Web.RegionalSettings.TimeZone.Description);
            }

            return req.CreateResponse(HttpStatusCode.OK, "The Time Zone is correctly configured");
        }
    }
}
