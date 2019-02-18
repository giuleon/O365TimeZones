using System;
using System.Collections.Generic;
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
    public static class GetTimeZones
    {
        [FunctionName("GetTimeZones")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");
            var siteURL = Environment.GetEnvironmentVariable("siteURL");
            var spTimeZones = new List<SPTimeZone>();
            var result = await Task.Run(async () => {
                using (var ctx = new ClientContext(siteURL))
                {
                    var user = Environment.GetEnvironmentVariable("spAdminUser");
                    var psw = Environment.GetEnvironmentVariable("password");
                    var passWord = new SecureString();
                    foreach (char c in psw.ToCharArray()) passWord.AppendChar(c);
                    ctx.Credentials = new SharePointOnlineCredentials(user, passWord);
                    Web web = ctx.Web;
                    TimeZoneCollection tzc = ctx.Web.RegionalSettings.TimeZones;
                    ctx.Load(tzc);
                    await ctx.ExecuteQueryAsync();

                    var timeZones = tzc.ToList();

                    foreach (var item in timeZones)
                    {
                        var spTimeZone = new SPTimeZone
                        {
                            Id = item.Id,
                            Description = item.Description
                        };
                        spTimeZones.Add(spTimeZone);
                    }

                    log.Info("Time Zones Collected");
                    return spTimeZones;
                }
            });

            return req.CreateResponse(HttpStatusCode.OK, result);
        }
    }

    public class SPTimeZone
    {
        public int Id { get; set; }
        public string Description { get; set; }
    }
}
