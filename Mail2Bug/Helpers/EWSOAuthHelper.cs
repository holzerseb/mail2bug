using Microsoft.Exchange.WebServices.Data;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Cache;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Mail2Bug.Helpers
{
    public class EWSOAuthHelper
    {
        public static ExchangeService OAuthConnectPost(Config.OAuthSecret oAuthCredentials, string emailAddress)
        {
            string LoginURL = String.Format("https://login.microsoftonline.com/{0}/oauth2/v2.0/token", oAuthCredentials.TenantID);

            var LogValues = new Dictionary<string, string>
            {
                { "grant_type", "client_credentials" },
                { "client_id", oAuthCredentials.ClientID },
                { "client_secret", oAuthCredentials.ClientSecret },
                { "scope", "https://outlook.office365.com/.default" }
            };
            string postData = "";
            foreach (var v in LogValues)
            {
                postData += (String.IsNullOrWhiteSpace(postData) ? "" : "&") + v.Key + "=" + v.Value;
            }
            var data = Encoding.ASCII.GetBytes(postData);

            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
               | SecurityProtocolType.Tls11
               | SecurityProtocolType.Tls12
               | SecurityProtocolType.Ssl3;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(LoginURL);
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Accept = "*/*";
            request.UserAgent = oAuthCredentials.UserAgentName;
            request.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
            request.ContentLength = data.Length;
            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            using (var response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (var reader = new StreamReader(stream))
            {
                var json = reader.ReadToEnd();
                var aToken = JObject.Parse(json)["access_token"].ToString();

                var ewsClient = new ExchangeService();
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(aToken);
                //Impersonate and include x-anchormailbox headers are required!
                ewsClient.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, emailAddress);
                ewsClient.HttpHeaders.Add("X-AnchorMailbox", emailAddress);
                ewsClient.Timeout = 60000;
                return ewsClient;
            }
        }
    }
}
