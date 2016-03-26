using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web;
using System.Configuration;

using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.DataService;
using Intuit.Ipp.Security;
using DevDefined.OAuth.Framework;


namespace QbAdd_inDotNetWeb
{
   
    public class QuickBooksController : ApiController
    {
        public class AccessToken
        {
            public string token;
            public string secret;
        }

        private String realmId, accessToken, accessTokenSecret, consumerKey, consumerSecret;

        [HttpGet]
        public IEnumerable<Purchase> GetPurchases(int n)
        {
            consumerKey = ConfigurationManager.AppSettings["consumerKey"].ToString();
            consumerSecret = ConfigurationManager.AppSettings["consumerSecret"].ToString();

            realmId = ConfigurationManager.AppSettings["RealmId"].ToString();
            accessToken = HttpContext.Current.Session["accessToken"].ToString();
            accessTokenSecret = HttpContext.Current.Session["accessTokenSecret"].ToString();

            IntuitServicesType intuitServicesType = IntuitServicesType.QBO;
            OAuthRequestValidator oauthValidator = new OAuthRequestValidator(accessToken, accessTokenSecret, consumerKey, consumerSecret);
            ServiceContext context = new ServiceContext(realmId, intuitServicesType, oauthValidator);
            context.IppConfiguration.BaseUrl.Qbo = ConfigurationManager.AppSettings["ServiceContext.BaseUrl.Qbo"].ToString();

            DataService dataService = new DataService(context);
            List<Purchase> purchases = dataService.FindAll(new Purchase(), 1, n).ToList();
            return purchases;
        }

        [HttpGet]
        public HttpResponseMessage SetToken(string t, string s)
        {
            HttpContext.Current.Session["accessToken"] = t;
            HttpContext.Current.Session["accessTokenSecret"] = s;

            return Request.CreateResponse(HttpStatusCode.OK, "Success");     
        }

        public HttpResponseMessage GetToken()
        {
            HttpStatusCode code = HttpStatusCode.NotFound;
            string message = "NotFound";
            if (null != HttpContext.Current.Session["accessToken"])
            {
                code = HttpStatusCode.OK;
                message = "Success";
            }

            return Request.CreateResponse(code, message);
        }

        [HttpGet]
        public HttpResponseMessage ClearToken()
        {
            HttpContext.Current.Session["accessToken"] = "";
            HttpContext.Current.Session["accessTokenSecret"] = "";

            return Request.CreateResponse(HttpStatusCode.OK, "Success");
        }
    }
}