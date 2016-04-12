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


namespace QbAdd_inDotNetWeb
{
   
    public class QuickBooksController : ApiController
    {

        private String realmId, accessToken, accessTokenSecret, consumerKey, consumerSecret;

        [HttpGet]
        public IEnumerable<Purchase> GetExpenses(int n)
        {
            consumerKey = ConfigurationManager.AppSettings["consumerKey"].ToString();
            consumerSecret = ConfigurationManager.AppSettings["consumerSecret"].ToString();

            realmId = HttpContext.Current.Session["realm"].ToString();
            accessToken = HttpContext.Current.Session["accessToken"].ToString();
            accessTokenSecret = HttpContext.Current.Session["accessTokenSecret"].ToString();

            IntuitServicesType intuitServicesType = IntuitServicesType.QBO;
            OAuthRequestValidator oauthValidator = new OAuthRequestValidator(accessToken, accessTokenSecret, consumerKey, consumerSecret);
            ServiceContext context = new ServiceContext(realmId, intuitServicesType, oauthValidator);
            context.IppConfiguration.BaseUrl.Qbo = ConfigurationManager.AppSettings["ServiceContext.BaseUrl.Qbo"].ToString();

            DataService dataService = new DataService(context);
            List<Purchase> expenses = dataService.FindAll(new Purchase(), 1, n).ToList();
            return expenses;
        }

        [HttpGet]
        public HttpResponseMessage SetToken(string token, string secret, string realm)
        {
            HttpContext.Current.Session["accessToken"] = token;
            HttpContext.Current.Session["accessTokenSecret"] = secret;
            HttpContext.Current.Session["realm"] = realm;

            return Request.CreateResponse(HttpStatusCode.OK, "Success");     
        }

        public HttpResponseMessage GetToken()
        {
            HttpStatusCode code = HttpStatusCode.NotFound;
            string message = "NotFound";
            if (null != HttpContext.Current.Session["accessToken"] && "" != HttpContext.Current.Session["accessToken"].ToString())
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