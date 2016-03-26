using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Configuration;
using DevDefined.OAuth.Consumer;
using DevDefined.OAuth.Framework;


namespace QbAdd_inDotNetWeb
{
    public partial class OAuthManager : System.Web.UI.Page
    {
        #region <<App Properties >>
        private string requestTokenUrl = ConfigurationManager.AppSettings["RequestTokenUrl"];
        private string accessTokenUrl = ConfigurationManager.AppSettings["AccessTokenUrl"];
        private string authorizeUrl = ConfigurationManager.AppSettings["AuthorizeUrl"];
        private string oauthUrl = ConfigurationManager.AppSettings["OauthLink"];
        private string consumerKey = ConfigurationManager.AppSettings["ConsumerKey"];
        private string consumerSecret = ConfigurationManager.AppSettings["ConsumerSecret"];
        private string oauthCallbackUrl = "https://localhost:44300/OauthManager.aspx?";
        private string GrantUrl = "https://localhost:44300/OauthManager.aspx?connect=true";
 
        #endregion
        /// <summary>
        /// Page Load with initialization of properties.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString.Count > 0)
            {
                List<string> queryKeys = new List<string>(Request.QueryString.AllKeys);
                if (queryKeys.Contains("connect"))
                {
                    FireAuth();
                }
                if (queryKeys.Contains("oauth_token"))
                {
                    ReadToken();
                }
            }

        }
        /// <summary>
        /// Initiate the ouath screen.
        /// </summary>
        private void FireAuth()
        {

            IOAuthSession session = CreateSession();
            IToken requestToken = session.GetRequestToken();
            HttpContext.Current.Session["requestToken"] = requestToken;
            var authUrl = string.Format("{0}?oauth_token={1}&oauth_callback={2}", authorizeUrl, requestToken.Token, UriUtility.UrlEncode(oauthCallbackUrl + "rt=" + requestToken.Token + "&rts=" + requestToken.TokenSecret));
            HttpContext.Current.Session["oauthLink"] = authUrl;

            HttpContext.Current.Response.Redirect(authUrl);
        }
        /// <summary>
        /// Read the values from the query string.
        /// </summary>
        private void ReadToken()
        {
            HttpContext.Current.Session["oauthToken"] = Request.QueryString["oauth_token"].ToString(); ;
            HttpContext.Current.Session["oauthVerifyer"] = Request.QueryString["oauth_verifier"].ToString();
            HttpContext.Current.Session["realm"] = Request.QueryString["realmId"].ToString();
            HttpContext.Current.Session["dataSource"] = Request.QueryString["dataSource"].ToString();
            //Stored in a session for demo purposes.
            //Production applications should securely store the Access Token
            IOAuthSession clientSession = CreateSession();
            IToken accessToken = clientSession.ExchangeRequestTokenForAccessToken((IToken)HttpContext.Current.Session["requestToken"], HttpContext.Current.Session["oauthVerifyer"].ToString());
            HttpContext.Current.Session["accessToken"] = accessToken.Token;
            HttpContext.Current.Session["accessTokenSecret"] = accessToken.TokenSecret;

        }

        protected IOAuthSession CreateSession()
        {
            var consumerContext = new OAuthConsumerContext
            {
                ConsumerKey = consumerKey,
                ConsumerSecret = consumerSecret,
                SignatureMethod = SignatureMethod.HmacSha1
            };
            return new OAuthSession(consumerContext,
                                    requestTokenUrl,
                                    oauthUrl,
                                    accessTokenUrl);
        }



    }
}