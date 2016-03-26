using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.SessionState;
using System.Web.Routing;
using System.Web;
using System.Web.Http.WebHost;

namespace QbAdd_inDotNetWeb
{
    public static class WebApiConfig
    {

        public static void Register(HttpConfiguration config)
        {
            config.MapHttpAttributeRoutes();

            /*config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );*/

            config.Routes.MapHttpRoute(
                name: "ActionApi",
                routeTemplate: "api/{action}/{id}",
                defaults: new { id = RouteParameter.Optional, controller = "QuickBooks"}
            );

            //RouteTable.Routes.MapHttpRoute(name: "DefaultApi", routeTemplate: "api/{controller}/{id}", defaults: new { id = RouteParameter.Optional }).RouteHandler = new SessionStateRouteHandler();

        }
    }
}
