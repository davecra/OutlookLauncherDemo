using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace OutlookLauncherDemoWeb.Controllers
{
    public class DefaultController : ApiController
    {
        /// <summary>
        /// Service Request - incoming
        /// </summary>
        public class WebServiceRequest
        {
            public string Command { get; set; }
            public string[] Params { get; set; }
        }

        /// <summary>
        /// Service Response - outgoing
        /// </summary>
        public class WebServiceResponse
        {
            public string Status { get; set; }
            public string Message { get; set; }
        }

        [HttpPost()]
        public WebServiceResponse Values(WebServiceRequest request)
        {
            WebServiceResponse response = null;
            switch (request.Command)
            {
                case "GetSprints":
                    return getSprints();
                case "GetConfigs":
                    return getConfigs();
            }

            response = new WebServiceResponse();
            response.Message = "Unknown command";
            return response;
        }

        private WebServiceResponse getConfigs()
        {
            WebServiceResponse response = new WebServiceResponse();
            response.Message = "Config1,Config2,Config3,Config4,Config5,Config6,Config7,Config8";
            return response;
        }

        private WebServiceResponse getSprints()
        {
            WebServiceResponse response = new WebServiceResponse();
            response.Message = "sprint1,sprint2,sprint3,sprint4";
            return response;
        }
    }
}
