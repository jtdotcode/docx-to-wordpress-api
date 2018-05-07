using Newtonsoft.Json;
using RestSharp;
using RestSharp.Authenticators;
using RestSharp.Validation;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace DocxEmailToWordPress
{
    class WordPressApi
    {
        const string BaseUrl = "***REMOVED***";

        readonly string _username;
        readonly string _password;

        public WordPressApi(string username, string password)
        {
            _username = username;
            _password = password;
        }

        public T Execute<T>(RestRequest request) where T : new()
        {
            var client = new RestClient();
            client.BaseUrl = new System.Uri(BaseUrl);
            client.Authenticator = new HttpBasicAuthenticator(_username, _password);
            
           // request.AddParameter("AccountSid", _accountSid, ParameterType.UrlSegment); // used on every request
            var response = client.Execute<T>(request);

            if (response.ErrorException != null)
            {
                const string message = "Error retrieving response.  Check inner details for more info.";
                var WpApiException = new ApplicationException(message, response.ErrorException);
                throw WpApiException;
            }
            return response.Data;
        }


        public JsonData Post(JsonData options)
        {
            Require.Argument("Title", options.Title);
            Require.Argument("Status", options.Status);
            Require.Argument("Url", options.Url);

            var request = new RestRequest(Method.POST);
            request.Resource = "Accounts/{AccountSid}/Calls";
            request.RootElement = "Calls";

            request.AddParameter("Caller", options.Caller);
            request.AddParameter("Called", options.Called);
            request.AddParameter("Url", options.Url);

            if (options.Method.HasValue) request.AddParameter("Method", options.Method);
            if (options.SendDigits.HasValue()) request.AddParameter("SendDigits", options.SendDigits);
            if (options.IfMachine.HasValue) request.AddParameter("IfMachine", options.IfMachine.Value);
            if (options.Timeout.HasValue) request.AddParameter("Timeout", options.Timeout.Value);

            return Execute<Call>(request);
        }


    }

}

