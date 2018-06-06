using RestSharp;
using RestSharp.Authenticators;
using System;

namespace DocxEmailToWordPress
{
    class WordPressApi
    {
        private const String status = "draft";
        
        private String[] categorie = { "43" };
        private const String WpApiUrl = "***REMOVED***";

        private const String username = "poster";
        private const String password = "***REMOVED***";

       
        public IRestResponse PostData(String contents, String title)
        {
            

            var client = new RestClient(WpApiUrl);
            client.Authenticator = new HttpBasicAuthenticator(username, password);
           

            var request = new RestRequest(Method.POST);
            request.AddHeader("Content-Type", "application/json");
            request.AddJsonBody(new JsonData
            {
                status = status,
                title = title,
                categories = categorie,
                content = contents,
                excerpt = "null!"
            });
           
            IRestResponse response = client.Execute(request);




            return response;
        }


    }
}


