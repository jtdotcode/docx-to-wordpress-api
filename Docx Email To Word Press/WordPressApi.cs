using RestSharp;
using RestSharp.Authenticators;
using System;

namespace DocxEmailToWordPress
{
    class WordPressApi
    {
        private String wpPostStatus = Properties.Settings.Default.wpPostStatus;

        private static String wpPostCategorie = Properties.Settings.Default.wpPostCategorie;
        private String[] wpPostCategories = { wpPostCategorie };
        private String wpApiUrl = Properties.Settings.Default.wpApiUrl;

        private String wpSiteUsername = Properties.Settings.Default.wpSiteUsername;
        private String wpSitePassword = Properties.Settings.Default.wpSitePassword;



        public IRestResponse PostData(String contents, String title)
        {
            

            var client = new RestClient(wpApiUrl);
            client.Authenticator = new HttpBasicAuthenticator(wpSiteUsername, wpSitePassword);
           

            var request = new RestRequest(Method.POST);
            request.AddHeader("Content-Type", "application/json");
            request.AddJsonBody(new JsonData
            {
                status = wpPostStatus,
                title = title,
                categories = wpPostCategories,
                content = contents,
                excerpt = "null!"
            });
           
            IRestResponse response = client.Execute(request);




            return response;
        }


    }
}


