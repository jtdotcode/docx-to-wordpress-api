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
        private int postTimeOut = Properties.Settings.Default.postTimeOut; 

        

        public IRestResponse PostData(String contents, String title)
        {

            var timeOut = postTimeOut;
            var client = new RestClient(wpApiUrl);
            client.Authenticator = new HttpBasicAuthenticator(wpSiteUsername, wpSitePassword);
            client.Timeout = 60;
            var apiTimeOut = client.Timeout;
            

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

            if (response.IsSuccessful)
            {
                return response;

            }else
            {
                do
                {
                   System.Threading.Thread.Sleep(1000);

                    Console.Write("Count is: " + timeOut);

                    if(apiTimeOut > 0)
                    {
                        response = client.Execute(request);
                    }

                    timeOut--;

                } while (!response.IsSuccessful && timeOut > 0);



            }


            return response;
            
        }


    }
}


