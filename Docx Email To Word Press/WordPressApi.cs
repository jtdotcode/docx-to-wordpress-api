using RestSharp;
using RestSharp.Authenticators;
using System;

namespace DocxEmailToWordPress
{
    class WordPressApi
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private String wpPostStatus = Properties.Settings.Default.wpPostStatus;

        private static int wpPostCategorie = Properties.Settings.Default.wpPostCategorie;
        private String[] wpPostCategories = { wpPostCategorie.ToString() };
        private String wpApiUrl = Properties.Settings.Default.wpApiUrl;

        private String wpSiteUsername = Properties.Settings.Default.wpSiteUsername;
        private String wpSitePassword = Properties.Settings.Default.wpSitePassword;
        private int apiPostTimeOut = Properties.Settings.Default.apiPostTimeOut;

        public bool testApi()
        {

            var client = new RestClient(wpApiUrl);
            client.Authenticator = new HttpBasicAuthenticator(wpSiteUsername, wpSitePassword);
            var request = new RestRequest(Method.GET);



            var connect = client.Execute(request);



            return connect.IsSuccessful;
        }
        

        public IRestResponse PostData(String contents, String title)
        {

            var timeOut = apiPostTimeOut;
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

            // blocking needs to be cleaned up, could use something from the library 

            if (response.IsSuccessful)
            {
                return response;

            }else
            {
                logger.Error("unable to Connect to API Trying again in: " + timeOut);

                do
                {
                   System.Threading.Thread.Sleep(1000);

                    
                    

                    if(timeOut == 0)
                    {
                        response = client.Execute(request);

                        if (!response.IsSuccessful)
                        {
                            logger.Error("Timeout reached unable to connect");
                            
                        }
                        
                    }

                    timeOut--;

                } while (!response.IsSuccessful && timeOut > 0);



            }


            return response;
            
        }


    }
}


