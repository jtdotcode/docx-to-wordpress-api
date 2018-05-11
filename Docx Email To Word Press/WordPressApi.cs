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
        private const String status = "draft";
        private const String title = "test-wordpress-api";
        private String[] categorie = { "43" };
        private static string fileName = @"C:\temp\test.docx";
        GetWordHtml getWordHtml;
        
        private const String username = "***REMOVED***";
        private const String password = "***REMOVED***";

       
        public Boolean PostData(String content)
        {
            var client = new RestClient("***REMOVED***");
            client.Authenticator = new HttpBasicAuthenticator(username, password);
            var jsonData = new JsonData() { Status = status,
            Title = title,
            Categories = categorie,
            Content = content
         };
            

            

            var request = new RestRequest(Method.POST);
            request.AddHeader("Postman-Token", "21d4bc9a-f980-404f-9667-1205061e76a0");
            request.AddHeader("Cache-Control", "no-cache");
            request.AddHeader("Authorization", "Basic am9obkBpb2l0LmNvbS5hdTpIYXNoQ2F0VGhpczE3Iw==");
            request.AddHeader("Content-Type", "application/json");
           
            IRestResponse response = client.Execute(request);


            return true;
        }


    }
}


