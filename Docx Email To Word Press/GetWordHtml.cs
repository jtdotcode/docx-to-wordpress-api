using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GoogleMapsApi.Entities.Common;
using GoogleMapsApi.Entities.PlacesText.Request;
using GoogleMapsApi;
using GoogleMapsApi.Entities.Geocoding.Request;
using System.Threading;
using GoogleMapsApi.Entities.Geocoding.Response;

namespace DocxEmailToWordPress
{
   
    class GetWordHtml : IDisposable
        
    {
        
        // log4net class log name
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        // Specify whether the instance is disposed. 
        private bool disposed = false;

        // The word package 
        private WordprocessingDocument package = null;

        /// <summary> 
        ///  Read Word Document 
        /// </summary> 
        /// <returns>Return Dictionary of schools and hours</returns> 
        /// 
        List<String> schoolsList = new List<string>();
        List<Double> HoursList = new List<Double>();
        Dictionary<String, Double> dic = new Dictionary<string, double>();
        String multiSchoolTitle = Properties.Settings.Default.multiSchoolTitle;
        String closingDate = string.Empty;
        String[] searchString = { "Monday,", "Tuesday,", "Wednesday,", "Thursday,", "Friday,", "Saturday,", "Sunday," };
        Boolean enableMaps = Properties.Settings.Default.enableMaps;
        public String AttachmentName { set; get; }
        Double mapCentreLng = Properties.Settings.Default.mapCentreLng;
        Double mapCentreLat = Properties.Settings.Default.mapCentreLat;
        int mapRadius = Properties.Settings.Default.mapRadius;



        public String ReadWordDocument(String filepath)
        {
            

            if (string.IsNullOrEmpty(filepath) || !File.Exists(filepath) )
            {
                throw new Exception("The file is invalid. Please select an existing file again");
            } 



            this.package = WordprocessingDocument.Open(filepath, true);


            OpenXmlElement element = package.MainDocumentPart.Document.Body;

            // make a copy to the only node we need to access 
            // this is the main table that contains the schools names and hours, it is table index 1
            OpenXmlElement ClonedNode = (OpenXmlElement)element.Elements<Table>().ElementAt(1).CloneNode(true);

            // set the elements to table row index 1, this is the middle row in the table
            OpenXmlElement row = ClonedNode.Elements<TableRow>().ElementAt(1);

            // set the elements to cell index 0 this contains the school names
            TableCell cell0 = row.Elements<TableCell>().ElementAt(0);

            // set the elements to cell index 1 this contains the school hours
            TableCell cell1 = row.Elements<TableCell>().ElementAt(1);

            // whole document
            OpenXmlElement wholeDocument = element;

            // call the GetPlainText method for cell index 0, this will create a list of the schools
            GetPlainText(cell0, 0);

            // call the GetPlainText method for cell index 1, this will create a list of the school hours
            GetPlainText(cell1, 1);

            // search whole document 
            GetPlainText(wholeDocument, 2);

            

            // 
            // if the lists don't match write to log file and exit.
            // if both list have the same amount of elements then assume the data is correct
            // and merge lists into a dictionary with school name as key can hours as value. 

            if (schoolsList.Count != HoursList.Count || HoursList.Count == 0.0 )
            {
                logger.Error("something Went Wrong the lists aren't even");
                logger.Error("Schools list = " + schoolsList.Count + " " + "Hours List = " + HoursList.Count);
                return null;


            } else
            {
                dic = schoolsList.Zip(HoursList, (k, v) => new { k, v })
              .ToDictionary(x => x.k, x => x.v );
            }

            var htmltable = GetHtmlData(dic);

            return htmltable;
        }


        /// <summary> 
        ///  Read Plain Text in selected XmlElements of word document
        ///  Adds each line of the schools name and hours to a list.
        /// </summary> 
        /// <param name="element">XmlElement in document</param>
        /// <param name="cell">Cell Element Index</param>
        ///  
        /// 
        

        private void GetPlainText(OpenXmlElement element, int cell)
        {
            
            // Enumerates each element of the cell 
            foreach (OpenXmlElement item in element.Elements())
               {

                // test switch based on elements local name
                switch (item.LocalName)
                {
                    // Text 
                    case "t":
                       
                        // check which cell, if cell index 0 add to schools list.
                        if (cell == 0)
                        {
                            schoolsList.Add(item.InnerText);
                        } else if (cell == 1)
                        {
                            // convert string to double, if the parse fails nothing added.
                            var t = Double.TryParse(item.InnerText, out double d);
                            if (t == true) {
                                HoursList.Add(d);
                            }

                        }

                        // bad solution need better options, but mah, searches document twice closing date is over written need 
                        // to fix.
                        if (cell == 2)
                        {

                            // check for closing date by checking for each day of the week.
                            foreach (var s in searchString)
                            {
                                if (item.InnerText.Contains(s))
                                {

                                    closingDate = item.InnerText;

                                    
                                }
                                
                                
                            }
                        }
                      

                        break;

                    case "cr":                          // Carriage return 
                    case "br":                          // Page break 

                        break;
                        
                    // Tab 
                    case "tab":

                        break;


                    // Paragraph 
                    case "p":
                        //call back method if paragraph is reached, continuing Enumeration .
                        GetPlainText(item, cell);

                        break;


                    default:
                        // no match, callback method and continue Enumeration
                        GetPlainText(item, cell);
                        break;

                } //end of switch

                


            } // end loop

            

        } // end get plain text

        private String GetHtmlData(Dictionary<String, Double> dictionary)
        {

            StringBuilder sbHours = new StringBuilder();
            StringBuilder sbSchools = new StringBuilder();
            

            Double totalHours = 0.0;

            foreach (var item in dictionary)
            {
               
                try
                {
                    var x = item.Value;
                    totalHours = x + totalHours;


                    var placeId = GetMapPlaceId(item.Key.ToString());
                    var schoolName = item.Key.ToString();




                    if (placeId != String.Empty && enableMaps)
                    {

                        sbSchools.Append("< a href=\"https://www.google.com/maps/dir/?api=1&origin=none&origin_place_id=" + placeId + "&travelmode=driving\"  target=\"_blank\" rel=\"noopener\">" + schoolName +  "</ a >");
                        sbSchools.Append("<br />");

                        sbHours.Append(item.Value.ToString());
                        sbHours.Append("<br />");


                    }
                    else
                    {
                        sbSchools.Append(item.Key.ToString());
                        sbSchools.Append("<br />");

                        sbHours.Append(item.Value.ToString());
                        sbHours.Append("<br />");


                    }
                    

                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                    //throw;
                }



            }

            String schools = sbSchools.ToString();
            String hours = sbHours.ToString();
            String totalHoursString = totalHours.ToString();
            String closingDateString = closingDate;

            HtmlString htmlString = new HtmlString($"<!--{AttachmentName}--><table width=\"624\" height=\"302\" border=\"1\" cellpadding=\"1\"><tr><td width=\"469\" height=\"44\" align=\"left\"><strong>School Name</strong></td><td width=\"139\" align=\"center\"><strong>Hours Per Week</strong></td></tr><tr><td height=\"217\" align=\"left\" valign=\"top\">{schools}</td><td align=\"center\" valign=\"top\">{hours}</td></tr><tr><td height=\"31\" align=\"right\"><strong>Total Hours</strong></td><td align=\"center\">{totalHoursString}</td></tr></table><p><strong>The closing date for this application is: {closingDateString} - 3:30PM</strong></p><p><small><i>Click on the School Name for Google Map directions.</i></small></p>");

            // log html string
            logger.Info(htmlString.ToString());
            

            return htmlString.ToString();

        }

        

        public String GetTitle()
        {

            if(dic.Count == 1)
            {

                var singleTitle = dic.Keys.First();

                return singleTitle.ToString();
            } else
            {
                return multiSchoolTitle;
            }

          
        }




        public String GetMapPlaceId(String schoolName)
        {


            String mapPlaceId = String.Empty;

            PlacesTextRequest request = new PlacesTextRequest() {
                ApiKey = Properties.Settings.Default.apiKey,
                Query = schoolName,
                Types = "school",
                Location = new Location(mapCentreLat, mapCentreLng),
                Radius = mapRadius



            };

            var result = GoogleMaps.PlacesText.QueryAsync(request).Result;


            switch (result.Status)
            {
                case (GoogleMapsApi.Entities.PlacesText.Response.Status.OVER_QUERY_LIMIT):
                    logger.Error("You have exceeded your Google API query limit.");
                    break;

                case (GoogleMapsApi.Entities.PlacesText.Response.Status.INVALID_REQUEST):
                    logger.Error("Google Maps API, Invalid Request");
                    break;

                case (GoogleMapsApi.Entities.PlacesText.Response.Status.OK):
                    logger.Info("Status Ok");
                    break;

                case (GoogleMapsApi.Entities.PlacesText.Response.Status.REQUEST_DENIED):
                    logger.Error("Google Maps API, Request Denied");
                    break;

                case (GoogleMapsApi.Entities.PlacesText.Response.Status.ZERO_RESULTS):
                    logger.Info("Zero results");
                    return String.Empty;
                    
                default:
                    break;
            }



            if (result.Status == GoogleMapsApi.Entities.PlacesText.Response.Status.OK)
            {
                //var mapPlaceIdSelect = from address in result.Results where address.FormattedAddress == "111111" select address.PlaceId;


                mapPlaceId = result.Results.FirstOrDefault().PlaceId;

            };

            GeocodingRequest geocoding = new GeocodingRequest()
            {
                ApiKey = Properties.Settings.Default.apiKey,
                PlaceId = mapPlaceId


            };

            var cts = new CancellationTokenSource();

            var georesult = GoogleMaps.Geocode.QueryAsync(geocoding, cts.Token).Result;


            bool isInArea = georesult.Results.Any(x => x.AddressComponents.Any(y => y.ShortName == "AU"));

            if (isInArea) { 
            isInArea = georesult.Results.Any(x => x.AddressComponents.Any(y => y.LongName == "Victoria"));
            }

            if (georesult.Status == Status.OK && isInArea)
            {
                return mapPlaceId;
            }
            else
            {

                return String.Empty;
            };

            

            
        }

        #region IDisposable interface  

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            // Protect from being called multiple times. 
            if (disposed)
            {
                return;
            }

            if (disposing)
            {
                // Clean up all managed resources. 
                if (this.package != null)
                {
                    this.package.Dispose();
                }
            }

            disposed = true;
        } 
        
        #endregion

    }

}

           

       
    

