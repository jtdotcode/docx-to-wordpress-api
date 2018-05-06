using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEmailToWordPress
{
    class GetWordPlainText : IDisposable
    {
        // Specify whether the instance is disposed. 
        private bool disposed = false;
        
        // The word package 
        private WordprocessingDocument package = null;

        /// <summary> 
        ///  Get the file name 
        /// </summary> 
        /// test path of document 
       

        private static string fileName = @"C:\temp\test.docx";
        
        /// <summary> 
        ///  Initialize the WordPlainTextManager instance 
        /// </summary> 
        /// <param name="filepath"></param> 
        public GetWordPlainText(string filepath)
        {

            if (string.IsNullOrEmpty(filepath) || !File.Exists(filepath))
            {
                throw new Exception("The file is invalid. Please select an existing file again");
            }

            this.package = WordprocessingDocument.Open(filepath, true);
        }

        /// <summary> 
        ///  Read Word Document 
        /// </summary> 
        /// <returns>Return Dictionary of schools and hours</returns> 
        /// 
        List<String> schoolsList = new List<string>();
        List<Double> HoursList = new List<Double>();
        Dictionary<String, Double> dic = new Dictionary<string, double>();


        public Dictionary<String, double> ReadWordDocument()
        {
            
            OpenXmlElement element = package.MainDocumentPart.Document.Body;

            // make a copy to the only node we need to access 
            // this is the main table that contains the schools names and hours, it is table index 1
            OpenXmlElement ClonedNode = (OpenXmlElement)element.Elements<Table>().ElementAt(1).CloneNode(true);

            // set the elemnts to table row index 1, this is the middle row in the table
            OpenXmlElement row = ClonedNode.Elements<TableRow>().ElementAt(1);

            // set the elements to cell index 0 this contains the school names
            TableCell cell0 = row.Elements<TableCell>().ElementAt(0);

            // set the elements to cell index 1 this contains the school hours
            TableCell cell1 = row.Elements<TableCell>().ElementAt(1);


            // call the GetPlainText method for cell index 0, this will create a list of the schools
            GetPlainText(cell0, 0);

            // call the GetPlainText method for cell index 1, this will create a list of the school hours
            GetPlainText(cell1, 1);

            // 
            // if the lists dont match write to to logfile and exit.
            // if both list have the same amount of elements then assume the data is correct
            // and merge lists into a dictionary with school name as key can hours as value. 

            if (schoolsList.Count != HoursList.Count)
            {
                Console.Write("something Went Wrong the lists arent even");

            } else
            {
                dic = schoolsList.Zip(HoursList, (k, v) => new { k, v })
              .ToDictionary(x => x.k, x => x.v );
            }


            return dic;
        }


        /// <summary> 
        ///  Read Plain Text in selected XmlElements of word document
        ///  Adds each line of the schools name and hours to a list.
        /// </summary> 
        /// <param name="element">XmlElement in document</param>
        /// <param name="cell">Cell Element Index</param>
        ///  
        /// 
        

        public void GetPlainText(OpenXmlElement element, int cell)
        {
            
            // Emumerates each element of the cell 
            foreach (OpenXmlElement item in element.Elements())
               {
                
                // test switch based on elements local name
                switch (item.LocalName)
                {
                    // Text 
                    case "t":
                        Console.WriteLine(item.InnerText);
                        // check which cell, if cell index 0 add to schools list.
                        if (cell == 0)
                        {
                            schoolsList.Add(item.InnerText);
                        } else
                        {
                            // convert string to double, if the parse fails nothing added.
                         var t = Double.TryParse(item.InnerText, out double d);
                            if(t == true) {
                                HoursList.Add(d);
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
                        //call back method if paragraph is reached, continuing Emumeration .
                        GetPlainText(item, cell);

                        break;




                    default:
                        // no match, callback method and continue Emumeration
                        GetPlainText(item, cell);
                        break;

                } //end of switch

                


            } // end loop

            

        } // end get plain text

        

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

           

       
    

