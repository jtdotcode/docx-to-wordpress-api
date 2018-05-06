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
        // private string FileName = string.Empty;

        private static string fileName = @"C:\temp\test.docx";


        private GetWordPlainText getWordPlainText;

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
        /// <returns>Plain Text in document </returns> 
        public string ReadWordDocument()
        {
            StringBuilder sb = new StringBuilder();
            OpenXmlElement element = package.MainDocumentPart.Document.Body;
            OpenXmlElement ClonedNode = (OpenXmlElement)element.Elements<Table>().ElementAt(1).CloneNode(true);

            OpenXmlElement row = ClonedNode.Elements<TableRow>().ElementAt(1);
            TableCell cell = row.Elements<TableCell>().ElementAt(1);

            if (element == null)
            {
                return string.Empty;
            }
            

            sb.Append(GetPlainText(cell));
            return sb.ToString();
        }


        /// <summary> 
        ///  Read Plain Text in all XmlElements of word document 
        /// </summary> 
        /// <param name="element">XmlElement in document</param> 
        /// <returns>Plain Text in XmlElement</returns> 
        /// 
        StringBuilder PlainTextInWord = new StringBuilder();
        List<String> schoolsList = new List<string>();

        public string GetPlainText(OpenXmlElement element)
        {
            
            foreach (OpenXmlElement item in element.Elements())
               {

                //  PlainTextInWord.Append(GetElement(item));
                var returnedElement = GetElement(item);

                if (returnedElement != null)
                {
                    schoolsList.Add(returnedElement);
                }

              

            } // end for loop
             

            return schoolsList.

        }

        /// <summary> 
        ///  Save the text to text file 
        /// </summary> 

        public String GetElement(OpenXmlElement element)
        {
            switch (element.LocalName)
            {
                // Text 
                case "t":
                    Console.WriteLine(element.InnerText);
                    return element.InnerText;
                
                case "cr":                          // Carriage return 
                case "br":                          // Page break 
                    
                    break;


                // Tab 
                case "tab":
                    
                    break;


                // Paragraph 
                case "p":
                  GetPlainText(element);
                    
                    break;

                
                    

                default:
                    GetPlainText(element);
                    break;

            } //end of switch

            return null;


        }
        


        public void saveToTxtTest()
        {

            /// <summary> 
            /// Get Plain Text from Word file 
          
           
                try
                {
                    
                    saveAsTxtFile(ReadWordDocument());
                
                }
                catch (Exception ex)
                {

                Console.WriteLine(ex.Message + "error");

                
                }
                finally
                {
                    if (getWordPlainText != null)
                    {
                        getWordPlainText.Dispose();
                    }
                }
            }

       
    

    public void saveAsTxtFile(string txt) {

            string path = @"c:\temp\MyTest.txt";
            string createText = txt;

            // This text is added only once to the file.
            if (!File.Exists(path))
            {
                // Create a file to write to.
                
                File.WriteAllText(path, createText);
            } else {

                // delete old file and create new one.

                File.Delete(path);
                File.WriteAllText(path, createText);


            }

           

            // Open the file to read from.
            string readText = File.ReadAllText(path);
            Console.WriteLine(readText);



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

           

       
    

