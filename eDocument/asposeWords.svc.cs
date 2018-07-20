using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.Xml;
using System.Data;
using System.Diagnostics;
using System.Configuration;
using asposeWR;
using System.IO;
using Aspose.Words.MailMerging;

namespace eDocument
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "asposeWords" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select asposeWords.svc or asposeWords.svc.cs at the Solution Explorer and start debugging.
    public class asposeWords : IasposeWords
    {

        //fills the Word template with XML Data
        public string fillWordTemplate(string ApplicationName, string InstanceCode, string TemplateName, string XMLData, string RemoveUnusedFields = "no", string RemoveUnusedRegions = "no", string RemoveEmptyParagraphs = "no")
        {

            // directory and URL of the final document storage
            string StorageDir = ConfigurationManager.AppSettings["DefaultStorageDir"];
            string StorageURL = ConfigurationManager.AppSettings["DefaultStorageURL"];

            // directories where filler workfiles will be stored
            string FillerInputDir = ConfigurationManager.AppSettings["DefaultInputDir"];
            string FillerOutputDir = ConfigurationManager.AppSettings["DefaultOutputDir"];
            string FillerOutputDocumentName = ConfigurationManager.AppSettings["DefaultOutputFileName"];
            string FillerOutputDocumentNameWithPath = FillerOutputDir + FillerOutputDocumentName;

            // generate the final generated document filename
            string documentFileName = RandomString(16) + ".docx";
            string documentFileNameWithPath = StorageDir + "/" + documentFileName;
            string documentURL = StorageURL + "/" + documentFileName;
            
            XMLData = RemoveBom(XMLData);
            XMLData = "<" + XMLData;            

            //first parse the XMLData
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(XMLData);

            // get the number of field nodes and declare arrays
            XmlNodeList fieldCounter = xml.GetElementsByTagName("fieldName");

            String[] names = new String[fieldCounter.Count];
            Object[] values = new Object[fieldCounter.Count];

            //get field names
            XmlNodeList elemList = xml.GetElementsByTagName("fieldName");
            for (int i = 0; i < elemList.Count; i++)
            {
                names[i] = elemList[i].InnerXml;
            }

            //get field values
            XmlNodeList elemList2 = xml.GetElementsByTagName("fieldValue");
            for (int i = 0; i < elemList2.Count; i++)
            {
                values[i] = elemList2[i].InnerXml;
            }

            LogMessageToFile("---------- START ----------");

            LogMessageToFile("ApplicationName: " + ApplicationName);
            LogMessageToFile("TemplateDir: " + FillerInputDir);
            LogMessageToFile("TemplateName: " + TemplateName);
            LogMessageToFile("OutputDir: " + FillerOutputDir);
            LogMessageToFile("OutputDocumentName: " + FillerOutputDocumentName);
            LogMessageToFile("---------------------------");

            // initialise wrapper
            asposeWordsWrapper AWW = new asposeWordsWrapper(ApplicationName, FillerInputDir, TemplateName, FillerOutputDir, FillerOutputDocumentName);

            LogMessageToFile("Wrapper inited");

            //loop over data tables

            //fetch data table from XML
            var dataTableNodes = xml.SelectNodes(@"//datatables/datatable");
            foreach (XmlNode item in dataTableNodes)
            {

                // fetch datatable name
                string tableName = item.SelectSingleNode("./tableName").InnerText;

                // fetch datatable data
                var tableDataNodes = item.SelectSingleNode("./tableData/" + tableName);

                // convert fethchet XML node into dataset                
                var dataTableDS = ConvertXmlNodeToDataSet(tableDataNodes);

                using (TextWriter oWriter = new StreamWriter(@"d:/temp/" + tableName + "ds.xml"))
                {
                    dataTableDS.WriteXml(oWriter);
                    oWriter.Flush();
                }

                AWW.ExecuteRegions(dataTableDS);

                LogMessageToFile("Region " + tableName + " executed");
            }
             
            //loop over queries
            var queryNodes = xml.SelectNodes(@"//queries/query");
            foreach (XmlNode item in queryNodes)
            {


                string sqlStatement = item.SelectSingleNode("./sql").InnerText;
                string tableName = item.SelectSingleNode("./tableName").InnerText;                
                AWW.ExecuteRegionsSQL(sqlStatement, tableName);

                LogMessageToFile("Query " + tableName + " executed");
            }

            // setup removal only for the last merge operation
            if (RemoveUnusedFields == "yes")
            {
                AWW.doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;

                if (RemoveUnusedFields == "yes")
                {
                    AWW.doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveUnusedFields;
                }

                if (RemoveEmptyParagraphs == "yes")
                {
                    AWW.doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveEmptyParagraphs;
                }
            }

            // execute simple mail merge as a last merge operation
            AWW.Execute(names, values);

            LogMessageToFile("Execute simple values runned");

            // remove unused fields other way
            AWW.doc.MailMerge.DeleteFields();

            LogMessageToFile("Removes executed");

            AWW.Save();

            LogMessageToFile("Document saved");
            LogMessageToFile(FillerOutputDocumentNameWithPath);
            LogMessageToFile(documentFileNameWithPath);
            LogMessageToFile("===========================");

            //copy documents to storage from which it will be downloaded
            System.IO.File.Move(FillerOutputDocumentNameWithPath, documentFileNameWithPath);
            System.IO.File.Move(FillerOutputDocumentNameWithPath + ".pdf", documentFileNameWithPath + ".pdf");
            System.IO.File.Move(FillerOutputDocumentNameWithPath + ".html", documentFileNameWithPath + ".html");

            return documentURL;
        }

        public string uploadFile(Stream fileContents)
        {
            FileStream targetStream = null;


            string FillerInputDir = ConfigurationManager.AppSettings["DefaultInputDir"];
            string FillerTemplateName = RandomString(16) + ".docx";


            string filePath = Path.Combine(FillerInputDir, FillerTemplateName);

            using (targetStream = new FileStream(filePath, FileMode.Create,
                                  FileAccess.Write, FileShare.None))
            {
                //read from the input stream in 65000 byte chunks

                const int bufferLen = 65000;
                byte[] buffer = new byte[bufferLen];
                int count = 0;
                while ((count = fileContents.Read(buffer, 0, bufferLen)) > 0)
                {
                    // save to output stream
                    targetStream.Write(buffer, 0, count);
                }
                targetStream.Close();
                fileContents.Close();
            }

            return FillerTemplateName;
        }

        public string getUserName()
        {
            return "Authenticated UserName" + ServiceSecurityContext.Anonymous.PrimaryIdentity;
        }

        private static string RemoveBom(string p)
        {
            string BOMMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
            if (p.StartsWith(BOMMarkUtf8))
                p = p.Remove(0, BOMMarkUtf8.Length);
            return p.Replace("\0", "");
        }

        private static DataSet ConvertXmlNodeToDataSet(XmlNode xmlnodeinput)
        {
            //declaring data set object
            DataSet dataset = null;
            if (xmlnodeinput != null)
            {
                XmlTextReader xtr = new XmlTextReader(xmlnodeinput.OuterXml, XmlNodeType.Element, null);
                dataset = new DataSet();
                dataset.ReadXml(xtr,XmlReadMode.Auto);
            }

            return dataset;
        }

        private void LogMessageToFile(string msg)
        {

            string LogDir = ConfigurationManager.AppSettings["DefaultLogDir"];
            string LogFile = LogDir + "/eDocumentLog.txt";

            System.IO.StreamWriter sw = System.IO.File.AppendText(LogFile);
            try
            {
                string logLine = System.String.Format(
                    "{0:G}: {1}", System.DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        private static Random random = new Random();
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }


    }

}
