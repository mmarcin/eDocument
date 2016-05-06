using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.Xml;
using System.Diagnostics;
using System.Configuration;
using asposeWR;


namespace eDocument
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "asposeWords" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select asposeWords.svc or asposeWords.svc.cs at the Solution Explorer and start debugging.
    public class asposeWords : IasposeWords
    {

        //fills the Word template with XML Data
        public string fillWordTemplate(string ApplicationName,string TemplateName, string OutputDocumentName, string XMLData, string TemplateDir = "", string OutputDir = "", string RemoveUnusedFields = "no", string RemoveUnusedRegions = "no", string RemoveEmptyParagraphs = "no")
        {
            
            if (TemplateDir == "")
            {
                TemplateDir = ConfigurationManager.AppSettings["DefaultTemplateDir"];
            }

            if (OutputDir == "")
            {
                OutputDir = ConfigurationManager.AppSettings["DefaultOutputDir"];
            }

            System.IO.File.WriteAllText("c:/temp/XMLData.xml", XMLData);
            XMLData = RemoveBom(XMLData);
            XMLData = "<" + XMLData;
            System.IO.File.WriteAllText("c:/temp/XMLData_NoBOM.xml", XMLData);


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

            // initialise wrapper
            asposeWordsWrapper AWW = new asposeWordsWrapper(ApplicationName, TemplateDir, TemplateName, OutputDir, OutputDocumentName);

            // remove unused fields
            if (RemoveUnusedFields == "yes")
            {
                AWW.RemoveUnusedFields();
            }

            // remove unused regions
            if (RemoveUnusedRegions == "yes")
            {
                AWW.RemoveUnusedRegions();
            }

            // remove empty paragraphs
            if (RemoveUnusedRegions == "yes")
            {
                AWW.RemoveEmptyParagraphs();
            }

            // execute simple mail merge
            AWW.Execute(names, values);

            //loop over queries
            var bookNodes = xml.SelectNodes(@"//queries/query");
            foreach (XmlNode item in bookNodes)
            {
                string sqlStatement = item.SelectSingleNode("./sql").InnerText;
                string tableName = item.SelectSingleNode("./tableName").InnerText;
                Console.WriteLine("title {0} price: {1}", sqlStatement, tableName); //just for demo
                AWW.ExecuteRegions(sqlStatement, tableName);
            }

            AWW.Save();
            return "Success";
        }


        private static string RemoveBom(string p)
        {
            string BOMMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
            if (p.StartsWith(BOMMarkUtf8))
                p = p.Remove(0, BOMMarkUtf8.Length);
            return p.Replace("\0", "");
        }

    }

}
