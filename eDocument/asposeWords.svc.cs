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
        public string EchoWithGet(string s)
        {
            return "You said " + s;
        }

        public string EchoWithPost(string s)
        {
            return "You said " + s;
        }

        //fills the Word template with XML Data
        public string fillWordTemplate(string ApplicationName,string TemplateName, string OutputDocumentName, string XMLData, string TemplateDir = "", string OutputDir = "")
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

            XmlNodeList counter = xml.GetElementsByTagName("fieldName");

            String[] names = new String[counter.Count];
            Object[] values = new Object[counter.Count];

            XmlNodeList elemList = xml.GetElementsByTagName("fieldName");
            for (int i = 0; i < elemList.Count; i++)
            {
                names[i] = elemList[i].InnerXml;
            }

            XmlNodeList elemList2 = xml.GetElementsByTagName("fieldValue");
            for (int i = 0; i < elemList2.Count; i++)
            {
                values[i] = elemList2[i].InnerXml;
            }

            asposeWordsWrapper AWW = new asposeWordsWrapper(ApplicationName, TemplateDir, TemplateName, OutputDir, OutputDocumentName);
            AWW.Execute(names, values);
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
