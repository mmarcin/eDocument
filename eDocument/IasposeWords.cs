using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;


namespace eDocument
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IasposeWords" in both code and config file together.
    [ServiceContract]
    public interface IasposeWords
    {

        [OperationContract]
        string fillWordTemplate(string ApplicationName, string TemplateName, string OutputDocumentName, string XMLData, string TemplateDir = "", string OutputDir = "", string RemoveUnusedFields = "no", string RemoveUnusedRegions = "no", string RemoveEmptyParagraphs = "no");
    }
}
