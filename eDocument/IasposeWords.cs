using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;


namespace eDocument
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IasposeWords" in both code and config file together.
    [ServiceContract]
    public interface IasposeWords
    {

        [OperationContract]
        string fillWordTemplate(string ApplicationName, string InstanceCode, string TemplateName, string XMLData, string RemoveUnusedFields = "no", string RemoveUnusedRegions = "no", string RemoveEmptyParagraphs = "no");

        [OperationContract]
        [WebInvoke(UriTemplate = "UploadFile /{fileName}")]
        string uploadFile(Stream fileContents);

    [OperationContract]
        string getUserName();
    }
}
