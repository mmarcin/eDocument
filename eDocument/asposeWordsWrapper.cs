using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Aspose.Words;
using Aspose.Words.MailMerging;
using System.Configuration;


namespace asposeWR
{

    public class asposeWordsWrapper
    {
        public string ApplicationName;      // Name of application / ConnectionString
        public string TemplateDir;          // Dir, where the template file is located      
        public string TemplateName;         // Template.docx

        public string OutputDir;            // Dir, where the output file will be written - c:/temp/
        public string OutputDocumentName;   // OutputFile.docx
        public string OutputPDFName;        // OttputFIle.docx.pdf
        public string OutputHTMLName;        // OttputFIle.docx.html

        public Document doc = new Document();

        public asposeWordsWrapper(string ApplicationName, string TemplateDir, string TemplateName, string OutputDir, string OutputDocumentName)
        {
            this.ApplicationName = ApplicationName;
            this.TemplateDir = TemplateDir;
            this.TemplateName = TemplateName;
            this.OutputDir = OutputDir;
            this.OutputDocumentName = OutputDocumentName;
            this.OutputPDFName = OutputDocumentName + ".pdf";
            this.OutputHTMLName = OutputDocumentName + ".html";

            Aspose.Words.License awLic = new Aspose.Words.License();
            awLic.SetLicense("Aspose.Total.lic");

            doc = new Document(this.TemplateDir + this.TemplateName);  // Create Aspose document object
            
        }

        //Simple mail merge
        public void Execute(string[] names, object[] values)
        {
            doc.MailMerge.Execute(names, values);            
        }

        //region mail merge with datasource
        public void ExecuteRegions(DataSet DSRegion)
        {            
            doc.MailMerge.ExecuteWithRegions(DSRegion);                        // Make MailMerge With dataset            
        }

        //region mail merge with SQL
        public void ExecuteRegionsSQL(string SelectString, string TableName)
        {
            DataTable TableWithData = GetDatabaseResults(ApplicationName, SelectString, TableName);  // Get Data
            doc.MailMerge.ExecuteWithRegions(TableWithData);                        // Make MailMerge With regions            
        }

        // Remove unused regions
        public void RemoveUnusedFields()
        {
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        }

        // Remove unused regions
        public void RemoveUnusedRegions() {
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;
        }

        // Remove unused regions
        public void RemoveEmptyParagraphs()
        {
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        }

        //save of the filled document
        public void Save()
        {
            doc.Save(this.OutputDir + this.OutputDocumentName); // Save the result

            Document PDFdoc = new Document(this.OutputDir + this.OutputDocumentName);
            PDFdoc.Save(this.OutputDir + this.OutputPDFName, SaveFormat.Pdf); //Save the result as PDF

            Document HTMLdoc = new Document(this.OutputDir + this.OutputDocumentName);
            HTMLdoc.Save(this.OutputDir + this.OutputHTMLName, SaveFormat.Html); //Save the result as HTML

        }


        private static DataTable GetDatabaseResults(string ApplicationName, string SelectString, string TableNameString)
        {
            DataTable table = ExecuteDataTable(ApplicationName,SelectString);
            table.TableName = TableNameString;
            return table;
        }

        private static DataTable ExecuteDataTable(string ApplicationName, string commandText)
        {
            // Open the database connection.
            // string connString = "Server=local.ebiz.sk;Database=eProcurement;User Id=sa;Password=Lomtec2000;";
            string connString = ConfigurationManager.ConnectionStrings[ApplicationName].ConnectionString;

            SqlConnection conn = new SqlConnection(connString);
            conn.Open();

            // Create and execute a command.
            SqlCommand cmd = new SqlCommand(commandText, conn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable table = new DataTable();
            da.Fill(table);

            // Close the database.
            conn.Close();

            return table;
        }

    }

}

