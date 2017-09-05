using System;
using ConnectAndExportCsv.Properties;
using DocuWare.Platform.ServerClient;
using Jitbit.Utils;
using System.Collections.Generic;

namespace DocuWare
{
    public class ConnectAndExport
    {

        public static Uri uri = new Uri("https://docuware-online.com:443/DocuWare/Platform");

        public static ServiceConnection ConnectWithOrg()
        {
            return ServiceConnection.Create(uri, Settings.Default.initial.ToString(), Settings.Default.entry.ToString(), organization: Settings.Default.org.ToString());
        }

        public static DocumentIndexFields Update(Document document)
        {
            var fields = new DocumentIndexFields()
            {
                Field = new List<DocumentIndexField>()
                {
                    DocumentIndexField.Create("STATUS", "Exported"),
                }
            };
            return document.PutToFieldsRelationForDocumentIndexFields(fields);
        }

        public static DocumentsQueryResult ListAllDocuments(ServiceConnection conn, string fileCabinetId, int? count = 100)
        {
            CsvExport myExport = new CsvExport();
            DocumentsQueryResult queryResult = conn.GetFromDocumentsForDocumentsQueryResultAsync(
                fileCabinetId,
                count: count).Result;
            foreach (var document in queryResult.Items)
            {
                if (document["STATUS"].Item.ToString() == "Processed")
                {
                    myExport.AddRow();
                    myExport["Type"] = document["DOCUMENT_TYPE"].ToString().Replace("DOCUMENT_TYPE (String): ","");
                    myExport["DWDOCID"] = document["DWDOCID"].ToString().Replace("DWDOCID (Int): ", "");
                    myExport["Invoice Date"] = document["INVOICE_DATE"].ToString().Replace("INVOICE_DATE (String): ", "").Replace("0:00","");
                    myExport["Post Period"] = "";
                    myExport["VendorID"] = document["VENDOR_ID"].ToString().Replace("VENDOR_ID (String): ", "");
                    myExport["VendorRef"] = "";
                    myExport["Location"] = "";
                    myExport["Terms"] = "";
                    myExport["Due Date"] = "";
                    myExport["Invoice Description"] = "";
                    myExport["Originating Branch"] = document["COMPANY"].ToString().Replace("COMPANY (String): ", "");
                    myExport["LineNbr"] = document["LINE_NUMBER"].ToString().Replace("LINE_NUMBER (String): ", "").Replace("LINE_NUMBER (Date): ", "");
                    myExport["Ext Cost"] = document["AMOUNT"].ToString().Replace("AMOUNT (String): ", "");
                    myExport["Account"] = document["GL_CODE"].ToString().Replace("GL_CODE (String): ", "").Replace("GL_CODE (Date): ", "");
                    myExport["Subaccount"] = document["GL_SUBCODE"].ToString().Replace("GL_SUBCODE (String): ", "").Replace("GL_SUBCODE (Date): null", "");
                    myExport["Destination Branch"] = document["COMPANY"].ToString().Replace("COMPANY (String): ", "");
                    myExport["Line Description"] = "";
                    Update(document);
                }
            }
            myExport.ExportToFile(string.Format(Settings.Default.exportpath, DateTime.Now));

            return queryResult;
        }

        public static void Main()
        {
            ListAllDocuments(ConnectWithOrg(), Settings.Default.fcguid);
            Console.WriteLine("Hit ENTER to exit...");
            Console.ReadLine();
        }
    }
}