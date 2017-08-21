using System;
using DocuWare.Platform.ServerClient;
using Jitbit.Utils;
using System.Collections.Generic;

namespace DocuWare
{
    public static class ConnectAndExport
    {
        public static Uri uri = new Uri("https://docuware-online.com:443/DocuWare/Platform");

        public static ServiceConnection ConnectWithOrg()
        {
            return ServiceConnection.Create(uri, "DWAdmin", "Tosh1ba!", organization: "SmartMFP-TBSNE-54");
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
                if (document["STATUS"].Item.ToString() == "Ready")
                {
                    myExport.AddRow();
                    myExport["DWDOCID"] = document["DWDOCID"];
                    myExport["FIRST_NAME"] = document["FIRST_NAME"];
                    myExport["LAST_NAME"] = document["LAST_NAME"];
                    Update(document);
                }
            }
            myExport.ExportToFile(string.Format("C:\\Docuware Project\\Import-{0:yyyy-MM-dd_hh-mm-ss-tt}.csv", DateTime.Now));
            //myExport.ExportToFile("C:\\Docuware Project\\Import.csv");

            return queryResult;
        }

        public static void Main()
        {
            ListAllDocuments(ConnectWithOrg(), "21d44bf7-7c1a-4620-baa2-a1f21d27bbd4");
            Console.WriteLine("Hit ENTER to exit...");
            Console.ReadLine();
        }
    }
}