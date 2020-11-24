using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;

namespace MHAGConverter
{
    internal class Converters
    {
        public bool Ver1(string path)
        {
            // Connect to Excel files
            var myConnection = new OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                                                     path +
                                                                     ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";");
            var myCommand = new OleDbDataAdapter("select * from [Sheet1$]", myConnection);
            var ds = new System.Data.DataSet();
            myCommand.Fill(ds);
            myConnection.Close();
            var rows = ds.Tables[0].Rows;
            var file =
                new System.IO.StreamWriter(@"E:\strings.xml",false, Encoding.UTF8);
            using (file)
            {
                const string header = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
                file.WriteLine(header);
                file.WriteLine("<resources>");
                for (var i = 0; i < rows.Count; i++)
                {
                    var row = rows[i];
                    var items = row.ItemArray;
                    if (items[1] == null) return false;
                   
                    file.WriteLine("<string name=\"{0}\">{1}</string>",
                        items[0],
                        ConverterAscii(items));
                }
                file.WriteLine("</resources>");
                return true;
            }

          
        }
        private static string ConverterAscii(IReadOnlyList<object> items) => items[1] == null ? "" : (string) items[1];
    }
}
