using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Sbot
{
    class excelData
    {
        private static string path = "";
        public static string DosyaYoluSec()
        {
            string Platform = "";
            OpenFileDialog ac = new OpenFileDialog();

            if (Environment.Is64BitProcess)
               
            { ac.Filter = "Excel filtresi (*.xls;*.xlsx)|*.xls;*.xlsx"; Platform = "x64. Sadece XLSX Dosyaları"; }
            else
               
            { ac.Filter = "Excel filtresi(*.xls)|*.xls"; Platform = "x86. Sadece XLS Dosyaları"; }

            ac.Title = "Platform " + Platform;
            ac.ShowDialog();
            path = ac.FileName.ToString();
            return path;
        }
        public static DataTable DosyaOku()
        {
            DataTable dtexcel = new DataTable();

            if (path.Trim().Length > 0)
            {
                DataTable schemaTable = new DataTable();
                string strConn = "";


                if (Environment.Is64BitProcess) { 
                   
                strConn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + path + "; Extended Properties = \"Excel 12.0; HDR = Yes; IMEX = 0\"";
                }

                else
                {
                    strConn = "Provider = Microsoft.Jet.OLEDB.4.0;  Data Source=" + path + "; Extended Properties = \"Excel 8.0; HDR = Yes; IMEX = 0\"";
                }

                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                DataRow schemaRow = schemaTable.Rows[0];
                string sheet = schemaRow["TABLE_NAME"].ToString();
                if (!sheet.EndsWith("_"))
                {
                    string query = "SELECT  * FROM [" + sheet + "]";
                    OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);
                    dtexcel.Locale = CultureInfo.CurrentCulture;
                    daexcel.Fill(dtexcel);
                }
                conn.Close();

            }
            else
            {
                MessageBox.Show("Okunacak EXCEL dosyası bulunamadı. Lütfen önce okunacak EXCEL dosyası seçin.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            return dtexcel;
        }

    }
}
