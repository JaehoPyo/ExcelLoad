using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ExcelLoad.classLib
{
    static class ExcelUtils
    {
        /// <summary>
        /// Excel용 OleDbConnectionString 구성
        /// </summary>
        /// <param name="fileName">엑셀파일명(경로)</param>
        /// <param name="fileExtenstion">엑셀파일확장자(xls or xlsx)</param>
        /// <returns>OleDbConnectionString 확장자가 다르면 Empty</returns>
        public static string getOleDbConnectionStringOrWhiteSpace(string fileName, string fileExtenstion)
        {
            // 확장자로 구분하여 커넥션 스트링을 구성
            string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
            string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
            string connectionString = string.Empty;
            string header = "Yes";
            switch (fileExtenstion)
            {
                case ".xls":  //Excel 97-03
                    connectionString = string.Format(Excel03ConString, fileName, header);
                    break;
                case ".xlsx": //Excel 07
                    connectionString = string.Format(Excel07ConString, fileName, header);
                    break;
                default:
                    connectionString = string.Empty;
                    break;

            }
            return connectionString;
        }

        /// <summary>
        /// Excel의 첫 번째 Sheet의 이름을 가져옴
        /// </summary>
        /// <param name="connectionString">OleDbConnectionString</param>
        /// <returns>첫 번재 Sheet의 이름 또는 null</returns>
        public static string getSheetNameOrNull(string connectionString)
        {
            string sheetName = string.Empty;
            // 첫 번째 시트의 이름을 가져옮
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;
                        connection.Open();
                        DataTable dtExcelSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                        connection.Close();
                    }
                }
                return sheetName;
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("공급자"))
                {
                    string errStr = "엑셀2010 Provider를 설치해주십시오." + Environment.NewLine +
                                    "[AccessDatabaseEngine 32비트 설치요망]" + Environment.NewLine +
                                    ex.Message;

                    MessageBox.Show(errStr, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return null;
            }
        }

        /// <summary>
        /// 엑셀 시트를 테이블로 반환
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="sheetName"></param>
        /// <returns>Excel DataTable or null</returns>
        public static DataTable getExcelTableOrNull(string connectionString, string sheetName)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    using (OleDbCommand command = new OleDbCommand())
                    {
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter())
                        {
                            DataTable table = new DataTable();
                            command.CommandText = "SELECT * From [" + sheetName + "]";
                            command.Connection = connection;
                            connection.Open();
                            adapter.SelectCommand = command;
                            adapter.Fill(table);
                            connection.Close();

                            return table;
                        }
                    }
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
