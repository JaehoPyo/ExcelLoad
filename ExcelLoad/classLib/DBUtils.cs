using System;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Windows.Forms;

namespace ExcelLoad.classLib
{
    static class DBUtils
    {
        /// <summary>
        /// Config.ini로부터 읽어와서 구성한 ConnectionString
        /// </summary>
        public static string ConnectionString
        {
            get
            {
                string DataBase = FileUtils.IniRead("Database", "DataBase");
                string ID = FileUtils.IniRead("Database", "ID");
                string PW = FileUtils.IniRead("Database", "PWD");
                string DataSource = FileUtils.IniRead("Database", "Datasource");

                return @"Persist Security Info=True;" +
                       @"Initial Catalog=" + DataBase + @";" +
                       @"User ID=" + ID + @";" +
                       @"Password=" + PW + @";" +
                       @"Data Source=" + DataSource + @";";
            }
        }

        public static SqlConnection sqlConnection { get; set; }

        /// <summary>
        /// Config.ini파일의 데이터베이스 접속 정보를 읽어서
        /// 데이터베이스 접속
        /// </summary>
        /// <returns>접속성공여부</returns>
        public static bool DatabaseConnect()
        {
            try
            {
                string DataBase = FileUtils.IniRead("Database", "DataBase");
                string ID = FileUtils.IniRead("Database", "ID");
                string PW = FileUtils.IniRead("Database", "PWD");
                string DataSource = FileUtils.IniRead("Database", "Datasource");
                string ConnectionString = @"Persist Security Info=True;" +
                                          @"Initial Catalog=" + DataBase + @";" +
                                          @"User ID=" + ID + @";" +
                                          @"Password=" + PW + @";" +
                                          @"Data Source=" + DataSource + @";";
                sqlConnection = new SqlConnection(ConnectionString);
                sqlConnection.Open();
                return true;
            }
            catch (SqlException ex)
            {
                FileUtils.WriteLog("[" + ex.Source + "-" + ex.Procedure + "]" + ex.Message);
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// ConnectionString을 매개변수로 넣어서 데이터베이스 접속
        /// </summary>
        /// <param name="ConnectionString">ConnectionString</param>
        /// <returns>성공여부</returns>
        public static bool DatabaseConnect(string ConnectionString)
        {
            try
            {
                sqlConnection = new SqlConnection(ConnectionString);
                sqlConnection.Open();
                return true;
            }
            catch (SqlException ex)
            {
                FileUtils.WriteLog("[" + ex.Source + "-" + ex.Procedure + "]" + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 연결된 데이터베이스 접속을 끊음
        /// </summary>
        public static void DatabaseDisConnect()
        {
            try
            {
                if (sqlConnection == null) return;

                if (sqlConnection.State == ConnectionState.Open)
                {
                    sqlConnection.Close();
                }
            }
            catch (SqlException ex)
            {
                FileUtils.WriteLog("[" + ex.Source + "-" + ex.Procedure + "]" + ex.Message);
                //MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Select 쿼리로 조회한 결과 테이블을 반환한다.(트랜잭션)
        /// </summary>
        /// <param name="SelectSQL">Select 쿼리</param>
        /// <returns>결과 테이블 또는 null</returns>
        public static DataTable getSelectResultTableOrNull(string SelectSQL, SqlTransaction transaction)
        {
            if (sqlConnection == null) return null;
            if (sqlConnection.State != ConnectionState.Open) return null;
            
            try
            {
                using (SqlCommand command = new SqlCommand())
                {
                    command.Transaction = transaction;
                    command.Connection = sqlConnection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = SelectSQL;
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet);
                    return dataSet.Tables[0];        
                }
            }
            catch(Exception ex)
            {
                FileUtils.WriteLog("[DBUtils.getSelectResultTableOrNull]" + ex.Message);
                MessageBox.Show("[DBUtils.getSelectResultTableOrNull]" + ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Select 쿼리로 조회한 결과 테이블을 반환한다.
        /// </summary>
        /// <param name="SelectSQL">Select 쿼리</param>
        /// <returns>결과 테이블 또는 null</returns>
        public static DataTable getSelectResultTableOrNull(string SelectSQL)
        {
            if (sqlConnection == null) return null;
            if (sqlConnection.State != ConnectionState.Open) return null;

            try
            {
                using (SqlCommand command = new SqlCommand())
                {
                    command.Connection = sqlConnection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = SelectSQL;
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet);
                    return dataSet.Tables[0];
                }
            }
            catch (Exception ex)
            {
                FileUtils.WriteLog("[DBUtils.getSelectResultTableOrNull]" + ex.Message);
                MessageBox.Show("[DBUtils.getSelectResultTableOrNull]" + ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Insert 쿼리로 테이블에 데이터를 넣는다.
        /// </summary>
        /// <param name="InsertSQL">Insert 쿼리</param>
        /// <returns>영향받은 행수</returns>
        public static int InsertData(string InsertSQL, SqlTransaction transaction)
        {
            if (sqlConnection == null) return -1;
            if (sqlConnection.State != ConnectionState.Open) return -1;

            try
            {
                using (SqlCommand command = new SqlCommand())
                {
                    command.Transaction = transaction;
                    command.Connection = sqlConnection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = InsertSQL;
                    int result = command.ExecuteNonQuery();

                    return result;
                }
            }
            catch (Exception ex)
            {
                FileUtils.WriteLog("[DBUtils.InsertData]" + ex.Message);
                MessageBox.Show("[DBUtils.InsertData]" + ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw ex;
            }
        }

        /// <summary>
        /// DB에 접속한 현재 PC의 IP를 가져옴
        /// </summary>
        public static string ClientIPAddress
        {
            get
            {
                IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
                string ip = string.Empty;
                for (int i = 0; i < host.AddressList.Length; i++)
                {
                    if (host.AddressList[i].AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        ip = host.AddressList[i].ToString();
                    }
                }
                return ip;
            }
        }

        /// <summary>
        /// B에 접속한 현재 PC의 이름을 가져옴
        /// </summary>
        public static string ClientPcName
        {
            get
            {
                return System.Windows.Forms.SystemInformation.ComputerName;
            }
        }
    }
}
