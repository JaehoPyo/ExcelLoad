using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Zebra.Sdk.Comm;
using Zebra.Sdk.Printer;
using Zebra.Sdk.Printer.Discovery;

namespace ExcelLoad.classLib
{
    public static class LabelPrint
    {
        public static Connection PrinterConnection { get; set; }
        public static ZebraPrinter Printer { get; set; }
        /// <summary>
        /// 프린터 커넥션 활성화
        /// </summary>
        public static bool PrinterConnectionOpen()
        {
            try
            {
                List<string> list = new List<string>();
                foreach (DiscoveredUsbPrinter usbPrinter in UsbDiscoverer.GetZebraUsbPrinters(new ZebraPrinterFilter()))
                {
                    list.Add(usbPrinter.ToString());
                }

                if (list.Count == 0)
                {
                    MessageBox.Show("연결된 라벨 프린터가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                
                // 커넥션이 없으면 커넥션 만들어서 Open하고 return true
                if(PrinterConnection == null)
                {
                   PrinterConnection = new UsbConnection(list[0]);
                   PrinterConnection.Open();
                   return true; 
                }
                else 
                {
                    return true;
                }
            }
            catch(ConnectionException ex)
            {
                MessageBox.Show("프린터 연결실패" + Environment.NewLine + ex.ToString(), "실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("프린터 연결실패" + Environment.NewLine + ex.ToString(), "실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 프린터 커넥션 비활성화
        /// </summary>
        public static void PrinterConnectionClose()
        {
            try
            {
                if (PrinterConnection.Connected)
                    PrinterConnection.Close();
            }
            catch (Exception ex)
            {
                FileUtils.WriteLog("[" + ex.Source + "]" + ex.Message);
                MessageBox.Show(ex.Message, "에러", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 라벨 프린트
        /// </summary>
        /// <param name="command">ZPL 라벨 프린트 커맨드</param>
        public static void PrintByCommand(string command)
        {
            try
            {
                if (!LabelPrint.PrinterConnectionOpen())
                    return;
                if (Printer == null) Printer = ZebraPrinterFactory.GetInstance(PrinterConnection);
                Printer.SendCommand(command);
            }
            catch(ConnectionException ex)
            {
                MessageBox.Show("프린트 실패" + Environment.NewLine + ex.ToString(), "실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public struct PrintData
        {
            public string LBL_NO;
            public string LOT_NO;
            public string ITEM_NM;
            public string ITEM_NO;
            public string VENDOR_NM;
            public string KEEP_CONDITION;
            public string UNLOAD_DATE;
            public string ORDER_QTY;
            public string UNIT;
            public string USE_DEADLN;
            public string LABEL_QTY;
            public string LABEL_SEQ;
        }
        public static void PrintLabel(string EventNM, string PGM_ID, string PGM_NM, PrintData data)
        {
            try
            {
                if (!LabelPrint.PrinterConnectionOpen())
                    return;

                if (Printer == null) Printer = ZebraPrinterFactory.GetInstance(PrinterConnection);

                string BOTTOM_LEFT = "";
                string BOTTOM_RIGHT = "";
                if (DBUtils.sqlConnection.State == ConnectionState.Open)
                {
                    using (SqlCommand SqlCommand = new SqlCommand())
                    {
                        SqlCommand.Connection = DBUtils.sqlConnection;
                        SqlCommand.CommandText = "P_8190_R_CFG";
                        SqlCommand.CommandType = CommandType.StoredProcedure;
                        SqlCommand.Parameters.Add("@rCOMP", SqlDbType.NVarChar).Value = "100";
                        SqlCommand.Parameters.Add("@rCFG_KEY", SqlDbType.NVarChar).Value = "";
                        SqlCommand.Parameters.Add("@rCFG_NM", SqlDbType.NVarChar).Value = "";
                        SqlCommand.Parameters.Add("@rCFG_VAL", SqlDbType.NVarChar).Value = "";
                        SqlCommand.Parameters.Add("@rCFG_TP", SqlDbType.NVarChar).Value = "LBLNO";
                        SqlCommand.Parameters.Add("@rETC", SqlDbType.NVarChar).Value = "";

                        SqlCommand.Parameters.Add("@rPGM_ID", SqlDbType.NVarChar).Value = PGM_ID;
                        SqlCommand.Parameters.Add("@rPGM_NM", SqlDbType.NVarChar).Value = PGM_NM;
                        SqlCommand.Parameters.Add("@rEVENT_NM", SqlDbType.NVarChar).Value = EventNM;
                        SqlCommand.Parameters.Add("@rMSG", SqlDbType.NVarChar).Value = "";
                        SqlCommand.Parameters.Add("@rCRT_USR", SqlDbType.NVarChar).Value = "System";
                        SqlCommand.Parameters.Add("@rCRT_PC", SqlDbType.NVarChar).Value = DBUtils.ClientPcName;
                        SqlCommand.Parameters.Add("@rCRT_IP", SqlDbType.NVarChar).Value = DBUtils.ClientIPAddress;
                        SqlCommand.Parameters.Add("@rCRT_MENU", SqlDbType.NVarChar).Value = "PC";
                        SqlCommand.Parameters.Add("@rOK", SqlDbType.NVarChar).Value = "";

                        SqlDataAdapter dataAdapter = new SqlDataAdapter(SqlCommand);
                        DataSet dataSet = new DataSet();
                        dataAdapter.Fill(dataSet);
                        if (dataSet.Tables[0].Rows.Count == 2)
                        {
                            BOTTOM_LEFT = dataSet.Tables[0].Rows[0]["CFG_VAL"].ToString();
                            BOTTOM_RIGHT = dataSet.Tables[0].Rows[1]["CFG_VAL"].ToString();
                        }                
                    }
                }
                else
                {
                    MessageBox.Show("DB연결이 끊어졌습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                }

                int DefaultX = 30;
                int DefaultY = 20;
                string command = "^XA" +
                                 "^SEE:UHANGUL.DAT^FS" +
                                 "^CW1,E:KFONT3.TTF" +
                                 "^CI28^FS" +
                                 "^FWB" + // 문자 방향
                                 "^BY2" + // 바코드 크기
                                 "^FO" + DefaultX.ToString() + "," + DefaultY.ToString() + "^GB532,740,3^FS" + // BOX 그리기
                                 "^CF1,50" + // 폰트설정
                                 "^FT" + (DefaultX + 50).ToString() + "," + (DefaultY + 730).ToString() + "^FD" + data.LBL_NO + "^FS" + "^BC,60,N,N,N" + "^FO" + (DefaultX + 10).ToString() + "," + (DefaultY + 40).ToString() + "^FD" + data.LBL_NO + "^FS" + 
                                 "^CF1,50,40"; // 폰트설정
                if (data.ITEM_NM.Length > 20)
                {
                    command = command + "^FT" + (DefaultX + 150).ToString() + "," + (DefaultY + 730).ToString() + "^FD품목: " + data.ITEM_NM.Substring(0, 21) + "^FS" +
                                        "^FT" + (DefaultX + 200).ToString() + "," + (DefaultY + 730).ToString() + "^FD" + data.ITEM_NM.Substring(21, data.ITEM_NM.Length - 21) + "^FS";
                }
                else
                {
                    command = command + "^FT" + (DefaultX + 150).ToString() + "," + (DefaultY + 730).ToString() + "^FD품목: " + data.ITEM_NM + "^FS";
                }
                command = command +     "^A1,40,30^FT" + (DefaultX + 300).ToString() + "," + (DefaultY + 730).ToString() + "^FD품목번호: " + data.ITEM_NO + "^FS" +
                                        "^A1,40,30^FT" + (DefaultX + 340).ToString() + "," + (DefaultY + 730).ToString() + "^FD재고번호: " + data.LOT_NO  + "^FS" + "^BC,50,N,N,N" + "^FO" + (DefaultX + 300).ToString() + "," + (DefaultY + 40).ToString() + "^FD" + data.LOT_NO + "^FS";
                                            
                command = command +     "^A1,35,30^FT" + (DefaultX + 380).ToString() + "," + (DefaultY + 730).ToString()  + "^FD입고: "     + data.ORDER_QTY + " " + data.UNIT + "^FS" +
                                        "^CF1,30,25" + // 폰트설정
                                        "^FT" + (DefaultX + 430).ToString() + "," + (DefaultY + 730).ToString()  + "^FD구매처: "   + data.VENDOR_NM      + "^FS" +
                                        "^FT" + (DefaultX + 460).ToString() + "," + (DefaultY + 730).ToString()  + "^FD하역일자: " + data.UNLOAD_DATE    + "^FS" +
                                        "^FT" + (DefaultX + 490).ToString() + "," + (DefaultY + 730).ToString()  + "^FD사용기한: " + data.USE_DEADLN     + "^FS" +
                                        "^FT" + (DefaultX + 520).ToString() + "," + (DefaultY + 730).ToString()  + "^FD보관조건: " + data.KEEP_CONDITION + "^FS" +
                                        "^FT" + (DefaultX + 560).ToString() + "," + (DefaultY + 740).ToString() + "^FD" + BOTTOM_LEFT + "^FS" +
                                        "^FO" + (DefaultX + 540).ToString() + "," + (DefaultY + 330).ToString() + "^FD" + data.LABEL_SEQ + "/" + data.LABEL_QTY + "^FS" +
                                        "^FO" + (DefaultX + 540).ToString() + "," + (DefaultY + 10).ToString() + "^FD" + BOTTOM_RIGHT + "^FS";
                command = command + "^XZ";
                Printer.SendCommand(command);
            }
            catch (ConnectionException ex)
            {
                MessageBox.Show("프린트 실패" + Environment.NewLine + ex.ToString(), "실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch(SqlException ex)
            {
                MessageBox.Show("프린트 실패" + Environment.NewLine + ex.ToString(), "실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("프린트 실패" + Environment.NewLine + ex.ToString(), "실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        public static void PrintWareHouse(string WMS, string WMS_NM)
        {
            try
            {
                if (Printer == null) Printer = ZebraPrinterFactory.GetInstance(PrinterConnection);

                int DefaultX = 10;
                int DefaultY = 10;
                string command = "^XA" +
                                 "^SEE:UHANGUL.DAT^FS" +
                                 "^CW1,E:KFONT3.TTF" +
                                 "^CI28^FS" +
                                 "^CF1,35" + // 폰트설정
                                 "^FO" + (DefaultX + 140).ToString() + "," + (DefaultY + 50).ToString() + "^FD" + WMS + ":" + WMS_NM + "^FS" +
                                 "^BCN,70,N,N,N" + "^FO" + (DefaultX + 140).ToString() + "," + (DefaultY + 90).ToString() + "^FD" + WMS + "000001" + "^FS" +
                                 "^FO" + (DefaultX + 160).ToString() + "," + (DefaultY + 170).ToString() + "^FD" + WMS + "000001" + "^FS";
                command = command + "^XZ";
                Printer.SendCommand(command);
            }
            catch (ConnectionException ex)
            {
                MessageBox.Show("프린트 실패" + Environment.NewLine + ex.ToString(), "실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("프린트 실패" + Environment.NewLine + ex.ToString(), "실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

    }
}
