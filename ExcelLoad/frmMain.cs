using ExcelLoad.classLib;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace ExcelLoad
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        /* 엑셀파일 로드 */
        private void btnLoadExcel_Click(object sender, EventArgs e)
        {
            try
            {
                string excelFilePath = string.Empty;
                string extension = string.Empty;
                using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
                {
                    // 파일 불러오기
                    dialog.Filters.Add(new CommonFileDialogFilter("엑셀", ".xlsx,.xls"));
                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        // 파일명(경로)와 확장자 구하기
                        excelFilePath = dialog.FileName;
                        extension = Path.GetExtension(excelFilePath);

                        // OleDbConnectionString 구성
                        string connectionString = ExcelUtils.getOleDbConnectionStringOrWhiteSpace(excelFilePath, extension);

                        // Sheet 이름 가져오기
                        string sheetName = ExcelUtils.getSheetNameOrNull(connectionString);

                        // Sheet의 데이터를 읽어서 Grid에 보이게함
                        gridControl1.DataSource = ExcelUtils.getExcelTableOrNull(connectionString, sheetName);
                    }
                    else // 파일을 선택하지 않았을 경우 빠져나감
                    {
                        return;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        bool isRunning = true;
        bool isComplete = true;
        /* DB에 데이터 삽입 */
        private void btnInsertData_Click(object sender, EventArgs e)
        {
            if (DBUtils.sqlConnection == null) return;
            if (gridView1.RowCount < 1) return;

            btnCancel.Visible = true;

            isRunning = true;
            isComplete = true;
            string ITEM_NO = string.Empty;
            string LOT_NO = string.Empty;
            string strSQL = string.Empty;
            DataTable table = null;

            // 중복데이터 표시 그리드 초기화
            tableForGrid2.Rows.Clear();
            gridView2.RefreshData();
            // 품목코드 없는 데이터 표시 그리드 초기화
            tableForGrid4.Rows.Clear();
            gridView4.RefreshData();

            // 엑셀데이터 한 행씩 수행
            SqlTransaction transaction = DBUtils.sqlConnection.BeginTransaction();
            try
            {
                for (int i = 0; i < gridView1.RowCount; i++)
                {
                    if (!isRunning)
                    {
                        transaction.Rollback();
                        btnCancel.Visible = false;
                        Application.DoEvents();
                        break;
                    }

                    // 품목마스터(W_ITEM)에 있는 품목인지 확인
                    LOT_NO = gridView1.GetRowCellValue(i, "LOT_NO").ToString();
                    ITEM_NO = gridView1.GetRowCellValue(i, "ITEM_NO").ToString();
                    strSQL = "SELECT COUNT(*) AS CNT FROM W_ITEM " +
                             " WHERE ITEM_NO ='" + ITEM_NO + "'";
                    table = DBUtils.getSelectResultTableOrNull(strSQL, transaction);

                    // 품목마스터(W_ITEM)에 없는 품목이면 다음 자료로 넘어감
                    if (table.Rows[0][0].ToString() == "0")
                    {
                        AddDataToGrid(i, tableForGrid4, gridView4);
                        continue;
                    }

                    // 창고(입출고장)에 이미 있는 재고인지 확인
                    strSQL = "SELECT COUNT(*) AS CNT FROM W_STOCK " +
                             " WHERE ITEM_NO ='" + ITEM_NO + "'" +
                             "   AND LOT_NO = '" + LOT_NO + "'" +
                             "   AND WMS = 1001";
                    table = DBUtils.getSelectResultTableOrNull(strSQL, transaction);

                    // 창고에 이미 있는 재고라면 오른쪽 그리드로 데이터를 복사
                    if (table.Rows[0][0].ToString() != "0")
                    {
                        // 오른쪽 그리드에 데이터 복사
                        AddDataToGrid2(i);
                        gridView1.FocusedRowHandle = i + 1;
                        FileUtils.WriteLog("중복데이터 : " + ITEM_NO + ", " + LOT_NO);
                    }
                    // 창고에 없는 재고라면 W_STOCK에 데이터 INSERT
                    else
                    {
                        InsertStockData(i, transaction);
                        gridView1.FocusedRowHandle = i + 1;
                    }
                    Application.DoEvents();
                }
            }
            catch(Exception ex)
            {
                transaction.Rollback();
                isComplete = false;
                MessageBox.Show(ex.Message);
            }

            if(isComplete)
            {
                // commit
                transaction.Commit();
            }
        }

        // 삽입취소
        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (isRunning && isComplete)
            {
                isRunning = false;
                isComplete = false;
            }
        }

        // 라벨발행 그리드 조회
        private void btnInquiry_Click(object sender, EventArgs e)
        {
            try
            {
                string excelFilePath = string.Empty;
                string extension = string.Empty;
                using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
                {
                    // 파일 불러오기
                    dialog.Filters.Add(new CommonFileDialogFilter("엑셀", ".xlsx,.xls"));
                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        // 파일명(경로)와 확장자 구하기
                        excelFilePath = dialog.FileName;
                        extension = Path.GetExtension(excelFilePath);

                        // OleDbConnectionString 구성
                        string connectionString = ExcelUtils.getOleDbConnectionStringOrWhiteSpace(excelFilePath, extension);

                        // Sheet 이름 가져오기
                        string sheetName = ExcelUtils.getSheetNameOrNull(connectionString);

                        // Sheet의 데이터를 읽어서 Grid에 보이게함
                        gridControl3.DataSource = ExcelUtils.getExcelTableOrNull(connectionString, sheetName);
                    }
                    else // 파일을 선택하지 않았을 경우 빠져나감
                    {
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // 라벨발행 그리드 라벨 수량 입력
        Dictionary<int, int> hash = new Dictionary<int, int>();
        Dictionary<int, string> hash2 = new Dictionary<int, string>();
        private void gridView3_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
            if (e.IsSetData)
            {
                if (e.Column.FieldName == "LBL_QTY")
                {
                    if (hash.ContainsKey(e.ListSourceRowIndex))
                    {
                        hash[e.ListSourceRowIndex] = (int)e.Value;
                    }
                    else
                    {
                        hash.Add(e.ListSourceRowIndex, (int)e.Value);
                    }
                }

                if (e.Column.FieldName == "LBL_NO")
                {
                    if (hash2.ContainsKey(e.ListSourceRowIndex))
                    {
                        hash2[e.ListSourceRowIndex] = (string)e.Value;
                    }
                    else
                    {
                        hash2.Add(e.ListSourceRowIndex, (string)e.Value);
                    }
                }
            }
            else if (e.IsGetData)
            {
                if (e.Column.FieldName == "LBL_QTY")
                {
                    if (hash.ContainsKey(e.ListSourceRowIndex))
                    {
                        e.Value = hash[e.ListSourceRowIndex];
                    }
                    else
                    {
                        e.Value = 0;
                    }
                }

                if (e.Column.FieldName == "LBL_NO")
                {
                    if (hash2.ContainsKey(e.ListSourceRowIndex))
                    {
                        e.Value = hash2[e.ListSourceRowIndex];
                    }
                    else
                    {
                        string ITEM_NO = (e.Row as DataRowView)["ITEM_NO"].ToString();
                        string LOT_NO = (e.Row as DataRowView)["LOT_NO"].ToString();
                        string LBL_NO = getLabelNoOrNullFromStock(ITEM_NO, LOT_NO);
                        e.Value = (LBL_NO == null ? "발행불가" : LBL_NO);
                    }
                }
            }
        }

        /// <summary>
        /// W_STOCK테이블에서 라벨번호 가져옴
        /// </summary>
        /// <param name="ITEM_NO"></param>
        /// <param name="LOT_NO"></param>
        /// <returns>LBL_NO or null</returns>
        private string getLabelNoOrNullFromStock(string ITEM_NO, string LOT_NO)
        {
            string strSQL = "SELECT TOP 1 LBL_NO " +
                                    "  FROM W_STOCK " +
                                    " WHERE ITEM_NO = '" + ITEM_NO + "'" +
                                    "   AND LOT_NO = '" + LOT_NO + "'";
            DataTable table = DBUtils.getSelectResultTableOrNull(strSQL);

            if (table == null) return null;

            if (table.Rows.Count > 0)
                return table.Rows[0]["LBL_NO"].ToString();
            else
                return null;
        }

        // 초기화 버튼 클릭
        private void btnClear_Click(object sender, EventArgs e)
        {
            gridControl3.DataSource = null;
            hash.Clear();
            hash2.Clear();
        }

        bool isPrinting = false;
        LabelPrint.PrintData data;
        // 라벨프린트 클릭
        private void btnPrint_Click(object sender, EventArgs e)
        {
            isPrinting = true;
            if (!LabelPrint.PrinterConnectionOpen())
                return;

            // 라벨발행
            int index;
            for (int i = 0; i < gridView3.SelectedRowsCount; i++)
            {
                index = gridView3.GetSelectedRows()[i];
                gridView3.FocusedRowHandle = index;
                Application.DoEvents();
                bool isPrintable = (gridView3.GetRowCellValue(index, "LBL_NO").ToString().Contains("불가") ? false : true);                
                if (!isPrintable) continue;

                int LBL_QTY = Convert.ToInt32(gridView3.GetRowCellValue(index, "LBL_QTY").ToString());
                for (int j = 1; j < LBL_QTY + 1; j++)
                {
                    if (!isPrinting) return;
                    data.ITEM_NM = gridView3.GetRowCellValue(index, "ITEM_NM").ToString();
                    data.ITEM_NO = gridView3.GetRowCellValue(index, "ITEM_NO").ToString();
                    data.KEEP_CONDITION = gridView3.GetRowCellValue(index, "STORE_CONDI").ToString();
                    data.LBL_NO = gridView3.GetRowCellValue(index, "LBL_NO").ToString();
                    data.LOT_NO = gridView3.GetRowCellValue(index, "LOT_NO").ToString();
                    data.VENDOR_NM = gridView3.GetRowCellValue(index, "VENDOR_NM").ToString();
                    data.ORDER_QTY = string.Format("{0:#,#.#####}", Convert.ToDecimal(gridView3.GetRowCellValue(index, "ITEM_QTY")));
                    data.LABEL_QTY = gridView3.GetRowCellValue(index, "LBL_QTY").ToString();
                    data.UNIT = gridView3.GetRowCellValue(index, "UNIT").ToString();
                    data.LABEL_SEQ = j.ToString();
                    if ((gridView3.GetRowCellValue(index, "IN_DATE") == null) ||
                        (gridView3.GetRowCellValue(index, "IN_DATE").ToString() == ""))
                    {
                        data.UNLOAD_DATE = "";
                    }
                    else
                    {
                        data.UNLOAD_DATE = DateTime.ParseExact(gridView3.GetRowCellValue(index, "IN_DATE").ToString(), "yyyyMMdd", null).ToString("yyyy-MM-dd");
                    }

                    if ((gridView3.GetRowCellValue(index, "USE_DEADLN") == null) ||
                        (gridView3.GetRowCellValue(index, "USE_DEADLN").ToString() == ""))
                    {
                        data.USE_DEADLN = "";
                    }
                    else
                    {
                        data.USE_DEADLN = DateTime.ParseExact(gridView3.GetRowCellValue(index, "USE_DEADLN").ToString(), "yyyyMMdd", null).ToString("yyyy-MM-dd");
                    }

                    LabelPrint.PrintLabel("전체라벨발행", "재고이관", "재고이관", data);

                    System.Threading.Thread.Sleep(1000);
                }
            }
        }

        // 중지 버튼 클릭
        private void btnStop_Click(object sender, EventArgs e)
        {
            if (isPrinting) isPrinting = false;
        }

        string strSQL;
        string ITEM_NO, LOT_NO, ITEM_QTY, MGM_STATUS, USE_DEADLN, MADEIN_NO, MADEIN_DATE, MADEIN_NM, VENDOR_NM;
        string STOCK_SEQ, LBL_NO, UNIT, IN_DATE;
        /// <summary>
        /// W_STOCK 테이블에 데이터 삽입
        /// </summary>
        /// <param name="grid1RowHandle">그리드1의 행 번호</param>
        public void InsertStockData(int grid1RowHandle, SqlTransaction transaction)
        {
            if (DBUtils.sqlConnection == null) return;
            try
            {

                ITEM_NO     = gridView1.GetRowCellValue(grid1RowHandle, "ITEM_NO").ToString();
                LOT_NO      = gridView1.GetRowCellValue(grid1RowHandle, "LOT_NO").ToString();
                ITEM_QTY    = gridView1.GetRowCellValue(grid1RowHandle, "ITEM_QTY").ToString();
                MGM_STATUS  = gridView1.GetRowCellValue(grid1RowHandle, "MGM_STATUS").ToString();
                USE_DEADLN  = gridView1.GetRowCellValue(grid1RowHandle, "USE_DEADLN").ToString();
                MADEIN_NO   = gridView1.GetRowCellValue(grid1RowHandle, "MADEIN_NO").ToString();
                MADEIN_DATE = gridView1.GetRowCellValue(grid1RowHandle, "MADEIN_DATE").ToString();
                MADEIN_NM   = gridView1.GetRowCellValue(grid1RowHandle, "MADEIN_NM").ToString();
                VENDOR_NM   = gridView1.GetRowCellValue(grid1RowHandle, "VENDOR_NM").ToString();
                STOCK_SEQ   = getSeqNo(transaction);
                UNIT        = getUnit(ITEM_NO, transaction);
                IN_DATE     = gridView1.GetRowCellValue(grid1RowHandle, "IN_DATE").ToString();

                IN_DATE     = IN_DATE.Contains("N/A") || IN_DATE == ""                ? "20180101" : IN_DATE;
                USE_DEADLN  = USE_DEADLN.Contains("N/A") || USE_DEADLN.Contains("")   ? "" : USE_DEADLN;
                MADEIN_DATE = MADEIN_DATE.Contains("N/A") || MADEIN_DATE.Contains("") ? "" : MADEIN_DATE;

                // 작은 따옴표 있는지 확인 후 있으면 작은 따옴표를 2개 붙임
                ITEM_NO     = ITEM_NO.Contains("'")     ? withSingleQuotedString(ITEM_NO)     : ITEM_NO;
                LOT_NO      = LOT_NO.Contains("'")      ? withSingleQuotedString(LOT_NO)      : LOT_NO;
                ITEM_QTY    = ITEM_QTY.Contains("'")    ? withSingleQuotedString(ITEM_QTY)    : ITEM_QTY;
                MGM_STATUS  = MGM_STATUS.Contains("'")  ? withSingleQuotedString(MGM_STATUS)  : MGM_STATUS;
                USE_DEADLN  = USE_DEADLN.Contains("'")  ? withSingleQuotedString(USE_DEADLN)  : USE_DEADLN;
                MADEIN_NO   = MADEIN_NO.Contains("'")   ? withSingleQuotedString(MADEIN_NO)   : MADEIN_NO;
                MADEIN_DATE = MADEIN_DATE.Contains("'") ? withSingleQuotedString(MADEIN_DATE) : MADEIN_DATE;
                MADEIN_NM   = MADEIN_NM.Contains("'")   ? withSingleQuotedString(MADEIN_NM)   : MADEIN_NM;
                VENDOR_NM   = VENDOR_NM.Contains("'")   ? withSingleQuotedString(VENDOR_NM)   : VENDOR_NM;

                // 제품의 라벨번호는 ITEM_NO + LOTNO
                if (gridView1.GetRowCellValue(grid1RowHandle, "ASSET_CLASS").ToString().Contains("제품"))
                {
                    LBL_NO = ITEM_NO + LOT_NO;
                    strSQL = "INSERT INTO W_LABEL(COMP,       ITEM_NO,           LOT_NO,           LBL_NO) " + Environment.NewLine +
                             "             VALUES('100', '" + ITEM_NO + "', '" + LOT_NO + "', '" + LBL_NO + "')";
                    DBUtils.InsertData(strSQL, transaction);
                }
                else // 원부자재의 라벨번호는 프로시저로 가져옴
                {
                    LBL_NO = (ITEM_NO == "PALLET" ? LOT_NO : getLabelNo(ITEM_NO, LOT_NO, transaction));
                }

                strSQL = " INSERT INTO W_STOCK ( WMS,               LOC,              ITEM_NO,             LOT_NO,              ITEM_QTY, " + Environment.NewLine +
                         "                       MGM_STATUS,        USE_DEADLN,       IN_DATE,             MADEIN_NO,           MADEIN_DATE, " + Environment.NewLine +
                         "                       MADEIN_NM,         VENDOR_NM,        STOCK_SEQ,           PRE_STOCK_SEQ,       LBL_NO, " + Environment.NewLine +
                         "                       REQ_NO,            REQ_SEQ,          REQ_ERP_CD,          REQ_ERP_SEQ,         PLAN_SEQ, " + Environment.NewLine +
                         "                       REQ_DATE,          IN_TYPE,          BIGO,                UNIT,                CRT_IP, " + Environment.NewLine +
                         "                       CRT_PC,            CRT_MENU,         IN_DT,               TEST_NO,             TEST_REQ_NO) " + Environment.NewLine +
                         "              VALUES ('1001',               '1001000001', '"         + ITEM_NO   + "',     '" + LOT_NO         + "', '" + ITEM_QTY + "', " + Environment.NewLine +
                         "                      '" + MGM_STATUS + "', '" + USE_DEADLN + "', '" + IN_DATE   + "',     '" + MADEIN_NO      + "', '" + MADEIN_DATE + "'," + Environment.NewLine +
                         "                      '" + MADEIN_NM  + "', '" + VENDOR_NM  + "', '" + STOCK_SEQ + "',     '" + STOCK_SEQ      + "', '" + LBL_NO + "', " + Environment.NewLine +
                         "                      '',                   '',                   '',                      '',                   '',  " + Environment.NewLine +
                         "                      '',                   '29',                 '재고이관',              '" + UNIT + "',  '" + DBUtils.ClientIPAddress + "', " +  Environment.NewLine +
                         "                      '" + DBUtils.ClientPcName + "', 'PC', CONVERT(DATETIME, '" + IN_DATE + "'), '', '')";
                // W_STOCK에 INSERT
                DBUtils.InsertData(strSQL, transaction);

                // W_STOCK_HIST에 이력 INSERT
                using (SqlCommand command = new SqlCommand())
                {
                    command.Transaction = transaction;
                    command.Connection = DBUtils.sqlConnection;
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "A_LOG_STOCK";
                    command.Parameters.Add("@rCOMP", SqlDbType.NVarChar).Value = "100";
                    command.Parameters.Add("@rSTOCK_SEQ", SqlDbType.NVarChar).Value = STOCK_SEQ;
                    command.Parameters.Add("@rIO_KIND", SqlDbType.NVarChar).Value = "조정";
                    command.Parameters.Add("@rIO_NO", SqlDbType.NVarChar).Value = "";
                    command.Parameters.Add("@rIO_SEQ", SqlDbType.NVarChar).Value = "";
                    command.Parameters.Add("@rIO_MSG", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rQTY", SqlDbType.NVarChar).Value = ITEM_QTY;
                    command.Parameters.Add("@rLOC_FROM", SqlDbType.NVarChar).Value = "1001000001";
                    command.Parameters.Add("@rLOC_TO", SqlDbType.NVarChar).Value = "1001000001";

                    command.Parameters.Add("@rPGM_ID", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rPGM_NM", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rEVENT_NM", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rMSG", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rCRT_USR", SqlDbType.NVarChar).Value = "SYSTEM";
                    command.Parameters.Add("@rCRT_PC", SqlDbType.NVarChar).Value = "PC";
                    command.Parameters.Add("@rCRT_IP", SqlDbType.NVarChar).Value = "";
                    command.Parameters.Add("@rCRT_MENU", SqlDbType.NVarChar).Value = "PC";
                    command.Parameters.Add("@rRET_MSG", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output;
                    command.ExecuteNonQuery();
                }
                FileUtils.WriteLog("데이터 삽입 : " + ITEM_NO + ", " + LOT_NO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 작은따옴표가 있는 문자열의 경우 작은따옴표를 하나 더 붙임
        /// </summary>
        /// <param name="TargetString"></param>
        /// <returns></returns>
        private string withSingleQuotedString(string TargetString)
        {
            if (!TargetString.Contains("'")) return null;

            var StringArr = TargetString.Split('\'');
            string NewString = string.Empty;
            for(int i = 0; i < StringArr.Length - 1 ; i++)
            {
                NewString = NewString + StringArr[i] + "\'\'";
            }
            NewString = NewString + StringArr[StringArr.Length - 1];
            return NewString;
        }

        // 라벨발행 그리드 행 색상 변경
        private void gridView3_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            var view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            if (view.RowCount < 1) return;
            if (view.GetRowCellValue(e.RowHandle, "LBL_NO") == null) return;
            
            // 라벨번호 없는 것 (발행불가인 것) 색깔 다르게 표시
            if(view.GetRowCellValue(e.RowHandle, "LBL_NO").ToString().Contains("불가"))
            {
                e.Appearance.ForeColor = System.Drawing.Color.Red;
            }
        }

        // 라벨수량 일괄적용
        private void btnLBL_QTY_Click(object sender, EventArgs e)
        {
            if (txtLBL_QTY.Text == "") return;
            for (int i = 0; i < gridView3.RowCount; i++)
            {
                gridView3.SetRowCellValue(i, "LBL_QTY", txtLBL_QTY.Text);
            }
        }

        /// <summary>
        /// W_ITEM 테이블에서 단위를 가져옴
        /// </summary>
        /// <param name="ITEM_NO">품목코드</param>
        /// <param name="transaction">트랜잭션</param>
        /// <returns></returns>
        public string getUnit(string ITEM_NO, SqlTransaction transaction)
        {
            if (DBUtils.sqlConnection == null) return null;
            string strSQL = "SELECT UNIT " +
                            "  FROM W_ITEM " +
                            " WHERE ITEM_NO = '" + ITEM_NO + "'";
            DataTable table = DBUtils.getSelectResultTableOrNull(strSQL, transaction);
            string unit = table.Rows[0][0].ToString();

            return unit;
        }

        /// <summary>
        /// 라벨번호를 가져옴
        /// </summary>
        /// <param name="ITEM_NO">품목코드</param>
        /// <param name="LOT_NO">재고번호</param>
        /// <param name="transaction">트랜잭션</param>
        /// <returns>라벨번호</returns>
        public string getLabelNo(string ITEM_NO, string LOT_NO, SqlTransaction transaction)
        {
            if (DBUtils.sqlConnection == null) return null;
            try
            {
                using (SqlCommand command = new SqlCommand())
                {
                    command.Transaction = transaction;
                    command.Connection = DBUtils.sqlConnection;
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "A_GET_LABEL";
                    command.Parameters.Add("@rCOMP", SqlDbType.NVarChar).Value = "100";
                    command.Parameters.Add("@rITEM_NO", SqlDbType.NVarChar).Value = ITEM_NO;
                    command.Parameters.Add("@rLOT_NO", SqlDbType.NVarChar).Value = LOT_NO;

                    command.Parameters.Add("@rPGM_ID", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rPGM_NM", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rEVENT_NM", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rMSG", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rCRT_USR", SqlDbType.NVarChar).Value = "SYSTEM";
                    command.Parameters.Add("@rCRT_PC", SqlDbType.NVarChar).Value = "PC";
                    command.Parameters.Add("@rCRT_IP", SqlDbType.NVarChar).Value = "";
                    command.Parameters.Add("@rCRT_MENU", SqlDbType.NVarChar).Value = "PC";
                    command.Parameters.Add("@rRET_MSG", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output;
                    command.ExecuteNonQuery();

                    return command.Parameters["@rRET_MSG"].Value.ToString().Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// STOCK_SEQ 를 생성함
        /// </summary>
        /// <param name="transaction">트랜잭션</param>
        /// <returns>STOCK_SEQ</returns>
        public string getSeqNo(SqlTransaction transaction)
        {
            if (DBUtils.sqlConnection == null) return null;
            try
            {
                using (SqlCommand command = new SqlCommand())
                {
                    command.Transaction = transaction;
                    command.Connection = DBUtils.sqlConnection;
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "A_SEQ_GET";
                    command.Parameters.Add("@rCOMP", SqlDbType.NVarChar).Value = "100";
                    command.Parameters.Add("@rNEW_YN", SqlDbType.NVarChar).Value = "1"; // 생성여부 1
                    command.Parameters.Add("@rSEQ_TYPE", SqlDbType.NVarChar).Value = "W_STOCK"; // 테이블 명
                    command.Parameters.Add("@rSEQ_DATE", SqlDbType.NVarChar).Value = DateTime.Now.ToString("yyyyMMdd"); // 일자

                    command.Parameters.Add("@rPGM_ID", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rPGM_NM", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rEVENT_NM", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rMSG", SqlDbType.NVarChar).Value = "재고이관";
                    command.Parameters.Add("@rCRT_USR", SqlDbType.NVarChar).Value = "SYSTEM";
                    command.Parameters.Add("@rCRT_PC", SqlDbType.NVarChar).Value = DBUtils.ClientPcName;
                    command.Parameters.Add("@rCRT_IP", SqlDbType.NVarChar).Value = DBUtils.ClientIPAddress;
                    command.Parameters.Add("@rCRT_MENU", SqlDbType.NVarChar).Value = "PC";
                    command.Parameters.Add("@rRET_MSG", SqlDbType.NVarChar, 4000).Direction = ParameterDirection.Output;
                    command.ExecuteNonQuery();

                    return command.Parameters["@rRET_MSG"].Value.ToString().Trim();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        // 오른쪽 그리드에 데이터 복사
        public void AddDataToGrid2(int grid1RowHandle)
        {
            if (gridControl2.DataSource == null) return;

            DataRow dataRow = gridView1.GetDataRow(grid1RowHandle);
            tableForGrid2.ImportRow(dataRow);
            gridView2.RefreshData();
            Application.DoEvents();
        }

        // 그리드에 데이터 복사
        public void AddDataToGrid(int grid1RowHandle, DataTable table, DevExpress.XtraGrid.Views.Grid.GridView gridView)
        {
            DataRow dataRow = gridView1.GetDataRow(grid1RowHandle);
            table.ImportRow(dataRow);
            gridView.RefreshData();
            Application.DoEvents();
        }

        DataTable tableForGrid2;
        DataTable tableForGrid4;
        /* 메인폼 Load */
        private void frmMain_Load(object sender, EventArgs e)
        {
            if (DBUtils.sqlConnection != null) return;

            if(DBUtils.DatabaseConnect())
            {
                lblDBCon.Text = "DB접속";
                tableForGrid2 = new DataTable();
                foreach(var column in gridView1.Columns)
                {
                    tableForGrid2.Columns.Add((column as DevExpress.XtraGrid.Columns.GridColumn).FieldName);
                }
                gridControl2.DataSource = tableForGrid2;

                tableForGrid4 = new DataTable();
                foreach (var column in gridView1.Columns)
                {
                    tableForGrid4.Columns.Add((column as DevExpress.XtraGrid.Columns.GridColumn).FieldName);
                }
                gridControl4.DataSource = tableForGrid4;
            }
            else
            {
                lblDBCon.Text = "DB접속실패";
            }
        }
        
        /* 메인폼 Close */
        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (DBUtils.sqlConnection == null) return;

            DBUtils.DatabaseDisConnect();
        }

        /* 엑셀로 저장 */
        private void btnSaveExcel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("엑셀파일로 저장하시겠습니까?", "알림", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == DialogResult.No)
                return;

            var btn = (sender as DevExpress.XtraEditors.SimpleButton);

            switch(btn.Tag.ToString())
            {
                case "1": FileUtils.SaveToExcel("재고이관 중복자료", gridControl2); break;
                case "2": FileUtils.SaveToExcel("품목코드 없는자료", gridControl4); break;
            }
        }

        /* 행번호 */
        private void gridView1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.RowHandle < 0) return;

            e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }
    }
}
