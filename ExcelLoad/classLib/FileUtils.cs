using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ExcelLoad.classLib
{
    static class FileUtils
    {
        private static string INIPath = Application.StartupPath + @"\Config.ini";

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);


        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        /// <summary>
        /// INI파일에 쓰기
        /// </summary>
        /// <param name="section">대괄호(섹션)</param>
        /// <param name="key">대괄호안의 Key</param>
        /// <param name="value">Key에 들어갈 값</param>
        public static void IniWrite(string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, INIPath);
        }

        /// <summary>
        /// INI파일로부터 읽기
        /// </summary>
        /// <param name="section">대괄호(섹션)</param>
        /// <param name="key">대괄호안의 Key</param>
        /// <returns>Key의 Value</returns>
        public static string IniRead(string section, string key)
        {
            StringBuilder stringBuilder = new StringBuilder(255);
            GetPrivateProfileString(section, key, "", stringBuilder, stringBuilder.Capacity, INIPath);

            return stringBuilder.ToString();
        }
        
        // 로그남기는 메서드
        public static void WriteLog(string message)
        {
            string LogFileName = String.Format("{0}.txt", DateTime.Now.ToString("yyyyMMdd"));
            string LogFilePath = Application.StartupPath + @"\log\";
            string CurrentTime = string.Format("{0:S10} {1:D2}:{2:D2}:{3:D2}", DateTime.Now.ToShortDateString(), DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
            string WriteData;

            WriteData = CurrentTime + " ==> " + message;
            //LogFile_Path = Check_LogFolder();
            //LogFile_Name = LogFile_Path + "\\" + LogFile_Name;
            string filePath = LogFilePath + LogFileName;
            DirectoryInfo directory = new DirectoryInfo(LogFilePath);
            directory.Create();
            
            if (!File.Exists(filePath))
            {
                FileStream LogFile = new FileStream(filePath, FileMode.Append, FileAccess.Write);

                StreamWriter SW_File = new StreamWriter(LogFile);
                SW_File.WriteLine(WriteData);
                SW_File.Close();
            }
        }

        public static void SaveToExcel(string fileName, DevExpress.XtraGrid.GridControl gridControl)
        {
            try
            {
                using (CommonSaveFileDialog dialog = new CommonSaveFileDialog())
                {
                    dialog.DefaultFileName = fileName + "_" + string.Format("{0:yyMMdd_HHmmss}", DateTime.Now);
                    dialog.DefaultExtension = "xlsx";
                    dialog.Filters.Add(new CommonFileDialogFilter("Excel 통합문서", "xlsx"));
                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        DevExpress.XtraPrinting.XlsxExportOptionsEx op = new DevExpress.XtraPrinting.XlsxExportOptionsEx();
                        op.AllowSortingAndFiltering = DevExpress.Utils.DefaultBoolean.False;
                        gridControl.ExportToXlsx(dialog.FileName, op);
                    }
                    else
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
    }
}
