using ExcelExport.Helper;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcelExport.Exporter
{
    public abstract class BaseExporter
    {
        public void SetExportPath(string exprotPath)
        {
            m_ExportPath = exprotPath;
        }

        public void AddExcel(string excelPath)
        {
            if (m_ExcelList == null)
            {
                m_ExcelList = new List<string>();
            }

            m_ExcelList.Add(excelPath);
        }

        public void ResetExcel()
        {
            if (m_ExcelList != null)
            {
                m_ExcelList.Clear();
            }
        }

        public void Export(List<bool> canExportList)
        {
            if (m_ExcelList == null || m_ExcelList.Count < 1)
            {
                return;
            }

            if (Directory.Exists(GetLanguageDataExprotPath()))
            {
                Directory.Delete(GetLanguageDataExprotPath(), true);
                Directory.CreateDirectory(GetLanguageDataExprotPath());
            }
            else
            {
                Directory.CreateDirectory(GetLanguageDataExprotPath());
            }

            CreateExportPath();


            if (m_DataTableNameList == null)
            {
                m_DataTableNameList = new List<string>();
            }

            if (m_ListLanguageDataTable == null)
            {
                m_ListLanguageDataTable = new List<DataTable>();
            }

            m_DataTableNameList.Clear();
            m_ListLanguageDataTable.Clear();

            for (int i = 0; i < m_ExcelList.Count; i++)
            {

                if (canExportList != null && i < canExportList.Count && canExportList[i])
                {
                    BeforeExport(m_ExcelList[i]);
                }
            }

            CreateConfigDataSheetScript();
            ExportLanguageKeys();
        }

        private void BeforeExport(string filePath)
        {
            DataTable[] dts = ExcelHelper.ExcelToTable(filePath);
            string allLanguageContent = string.Empty;
            if (dts == null || dts.Length < 1)
            {
                return;
            }

            for (int i = 0; i < dts.Length; i++)
            {
                DataTable dt = dts[i];

                if (dt.Rows.Count < 4 || dt.Columns.Count < 1)
                {
                    continue;
                }

                if (dt.Rows[3][0].ToString().ToLower().Equals("ban"))
                {
                    continue;
                }

                //每行第一列如果填入BAN则此行不导出
                for (int row = dt.Rows.Count - 1; row > 3; row--)
                {
                    if (dt.Rows[row][0].ToString().ToLower().Equals("ban"))
                    {
                        dt.Rows.RemoveAt(row);
                    }
                }

                //每列的第三行如果填入BAN则此列不导出(第一列为id，强制导出)
                for (int col = dt.Columns.Count - 1; col > 0; col--)
                {
                    if (col > 1 && dt.Rows[3][col].ToString().ToLower().Equals("ban"))
                    {
                        dt.Columns.RemoveAt(col);
                    }
                }

                string excelName = Path.GetFileName(filePath);
                string sheetName = dt.TableName;
                string dataTableName = dt.Rows[1][0].ToString();

                if (dt.Rows[3][0].ToString().ToLower().Equals("language"))
                {
                    m_ListLanguageDataTable.Add(dt);
                    ExportLanguageData(dt, excelName, sheetName);
                }
                else
                {
                    m_DataTableNameList.Add(dataTableName);
                    ExportData(dt, excelName, sheetName);
                }
            }
        }

        private void ExportLanguageKeys()
        {
            if (m_ListLanguageDataTable == null)
            {
                return;
            }

            StringBuilder allKeysSB = new StringBuilder();
            StringBuilder allContentSB = new StringBuilder();

            bool hasAddAllKeys = false;
            for (int i = 0; i < m_ListLanguageDataTable.Count; i++)
            {
                DataTable dt = m_ListLanguageDataTable[i];

                for (int j = 4; j < dt.Rows.Count; j++)
                {
                    if (!hasAddAllKeys)
                    {
                        string keyName = dt.Rows[j][2].ToString().Trim();

                        if (j < dt.Rows.Count - 1)
                        {
                            allKeysSB.AppendLine(keyName);
                        }
                        else
                        {
                            allKeysSB.Append(keyName);
                        }
                    }

                    allContentSB.Append(dt.Rows[j][3].ToString().Trim());
                }

                hasAddAllKeys = true;
            }

            allContentSB.Append("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_+[]{}|;':\",.<>?/`~");
            try
            {
                ExportLanguageFile(GetLanguageDataExprotPath(GetLanguageKeysName()), allKeysSB.ToString());
                ExportLanguageFile(GetLanguageDataExprotPath(GetLanguageContentName()), allContentSB.ToString());
                allKeysSB.Clear();
                allContentSB.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ExportLanguageFile(string path, string content)
        {
            using FileStream fs = new FileStream(path, FileMode.Create);
            using StreamWriter sw = new StreamWriter(fs);
            sw.Write(content);
        }


        protected string GetLanguageDataExprotPath(string fileName = "")
        {
            return string.Format("{0}/{1}/LanguageDatas/{2}", m_ExportPath, m_ExportRootPath, fileName);
        }

        protected string GetLanguageDataName(string fileName, string ext)
        {
            return string.Format("{0}LanguageData{1}", fileName, ext);
        }

        private string GetLanguageKeysName()
        {
            return "LanguageKeys.txt";
        }

        private string GetLanguageContentName()
        {
            return "LanguageContent.txt";
        }

        protected abstract void CreateExportPath();
        protected abstract void ExportData(DataTable dt, string excelName, string sheetName);
        protected abstract void ExportLanguageData(DataTable dt, string excelName, string sheetName);
        protected abstract void CreateConfigDataSheetScript();

        protected string m_ExportPath = string.Empty;

        protected List<string> m_DataTableNameList = null;
        protected List<string> m_ExcelList = null;

        private List<DataTable> m_ListLanguageDataTable = null;

        protected const string m_ExportRootPath = "DataExport";
    }
}
