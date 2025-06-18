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

            CreateExportPath();


            if (m_DataTableNameList == null)
            {
                m_DataTableNameList = new List<string>();
            }

            m_DataTableNameList.Clear();


            for (int i = 0; i < m_ExcelList.Count; i++)
            {

                if (canExportList != null && i < canExportList.Count && canExportList[i])
                {
                    ExportData(m_ExcelList[i]);
                }
            }

            CreateDataHelperScript();
            ExportLanguageKeys();
        }

        private void ExportData(string filePath)
        {
            DataTable[] dts = ExcelHelper.ExcelToTable(filePath);

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
                    m_LanguageDataTable = dt;
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
            if (m_LanguageDataTable == null)
            {
                return;
            }

            StringBuilder sb = new StringBuilder();

            for (int i = 4; i < m_LanguageDataTable.Rows.Count; i++)
            {
                string keyName = m_LanguageDataTable.Rows[i][2].ToString().Trim();

                if (i < m_LanguageDataTable.Rows.Count - 1)
                {
                    sb.AppendLine(keyName);
                }
                else
                {
                    sb.Append(keyName);
                }
            }

            try
            {
                using FileStream fs = new FileStream(string.Format("{0}/C#/Data/LanguageKeys.txt", m_ExportPath), FileMode.Create);
                using StreamWriter sw = new StreamWriter(fs);
                sw.Write(sb.ToString());
                sb.Clear();
                m_LanguageDataTable = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        protected abstract void CreateExportPath();
        protected abstract void ExportData(DataTable dt, string excelName, string sheetName);
        protected abstract void ExportLanguageData(DataTable dt, string excelName, string sheetName);
        protected abstract void CreateDataHelperScript();

        protected string m_ExportPath = string.Empty;

        protected List<string> m_DataTableNameList = null;
        protected List<string> m_ExcelList = null;

        private DataTable m_LanguageDataTable = null;
    }
}
