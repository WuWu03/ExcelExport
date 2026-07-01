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
        private List<DataTable> m_LanguageDataTables = null;
        protected string exportPath { get; private set; }
        protected string authorName { get; private set; }
        protected List<string> dataTableNames { get; private set; }
        protected List<string> excels { get; private set; }

        public abstract string exporterName { get; }

        public void SetExportPath(string exprotPath)
        {
            exportPath = exprotPath;
        }

        public void SetAuthorName(string authorName)
        {
            this.authorName = authorName;
        }

        public void AddExcel(string excelPath)
        {
            excels ??= new List<string>();
            excels.Add(excelPath);
        }

        public void ResetExcel()
        {
            excels?.Clear();
        }

        public void Export(List<bool> canExportList)
        {
            if (excels == null || excels.Count < 1)
            {
                return;
            }

            CreateExportPath();

            dataTableNames ??= new List<string>();
            m_LanguageDataTables ??= new List<DataTable>();

            dataTableNames.Clear();
            m_LanguageDataTables.Clear();

            for (int i = 0; i < excels.Count; i++)
            {

                if (canExportList != null && i < canExportList.Count && canExportList[i])
                {
                    Export(excels[i]);
                }
            }

            CreateConfigDataSheetScript();
            ExportLanguageKeys();
        }

        private void Export(string filePath)
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
                    bool isColNull = dt.Rows[0][col] == null || string.IsNullOrEmpty(dt.Rows[0][col].ToString());

                    if (col > 1 && (dt.Rows[3][col].ToString().ToLower().Equals("ban") || isColNull))
                    {
                        dt.Columns.RemoveAt(col);
                    }
                }

                string excelName = Path.GetFileName(filePath);
                string sheetName = dt.TableName;
                string dataTableName = dt.Rows[1][0].ToString();

                if (dt.Rows[3][0].ToString().ToLower().Equals("language"))
                {
                    m_LanguageDataTables.Add(dt);
                    ExportLanguageData(dt, excelName, sheetName);
                }
                else
                {
                    dataTableNames.Add(dataTableName);
                    ExportData(dt, excelName, sheetName);
                }
            }
        }

        private void ExportLanguageKeys()
        {
            if (m_LanguageDataTables == null)
            {
                return;
            }

            StringBuilder allKeysSB = new();
            StringBuilder allContentSB = new();

            bool hasAddAllKeys = false;
            for (int i = 0; i < m_LanguageDataTables.Count; i++)
            {
                DataTable dt = m_LanguageDataTables[i];

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

            allContentSB.Append("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ`1234567890-=~!@#$%^&*()_+[]\\{}|;':\",./<>?+-*/");
            try
            {
                CreateLanguageKeyFile(allKeysSB.ToString());
                CreateLanguageContentFile(allContentSB.ToString());
                allKeysSB.Clear();
                allContentSB.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //protected abstract void ExportLanguageFile(string path, string content);
        //protected abstract string GetLanguageDataExprotPath(string fileName = "");
        //protected abstract string GetLanguageDataName(string fileName, string ext);
        protected abstract void CreateExportPath();
        protected abstract void ExportData(DataTable dt, string excelName, string sheetName);
        protected abstract void ExportLanguageData(DataTable dt, string excelName, string sheetName);
        protected abstract void CreateConfigDataSheetScript();
        protected abstract void CreateLanguageKeyFile(string content);
        protected abstract void CreateLanguageContentFile(string content);
    }
}
