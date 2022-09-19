using ExcelExport.Helper;
using System.Collections.Generic;
using System.Data;
using System.IO;

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


                if (dt.Rows[3][0].ToString().Contains("BAN"))
                {
                    continue;
                }

                string excelName = Path.GetFileName(filePath);
                string sheetName = dt.TableName;
                string dataTableName = dt.Rows[1][0].ToString();
         
                m_DataTableNameList.Add(dataTableName);

                //每行第一列如果填入BAN则此行不导出
                for (int row = dt.Rows.Count - 1; row > 3; row--)
                {
                    if (dt.Rows[row][0].ToString().Contains("BAN"))
                    {
                        dt.Rows.RemoveAt(row);
                    }
                }

                //每列的第三行如果填入BAN则此列不导出(第一列为id，强制导出)
                for (int col = dt.Columns.Count - 1; col > 0; col--)
                {
                    if (col > 1 && dt.Rows[3][col].ToString().Contains("BAN"))
                    {
                        dt.Columns.RemoveAt(col);
                    }
                }

                ExportData(dt, excelName, sheetName);
            }
        }

        protected abstract void CreateExportPath();
        protected abstract void ExportData(DataTable dt, string excelName, string sheetName);
        protected abstract void CreateDataHelperScript();

        protected string m_ExportPath = string.Empty;
        protected List<string> m_DataTableNameList = null;
        protected List<string> m_ExcelList = null;
    }
}
