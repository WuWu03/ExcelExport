using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelExport.Helper
{
    public static class ExcelHelper
    {
        /// <summary>
        /// 根据文件路径读取Excel文件
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static DataTable[] ExcelToTable(string file)
        {
            IWorkbook workbook = null;
            string fileExt = Path.GetExtension(file).ToLower();

            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (fileExt == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileExt == ".xls")
                {
                    workbook = new HSSFWorkbook(fs);
                }

                if (workbook == null) 
                { 
                    return null; 
                }

                if (workbook.NumberOfSheets < 1)
                {
                    return null;
                }

                DataTable[] dts = new DataTable[workbook.NumberOfSheets];//一张excel中可能有许多张表，要把这些表全部读出来

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    dts[i] = new DataTable();
                    ISheet sheet = workbook.GetSheetAt(i);
                    IRow header = sheet.GetRow(sheet.FirstRowNum);

                    if (header == null)
                    {
                        continue;
                    }

                    List<int> columns = new List<int>();

                    for (int j = 0; j < header.LastCellNum; j++)
                    {
                        object obj = GetValueType(header.GetCell(j));
                        if (obj == null || obj.ToString() == string.Empty)//中间出现空列也要读取
                        {
                            dts[i].Columns.Add(new DataColumn("Columns" + j.ToString()));
                        }
                        else
                        {
                            dts[i].Columns.Add(new DataColumn(obj.ToString()));
                        }

                        columns.Add(j);
                    }

                    //数据
                    int rowIndex = 0;
                    for (int j = sheet.FirstRowNum; j <= sheet.LastRowNum; j++)
                    {
                        DataRow dr = dts[i].NewRow();
                        bool hasValue = false;
                        foreach (int k in columns)
                        {
                            IRow row = sheet.GetRow(j); 

                            if(row != null)
                            {
                                dr[k] = GetValueType(row.GetCell(k));
                            }
                      
                            if (dr[k] != null && dr[k].ToString() != string.Empty)
                            {
                                hasValue = true;
                            }
                        }

                        if (hasValue || rowIndex == 3)//第4行为标记行，可能为空
                        {
                            dts[i].Rows.Add(dr);
                        }

                        rowIndex++;
                    }

                    dts[i].TableName = sheet.SheetName;
                }

                return dts;
            }
        }

        private static object GetValueType(ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            switch (cell.CellType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                case CellType.Formula:
                    return GetCachedFormulaResult(cell);
                default:
                    return string.Empty;
            }
        }

        private static object GetCachedFormulaResult(ICell cell)
        {
            switch (cell.CachedFormulaResultType)
            {
                case CellType.Unknown:
                    return null;
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                default:
                    return null;
            }
        }
    }
}
