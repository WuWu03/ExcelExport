using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Exporter
{
    public class LuaExporter : BaseExporter
    {
        protected override void CreateExportPath()
        {
            if (Directory.Exists(string.Format("{0}/Lua", m_ExportPath)))
            {
                Directory.Delete(string.Format("{0}/Lua", m_ExportPath), true);
                Directory.CreateDirectory(string.Format("{0}/Lua", m_ExportPath));
            }
            else
            {
                Directory.CreateDirectory(string.Format("{0}/Lua", m_ExportPath));
            }
        }

        /// <summary>
        /// 导出数据
        /// </summary>
        protected override void ExportData(DataTable dt, string excelName, string sheetName)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("local data = {\n");

            for (int i = 4; i < dt.Rows.Count; i++)
            {
                sb.AppendFormat("\t[{0}] = ", i - 3);
                sb.Append("{\n");

                for (int j = 1; j < dt.Columns.Count; j++)
                {
                    string fieldName = dt.Rows[0][j].ToString().Trim();
                    string fieldType = dt.Rows[1][j].ToString().Trim();
                    string fieldValue = dt.Rows[i][j].ToString().Trim();
                    string fieldStr = GetFieldStr(fieldName, fieldValue, fieldType);

                    if (!string.IsNullOrEmpty(fieldStr))
                    {
                        sb.AppendFormat("\t\t{0}\n", fieldStr);
                    }
                }

                sb.Append("\t},\n");
            }

            sb.Append("}\n");
            sb.AppendFormat("--excelName = {0}\n", excelName);
            sb.AppendFormat("--sheetName = {0}\n", sheetName);
            sb.Append("return data");

            File.WriteAllText(string.Format("{0}/Lua/{1}Data.lua", m_ExportPath, dt.Rows[1][0].ToString()), sb.ToString());
        }

        /// <summary>
        /// 创建数据总表
        /// </summary>
        protected override void CreateDataHelperScript()
        {
            
        }

        private string GetFieldStr(string fieldName, string fieldValue, string fieldType)
        {
            if (fieldType.Equals("string"))
            {
                return string.Format("{0} = \"{1}\",", fieldName, fieldValue);
            }
            else if (fieldType.Equals("json"))
            {
                if (string.IsNullOrEmpty(fieldValue))
                {
                    return string.Empty;
                }

                object jsonObj = LitJson.JsonMapper.ToObject(fieldValue);

                if (jsonObj == null)
                {
                    return string.Empty;
                }

                StringBuilder jsonSB = new StringBuilder();
                ParseJson(jsonObj, jsonSB);
                return string.Format("{0} = {1}\n{2}\n\t\t{3},", fieldName, "{", jsonSB.ToString(), "}");
            }

            return string.Format("{0} = {1},", fieldName, fieldValue);
        }

        private void ParseJson(object json, StringBuilder sb, int tCount = 3)
        {
            if (json is List<object> jsonList)
            {
                for (int i = 0; i < jsonList.Count; i++)
                {
                    if (!JsonFieldIsBaseValueType(string.Format("[{0}]", i + 1), jsonList[i], tCount, sb))
                    {
                        for (int j = 0; j < tCount; j++)
                        {
                            sb.Append("\t");
                        }

                        sb.AppendFormat("[{0}] = ", i + 1);
                        sb.Append("{\n");

                        ParseJson(jsonList[i], sb, tCount + 1);

                        for (int j = 0; j < tCount; j++)
                        {
                            sb.Append("\t");
                        }

                        sb.Append("},\n");
                    }
                }
            }
            else if (json is Dictionary<string, object> jsonDic)
            {
                foreach (KeyValuePair<string, object> kvp in jsonDic)
                {
                    string key = kvp.Key;
                    object val = kvp.Value;

                    if (!JsonFieldIsBaseValueType(key, val, tCount,sb))
                    {
                        for (int i = 0; i < tCount; i++)
                        {
                            sb.Append("\t");
                        }

                        sb.AppendFormat("{0} = ", key);
                        sb.Append("{\n");

                        ParseJson(val, sb, tCount + 1);

                        for (int i = 0; i < tCount; i++)
                        {
                            sb.Append("\t");
                        }

                        sb.Append("},\n");
                    }
                }
            }
        }

        private bool JsonFieldIsBaseValueType(string fieldName,object fieldValue,int tCount, StringBuilder sb)
        {
            string fieldValueStr = fieldValue.ToString();
            string fieldType = string.Empty;

            if (fieldValue is int)
            {
                fieldType = "int";
            }
            else if (fieldValue is long)
            {
                fieldType = "long";
            }
            else if (fieldValue is float)
            {
                fieldType = "float";
            }
            else if (fieldValue is double)
            {
                fieldType = "double";
            }
            else if (fieldValue is bool)
            {
                fieldType = "bool";
                fieldValueStr = fieldValueStr.ToLower();
            }
            else if (fieldValue is string)
            {
                fieldType = "string";
            }

            if (string.IsNullOrEmpty(fieldType))
            {
                return false;
            }

            for (int i = 0; i < tCount; i++)
            {
                sb.Append("\t");
            }

            sb.AppendFormat("{0}\n", GetFieldStr(fieldName, fieldValueStr, fieldType));
            return true;
        }

    }
}