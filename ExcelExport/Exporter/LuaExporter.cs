using LitJson;
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
                if (!string.IsNullOrEmpty(fieldValue))
                {
                    JsonData jsonData = LitJson.JsonMapper.ToObject(fieldValue);

                    if (jsonData != null)
                    {
                        StringBuilder jsonSB = new StringBuilder();
                        ParseJson(jsonData, jsonSB);
                        return string.Format("{0} = {1}\n{2}\t\t{3},", fieldName, "{", jsonSB.ToString(), "}");
                    }
                }

                return string.Format("{0} = nil,", fieldName);
            }

            if (string.IsNullOrEmpty(fieldValue))
            {
                fieldValue = fieldType.Contains("bool") ? "false" : "nil";
            }
            else if (fieldType.Contains("[]"))
            {
                string result = fieldName + " = {\n\t\t\t";
                string fieldValueTemp = fieldValue.Replace(" ", "").Replace(",", ",\n\t\t\t");

                if (fieldType.Contains("string"))
                {
                    fieldValueTemp = "\"" + fieldValue.Replace(" ", "").Replace(",", "\",\n\t\t\t\"") + "\"";
                }
                else if(fieldType.Contains("bool"))
                {
                    fieldValueTemp = fieldValueTemp.ToLower();
                }

                return result + fieldValueTemp + ",\n\t\t},";
            }
            else if (fieldType.Contains("Vector"))
            {
                string[] vectorValues = fieldValue.Split(',');
                string[] vectorFieldName = { "x", "y", "z" };
                string result = fieldName + " = {";

                for (int i = 0; i < vectorValues.Length; i++)
                {
                    result += string.Format("{0} = {1}", vectorFieldName[i], vectorValues[i]);

                    if(i < vectorValues.Length - 1)
                    {
                        result += ",";
                    }
                }
                

                return result + "},";
            }

            return string.Format("{0} = {1},", fieldName, fieldValue);
        }

        private void ParseJson(JsonData jsonData, StringBuilder sb, int tCount = 3)
        {
            if (jsonData.IsArray)
            {
                for (int i = 0; i < jsonData.Count; i++)
                {
                    if (!JsonFieldIsBaseValueType(string.Format("[{0}]", i + 1), jsonData[i], tCount, sb))
                    {
                        for (int j = 0; j < tCount; j++)
                        {
                            sb.Append("\t");
                        }

                        sb.AppendFormat("[{0}] = ", i + 1);
                        sb.Append("{\n");

                        ParseJson(jsonData[i], sb, tCount + 1);

                        for (int j = 0; j < tCount; j++)
                        {
                            sb.Append("\t");
                        }

                        sb.Append("},\n");
                    }
                }
            }
            else if (jsonData.Keys.Count > 0)
            {
                foreach (KeyValuePair<string, LitJson.JsonData> kvp in jsonData)
                {
                    string key = kvp.Key;
                    JsonData val = kvp.Value;

                    if (!JsonFieldIsBaseValueType(key, val, tCount, sb))
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

        private bool JsonFieldIsBaseValueType(string fieldName, JsonData jsonData, int tCount, StringBuilder sb)
        {
            string fieldValueStr = jsonData.ToString();
            string fieldType = string.Empty;

            if (jsonData.IsInt)
            {
                fieldType = "int";
            }
            else if (jsonData.IsLong)
            {
                fieldType = "long";
            }
            else if (jsonData.IsDouble)
            {
                fieldType = "double";
            }
            else if (jsonData.IsBoolean)
            {
                fieldType = "bool";
                fieldValueStr = fieldValueStr.ToLower();
            }
            else if (jsonData.IsString)
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