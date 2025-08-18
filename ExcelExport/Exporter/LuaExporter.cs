using ExcelExport.Helper;
using ExcelExport.LitJson;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcelExport.Exporter
{
    public class LuaExporter : BaseExporter
    {
        protected override void CreateExportPath()
        {
            ExportHelper.VerifyPath(GetDataExportPath());
            ExportHelper.VerifyPath(GetLanguageDataExprotPath());
        }

        /// <summary>
        /// 导出数据
        /// </summary>
        protected override void ExportData(DataTable dt, string excelName, string sheetName)
        {
            string dataTableName = dt.Rows[1][0].ToString();
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("--excelName = {0}\r\n", excelName);
            sb.AppendFormat("--sheetName = {0}\r\n", sheetName);
            sb.Append("\r\n");
            sb.Append("local data = {\r\n");

            for (int i = 4; i < dt.Rows.Count; i++)
            {
                sb.AppendFormat("\t[{0}] = ", i - 3);
                sb.Append("{\r\n");

                for (int j = 1; j < dt.Columns.Count; j++)
                {
                    string fieldName = dt.Rows[0][j].ToString().Trim();
                    string fieldType = dt.Rows[1][j].ToString().Trim();
                    string fieldValue = dt.Rows[i][j].ToString().Trim();

                    string fieldStr = GetFieldStr(fieldName, fieldValue, fieldType);

                    if (!string.IsNullOrEmpty(fieldStr))
                    {
                        sb.AppendFormat("\t\t{0}\r\n", fieldStr);
                    }
                }

                sb.Append("\t},\r\n");
            }

            sb.Append("}\r\n");
            sb.Append("return data");

            File.WriteAllText(GetDataExportPath(GetConfigDataName(dataTableName)), sb.ToString());
        }

        /// <summary>
        /// 创建数据总表
        /// </summary>
        protected override void CreateConfigDataSheetScript()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("//--[[\r\n");
            sb.Append("数据总表\r\n");
            if (!string.IsNullOrEmpty(m_AuthorName))
            {
                sb.AppendFormat("作者：{0}", m_AuthorName);
                sb.Append("\r\n");
            }
            sb.AppendFormat("创建时间：{0}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.Append("备注：此代码为工具生成 请勿手工修改\r\n");
            sb.Append("]]\r\n");
            sb.Append("\r\nConfigDataSheet = {}\r\n");
            sb.AppendFormat("function ConfigDataSheet:Init(filePath)\r\n");

            for (int i = 0; i < m_DataTableNames.Count; i++)
            {
                if (!string.IsNullOrEmpty(m_DataTableNames[i]))
                {
                    sb.AppendFormat("\t\tself.{0}ConfigData = require(string.format(%s/%s,filePath,\"{0}ConfigData\"))\r\n", m_DataTableNames[i]);
                }
            }

            sb.Append("end");
            sb.Append('}');

            try
            {
                File.WriteAllText(GetDataExportPath("ConfigDataSheet.lua"), sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            sb.Clear();
        }

        protected override void ExportLanguageData(DataTable dt, string excelName, string sheetName)
        {
            string dataTableName = dt.Rows[1][0].ToString();
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("--excelName = {0}\r\n", excelName);
            sb.AppendFormat("--sheetName = {0}\r\n", sheetName);
            sb.Append("\r\nlocal data = {\r\n");

            for (int i = 4; i < dt.Rows.Count; i++)
            {
                string key = dt.Rows[i][2].ToString().Trim();
                string content = dt.Rows[i][3].ToString().Trim();

                sb.AppendFormat("\t[\"{0}\"] = \"{1}\"\r\n", key, content);
            }

            sb.Append("}\r\n");
            sb.Append("return data");

            File.WriteAllText(GetLanguageDataExprotPath(GetLanguageDataName(dataTableName)), sb.ToString());
        }

        protected override void CreateLanguageKeyFile(string content)
        {
            File.WriteAllText(GetLanguageDataExprotPath("LanguageKeys.txt"), content);
        }

        protected override void CreateLanguageContentFile(string content)
        {
            File.WriteAllText(GetLanguageDataExprotPath("LanguageContent.txt"), content);
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
                else if (fieldType.Contains("bool"))
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

                    if (i < vectorValues.Length - 1)
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
                sb.Append('\t');
            }

            sb.AppendFormat("{0}\n", GetFieldStr(fieldName, fieldValueStr, fieldType));
            return true;
        }

        private string GetDataExportPath(string fileName = "")
        {
            return string.Format("{0}Lua\\Datas\\{1}", m_ExportPath, fileName);
        }

        private string GetConfigDataName(string fileName)
        {
            return string.Format("{0}ConfigData.lua", fileName);
        }

        private string GetLanguageDataExprotPath(string fileName = "")
        {
            return string.Format("{0}Lua\\LanguageDatas\\{1}", m_ExportPath, fileName);
        }

        private string GetLanguageDataName(string fileName)
        {
            return string.Format("{0}LanguageData.lua", fileName);
        }
    }
}