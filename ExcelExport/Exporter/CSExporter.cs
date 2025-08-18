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
    public class CSExporter : BaseExporter
    {
        class JsonStruct
        {
            public string className;
            public Dictionary<string, string> fields = new Dictionary<string, string>();
            public List<JsonStruct> children;
        }

        protected override void CreateExportPath()
        {
            ExportHelper.VerifyPath(GetDataExportPath());
            ExportHelper.VerifyPath(GetScriptsExportPath());
            ExportHelper.VerifyPath(GetLanguageDataExprotPath());
        }

        protected override void ExportData(DataTable dt, string excelName, string sheetName)
        {
            string dataTableName = dt.Rows[1][0].ToString();
            byte[] buffer = GetDataBuffer(dt);
            //写入文件
            try
            {
                FileStream fs = new FileStream(GetDataExportPath(GetConfigDataName(dataTableName)), FileMode.Create);
                fs.Write(buffer, 0, buffer.Length);
                fs.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            CreateDataScript(dt, excelName, sheetName);
        }

        protected override void ExportLanguageData(DataTable dt, string excelName, string sheetName)
        {
            string dataTableName = dt.Rows[1][0].ToString();
            byte[] buffer = GetDataBuffer(dt);

            //写入文件
            try
            {
                File.WriteAllBytes(GetLanguageDataExprotPath(GetLanguageDataName(dataTableName)), buffer);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        protected override void CreateLanguageKeyFile(string content)
        {
            File.WriteAllText(GetLanguageDataExprotPath("LanguageKeys.txt"), content);
        }

        protected override void CreateLanguageContentFile(string content)
        {
            File.WriteAllText(GetLanguageDataExprotPath("LanguageContent.txt"), content);
        }

        /// <summary>
        /// 生成C#代码
        /// </summary>
        private void CreateDataScript(DataTable dt, string excelName, string sheetName)
        {
            string dataTableName = dt.Rows[1][0].ToString();
            string[,] dataArr = new string[dt.Columns.Count - 1, 3];

            for (int i = 0; i < 3; i++)
            {
                for (int j = 1; j < dt.Columns.Count; j++)
                {
                    dataArr[j - 1, i] = dt.Rows[i][j].ToString().Trim();
                }
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("//===================================================\r\n");

            if (!string.IsNullOrEmpty(m_AuthorName))
            {
                sb.AppendFormat("//作者：{0}", m_AuthorName);
                sb.Append("\r\n");
            }

            sb.AppendFormat("//创建时间：{0}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.Append("//备注：此代码为工具生成 请勿手工修改\r\n");
            sb.Append("//===================================================\r\n");
            sb.Append("using GameFrameWork;\r\n");
            sb.Append("using GameFrameWork.ConfigData;\r\n");
            sb.Append("using LitJson;\r\n");
            sb.Append("using System;\r\n");
            sb.Append("using System.Collections;\r\n");
            sb.Append("using UnityEngine;\r\n");
            sb.Append("\r\n");
            sb.Append("/// <summary>\r\n");
            sb.AppendFormat("/// {0}数据表\r\n", excelName);
            sb.AppendFormat("/// SheetName:{0}\r\n", sheetName);
            sb.Append("/// </summary>\r\n");
            sb.AppendFormat("public class {0}ConfigData : BaseConfigData\r\n", dataTableName);
            sb.Append("{\r\n");

            //生成Json实体类代码
            Dictionary<string, string> jsonDic = new Dictionary<string, string>();

            for (int i = 1; i < dataArr.GetLength(0); i++)
            {
                string typeName = GetTypeName(dataArr[i, 1]);

                if (typeName.Contains("json"))
                {
                    typeName = string.Concat(dataArr[i, 0][..1].ToUpper(), dataArr[i, 0].AsSpan(1));

                    for (int j = 4; j < dt.Rows.Count; j++)
                    {
                        string jsonStr = dt.Rows[j][i + 1].ToString();

                        if (!string.IsNullOrEmpty(jsonStr))
                        {
                            jsonDic.Add(typeName, jsonStr);
                            break;
                        }
                    }
                }
            }

            foreach (KeyValuePair<string, string> kvp in jsonDic)
            {
                JsonData obj = JsonMapper.ToObject(kvp.Value);
                JsonStruct jsonStruct = new JsonStruct
                {
                    className = kvp.Key
                };
                ParseJson(obj, jsonStruct);
                CreateJsonCode(jsonStruct, sb);
            }

            //生成字段代码
            for (int i = 1; i < dataArr.GetLength(0); i++)
            {
                string typeName = GetTypeName(dataArr[i, 1]);

                if (typeName.Contains("json"))
                {
                    string listSymbol = string.Empty;

                    if (typeName.Contains("[]"))
                    {
                        listSymbol = "[]";
                    }

                    typeName = string.Concat(dataArr[i, 0][..1].ToUpper(), dataArr[i, 0].AsSpan(1), listSymbol);
                }

                sb.Append("\t/// <summary>\r\n");
                sb.AppendFormat("\t/// {0}\r\n", dataArr[i, 2]);
                sb.Append("\t/// </summary>\r\n");
                sb.AppendFormat("\tpublic {0} {1} {{ get; private set; }}\r\n", typeName, dataArr[i, 0]);
                sb.Append("\r\n");
            }

            //生成克隆代码

            string variableName = string.Concat(dataTableName[..1].ToLower(), dataTableName.AsSpan(1));

            sb.AppendFormat("\tpublic {0}ConfigData Clone()\r\n", dataTableName);
            sb.Append("\t{\r\n");
            sb.AppendFormat("\t\t{0}ConfigData {1}ConfigData = new {2}ConfigData();\r\n", dataTableName, variableName, dataTableName);

            for (int i = 1; i < dataArr.GetLength(0); i++)
            {
                sb.AppendFormat("\t\t{0}ConfigData.{1} = this.{2};", variableName, dataArr[i, 0], dataArr[i, 0]);
                sb.Append("\r\n");
            }

            sb.AppendFormat("\t\treturn {0}ConfigData;\r\n", variableName);
            sb.Append("\t}\r\n");
            sb.Append("\r\n");

            //生成解析代码
            sb.AppendFormat("\tpublic override void Read(ConfigDataParser parser)\r\n");
            sb.Append("\t{\r\n");

            for (int i = 0; i < dataArr.GetLength(0); i++)
            {
                if (string.IsNullOrEmpty(dataArr[i, 0]))
                {
                    continue;
                }

                string fieldName = string.Concat(dataArr[i, 0][..1].ToLower(), dataArr[i, 0].AsSpan(1));
                string typeName = GetTypeName(dataArr[i, 1]);

                if (typeName.Contains("json"))
                {
                    string listSymbol = string.Empty;

                    if (typeName.Contains("[]"))
                    {
                        listSymbol = "[]";
                    }

                    typeName = string.Concat(dataArr[i, 0][..1].ToUpper(), dataArr[i, 0].AsSpan(1), listSymbol);
                    sb.AppendFormat("\t\tthis.{0} = JsonMapper.ToObject<{1}>(parser.GetFieldValue(\"{0}\"));\r\n", fieldName, typeName);
                }
                else
                {
                    sb.AppendFormat("\t\tthis.{0} = parser.GetFieldValue(\"{0}\"){1};\r\n", fieldName, GetTypeParseStr(typeName));
                }
            }

            sb.Append("\t}\r\n");
            sb.Append('}');

            try//写入文件
            {
                File.WriteAllText(GetScriptsExportPath(GetScriptName(dataTableName)), sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 创建数据总表
        /// </summary>
        protected override void CreateConfigDataSheetScript()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("//===================================================\r\n");
            if (!string.IsNullOrEmpty(m_AuthorName))
            {
                sb.AppendFormat("//作者：{0}", m_AuthorName);
                sb.Append("\r\n");
            }
            sb.AppendFormat("//创建时间：{0}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.Append("//备注：此代码为工具生成 请勿手工修改\r\n");
            sb.Append("//===================================================\r\n");
            sb.Append("using System.Collections;\r\n");
            sb.Append("using GameFrameWork.ConfigData;\r\n");
            sb.Append("\r\n");
            sb.Append("/// <summary>\r\n");
            sb.Append("///数据总表\r\n");
            sb.Append("/// </summary>\r\n");
            sb.AppendFormat("public static class ConfigDataSheet\r\n");
            sb.Append("{\r\n");

            for (int i = 0; i < m_DataTableNames.Count; i++)
            {
                if (!string.IsNullOrEmpty(m_DataTableNames[i]))
                {
                    string fieldName = string.Concat(m_DataTableNames[i][..1].ToLower(), m_DataTableNames[i].AsSpan(1));
                    sb.AppendFormat("\tpublic static {0}ConfigData[] {1}ConfigDatas = null;", m_DataTableNames[i], fieldName);
                    sb.Append("\r\n");
                }
            }

            sb.Append("\r\n");
            sb.Append("\tpublic static void Init(string filePath)\r\n");
            sb.Append("\t{\r\n");

            for (int i = 0; i < m_DataTableNames.Count; i++)
            {
                if (!string.IsNullOrEmpty(m_DataTableNames[i]))
                {
                    string fieldName = string.Concat(m_DataTableNames[i][..1].ToLower(), m_DataTableNames[i].AsSpan(1));
                    sb.AppendFormat("\t\t{0}ConfigDatas = LoadConfigData<{1}ConfigData>(filePath, \"{2}ConfigData\");\r\n", fieldName, m_DataTableNames[i], m_DataTableNames[i]);
                }
            }

            sb.Append("\t}\r\n");
            sb.Append('}');

            try
            {
                File.WriteAllText(GetScriptsExportPath("ConfigDataSheet.cs"), sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            sb.Clear();
        }

        private byte[] GetDataBuffer(DataTable dt)
        {
            byte[] buffer = null;

            using (MemoryStreamEx mse = new MemoryStreamEx())
            {
                mse.WriteInt(dt.Rows.Count - 3);
                mse.WriteInt(dt.Columns.Count - 1);

                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    mse.WriteUTF8String(dt.Rows[0][i].ToString().Trim());
                }

                for (int i = 4; i < dt.Rows.Count; i++)
                {
                    for (int j = 1; j < dt.Columns.Count; j++)
                    {
                        mse.WriteUTF8String(dt.Rows[i][j].ToString().Trim());
                    }
                }

                buffer = mse.ToArray();
            }

            //压缩
            buffer = ZlibHelper.CompressBytes(buffer);
            return buffer;
        }

        private void ParseJson(LitJson.JsonData jsonData, JsonStruct jsonStruct)
        {
            if (jsonData.IsArray)
            {
                for (int i = 0; i < jsonData.Count; i++)
                {
                    ParseJson(jsonData[i], jsonStruct);
                }
            }
            else if (jsonData.Keys.Count > 0)
            {
                foreach (KeyValuePair<string, LitJson.JsonData> kvp in jsonData)
                {
                    string key = kvp.Key.Trim();
                    LitJson.JsonData val = kvp.Value;
                    string fieldType = JsonFieldType(val);
                    string fieldName = string.Concat(key[..1].ToLower(), key.AsSpan(1));

                    if (string.IsNullOrEmpty(fieldType))
                    {
                        fieldType = string.Concat(key[..1].ToUpper(), key.AsSpan(1));

                        jsonStruct.children ??= new List<JsonStruct>();

                        JsonStruct childJsonStruct = null;

                        for (int i = 0; i < jsonStruct.children.Count; i++)
                        {
                            if (jsonStruct.children[i].className.Equals(fieldType))
                            {
                                childJsonStruct = jsonStruct.children[i];
                                break;
                            }
                        }

                        if (childJsonStruct == null)
                        {
                            childJsonStruct = new JsonStruct
                            {
                                className = fieldType
                            };
                            jsonStruct.children.Add(childJsonStruct);
                        }

                        if (!jsonStruct.fields.ContainsKey(fieldName))
                        {
                            jsonStruct.fields.Add(fieldName, fieldType + (val.IsArray ? "[]" : string.Empty));
                            ParseJson(val, childJsonStruct);
                        }
                    }
                    else jsonStruct.fields.TryAdd(fieldName, fieldType);
                }
            }
        }

        private string JsonFieldType(LitJson.JsonData fieldValue)
        {
            string fieldType = string.Empty;

            if (fieldValue.IsInt)
            {
                fieldType = "int";
            }
            else if (fieldValue.IsLong)
            {
                fieldType = "long";
            }
            else if (fieldValue.IsDouble)
            {
                if (fieldValue.ToString().Split('.')[1].Length < 7)
                {
                    fieldType = "float";
                }
                else
                {
                    fieldType = "double";
                }
            }
            else if (fieldValue.IsBoolean)
            {
                fieldType = "bool";
            }
            else if (fieldValue.IsString)
            {
                fieldType = "string";
            }
            else if (fieldValue.IsArray && fieldValue.Count > 0)
            {
                string typeTemp = JsonFieldType(fieldValue[0]);
                if (!string.IsNullOrEmpty(typeTemp))
                {
                    return typeTemp + "[]";
                }
            }

            return fieldType;
        }

        private void CreateJsonCode(JsonStruct jsonStruct, StringBuilder sb, int tCount = 1)
        {
            for (int i = 0; i < tCount; i++)
            {
                sb.Append('\t');
            }
            sb.AppendFormat("public class {0}\r\n", jsonStruct.className);

            for (int i = 0; i < tCount; i++)
            {
                sb.Append('\t');
            }

            sb.Append("{\r\n");

            if (jsonStruct.children != null && jsonStruct.children.Count > 0)
            {
                for (int i = 0; i < jsonStruct.children.Count; i++)
                {
                    CreateJsonCode(jsonStruct.children[i], sb, tCount + 1);
                }

                sb.Append("\r\n");
            }

            foreach (KeyValuePair<string, string> kvp in jsonStruct.fields)
            {
                for (int i = 0; i < tCount + 1; i++)
                {
                    sb.Append('\t');
                }

                sb.AppendFormat("public {0} {1} {{ get; set; }}\r\n", kvp.Value, kvp.Key);
            }

            for (int i = 0; i < tCount; i++)
            {
                sb.Append('\t');
            }

            sb.Append("}\r\n");
        }

        private string GetTypeName(string typeName)
        {
            return typeName.ToLower() switch
            {
                "int" => "int",
                "long" => "long",
                "float" => "float",
                "double" => "double",
                "bool" => "bool",
                "string" => "string",
                "vector2" => "Vector2",
                "vector3" => "Vector3",
                "int[]" => "int[]",
                "long[]" => "long[]",
                "float[]" => "float[]",
                "double[]" => "double[]",
                "bool[]" => "bool[]",
                "string[]" => "string[]",
                "json" => "json",
                "json[]" => "json[]",
                _ => string.Empty,
            };
        }

        private string GetTypeParseStr(string typeName)
        {
            return typeName.ToLower() switch
            {
                "int" => ".ToInt()",
                "long" => ".ToLong()",
                "float" => ".ToFloat()",
                "double" => ".ToDouble()",
                "bool" => ".ToBool()",
                "string" => string.Empty,
                "vector2" => ".ToVector2()",
                "vector3" => ".ToVector3()",
                "int[]" => ".ToIntArray()",
                "long[]" => ".ToLongArray()",
                "float[]" => ".ToFloatArray()",
                "double[]" => ".ToDoubleArray()",
                "bool[]" => ".ToBoolArray()",
                "string[]" => ".ToStringArray()",
                _ => string.Empty,
            };
        }

        private string GetDataExportPath(string fileName = "")
        {
            return string.Format("{0}C#\\Datas\\{1}", m_ExportPath, fileName);
        }

        private string GetScriptsExportPath(string fileName = "")
        {
            return string.Format("{0}C#\\Scripts\\{1}", m_ExportPath, fileName);
        }

        private string GetConfigDataName(string fileName)
        {
            return string.Format("{0}ConfigData.bytes", fileName);
        }

        private string GetScriptName(string fileName)
        {
            return string.Format("{0}ConfigData.cs", fileName);
        }

        private string GetLanguageDataExprotPath(string fileName = "")
        {
            return string.Format("{0}C#\\LanguageDatas\\{1}", m_ExportPath, fileName);
        }

        private string GetLanguageDataName(string fileName)
        {
            return string.Format("{0}LanguageData.bytes", fileName);
        }
    }
}
