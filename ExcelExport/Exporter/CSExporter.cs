using ExcelExport.Helper;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Exporter
{
    public class CSExporter : BaseExporter
    {
        protected override void CreateExportPath()
        {
            if (Directory.Exists(string.Format("{0}/C#/Data", m_ExportPath)))
            {
                Directory.Delete(string.Format("{0}/C#/Data", m_ExportPath), true);
                Directory.CreateDirectory(string.Format("{0}/C#/Data", m_ExportPath));
            }
            else
            {
                Directory.CreateDirectory(string.Format("{0}/C#/Data", m_ExportPath));
            }

            if (Directory.Exists(string.Format("{0}/C#/Script", m_ExportPath)))
            {
                Directory.Delete(string.Format("{0}/C#/Script", m_ExportPath), true);
                Directory.CreateDirectory(string.Format("{0}/C#/Script", m_ExportPath));
            }
            else
            {
                Directory.CreateDirectory(string.Format("{0}/C#/Script", m_ExportPath));
            }
        }

        protected override void ExportData(DataTable dt, string excelName, string sheetName)
        {
            byte[] buffer = null;
            string dataTableName = dt.Rows[1][0].ToString();

            using (MemoryStreamEx mse = new MemoryStreamEx())
            {
                mse.WriteInt(dt.Rows.Count - 3);
                mse.WriteInt(dt.Columns.Count - 1);

                for (int j = 1; j < dt.Columns.Count; j++)
                {
                    mse.WriteUTF8String(dt.Rows[0][j].ToString().Trim());
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

            //写入文件
            FileStream fs = new FileStream(string.Format("{0}/C#/Data/{1}Data.bytes", m_ExportPath, dataTableName), FileMode.Create);
            fs.Write(buffer, 0, buffer.Length);
            fs.Close();

            string[,] dataArr = new string[dt.Columns.Count - 1, 3];

            for (int i = 0; i < 3; i++)
            {
                for (int j = 1; j < dt.Columns.Count; j++)
                {
                    dataArr[j - 1, i] = dt.Rows[i][j].ToString().Trim();
                }
            }

            CreateDataScript(dataArr, dt, excelName, sheetName, dataTableName);
        }

        /// <summary>
        /// 生成C#代码
        /// </summary>
        private void CreateDataScript(string[,] dataArr, DataTable dt, string excelName, string sheetName, string dataTableName)
        {
            if (dataArr == null)
            {
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n");
            sb.Append("//===================================================\r\n");
            sb.Append("//作者：GQY                                          \r\n");
            sb.AppendFormat("//创建时间：{0}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.Append("//备注：此代码为工具生成 请勿手工修改\r\n");
            sb.Append("//===================================================\r\n");
            sb.Append("using GameFrameWork;\r\n");
            sb.Append("using GameFrameWork.LocalData;\r\n");
            sb.Append("using LitJson;\r\n");
            sb.Append("using System;\r\n");
            sb.Append("using System.Collections;\r\n");
            sb.Append("using UnityEngine;\r\n");
            sb.Append("\r\n");
            sb.Append("/// <summary>\r\n");
            sb.AppendFormat("/// {0}数据表\r\n", excelName);
            sb.AppendFormat("/// SheetName:{0}\r\n", sheetName);
            sb.Append("/// </summary>\r\n");
            sb.AppendFormat("public class {0}Data : BaseLocalData\r\n", dataTableName);
            sb.Append("{\r\n");

            //生成Json实体类代码
            Dictionary<string, string> jsonDic = new Dictionary<string, string>();

            for (int i = 1; i < dataArr.GetLength(0); i++)
            {
                string typeName = dataArr[i, 1];

                if (typeName.Equals("json"))
                {
                    typeName = dataArr[i, 0].Substring(0, 1).ToUpper() + dataArr[i, 0].Substring(1);

                    for (int j = 3; j < dt.Rows.Count; j++)
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
                LitJson.JsonData obj = LitJson.JsonMapper.ToObject(kvp.Value);
                JsonStruct jsonStruct = new JsonStruct();
                jsonStruct.className = kvp.Key;
                ParseJson(obj, jsonStruct);
                CreateJsonCode(jsonStruct, sb);
            }

            //生成字段代码
            for (int i = 1; i < dataArr.GetLength(0); i++)
            {
                string typeName = dataArr[i, 1];

                if (typeName.Equals("json"))
                {
                    typeName = dataArr[i, 0].Substring(0, 1).ToUpper() + dataArr[i, 0].Substring(1);
                }

                sb.Append("\t/// <summary>\r\n");
                sb.AppendFormat("\t/// {0}\r\n", dataArr[i, 2]);
                sb.Append("\t/// </summary>\r\n");
                sb.AppendFormat("\tpublic {0} {1} {{ get; private set; }}\r\n", typeName, dataArr[i, 0]);
                sb.Append("\r\n");
            }

            //生成克隆代码
            sb.AppendFormat("\tpublic {0}Data Clone()\r\n", dataTableName);
            sb.Append("\t{\r\n");
            sb.AppendFormat("\t\t{0}Data {1}Data = new {2}Data();\r\n", dataTableName, dataTableName.ToLower(), dataTableName);

            for (int i = 1; i < dataArr.GetLength(0); i++)
            {
                sb.AppendFormat("\t\t{0}Data.{1} = this.{2};", dataTableName.ToLower(), dataArr[i, 0], dataArr[i, 0]);
                sb.Append("\r\n");
            }

            sb.AppendFormat("\t\treturn {0}Data;\r\n", dataTableName.ToLower());
            sb.Append("\t}\r\n");
            sb.Append("\r\n");

            //生成解析代码
            sb.AppendFormat("\tpublic override void Read(LocalDataParser parser)\r\n");
            sb.Append("\t{\r\n");

            for (int i = 0; i < dataArr.GetLength(0); i++)
            {
                string fieldName = dataArr[i, 0].Substring(0, 1).ToLower() + dataArr[i, 0].Substring(1);
                string typeName = dataArr[i, 1];

                if (typeName.Equals("json"))
                {
                    typeName = dataArr[i, 0].Substring(0, 1).ToUpper() + dataArr[i, 0].Substring(1);
                    sb.AppendFormat("\t\tthis.{0} = JsonMapper.ToObject<{1}>(parser.GetFieldValue(\"{0}\"));\r\n", fieldName, typeName);
                }
                else
                {
                    sb.AppendFormat("\t\tthis.{0} = parser.GetFieldValue(\"{0}\"){1};\r\n", fieldName, ChangeTypeName(typeName));
                }
            }

            sb.Append("\t}\r\n");
            sb.Append("}\r\n");

            //写入文件
            using (FileStream fs = new FileStream(string.Format("{0}/C#/Script/{1}Data.cs", m_ExportPath, dataTableName), FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.Write(sb.ToString());
                }
            }
        }

        /// <summary>
        /// 创建数据总表
        /// </summary>
        protected override void CreateDataHelperScript()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("\r\n");
            sb.Append("//===================================================\r\n");
            sb.Append("//作者：GQY                                          \r\n");
            sb.AppendFormat("//创建时间：{0}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.Append("//备注：此代码为工具生成 请勿手工修改\r\n");
            sb.Append("//===================================================\r\n");
            sb.Append("using System.Collections;\r\n");
            sb.Append("using GameFrameWork.LocalData;\r\n");
            sb.Append("\r\n");
            sb.Append("/// <summary>\r\n");
            sb.AppendFormat("///数据总表\r\n");
            sb.Append("/// </summary>\r\n");
            sb.AppendFormat("public static class DataHelper\r\n");
            sb.Append("{\r\n");

            for (int i = 0; i < m_DataTableNameList.Count; i++)
            {
                if (!string.IsNullOrEmpty(m_DataTableNameList[i]))
                {
                    sb.AppendFormat("\tpublic static {0}Data[] {1}Data = null;", m_DataTableNameList[i], m_DataTableNameList[i]);
                    sb.Append("\r\n");
                }
            }

            sb.Append("\r\n");
            sb.Append("\tpublic static void Init(string filePath)\r\n");
            sb.Append("\t{\r\n");

            for (int i = 0; i < m_DataTableNameList.Count; i++)
            {
                if (!string.IsNullOrEmpty(m_DataTableNameList[i]))
                {
                    sb.AppendFormat("\t\t{0}Data = LoadData<{1}Data>(filePath, \"{2}Data.bytes\");\r\n", m_DataTableNameList[i], m_DataTableNameList[i], m_DataTableNameList[i]);
                }
            }

            sb.Append("\t}\r\n");
            sb.Append("\r\n");

            //sb.Append("\tpublic static T[] LoadData<T>(string filePath, string fileName) where T : BaseLocalData, new()\r\n");
            //sb.Append("\t{\r\n");
            //sb.Append("\t\tstring path = string.Format(filePath + \"/{0}\", fileName);\r\n");
            //sb.Append("\t\tT[] t = null;\r\n");
            //sb.Append("\t\tusing (LocalDataParser parser = new LocalDataParser(path))\r\n");
            //sb.Append("\t\t{\r\n");
            //sb.Append("\t\t\tt = new T[parser.row - 1];\r\n");
            //sb.Append("\t\t\tint index = 0;\r\n");
            //sb.Append("\t\t\twhile (!parser.eof)\r\n");
            //sb.Append("\t\t\t{\r\n");
            //sb.Append("\t\t\t\tt[index] = new T();\r\n");
            //sb.Append("\t\t\t\tt[index].Read(parser);\r\n");
            //sb.Append("\t\t\t\tparser.Next();\r\n");
            //sb.Append("\t\t\t\tindex++;\r\n");
            //sb.Append("\t\t\t}\r\n");
            //sb.Append("\t\t}\r\n");
            //sb.Append("\t\treturn t;\r\n");
            //sb.Append("\t}\r\n");

            sb.Append("}\r\n");

            using (FileStream fs = new FileStream(string.Format("{0}/C#/Script/DataHelper.cs", m_ExportPath), FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    sw.Write(sb.ToString());
                }
            }
        }

        private void ParseJson(LitJson.JsonData json, JsonStruct jsonStruct)
        {
            if (json.IsObject)
            {
                foreach (KeyValuePair<string, LitJson.JsonData> kvp in json)
                {
                    string key = kvp.Key;
                    LitJson.JsonData val = kvp.Value;
                    string fieldType = JsonFieldType(val);
                    string fieldName = key.Substring(0, 1).ToLower() + key.Substring(1);

                    if (string.IsNullOrEmpty(fieldType))
                    {
                        fieldType = key.Substring(0, 1).ToUpper() + key.Substring(1);

                        if(jsonStruct.jsonStructList == null)
                        {
                            jsonStruct.jsonStructList = new List<JsonStruct>();
                        }

                        JsonStruct childJsonStruct = new JsonStruct();
                        childJsonStruct.className = fieldType;
                        jsonStruct.jsonStructList.Add(childJsonStruct);

                        if (val.IsArray)
                        {
                            if (!jsonStruct.fields.ContainsKey(fieldName))
                            {
                                jsonStruct.fields.Add(fieldName, fieldType + "[]");
                            }
                            ParseJson(val, childJsonStruct);
                        }
                        else
                        {
                            if (!jsonStruct.fields.ContainsKey(fieldName))
                            {
                                jsonStruct.fields.Add(fieldName, fieldType);
                            }
                            ParseJson(val, childJsonStruct);
                        }
                    }
                    else
                    {
                        if (!jsonStruct.fields.ContainsKey(fieldName))
                        {
                            jsonStruct.fields.Add(fieldName, fieldType);
                        }
                    }
                }
            }
            else if(json.IsArray)
            {
                for (int i = 0; i < json.Count; i++)
                {
                    ParseJson(json[i], jsonStruct);
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
            else if(fieldValue.IsArray && fieldValue.Count > 0)
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
                sb.Append("\t");
            }
            sb.AppendFormat("public class {0} \r\n", jsonStruct.className);

            for (int i = 0; i < tCount; i++)
            {
                sb.Append("\t");
            }

            sb.Append("{\r\n");

            if (jsonStruct.jsonStructList != null && jsonStruct.jsonStructList.Count > 0)
            {
                for (int i = 0; i < jsonStruct.jsonStructList.Count; i++)
                {
                    CreateJsonCode(jsonStruct.jsonStructList[i], sb, tCount + 1);
                }

                sb.Append("\r\n");
            }

            foreach (KeyValuePair<string,string> kvp in jsonStruct.fields)
            {
                for (int i = 0; i < tCount + 1; i++)
                {
                    sb.Append("\t");
                }

                sb.AppendFormat("public {0} {1} {{ get; set; }}\r\n", kvp.Value, kvp.Key);
            }

            for (int i = 0; i < tCount; i++)
            {
                sb.Append("\t");
            }

            sb.Append("}\r\n");
        }

        private string ChangeTypeName(string type)
        {
            string str = string.Empty;

            switch (type)
            {
                case "int":
                    str = ".ToInt()";
                    break;
                case "long":
                    str = ".ToLong()";
                    break;
                case "float":
                    str = ".ToFloat()";
                    break;
                case "double":
                    str = ".ToDouble()";
                    break;
                case "bool":
                    str = ".ToBool()";
                    break;
                case "Vector2":
                    str = ".ToVector2()";
                    break;
                case "Vector3":
                    str = ".ToVector3()";
                    break;
                case "int[]":
                    str = ".ToIntArray()";
                    break;
                case "long[]":
                    str = ".ToLongArray()";
                    break;
                case "float[]":
                    str = ".ToFloatArray()";
                    break;
                case "double[]":
                    str = ".ToDoubleArray()";
                    break;
                case "bool[]":
                    str = ".ToBoolArray()";
                    break;
                case "string[]":
                    str = ".ToStringArray()";
                    break;
            }

            return str;
        }


        class JsonStruct
        {
            public string className;
            public Dictionary<string, string> fields = new Dictionary<string, string>();
            public List<JsonStruct> jsonStructList;
        }
    }
}
