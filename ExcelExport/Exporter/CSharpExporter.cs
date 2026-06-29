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
    public class CSharpExporter : BaseExporter
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
            sb.AppendLine("/*");
            sb.AppendFormat(" * @Desc: {0} 数据表，SheetName: {1}\r\n", excelName, sheetName);
            sb.AppendFormat(" * @Date: {0}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine(" * @Author: " + m_AuthorName);
            sb.AppendLine(" * @Note: 工具生成，请勿修改");
            sb.AppendLine(" */");
            sb.AppendLine();
            sb.Append("using WuWuFramework;\r\n");
            sb.Append("using WuWuFramework.ConfigData;\r\n");
            sb.Append("using LitJson;\r\n");
            sb.Append("using System;\r\n");
            sb.Append("using System.Collections;\r\n");
            sb.Append("using UnityEngine;\r\n");
            sb.Append("\r\n");
            sb.AppendFormat("public class {0}ConfigData : BaseConfigData\r\n", dataTableName);
            sb.Append("{\r\n");

            //生成Json实体类代码
            Dictionary<string, string> jsonDic = new Dictionary<string, string>();

            for (int i = 1; i < dataArr.GetLength(0); i++)
            {
                FieldType fieldType = GetFieldType(dataArr[i, 1]);

                if (fieldType == FieldType.Json || fieldType == FieldType.JsonArray)
                {
                    string typeName = string.Concat(dataArr[i, 0][..1].ToUpper(), dataArr[i, 0].AsSpan(1));

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
                string fieldName = string.Concat(dataArr[i, 0][..1].ToLower(), dataArr[i, 0].AsSpan(1));
                FieldType fieldType = GetFieldType(dataArr[i, 1]);
                string[] typeNames = GetTypeName(dataArr[i, 1]);
                string typeName = typeNames[0];

                if (fieldType == FieldType.Json)
                {
                    typeName = string.Concat(fieldName[..1].ToUpper(), fieldName.AsSpan(1));
                }
                else if (fieldType == FieldType.JsonArray)
                {
                    typeName = string.Concat(fieldName[..1].ToUpper(), fieldName.AsSpan(1), "[]");
                }
                else if (fieldType == FieldType.Dictionary)
                {
                    typeName = string.Concat(typeNames[0], "<", typeNames[1], ",", typeNames[2], ">");
                }

                sb.Append("\t/// <summary>\r\n");
                sb.AppendFormat("\t/// {0}\r\n", dataArr[i, 2]);
                sb.Append("\t/// </summary>\r\n");
                sb.AppendFormat("\tpublic {0} {1} {{ get; private set; }}\r\n", typeName, fieldName);
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
                string[] typeNames = GetTypeName(dataArr[i, 1]);
                FieldType fieldType = GetFieldType(dataArr[i, 1]);

                if (fieldType == FieldType.Json)
                {
                    string typeName = string.Concat(dataArr[i, 0][..1].ToUpper(), dataArr[i, 0].AsSpan(1));
                    sb.AppendFormat("\t\tthis.{0} = JsonMapper.ToObject<{1}>(parser{2});\r\n", fieldName, typeName, GetTypeParseStr(FieldType.UTF8String));
                }
                else if (fieldType == FieldType.JsonArray)
                {
                    string typeName = string.Concat(dataArr[i, 0][..1].ToUpper(), dataArr[i, 0].AsSpan(1));
                    sb.AppendFormat("\t\tthis.{0} = JsonMapper.ToObject<{1}[]>(parser{2});\r\n", fieldName, typeName, GetTypeParseStr(FieldType.UTF8String));
                }
                else
                {
                    sb.AppendFormat("\t\tthis.{0} = parser{1};\r\n", fieldName, GetTypeParseStr(fieldType));
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



        private byte[] GetDataBuffer(DataTable dt)
        {
            byte[] buffer = null;

            using (MemoryStreamEx mse = new MemoryStreamEx())
            {
                mse.WriteInt(dt.Rows.Count - 3);//写入行数
                mse.WriteInt(dt.Columns.Count - 1);//写入列数

                for (int i = 4; i < dt.Rows.Count; i++)
                {
                    for (int j = 1; j < dt.Columns.Count; j++)
                    {
                        string typeName = dt.Rows[1][j].ToString();
                        FieldType fieldType = GetFieldType(typeName);
                        bool fieldNull = dt.Rows[i][j] == null || dt.Rows[i][j] == DBNull.Value || string.IsNullOrEmpty(dt.Rows[i][j].ToString());
                        string fieldStr = dt.Rows[i][j].ToString().Trim();

                        try
                        {
                            if (fieldType == FieldType.Dictionary)
                            {
                                WriteDictionary(typeName, fieldNull, fieldStr, mse, dt.TableName, i, j);
                            }
                            else
                            {
                                WriteButter(fieldType, fieldNull, fieldStr, mse, dt.TableName, i, j);
                            }
                        }
                        catch
                        {

                        }
                    }
                }

                buffer = mse.ToArray();
            }

            buffer = ZlibHelper.CompressBytes(buffer);//压缩
            return buffer;
        }

        /// <summary>
        /// 写入字典数据
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="fieldNull"></param>
        /// <param name="fieldStr"></param>
        /// <param name="mse"></param>
        /// <param name="tableName"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <exception cref="Exception"></exception>
        private void WriteDictionary(string typeName, bool fieldNull, string fieldStr, MemoryStreamEx mse, string tableName, int row, int col)
        {
            if (fieldNull)
            {
                mse.WriteUShort(0);
            }
            else
            {
                string[] dicContents = fieldStr.Split('|');
                string[] dicTypeNames = GetTypeName(typeName);
                mse.WriteUShort((ushort)dicContents.Length);

                for (int dicIndex = 0; dicIndex < dicContents.Length; dicIndex++)
                {
                    string[] dicFields = dicContents[dicIndex].Split(",");

                    if (dicFields.Length != 2)
                    {
                        throw new Exception("[" + tableName + "] [" + row + "," + col + "] Dictionary 数值配置错误");
                    }

                    for (int dicFieldIndex = 0; dicFieldIndex < 2; dicFieldIndex++)
                    {
                        string dicField = dicFields[dicFieldIndex];
                        FieldType dicFieldType = GetFieldType(dicTypeNames[dicFieldIndex + 1]);//字典共有三个类型，0-Dictionary，1-key的类型，2-value的类型，这里要写入key和value所以要fieldIndex+1
                        WriteButter(dicFieldType, string.IsNullOrEmpty(dicField), dicField, mse, typeName, row, col);
                    }
                }
            }
        }

        /// <summary>
        /// 写入一般类型数据
        /// </summary>
        /// <param name="fieldType"></param>
        /// <param name="fieldNull"></param>
        /// <param name="fieldStr"></param>
        /// <param name="mse"></param>
        /// <param name="tableName"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <exception cref="Exception"></exception>
        private void WriteButter(FieldType fieldType, bool fieldNull, string fieldStr, MemoryStreamEx mse, string tableName, int row, int col)
        {
            switch (fieldType)
            {
                case FieldType.Byte:
                    mse.WriteByte(fieldNull ? (byte)0 : byte.Parse(fieldStr));
                    break;
                case FieldType.Short:
                    mse.WriteInt(fieldNull ? (short)0 : short.Parse(fieldStr));
                    break;
                case FieldType.Int:
                    mse.WriteInt(fieldNull ? 0 : int.Parse(fieldStr));
                    break;
                case FieldType.Long:
                    mse.WriteLong(fieldNull ? 0 : long.Parse(fieldStr));
                    break;
                case FieldType.Float:
                    mse.WriteFloat(fieldNull ? 0 : float.Parse(fieldStr));
                    break;
                case FieldType.Double:
                    mse.WriteDouble(fieldNull ? 0 : double.Parse(fieldStr));
                    break;
                case FieldType.Bool:
                    if (fieldNull)
                    {
                        mse.WriteBool(false);
                    }
                    else
                    {
                        string boolValue = fieldStr.ToLower();

                        if (boolValue != "false" && boolValue != "true" && boolValue != "0" && boolValue != "1")
                        {
                            throw new Exception("[" + tableName + "] [" + row + "," + col + "] bool 数值配置错误");
                        }

                        mse.WriteBool(boolValue == "true" || boolValue == "1");
                    }

                    break;
                case FieldType.UTF8String:
                    mse.WriteUTF8String(fieldStr);
                    break;
                case FieldType.Vector2:
                    if (fieldNull)
                    {
                        mse.WriteFloat(0);
                        mse.WriteFloat(0);
                    }
                    else
                    {
                        string[] vector2Value = fieldStr.Split(',');

                        if (vector2Value.Length != 2)
                        {
                            throw new Exception("[" + tableName + "] [" + row + "," + col + "] Vector2 数值配置错误");
                        }

                        mse.WriteFloat(float.Parse(vector2Value[0]));
                        mse.WriteFloat(float.Parse(vector2Value[1]));
                    }
                    break;
                case FieldType.Vector3:
                    if (fieldNull)
                    {
                        mse.WriteFloat(0);
                        mse.WriteFloat(0);
                        mse.WriteFloat(0);
                    }
                    else
                    {
                        string[] vector3Value = fieldStr.Split(',');

                        if (vector3Value.Length != 3)
                        {
                            throw new Exception("[" + tableName + "] [" + row + "," + col + "] Vector3数值配置错误");
                        }

                        mse.WriteFloat(float.Parse(vector3Value[0]));
                        mse.WriteFloat(float.Parse(vector3Value[1]));
                        mse.WriteFloat(float.Parse(vector3Value[2]));
                    }

                    break;
                case FieldType.Json:
                    mse.WriteUTF8String(fieldStr);
                    break;
                case FieldType.ByteArray:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] byteArray = fieldStr.Split(',');
                        mse.WriteUShort((ushort)byteArray.Length);

                        for (int index = 0; index < byteArray.Length; index++)
                        {
                            mse.WriteByte(byte.Parse(byteArray[index]));
                        }
                    }

                    break;
                case FieldType.ShortArray:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] shortArray = fieldStr.Split(',');
                        mse.WriteUShort((ushort)shortArray.Length);

                        for (int index = 0; index < shortArray.Length; index++)
                        {
                            mse.WriteShort(short.Parse(shortArray[index]));
                        }
                    }

                    break;
                case FieldType.IntArray:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] intArray = fieldStr.Split(',');
                        mse.WriteUShort((ushort)intArray.Length);

                        for (int index = 0; index < intArray.Length; index++)
                        {
                            mse.WriteInt(string.IsNullOrEmpty(intArray[index]) ? 0 : int.Parse(intArray[index]));
                        }
                    }

                    break;
                case FieldType.LongArray:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] longArray = fieldStr.Split(',');
                        mse.WriteUShort((ushort)longArray.Length);

                        for (int index = 0; index < longArray.Length; index++)
                        {
                            mse.WriteLong(long.Parse(longArray[index]));
                        }
                    }

                    break;
                case FieldType.FloatArray:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] floatArray = fieldStr.Split(',');
                        mse.WriteUShort((ushort)floatArray.Length);

                        for (int index = 0; index < floatArray.Length; index++)
                        {
                            mse.WriteFloat(float.Parse(floatArray[index]));
                        }
                    }

                    break;
                case FieldType.DoubleArray:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] doubleArray = fieldStr.Split(',');
                        mse.WriteUShort((ushort)doubleArray.Length);

                        for (int index = 0; index < doubleArray.Length; index++)
                        {
                            mse.WriteDouble(double.Parse(doubleArray[index]));
                        }
                    }

                    break;
                case FieldType.BoolArray:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] boolArray = fieldStr.Split(',');
                        mse.WriteUShort((ushort)boolArray.Length);

                        for (int index = 0; index < boolArray.Length; index++)
                        {
                            mse.WriteBool(bool.Parse(string.Concat(boolArray[index][..1].ToUpper(), boolArray[index].AsSpan(1))));
                        }
                    }

                    break;
                case FieldType.UTF8StringArray:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] stringArray = fieldStr.Split(',');
                        mse.WriteUShort((ushort)stringArray.Length);

                        for (int index = 0; index < stringArray.Length; index++)
                        {
                            mse.WriteUTF8String(stringArray[index]);
                        }
                    }

                    break;
                case FieldType.Vector2Array:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] vector2Array = fieldStr.Split('|');
                        mse.WriteUShort((ushort)vector2Array.Length);

                        for (int index = 0; index < vector2Array.Length; index++)
                        {
                            string[] vector2Value = vector2Array[index].Split(',');

                            if (vector2Value.Length != 2)
                            {
                                throw new Exception("[" + tableName + "] [" + row + "," + col + "] Vector2[] 数值配置错误");
                            }

                            mse.WriteFloat(float.Parse(vector2Value[0]));
                            mse.WriteFloat(float.Parse(vector2Value[1]));
                        }
                    }
                    break;
                case FieldType.Vector3Array:
                    if (fieldNull)
                    {
                        mse.WriteUShort(0);
                    }
                    else
                    {
                        string[] vector3Array = fieldStr.Split('|');
                        mse.WriteUShort((ushort)vector3Array.Length);

                        for (int index = 0; index < vector3Array.Length; index++)
                        {
                            string[] vector3Value = vector3Array[index].Split(',');

                            if (vector3Value.Length != 3)
                            {
                                throw new Exception("[" + tableName + "] [" + row + "," + col + "] Vector3[] 数值配置错误");
                            }

                            mse.WriteFloat(float.Parse(vector3Value[0]));
                            mse.WriteFloat(float.Parse(vector3Value[1]));
                            mse.WriteFloat(float.Parse(vector3Value[2]));
                        }
                    }

                    break;
                case FieldType.JsonArray:
                    mse.WriteUTF8String(fieldStr);
                    break;
            }
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

        enum FieldType : byte
        {
            Byte,
            Short,
            Int,
            Long,
            Float,
            Double,
            Bool,
            UTF8String,
            Vector2,
            Vector3,
            Json,
            ByteArray,
            ShortArray,
            IntArray,
            LongArray,
            FloatArray,
            DoubleArray,
            BoolArray,
            UTF8StringArray,
            Vector2Array,
            Vector3Array,
            JsonArray,
            Dictionary,
        }

        private string[] GetTypeName(string typeName)
        {
            typeName = typeName.ToLower();

            if (typeName.Contains("dictionary"))
            {
                int dicNameLength = "dictionary".Length;
                string[] dicTypes = typeName[(dicNameLength + 1)..(typeName.Length - 1)].Trim().Split(',');

                if (dicTypes.Length != 2)
                {
                    throw new Exception("字典类型配置错误！");
                }

                return new string[] { "Dictionary", dicTypes[0], dicTypes[1] };
            }

            return typeName switch
            {
                "byte" => new string[] { "byte" },
                "short" => new string[] { "short" },
                "int" => new string[] { "int" },
                "long" => new string[] { "long" },
                "float" => new string[] { "float" },
                "double" => new string[] { "double" },
                "bool" => new string[] { "bool" },
                "string" => new string[] { "string" },
                "vector2" => new string[] { "Vector2" },
                "vector3" => new string[] { "Vector3" },
                "byte[]" => new string[] { "byte[]" },
                "short[]" => new string[] { "short[]" },
                "int[]" => new string[] { "int[]" },
                "long[]" => new string[] { "long[]" },
                "float[]" => new string[] { "float[]" },
                "double[]" => new string[] { "double[]" },
                "bool[]" => new string[] { "bool[]" },
                "string[]" => new string[] { "string[]" },
                "vector2[]" => new string[] { "Vector2[]" },
                "vector3[]" => new string[] { "Vector3[]" },
                "json" => new string[] { "json" },
                "json[]" => new string[] { "json[]" },
                _ => throw new NotImplementedException(),
            };
        }

        private FieldType GetFieldType(string fieldType)
        {
            string[] typeNames = GetTypeName(fieldType);

            return typeNames[0] switch
            {
                "byte" => FieldType.Byte,
                "short" => FieldType.Short,
                "int" => FieldType.Int,
                "long" => FieldType.Long,
                "float" => FieldType.Float,
                "double" => FieldType.Double,
                "bool" => FieldType.Bool,
                "string" => FieldType.UTF8String,
                "Vector2" => FieldType.Vector2,
                "Vector3" => FieldType.Vector3,
                "byte[]" => FieldType.ByteArray,
                "short[]" => FieldType.ShortArray,
                "int[]" => FieldType.IntArray,
                "long[]" => FieldType.LongArray,
                "float[]" => FieldType.FloatArray,
                "double[]" => FieldType.DoubleArray,
                "bool[]" => FieldType.BoolArray,
                "string[]" => FieldType.UTF8StringArray,
                "Vector2[]" => FieldType.Vector2Array,
                "Vector3[]" => FieldType.Vector3Array,
                "json" => FieldType.Json,
                "json[]" => FieldType.JsonArray,
                "Dictionary" => FieldType.Dictionary,
                _ => throw new NotImplementedException(),
            };
        }

        private string GetTypeParseStr(FieldType fieldType)
        {
            return fieldType switch
            {
                FieldType.Byte => ".ReadByte()",
                FieldType.Short => ".ReadShort()",
                FieldType.Int => ".ReadInt()",
                FieldType.Long => ".ReadLong()",
                FieldType.Float => ".ReadFloat()",
                FieldType.Double => ".ReadDouble()",
                FieldType.Bool => ".ReadBool()",
                FieldType.UTF8String => ".ReadUTF8String()",
                FieldType.Vector2 => ".ReadVector2()",
                FieldType.Vector3 => ".ReadVector3()",
                FieldType.ByteArray => ".ReadByteArray()",
                FieldType.ShortArray => ".ReadShortArray()",
                FieldType.IntArray => ".ReadIntArray()",
                FieldType.LongArray => ".ReadLongArray()",
                FieldType.FloatArray => ".ReadFloatArray()",
                FieldType.DoubleArray => ".ReadDoubleArray()",
                FieldType.BoolArray => ".ReadBoolArray()",
                FieldType.UTF8StringArray => ".ReadUTF8StringArray()",
                FieldType.Vector2Array => ".ReadVector2Array()",
                FieldType.Vector3Array => ".ReadVector3Array()",
                FieldType.Dictionary => ".ReadDictionary()",
                _ => throw new NotImplementedException(),
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

        #region 废弃数据总表
        /// <summary>
        /// 创建数据总表
        /// </summary>
        protected override void CreateConfigDataSheetScript()
        {
            //StringBuilder sb = new StringBuilder();
            //sb.AppendLine("/*");
            //sb.AppendLine(" * @Desc: 数据实体类定义");
            //sb.AppendFormat(" * @Date: {0}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            //sb.AppendLine(" * @Author: " + m_AuthorName);
            //sb.AppendLine(" * @Note: 工具生成，请勿修改");
            //sb.AppendLine(" */");
            //sb.AppendLine();
            //sb.Append("using System.Collections;\r\n");
            //sb.Append("using WuWuFramework.ConfigData;\r\n");
            //sb.Append("\r\n");
            //sb.AppendFormat("public static class ConfigDataSheet\r\n");
            //sb.Append("{\r\n");

            //for (int i = 0; i < m_DataTableNames.Count; i++)
            //{
            //    if (!string.IsNullOrEmpty(m_DataTableNames[i]))
            //    {
            //        string fieldName = string.Concat(m_DataTableNames[i][..1].ToLower(), m_DataTableNames[i].AsSpan(1));
            //        sb.AppendFormat("\tpublic static {0}ConfigData[] {1}ConfigDatas = null;", m_DataTableNames[i], fieldName);
            //        sb.Append("\r\n");
            //    }
            //}

            //sb.Append("\r\n");
            //sb.Append("\tpublic static void Init(string filePath)\r\n");
            //sb.Append("\t{\r\n");

            //for (int i = 0; i < m_DataTableNames.Count; i++)
            //{
            //    if (!string.IsNullOrEmpty(m_DataTableNames[i]))
            //    {
            //        string fieldName = string.Concat(m_DataTableNames[i][..1].ToLower(), m_DataTableNames[i].AsSpan(1));
            //        sb.AppendFormat("\t\t{0}ConfigDatas = LoadConfigData<{1}ConfigData>(filePath, \"{2}ConfigData\");\r\n", fieldName, m_DataTableNames[i], m_DataTableNames[i]);
            //    }
            //}

            //sb.Append("\t}\r\n");
            //sb.Append('}');

            //try
            //{
            //    File.WriteAllText(GetScriptsExportPath("ConfigDataSheet.cs"), sb.ToString());
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

            //sb.Clear();
        }
        #endregion
    }
}
