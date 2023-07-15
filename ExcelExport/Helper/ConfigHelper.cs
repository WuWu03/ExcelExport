using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace ExcelExport.Helper
{
    public static class ConfigHelper
    {
        public static int CurrSelectIndex
        {
            get
            {
                return s_CurrSelectIndex;
            }
            set
            {
                s_CurrSelectIndex = value;
            }
        }

        public static List<string[]> ConfigData
        {
            get
            {
                return s_ConfigData;
            }
        }

        public static void InitConfig()
        {
            XmlNode xmlNode = GetXmlDocument();

            if (xmlNode.ChildNodes[1].ChildNodes.Count > 0)
            {
                for (int i = 0; i < xmlNode.ChildNodes[1].ChildNodes[0].ChildNodes.Count; i++)
                {
                    string[] config = new string[3];
                    config[0] = xmlNode.ChildNodes[1].ChildNodes[0].ChildNodes[i].ChildNodes[0].InnerText;
                    config[1] = xmlNode.ChildNodes[1].ChildNodes[0].ChildNodes[i].ChildNodes[1].InnerText;
                    config[2] = xmlNode.ChildNodes[1].ChildNodes[0].ChildNodes[i].ChildNodes[2].InnerText;

                    s_ConfigData.Add(config);
                }
            }

            if (xmlNode.ChildNodes[1].ChildNodes.Count >= 2)
            {
                s_CurrSelectIndex = int.Parse(xmlNode.ChildNodes[1].ChildNodes[1].InnerText);
            }
            else
            {
                s_CurrSelectIndex = 0;
            }
        }

        public static string[] GetCurrConfig()
        {
            if(s_ConfigData == null || s_ConfigData.Count < 1)
            {
                return null;
            }

            if(s_CurrSelectIndex < 0 || s_CurrSelectIndex >= s_ConfigData.Count)
            {
                return null;
            }

            return s_ConfigData[s_CurrSelectIndex]; ;
        }

        public static void AddPathConfig(string configName, string excelPath, string exportPath)
        {
            s_ConfigData.Add(new string[3] { configName, excelPath, exportPath });
            s_CurrSelectIndex = s_ConfigData.Count - 1;

            XmlDocument doc = GetXmlDocument();

            XmlNode pathNode = doc.CreateNode(XmlNodeType.Element, "Path", null);
            XmlNode nameNode = doc.CreateNode(XmlNodeType.Element, "Name", null);
            XmlNode excelNode = doc.CreateNode(XmlNodeType.Element, "ExcelPath", null);
            XmlNode exportNode = doc.CreateNode(XmlNodeType.Element, "ExportPath", null);

            nameNode.InnerText = configName;
            excelNode.InnerText = excelPath;
            exportNode.InnerText = exportPath;

            pathNode.AppendChild(nameNode);
            pathNode.AppendChild(excelNode);
            pathNode.AppendChild(exportNode);
            doc.ChildNodes[1].ChildNodes[0].AppendChild(pathNode);

            SetXmlNode(doc);
        }

        public static void DeletePathConfig()
        {
            XmlDocument doc = GetXmlDocument();
            XmlNode currNode = doc.ChildNodes[1].ChildNodes[0].ChildNodes[s_CurrSelectIndex];
            doc.ChildNodes[1].ChildNodes[0].RemoveChild(currNode);

            s_ConfigData.RemoveAt(s_CurrSelectIndex);
            s_CurrSelectIndex--;

            if(s_CurrSelectIndex < 0)
            {
                s_CurrSelectIndex = 0;
            }

            SetXmlNode(doc);
        }

        public static void ModifyPahtConfig(string configName, string excelPath,string exportPath)
        {
            s_ConfigData[s_CurrSelectIndex][0] = configName;
            s_ConfigData[s_CurrSelectIndex][1] = excelPath;
            s_ConfigData[s_CurrSelectIndex][2] = exportPath;

            XmlDocument doc = GetXmlDocument();

            XmlNode currNode = doc.ChildNodes[1].ChildNodes[0].ChildNodes[s_CurrSelectIndex];
            currNode.ChildNodes[0].InnerText = configName;
            currNode.ChildNodes[1].InnerText = excelPath;
            currNode.ChildNodes[2].InnerText = exportPath;

            SetXmlNode(doc);
        }

        private static void SetXmlNode(XmlDocument doc)
        {   
            XmlNode indexNode = null;

            if (doc.ChildNodes[1].ChildNodes.Count < 2)
            {
                indexNode = doc.CreateNode(XmlNodeType.Element, "CurrIndex", null);
                doc.ChildNodes[1].AppendChild(indexNode);
            }
            else
            {
                indexNode = doc.ChildNodes[1].ChildNodes[1];
            }

            indexNode.InnerText = s_CurrSelectIndex.ToString();
            doc.Save("PathConfig.xml");
        }

        private static XmlDocument GetXmlDocument()
        {
            XmlDocument doc = new XmlDocument();

            if (File.Exists("PathConfig.xml"))
            {
                doc.Load("PathConfig.xml");
            }
            else
            {
                XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "utf-8", null);
                doc.AppendChild(dec);

                XmlNode pathConfig = doc.CreateNode(XmlNodeType.Element, "PathConfig", null);
                XmlNode pathList = doc.CreateNode(XmlNodeType.Element, "PathList", null);
                pathConfig.AppendChild(pathList);
                doc.AppendChild(pathConfig);
            }

            return doc;
        }

        private static List<string[]> s_ConfigData = new List<string[]>();
        private static int s_CurrSelectIndex = -1;
    }
}
