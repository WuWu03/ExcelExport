using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace ExcelExport.Helper
{
    public static class ConfigHelper
    {
        public static int currSelectIndex
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

        public static List<string[]> configData
        {
            get
            {
                return s_ConfigDatas;
            }
        }

        public static void InitConfig()
        {
            XmlDocument xmlNode = GetXmlDocument();

            if (xmlNode.ChildNodes[1].ChildNodes.Count > 0)
            {
                for (int i = 0; i < xmlNode.ChildNodes[1].ChildNodes[0].ChildNodes.Count; i++)
                {
                    string[] config = new string[4];
                    config[0] = xmlNode.ChildNodes[1].ChildNodes[0].ChildNodes[i].ChildNodes[0].InnerText;
                    config[1] = xmlNode.ChildNodes[1].ChildNodes[0].ChildNodes[i].ChildNodes[1].InnerText;
                    config[2] = xmlNode.ChildNodes[1].ChildNodes[0].ChildNodes[i].ChildNodes[2].InnerText;
                    config[3] = xmlNode.ChildNodes[1].ChildNodes[0].ChildNodes[i].ChildNodes[3].InnerText;

                    s_ConfigDatas.Add(config);
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
            if (s_ConfigDatas == null || s_ConfigDatas.Count < 1)
            {
                return null;
            }

            if (s_CurrSelectIndex < 0 || s_CurrSelectIndex >= s_ConfigDatas.Count)
            {
                return null;
            }

            return s_ConfigDatas[s_CurrSelectIndex]; ;
        }

        public static void AddPathConfig(string excelPath, string exportPath, string authorName, string configName)
        {
            s_ConfigDatas.Add(new string[4] { excelPath, exportPath, authorName, configName });
            s_CurrSelectIndex = s_ConfigDatas.Count - 1;

            XmlDocument doc = GetXmlDocument();

            XmlNode pathNode = doc.CreateNode(XmlNodeType.Element, "Path", null);
            XmlNode excelPathNode = doc.CreateNode(XmlNodeType.Element, "ExcelPath", null);
            XmlNode exportPathNode = doc.CreateNode(XmlNodeType.Element, "ExportPath", null);
            XmlNode autherNameNode = doc.CreateNode(XmlNodeType.Element, "AuthorName", null);
            XmlNode configNameNode = doc.CreateNode(XmlNodeType.Element, "ConfigName", null);

            excelPathNode.InnerText = excelPath;
            exportPathNode.InnerText = exportPath;
            autherNameNode.InnerText = authorName;
            configNameNode.InnerText = configName;

            pathNode.AppendChild(excelPathNode);
            pathNode.AppendChild(exportPathNode);
            pathNode.AppendChild(autherNameNode);
            pathNode.AppendChild(configNameNode);
            doc.ChildNodes[1].ChildNodes[0].AppendChild(pathNode);

            SetXmlNode(doc);
        }

        public static void DeletePathConfig()
        {
            XmlDocument doc = GetXmlDocument();
            XmlNode currNode = doc.ChildNodes[1].ChildNodes[0].ChildNodes[s_CurrSelectIndex];
            doc.ChildNodes[1].ChildNodes[0].RemoveChild(currNode);

            s_ConfigDatas.RemoveAt(s_CurrSelectIndex);
            s_CurrSelectIndex--;

            if (s_CurrSelectIndex < 0)
            {
                s_CurrSelectIndex = 0;
            }

            SetXmlNode(doc);
        }

        public static void ModifyPahtConfig(string excelPath, string exportPath, string authorName, string configName)
        {
            s_ConfigDatas[s_CurrSelectIndex][0] = excelPath;
            s_ConfigDatas[s_CurrSelectIndex][1] = exportPath;
            s_ConfigDatas[s_CurrSelectIndex][2] = authorName;
            s_ConfigDatas[s_CurrSelectIndex][3] = configName;

            XmlDocument doc = GetXmlDocument();

            XmlNode currNode = doc.ChildNodes[1].ChildNodes[0].ChildNodes[s_CurrSelectIndex];

            currNode.ChildNodes[0].InnerText = excelPath;
            currNode.ChildNodes[1].InnerText = exportPath;
            currNode.ChildNodes[2].InnerText = authorName;
            currNode.ChildNodes[3].InnerText = configName;
            SetXmlNode(doc);
        }

        private static void SetXmlNode(XmlDocument doc)
        {
            XmlNode indexNode;

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
            XmlDocument doc = new();

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

        private static readonly List<string[]> s_ConfigDatas = new();
        private static int s_CurrSelectIndex = -1;
    }
}
