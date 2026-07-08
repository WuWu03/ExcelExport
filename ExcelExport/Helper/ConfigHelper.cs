using System.Xml;

namespace ExcelExport.Helper
{
    public static class ConfigHelper
    {
        private static readonly List<string[]> s_ConfigDatas = [];
        private static int s_CurrSelectIndex = -1;
        private const string PATH_CONFIG_NAME = "PathConfig.xml";

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
            s_CurrSelectIndex = 0;
            XmlDocument doc = LoadConfig();

            if (doc != null)
            {
                XmlNode configRoot = doc.ChildNodes[1];
                XmlNode pathConfigs = configRoot.ChildNodes[0];
                XmlNode indexConfig = configRoot.ChildNodes[1];

                for (int i = 0; i < pathConfigs.ChildNodes.Count; i++)
                {
                    XmlNode pathConfig = pathConfigs.ChildNodes[i];
                    string excelPath = string.IsNullOrEmpty(pathConfig.ChildNodes[0].InnerText) ? string.Empty : pathConfig.ChildNodes[0].InnerText;
                    string exportPath = string.IsNullOrEmpty(pathConfig.ChildNodes[1].InnerText) ? string.Empty : pathConfig.ChildNodes[1].InnerText;
                    string authorName = string.IsNullOrEmpty(pathConfig.ChildNodes[2].InnerText) ? string.Empty : pathConfig.ChildNodes[2].InnerText;
                    string configName = string.IsNullOrEmpty(pathConfig.ChildNodes[3].InnerText) ? string.Empty : pathConfig.ChildNodes[3].InnerText;
                    s_ConfigDatas.Add([excelPath, exportPath, authorName, configName]);
                }

                s_CurrSelectIndex = string.IsNullOrEmpty(indexConfig.InnerText) ? 0 : int.Parse(indexConfig.InnerText);
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
            s_ConfigDatas.Add([excelPath, exportPath, authorName, configName]);
            s_CurrSelectIndex = s_ConfigDatas.Count - 1;
        }

        public static void DeletePathConfig()
        {
            s_ConfigDatas.RemoveAt(s_CurrSelectIndex);
            s_CurrSelectIndex--;

            if (s_CurrSelectIndex < 0)
            {
                s_CurrSelectIndex = 0;
            }
        }

        public static void ModifyPathConfig(string excelPath, string exportPath, string authorName, string configName)
        {
            s_ConfigDatas[s_CurrSelectIndex][0] = excelPath;
            s_ConfigDatas[s_CurrSelectIndex][1] = exportPath;
            s_ConfigDatas[s_CurrSelectIndex][2] = authorName;
            s_ConfigDatas[s_CurrSelectIndex][3] = configName;
        }

        public static void SaveConfig()
        {
            XmlDocument doc = CreateConfig();
            XmlNode configRoot = doc.ChildNodes[1];
            XmlNode pathConfigs = configRoot.ChildNodes[0];
            XmlNode indexConfig = configRoot.ChildNodes[1];

            foreach (string[] configData in s_ConfigDatas)
            {
                XmlNode config = doc.CreateNode(XmlNodeType.Element, "Config", null);
                XmlNode excelPathNode = doc.CreateNode(XmlNodeType.Element, "ExcelPath", null);
                XmlNode exportPathNode = doc.CreateNode(XmlNodeType.Element, "ExportPath", null);
                XmlNode autherNameNode = doc.CreateNode(XmlNodeType.Element, "AuthorName", null);
                XmlNode configNameNode = doc.CreateNode(XmlNodeType.Element, "ConfigName", null);
                excelPathNode.InnerText = string.IsNullOrEmpty(configData[0]) ? string.Empty : configData[0];
                exportPathNode.InnerText = string.IsNullOrEmpty(configData[1]) ? string.Empty : configData[1];
                autherNameNode.InnerText = string.IsNullOrEmpty(configData[2]) ? string.Empty : configData[2];
                configNameNode.InnerText = string.IsNullOrEmpty(configData[3]) ? string.Empty : configData[3];
                config.AppendChild(excelPathNode);
                config.AppendChild(exportPathNode);
                config.AppendChild(autherNameNode);
                config.AppendChild(configNameNode);
                pathConfigs.AppendChild(config);
            }

            indexConfig.InnerText = s_CurrSelectIndex.ToString();
            doc.Save(PATH_CONFIG_NAME);
        }

        private static XmlDocument LoadConfig()
        {
            if (File.Exists(PATH_CONFIG_NAME))
            {
                XmlDocument doc = new();
                doc.Load(PATH_CONFIG_NAME);
                return doc;
            }

            return null;
        }

        private static XmlDocument CreateConfig()
        {
            XmlDocument doc = new();
            XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "utf-8", null);
            doc.AppendChild(dec);
            XmlNode configRoot = doc.CreateNode(XmlNodeType.Element, "ConfigRoot", null);
            XmlNode pathConfig = doc.CreateNode(XmlNodeType.Element, "PathConfigs", null);
            XmlNode indexConfig = doc.CreateNode(XmlNodeType.Element, "IndexConfig", null);
            configRoot.AppendChild(pathConfig);
            configRoot.AppendChild(indexConfig);
            doc.AppendChild(configRoot);
            return doc;
        }
    }
}
