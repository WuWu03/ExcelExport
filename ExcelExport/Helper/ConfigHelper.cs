using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelExport.Helper
{
    public static class ConfigHelper
    {
        public static void SaveExcelPath(string path)
        {
            SetXmlNode("ExcelPath", path);
        }

        public static void SaveExportPath(string path)
        {
            SetXmlNode("ExportPath", path);
        }

        public static string GetExcelPath()
        {
            XmlNode path = GetXmlNode("ExcelPath");

            if (path != null)
            {
                return path.InnerText;
            }

            return string.Empty;
        }

        public static string GetExportPath()
        {
            XmlNode path = GetXmlNode("ExportPath");

            if(path != null)
            {
                return path.InnerText;
            }

            return string.Empty;
        }

        private static XmlNode GetXmlNode(string nodeName)
        {
            XmlDocument doc = GetXmlDocument();
            XmlNode path = doc.ChildNodes[1].ChildNodes[0];

            foreach (XmlNode node in path.ChildNodes)
            {
                if (node.Name == nodeName)
                {
                    return node;
                }
            }

            return null;
        }
        private static void SetXmlNode(string nodeName, string content)
        {
            XmlDocument doc = GetXmlDocument();
            XmlNode path = doc.ChildNodes[1].ChildNodes[0];
            bool hasNode = false;

            foreach(XmlNode node in path.ChildNodes)
            {
                if(node.Name == nodeName)
                {
                    hasNode = true;
                    node.InnerText = content;
                    break;
                }
            }

            if (!hasNode)
            {
                XmlNode newNode = doc.CreateNode(XmlNodeType.Element, nodeName, null);
                newNode.InnerText = content;
                path.AppendChild(newNode);
            }

            doc.Save("PathConfig.xml");
        }


        private static XmlDocument GetXmlDocument()
        {
            XmlDocument doc = new XmlDocument();
            XmlNode pathList = null;

            if (File.Exists("PathConfig.xml"))
            {
                doc.Load("PathConfig.xml");
                pathList = doc.ChildNodes[1];
            }
            else
            {
                XmlDeclaration dec = doc.CreateXmlDeclaration("1.0", "utf-8", null);
                doc.AppendChild(dec);

                pathList = doc.CreateNode(XmlNodeType.Element, "PathList", null);
                doc.AppendChild(pathList);
            }

            if (pathList.ChildNodes.Count < 1)
            {
                XmlNode path = doc.CreateNode(XmlNodeType.Element, "Path", null);
                pathList.AppendChild(path);
            }

            return doc;
        }
    }
}
