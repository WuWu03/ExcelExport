using ExcelExport.Exporter;
using System.Collections.Generic;
using System.IO;

namespace ExcelExport.Helper
{
    public static class ExportHelper
    {
        public const string ExportRoot = "DataExport";

        public static BaseExporter currExproter
        {
            get
            {
                return m_Exporters[m_CurrExporterIndex];
            }
        }

        public static void AddExcel(string excelPath)
        {
            s_CanExportExcels ??= new List<bool>();
            s_CanExportExcels.Add(true);

            for (int i = 0; i < m_Exporters.Length; i++)
            {
                m_Exporters[i].AddExcel(excelPath);
            }
        }

        public static void ResetExcel()
        {
            s_CanExportExcels?.Clear();

            for (int i = 0; i < m_Exporters.Length; i++)
            {
                m_Exporters[i].ResetExcel();
            }
        }

        public static void Export(string exportPath, string authorName)
        {
            currExproter.SetExportPath(exportPath);
            currExproter.SetAuthorName(authorName);
            currExproter.Export(s_CanExportExcels);
        }

        public static void SetExcelCanExport(int index, bool value)
        {
            if (s_CanExportExcels == null)
            {
                return;
            }

            if (index < 0 || index >= s_CanExportExcels.Count)
            {
                return;
            }

            s_CanExportExcels[index] = value;
        }

        public static void SetCurrExporter(int index)
        {
            m_CurrExporterIndex = index;
        }

        public static void VerifyPath(string path)
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }

            Directory.CreateDirectory(path);
        }

        private static int m_CurrExporterIndex = 0;
        private static BaseExporter[] m_Exporters = new BaseExporter[] { new CSharpExporter(), new LuaExporter()};
        private static List<bool> s_CanExportExcels = null;
    }
}
