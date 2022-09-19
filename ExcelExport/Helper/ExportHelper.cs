using ExcelExport.Exporter;
using System.Collections.Generic;

namespace ExcelExport.Helper
{
    public static class ExportHelper
    {
        private static BaseExporter CurrExproter
        {
            get
            {
                return m_Exporters[m_CurrExporterIndex];
            }
        }

        public static void AddExcel(string excelPath)
        {
            if (s_CanExportList == null)
            {
                s_CanExportList = new List<bool>();
            }

            s_CanExportList.Add(true);

            for(int i = 0; i < m_Exporters.Length; i++)
            {
                m_Exporters[i].AddExcel(excelPath);
            }     
        }

        public static void ResetExcel()
        {
            if (s_CanExportList != null)
            {
                s_CanExportList.Clear();
            }

            for (int i = 0; i < m_Exporters.Length; i++)
            {
                m_Exporters[i].ResetExcel();
            }
        }

        public static void Export(string exportPath)
        {
            CurrExproter.SetExportPath(exportPath);
            CurrExproter.Export(s_CanExportList);
        }

        public static void SetExcelCanExport(int index, bool value)
        {
            if (s_CanExportList == null)
            {
                return;
            }

            if (index < 0 || index >= s_CanExportList.Count)
            {
                return;
            }

            s_CanExportList[index] = value;
        }

        public static void SetCurrExporter(int index)
        {
            m_CurrExporterIndex = index;
        }

        private static int m_CurrExporterIndex = 0;
        private static BaseExporter[] m_Exporters = new BaseExporter[] { new CSExporter(), new LuaExporter() };
        private static List<bool> s_CanExportList = null;
    }
}
