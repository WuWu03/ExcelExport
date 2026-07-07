using ExcelExport.Exporter;

namespace ExcelExport.Helper
{
    public static class ExportHelper
    {
        private static int s_CurrExporterIndex = 0;
        private readonly static List<BaseExporter> s_Exporters = [];
        private static List<bool> s_CanExportExcels = null;

        public const string ExportRoot = "DataExport";

        public static BaseExporter currExproter
        {
            get
            {
                return s_Exporters[s_CurrExporterIndex];
            }
        }

        public static List<BaseExporter> exporters
        {
            get
            {
                return s_Exporters;
            }
        }

        public static void AddExporter(BaseExporter exporter)
        {
            s_Exporters.Add(exporter);
        }

        public static void AddExcel(string excelPath)
        {
            s_CanExportExcels ??= [];
            s_CanExportExcels.Add(true);

            for (int i = 0; i < s_Exporters.Count; i++)
            {
                s_Exporters[i].AddExcel(excelPath);
            }
        }

        public static void ResetExcel()
        {
            s_CanExportExcels?.Clear();

            for (int i = 0; i < s_Exporters.Count; i++)
            {
                s_Exporters[i].ResetExcel();
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
            s_CurrExporterIndex = index;
        }

        public static void VerifyPath(string path)
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }

            Directory.CreateDirectory(path);
        }
    }
}
