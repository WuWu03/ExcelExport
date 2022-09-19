using ExcelExport.Helper;
using System;
using System.IO;
using System.Windows.Forms;

namespace ExcelExport
{
    public partial class ExcelExport : Form
    {
        public ExcelExport()
        {
            InitializeComponent();
            this.codeTypeComboBox.SelectedIndex = 0;

            textBoxExcel.Text = ConfigHelper.GetExcelPath();
            textBoxExport.Text = ConfigHelper.GetExportPath();
        }


        /// <summary>
        /// 选择表格按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnSelectClick(object sender, EventArgs e)
        {
            string excelPath = textBoxExcel.Text;

            if (!string.IsNullOrEmpty(excelPath) && Directory.Exists(excelPath))
            {
                LoadExcelFiles(excelPath);
                return;
            }

            using(FolderBrowserDialog fbDlg = new FolderBrowserDialog())
            {
                if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBoxExcel.Text = fbDlg.SelectedPath;
                    ConfigHelper.SaveExcelPath(fbDlg.SelectedPath);
                    LoadExcelFiles(fbDlg.SelectedPath);
                }
            }
        }

        private void LoadExcelFiles(string path)
        {
            ExportHelper.ResetExcel();
            excelList.Items.Clear();

            string[] files = Directory.GetFiles(path + "\\", "*", SearchOption.AllDirectories);

            foreach (string strName in files)
            {
                if (!Path.GetExtension(strName).Contains("xls") || strName.Contains("$"))//非excel文件或excel的缓存文件不进行读取
                {
                    continue;
                }

                ExportHelper.AddExcel(strName);
                excelList.Items.Add(strName);
                excelList.SetItemChecked(excelList.Items.Count - 1, true);
            }

            MessageBox.Show(this, "读取成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 创建按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnCreateClick(object sender, EventArgs e)
        {
            string exportPath = textBoxExport.Text;

            if (!string.IsNullOrEmpty(exportPath) && Directory.Exists(exportPath))
            {
                ExportExcel(exportPath);
                return;
            }

            using (FolderBrowserDialog fbDlg = new FolderBrowserDialog())
            {
                if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBoxExport.Text = fbDlg.SelectedPath;
                    ConfigHelper.SaveExportPath(fbDlg.SelectedPath);
                    ExportExcel(fbDlg.SelectedPath);
                }
            }
        }

        private void ExportExcel(string path)
        {
            ExportHelper.Export(path);
            MessageBox.Show(this, "创建成功");
        }

        private void OnCodeTypeComboBoxSelectedIndexChanged(object sender, EventArgs e)
        {
            ExportHelper.SetCurrExporter(codeTypeComboBox.SelectedIndex);
        }

        private void OnExcelListItemCheck(object sender, ItemCheckEventArgs e)
        {
            ExportHelper.SetExcelCanExport(e.Index, !excelList.GetItemChecked(e.Index));
        }

        private void OnTextBoxExcelDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void OnTextBoxExcelDragDrop(object sender, DragEventArgs e)
        {
            string excelPath = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

            textBoxExcel.Text = excelPath;
            ConfigHelper.SaveExcelPath(excelPath);
        }

        private void OnTextBoxExportDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void OnTextBoxExportDragDrop(object sender, DragEventArgs e)
        {
            string exportPath = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString(); 

            textBoxExport.Text = exportPath;
            ConfigHelper.SaveExportPath(exportPath);
        }

        //异或因子
        //private byte[] xorScale = new byte[] { 45, 66, 38, 55, 23, 254, 9, 165, 90, 19, 41, 45, 201, 58, 55, 37, 254, 185, 165, 169, 19, 171 };//.data文件的xor加解密因子
        //private List<string> _allTalbeName = new List<string>();
    }
}