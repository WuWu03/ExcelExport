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
            ConfigHelper.InitConfig();

            this.codeTypeComboBox.SelectedIndex = 0;
 
            for (int i = 0;i < ConfigHelper.ConfigData.Count;i++)
            {
                this.configListComboBox.Items.Add(ConfigHelper.ConfigData[i][0]);
            }

            this.configListComboBox.Items.Add("添加配置");
            this.configListComboBox.SelectedIndex = ConfigHelper.CurrSelectIndex;

            bool showAddPathBtn = ConfigHelper.ConfigData.Count < 1 || configListComboBox.SelectedIndex == ConfigHelper.ConfigData.Count;
            this.btnModifyPathConfig.Visible = !showAddPathBtn;
            this.btnAddPathConfig.Visible = showAddPathBtn;
        }


        /// <summary>
        /// 选择表格按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnSelectClick(object sender, EventArgs e)
        {
            string excelPath = string.Empty;
            string[] configData = ConfigHelper.GetCurrConfig();

            if(configData != null)
            {
                excelPath = configData[1];
            }

            if (!string.IsNullOrEmpty(excelPath) && Directory.Exists(excelPath))
            {
                LoadExcelFiles(excelPath);
                return;
            }

            using(FolderBrowserDialog fbDlg = new FolderBrowserDialog())
            {
                if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    LoadExcelFiles(fbDlg.SelectedPath);
                }
            }
        }

        /// <summary>
        /// 创建按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnCreateClick(object sender, EventArgs e)
        {
            string exportPath = string.Empty;
            string[] configData = ConfigHelper.GetCurrConfig();

            if (configData != null)
            {
                exportPath = configData[2];
            }

            if (!string.IsNullOrEmpty(exportPath) && Directory.Exists(exportPath))
            {
                ExportExcel(exportPath);
                return;
            }

            using (FolderBrowserDialog fbDlg = new FolderBrowserDialog())
            {
                if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    ExportExcel(fbDlg.SelectedPath);
                }
            }
        }

        private void OnBtnSelectExcelClick(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbDlg = new FolderBrowserDialog())
            {
                if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBoxExcel.Text = fbDlg.SelectedPath;
                }
            }
        }

        private void OnBtnSelectExportClick(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbDlg = new FolderBrowserDialog())
            {
                if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBoxExport.Text = fbDlg.SelectedPath;
                }
            }
        }

        private void OnBtnModifyPathConfigClick(object sender, EventArgs e)
        {
            ConfigHelper.ModifyPahtConfig(textBoxPathName.Text, textBoxExcel.Text, textBoxExport.Text);
            configListComboBox.Items[configListComboBox.SelectedIndex] = ConfigHelper.GetCurrConfig()[0];
            MessageBox.Show(this, "修改成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OnBtnAddPathConfigClick(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBoxPathName.Text))
            {
                MessageBox.Show(this, "名称不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            configListComboBox.Items[configListComboBox.Items.Count - 1] = textBoxPathName.Text;
            configListComboBox.Items.Add("添加配置");

            ConfigHelper.CurrSelectIndex = configListComboBox.Items.Count - 2;
            ConfigHelper.AddPathConfig(textBoxPathName.Text, textBoxExcel.Text, textBoxExport.Text);

            configListComboBox.SelectedIndex = ConfigHelper.CurrSelectIndex;
            OnConfigListComboBoxChanged(configListComboBox, null);

            MessageBox.Show(this, "添加成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OnBtnDelectPathConfigClick(object sender, EventArgs e)
        {
            if(MessageBox.Show("确认删除本条配置？","警告",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int selectedIndex = configListComboBox.SelectedIndex;
                configListComboBox.Items.RemoveAt(configListComboBox.SelectedIndex);
                selectedIndex--;

                if (selectedIndex < 0)
                {
                    selectedIndex = 0;
                }

                textBoxPathName.Text = string.Empty;
                ConfigHelper.DeletePathConfig();
                configListComboBox.SelectedIndex = selectedIndex;

                MessageBox.Show(this, "删除成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void ExportExcel(string path)
        {
            ExportHelper.Export(path);
            MessageBox.Show(this, "创建成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        }

        private void OnConfigListComboBoxChanged(object sender, EventArgs e)
        {
            bool showAddPathBtn = ConfigHelper.ConfigData.Count < 1 || this.configListComboBox.SelectedIndex == ConfigHelper.ConfigData.Count;
            btnModifyPathConfig.Visible = !showAddPathBtn;
            btnDelectPathConfig.Visible = !showAddPathBtn;
            btnAddPathConfig.Visible = showAddPathBtn;
       
            if (showAddPathBtn)
            {
                textBoxExcel.Text = string.Empty;
                textBoxExport.Text = string.Empty;
                textBoxPathName.Text = string.Empty;
            }
            else
            {
                ConfigHelper.CurrSelectIndex = configListComboBox.SelectedIndex;
                string[] config = ConfigHelper.GetCurrConfig();
                textBoxPathName.Text = config[0];
                textBoxExcel.Text = config[1];
                textBoxExport.Text = config[2];
            }
        }

    

        //异或因子
        //private byte[] xorScale = new byte[] { 45, 66, 38, 55, 23, 254, 9, 165, 90, 19, 41, 45, 201, 58, 55, 37, 254, 185, 165, 169, 19, 171 };//.data文件的xor加解密因子
        //private List<string> _allTalbeName = new List<string>();
    }
}