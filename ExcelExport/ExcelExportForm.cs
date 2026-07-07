using ExcelExport.Exporter;
using ExcelExport.Helper;

namespace ExcelExport
{
    public partial class ExcelExportForm : Form
    {
        /// <summary>
        /// 初始化
        /// </summary>
        public ExcelExportForm()
        {
            InitializeComponent();
            ConfigHelper.InitConfig();
            ExportHelper.AddExporter(new CSharpExporter());
            // ExportHelper.AddExporter(new LuaExporter()));

            for(int i = 0; i < ExportHelper.exporters.Count; i++)
            {
                codeTypeComboBox.Items.Add(ExportHelper.exporters[i].exporterName);
            }

            codeTypeComboBox.SelectedIndex = 0;

            for (int i = 0; i < ConfigHelper.configData.Count; i++)
            {
                configListComboBox.Items.Add(ConfigHelper.configData[i][3]);
            }

            configListComboBox.Items.Add("添加配置");
            configListComboBox.SelectedIndex = ConfigHelper.currSelectIndex;

            bool showAddPathBtn = ConfigHelper.configData.Count < 1 || configListComboBox.SelectedIndex == ConfigHelper.configData.Count;
            btnModifyPathConfig.Visible = !showAddPathBtn;
            btnAddPathConfig.Visible = showAddPathBtn;
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

            if (configData != null)
            {
                excelPath = configData[0];
            }

            if (!string.IsNullOrEmpty(excelPath) && Directory.Exists(excelPath))
            {
                LoadExcelFiles(excelPath);
                return;
            }

            using FolderBrowserDialog fbDlg = new();

            if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                LoadExcelFiles(fbDlg.SelectedPath);
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
            string authorName = string.Empty;
            string[] configData = ConfigHelper.GetCurrConfig();

            if (configData != null)
            {
                exportPath = configData[1];
                authorName = configData[2];
            }

            if (!string.IsNullOrEmpty(exportPath) && Directory.Exists(exportPath))
            {
                ExportExcel(exportPath, authorName);
                return;
            }

            using FolderBrowserDialog fbDlg = new();

            if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ExportExcel(fbDlg.SelectedPath, authorName);
            }
        }

        /// <summary>
        /// excel文件路径选择按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnSelectExcelClick(object sender, EventArgs e)
        {
            using FolderBrowserDialog fbDlg = new();

            if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBoxExcel.Text = fbDlg.SelectedPath;
            }
        }

        /// <summary>
        /// 导出路径选择按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnSelectExportClick(object sender, EventArgs e)
        {
            using FolderBrowserDialog fbDlg = new();

            if (fbDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBoxExport.Text = fbDlg.SelectedPath;
            }
        }

        /// <summary>
        /// 修改配置按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnModifyPathConfigClick(object sender, EventArgs e)
        {
            ConfigHelper.ModifyPahtConfig(textBoxExcel.Text, textBoxExport.Text, textBoxAuthorName.Text, textBoxConfigName.Text);
            configListComboBox.Items[configListComboBox.SelectedIndex] = ConfigHelper.GetCurrConfig()[3];
            MessageBox.Show(this, "修改成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 添加配置按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnAddPathConfigClick(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBoxConfigName.Text))
            {
                MessageBox.Show(this, "名称不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string excelPath = textBoxExcel.Text;
            string exportPath = textBoxExport.Text;
            string authorName = textBoxAuthorName.Text;
            string configName = textBoxConfigName.Text;
            configListComboBox.Items[^1] = configName;
            configListComboBox.Items.Add("添加配置");
            ConfigHelper.currSelectIndex = configListComboBox.Items.Count - 2;
            ConfigHelper.AddPathConfig(excelPath, exportPath, authorName, configName);
            configListComboBox.SelectedIndex = ConfigHelper.currSelectIndex;
            OnConfigListComboBoxChanged(configListComboBox, null);
            MessageBox.Show(this, "添加成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 删除配置按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnBtnDeletePathConfigClick(object sender, EventArgs e)
        {
            if (MessageBox.Show("确认删除本条配置？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int selectedIndex = configListComboBox.SelectedIndex;
                configListComboBox.Items.RemoveAt(configListComboBox.SelectedIndex);
                selectedIndex--;

                if (selectedIndex < 0)
                {
                    selectedIndex = 0;
                }

                textBoxConfigName.Text = string.Empty;
                ConfigHelper.DeletePathConfig();
                configListComboBox.SelectedIndex = selectedIndex;

                MessageBox.Show(this, "删除成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 读取excel
        /// </summary>
        /// <param name="path"></param>
        private void LoadExcelFiles(string path)
        {
            ExportHelper.ResetExcel();
            excelList.Items.Clear();

            string[] files = Directory.GetFiles(path + "\\", "*", SearchOption.AllDirectories);

            if (files != null && files.Length > 0)
            {
                foreach (string strName in files)
                {
                    if (!Path.GetExtension(strName).Contains("xls") || strName.Contains('$'))//非excel文件或excel的缓存文件不进行读取
                    {
                        continue;
                    }

                    ExportHelper.AddExcel(strName);
                    excelList.Items.Add(strName);
                    excelList.SetItemChecked(excelList.Items.Count - 1, true);
                }

                MessageBox.Show(this, "读取成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(this, "读取失败，该路径下不存在Excel文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导出excel
        /// </summary>
        /// <param name="path"></param>
        private void ExportExcel(string path, string authorName)
        {
            if (excelList.Items != null && excelList.Items.Count > 0)
            {
                string exportPath = (!path.EndsWith("\\DataExport\\")) ? string.Format("{0}\\DataExport\\", path) : path;
                ExportHelper.Export(exportPath, authorName);
                MessageBox.Show(this, "创建成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                System.Diagnostics.Process.Start("explorer.exe", exportPath);
            }
            else
            {
                MessageBox.Show(this, "文件列表为空，无法导出", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导出语言选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnCodeTypeComboBoxSelectedIndexChanged(object sender, EventArgs e)
        {
            ExportHelper.SetCurrExporter(codeTypeComboBox.SelectedIndex);
        }

        /// <summary>
        /// 文件列表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnExcelListItemCheck(object sender, ItemCheckEventArgs e)
        {
            ExportHelper.SetExcelCanExport(e.Index, !excelList.GetItemChecked(e.Index));
        }

        /// <summary>
        /// 文本框文件路径拖拽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnTextBoxExcelDragEnter(object sender, DragEventArgs e)
        {
            if (e == null || e.Data == null)
            {
                return;
            }

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        /// <summary>
        /// 文本框文件路径拖拽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnTextBoxExcelDragDrop(object sender, DragEventArgs e)
        {
            if(e == null || e.Data == null)
            {
                return;
            }

            object data = e.Data.GetData(DataFormats.FileDrop);

            if (data is Array array)
            {
                object value = array.GetValue(0);
                textBoxExcel.Text = value != null ? value.ToString() : string.Empty;
            }
        }


        /// <summary>
        /// 文本框文件路径拖拽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnTextBoxExportDragEnter(object sender, DragEventArgs e)
        {
            if (e == null || e.Data == null)
            {
                return;
            }

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }


        /// <summary>
        /// 文本框文件路径拖拽
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnTextBoxExportDragDrop(object sender, DragEventArgs e)
        {
            if (e == null || e.Data == null)
            {
                return;
            }

            object data = e.Data.GetData(DataFormats.FileDrop);

            if (data is Array array) 
            {
                object value = array.GetValue(0);
                textBoxExport.Text = value != null ? value.ToString() : string.Empty; ;
            }
        }

        /// <summary>
        /// 路径配置列表选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnConfigListComboBoxChanged(object sender, EventArgs e)
        {
            bool showAddPathBtn = ConfigHelper.configData.Count < 1 || this.configListComboBox.SelectedIndex == ConfigHelper.configData.Count;
            btnModifyPathConfig.Visible = !showAddPathBtn;
            btnDelectPathConfig.Visible = !showAddPathBtn;
            btnAddPathConfig.Visible = showAddPathBtn;

            if (showAddPathBtn)
            {
                textBoxExcel.Text = string.Empty;
                textBoxExport.Text = string.Empty;
                textBoxAuthorName.Text = string.Empty;
                textBoxConfigName.Text = string.Empty;
            }
            else
            {
                ConfigHelper.currSelectIndex = configListComboBox.SelectedIndex;
                string[] config = ConfigHelper.GetCurrConfig();
                textBoxExcel.Text = config[0];
                textBoxExport.Text = config[1];
                textBoxAuthorName.Text = config[2];
                textBoxConfigName.Text = config[3];
            }
        }



        //异或因子
        //private byte[] xorScale = new byte[] { 45, 66, 38, 55, 23, 254, 9, 165, 90, 19, 41, 45, 201, 58, 55, 37, 254, 185, 165, 169, 19, 171 };//.data文件的xor加解密因子
        //private List<string> _allTalbeName = new List<string>();
    }
}

