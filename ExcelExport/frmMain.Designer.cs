using System;

namespace ExcelExport
{
    partial class ExcelExport
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCreate = new System.Windows.Forms.Button();
            this.codeTypeComboBox = new System.Windows.Forms.ComboBox();
            this.btnSelect = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.excelList = new System.Windows.Forms.CheckedListBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.btnModifyPathConfig = new System.Windows.Forms.Button();
            this.btnAddPathConfig = new System.Windows.Forms.Button();
            this.textBoxPathName = new System.Windows.Forms.TextBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.btnSelectExport = new System.Windows.Forms.Button();
            this.textBoxExport = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnSelectExcel = new System.Windows.Forms.Button();
            this.textBoxExcel = new System.Windows.Forms.TextBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.configListComboBox = new System.Windows.Forms.ComboBox();
            this.btnDelectPathConfig = new System.Windows.Forms.Button();
            this.groupBox2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCreate
            // 
            this.btnCreate.Location = new System.Drawing.Point(517, 17);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(145, 36);
            this.btnCreate.TabIndex = 0;
            this.btnCreate.Text = "导出数据";
            this.btnCreate.UseVisualStyleBackColor = true;
            this.btnCreate.Click += new System.EventHandler(this.OnBtnCreateClick);
            // 
            // codeTypeComboBox
            // 
            this.codeTypeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.codeTypeComboBox.FormattingEnabled = true;
            this.codeTypeComboBox.Items.AddRange(new object[] {
            "C#",
            "Lua"});
            this.codeTypeComboBox.Location = new System.Drawing.Point(6, 23);
            this.codeTypeComboBox.Name = "codeTypeComboBox";
            this.codeTypeComboBox.Size = new System.Drawing.Size(336, 20);
            this.codeTypeComboBox.TabIndex = 4;
            this.codeTypeComboBox.SelectedIndexChanged += new System.EventHandler(this.OnCodeTypeComboBoxSelectedIndexChanged);
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(366, 17);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(145, 36);
            this.btnSelect.TabIndex = 2;
            this.btnSelect.Text = "读取Excel文件";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.OnBtnSelectClick);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.codeTypeComboBox);
            this.groupBox2.Controls.Add(this.btnSelect);
            this.groupBox2.Controls.Add(this.btnCreate);
            this.groupBox2.Location = new System.Drawing.Point(9, 409);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(668, 62);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "读取|导出";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(-1, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(696, 505);
            this.tabControl1.TabIndex = 7;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(688, 479);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "读取&导出";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.groupBox1.Controls.Add(this.excelList);
            this.groupBox1.Location = new System.Drawing.Point(9, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(668, 399);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "选择文件";
            // 
            // excelList
            // 
            this.excelList.CheckOnClick = true;
            this.excelList.FormattingEnabled = true;
            this.excelList.Location = new System.Drawing.Point(6, 16);
            this.excelList.Name = "excelList";
            this.excelList.Size = new System.Drawing.Size(656, 372);
            this.excelList.TabIndex = 3;
            this.excelList.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.OnExcelListItemCheck);
            // 
            // tabPage2
            // 
            this.tabPage2.AutoScroll = true;
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Controls.Add(this.groupBox7);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(688, 479);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "路径配置";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.BackColor = System.Drawing.Color.Transparent;
            this.groupBox3.Controls.Add(this.groupBox6);
            this.groupBox3.Controls.Add(this.groupBox5);
            this.groupBox3.Controls.Add(this.groupBox4);
            this.groupBox3.Location = new System.Drawing.Point(6, 67);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(676, 243);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "路径配置";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.btnDelectPathConfig);
            this.groupBox6.Controls.Add(this.btnModifyPathConfig);
            this.groupBox6.Controls.Add(this.btnAddPathConfig);
            this.groupBox6.Controls.Add(this.textBoxPathName);
            this.groupBox6.Location = new System.Drawing.Point(6, 138);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(664, 75);
            this.groupBox6.TabIndex = 4;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "路径名称";
            // 
            // btnModifyPathConfig
            // 
            this.btnModifyPathConfig.Location = new System.Drawing.Point(499, 17);
            this.btnModifyPathConfig.Name = "btnModifyPathConfig";
            this.btnModifyPathConfig.Size = new System.Drawing.Size(159, 23);
            this.btnModifyPathConfig.TabIndex = 2;
            this.btnModifyPathConfig.Text = "修改配置";
            this.btnModifyPathConfig.UseVisualStyleBackColor = true;
            this.btnModifyPathConfig.Click += new System.EventHandler(this.OnBtnModifyPathConfigClick);
            // 
            // btnAddPathConfig
            // 
            this.btnAddPathConfig.Location = new System.Drawing.Point(499, 17);
            this.btnAddPathConfig.Name = "btnAddPathConfig";
            this.btnAddPathConfig.Size = new System.Drawing.Size(159, 23);
            this.btnAddPathConfig.TabIndex = 1;
            this.btnAddPathConfig.Text = "添加配置";
            this.btnAddPathConfig.UseVisualStyleBackColor = true;
            this.btnAddPathConfig.Click += new System.EventHandler(this.OnBtnAddPathConfigClick);
            // 
            // textBoxPathName
            // 
            this.textBoxPathName.AllowDrop = true;
            this.textBoxPathName.Location = new System.Drawing.Point(6, 18);
            this.textBoxPathName.Name = "textBoxPathName";
            this.textBoxPathName.Size = new System.Drawing.Size(487, 21);
            this.textBoxPathName.TabIndex = 0;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.btnSelectExport);
            this.groupBox5.Controls.Add(this.textBoxExport);
            this.groupBox5.Location = new System.Drawing.Point(6, 79);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(664, 53);
            this.groupBox5.TabIndex = 3;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "导出路径";
            // 
            // btnSelectExport
            // 
            this.btnSelectExport.Location = new System.Drawing.Point(499, 17);
            this.btnSelectExport.Name = "btnSelectExport";
            this.btnSelectExport.Size = new System.Drawing.Size(159, 23);
            this.btnSelectExport.TabIndex = 1;
            this.btnSelectExport.Text = "选择";
            this.btnSelectExport.UseVisualStyleBackColor = true;
            this.btnSelectExport.Click += new System.EventHandler(this.OnBtnSelectExportClick);
            // 
            // textBoxExport
            // 
            this.textBoxExport.AllowDrop = true;
            this.textBoxExport.Location = new System.Drawing.Point(6, 18);
            this.textBoxExport.Name = "textBoxExport";
            this.textBoxExport.ReadOnly = true;
            this.textBoxExport.Size = new System.Drawing.Size(487, 21);
            this.textBoxExport.TabIndex = 0;
            this.textBoxExport.DragDrop += new System.Windows.Forms.DragEventHandler(this.OnTextBoxExportDragDrop);
            this.textBoxExport.DragEnter += new System.Windows.Forms.DragEventHandler(this.OnTextBoxExportDragEnter);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btnSelectExcel);
            this.groupBox4.Controls.Add(this.textBoxExcel);
            this.groupBox4.Location = new System.Drawing.Point(6, 20);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(664, 53);
            this.groupBox4.TabIndex = 2;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Excel路径";
            // 
            // btnSelectExcel
            // 
            this.btnSelectExcel.Location = new System.Drawing.Point(499, 17);
            this.btnSelectExcel.Name = "btnSelectExcel";
            this.btnSelectExcel.Size = new System.Drawing.Size(159, 23);
            this.btnSelectExcel.TabIndex = 1;
            this.btnSelectExcel.Text = "选择";
            this.btnSelectExcel.UseVisualStyleBackColor = true;
            this.btnSelectExcel.Click += new System.EventHandler(this.OnBtnSelectExcelClick);
            // 
            // textBoxExcel
            // 
            this.textBoxExcel.AllowDrop = true;
            this.textBoxExcel.Location = new System.Drawing.Point(6, 18);
            this.textBoxExcel.Name = "textBoxExcel";
            this.textBoxExcel.ReadOnly = true;
            this.textBoxExcel.Size = new System.Drawing.Size(487, 21);
            this.textBoxExcel.TabIndex = 0;
            this.textBoxExcel.DragDrop += new System.Windows.Forms.DragEventHandler(this.OnTextBoxExcelDragDrop);
            this.textBoxExcel.DragEnter += new System.Windows.Forms.DragEventHandler(this.OnTextBoxExcelDragEnter);
            // 
            // groupBox7
            // 
            this.groupBox7.BackColor = System.Drawing.Color.Transparent;
            this.groupBox7.Controls.Add(this.configListComboBox);
            this.groupBox7.Location = new System.Drawing.Point(6, 6);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(676, 55);
            this.groupBox7.TabIndex = 4;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "当前路径";
            // 
            // configListComboBox
            // 
            this.configListComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.configListComboBox.FormattingEnabled = true;
            this.configListComboBox.Location = new System.Drawing.Point(6, 20);
            this.configListComboBox.Name = "configListComboBox";
            this.configListComboBox.Size = new System.Drawing.Size(664, 20);
            this.configListComboBox.TabIndex = 7;
            this.configListComboBox.SelectedIndexChanged += new System.EventHandler(this.OnConfigListComboBoxChanged);
            // 
            // btnDelectPathConfig
            // 
            this.btnDelectPathConfig.Location = new System.Drawing.Point(499, 46);
            this.btnDelectPathConfig.Name = "btnDelectPathConfig";
            this.btnDelectPathConfig.Size = new System.Drawing.Size(159, 23);
            this.btnDelectPathConfig.TabIndex = 3;
            this.btnDelectPathConfig.Text = "删除配置";
            this.btnDelectPathConfig.UseVisualStyleBackColor = true;
            this.btnDelectPathConfig.Click += new System.EventHandler(this.OnBtnDeletePathConfigClick);
            // 
            // ExcelExport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(692, 503);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "ExcelExport";
            this.Text = "Excel导出工具  WuWu";
            this.groupBox2.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCreate;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.ComboBox codeTypeComboBox;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.CheckedListBox excelList;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button btnSelectExport;
        private System.Windows.Forms.TextBox textBoxExport;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button btnSelectExcel;
        private System.Windows.Forms.TextBox textBoxExcel;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.Button btnAddPathConfig;
        private System.Windows.Forms.TextBox textBoxPathName;
        private System.Windows.Forms.ComboBox configListComboBox;
        private System.Windows.Forms.Button btnModifyPathConfig;
        private System.Windows.Forms.Button btnDelectPathConfig;
    }
}

