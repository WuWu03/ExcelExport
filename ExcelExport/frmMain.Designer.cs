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
            btnCreate = new System.Windows.Forms.Button();
            codeTypeComboBox = new System.Windows.Forms.ComboBox();
            btnSelect = new System.Windows.Forms.Button();
            groupBox2 = new System.Windows.Forms.GroupBox();
            tabControl1 = new System.Windows.Forms.TabControl();
            tabPage1 = new System.Windows.Forms.TabPage();
            groupBox1 = new System.Windows.Forms.GroupBox();
            excelList = new System.Windows.Forms.CheckedListBox();
            tabPage2 = new System.Windows.Forms.TabPage();
            groupBox7 = new System.Windows.Forms.GroupBox();
            groupBox5 = new System.Windows.Forms.GroupBox();
            textBoxAuthorName = new System.Windows.Forms.TextBox();
            groupBox6 = new System.Windows.Forms.GroupBox();
            btnDelectPathConfig = new System.Windows.Forms.Button();
            btnModifyPathConfig = new System.Windows.Forms.Button();
            btnAddPathConfig = new System.Windows.Forms.Button();
            textBoxConfigName = new System.Windows.Forms.TextBox();
            groupBox9 = new System.Windows.Forms.GroupBox();
            btnSelectExport = new System.Windows.Forms.Button();
            textBoxExport = new System.Windows.Forms.TextBox();
            groupBox8 = new System.Windows.Forms.GroupBox();
            btnSelectExcel = new System.Windows.Forms.Button();
            textBoxExcel = new System.Windows.Forms.TextBox();
            groupBox3 = new System.Windows.Forms.GroupBox();
            configListComboBox = new System.Windows.Forms.ComboBox();
            groupBox4 = new System.Windows.Forms.GroupBox();
            groupBox2.SuspendLayout();
            tabControl1.SuspendLayout();
            tabPage1.SuspendLayout();
            groupBox1.SuspendLayout();
            tabPage2.SuspendLayout();
            groupBox7.SuspendLayout();
            groupBox5.SuspendLayout();
            groupBox6.SuspendLayout();
            groupBox9.SuspendLayout();
            groupBox8.SuspendLayout();
            groupBox3.SuspendLayout();
            groupBox4.SuspendLayout();
            SuspendLayout();
            // 
            // btnCreate
            // 
            btnCreate.Location = new System.Drawing.Point(603, 24);
            btnCreate.Margin = new System.Windows.Forms.Padding(4);
            btnCreate.Name = "btnCreate";
            btnCreate.Size = new System.Drawing.Size(169, 51);
            btnCreate.TabIndex = 0;
            btnCreate.Text = "导出数据";
            btnCreate.UseVisualStyleBackColor = true;
            btnCreate.Click += OnBtnCreateClick;
            // 
            // codeTypeComboBox
            // 
            codeTypeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            codeTypeComboBox.FormattingEnabled = true;
            codeTypeComboBox.Location = new System.Drawing.Point(7, 38);
            codeTypeComboBox.Margin = new System.Windows.Forms.Padding(4);
            codeTypeComboBox.Name = "codeTypeComboBox";
            codeTypeComboBox.Size = new System.Drawing.Size(392, 25);
            codeTypeComboBox.TabIndex = 4;
            codeTypeComboBox.SelectedIndexChanged += OnCodeTypeComboBoxSelectedIndexChanged;
            // 
            // btnSelect
            // 
            btnSelect.Location = new System.Drawing.Point(427, 24);
            btnSelect.Margin = new System.Windows.Forms.Padding(4);
            btnSelect.Name = "btnSelect";
            btnSelect.Size = new System.Drawing.Size(169, 51);
            btnSelect.TabIndex = 2;
            btnSelect.Text = "读取Excel文件";
            btnSelect.UseVisualStyleBackColor = true;
            btnSelect.Click += OnBtnSelectClick;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(codeTypeComboBox);
            groupBox2.Controls.Add(btnSelect);
            groupBox2.Controls.Add(btnCreate);
            groupBox2.Location = new System.Drawing.Point(10, 579);
            groupBox2.Margin = new System.Windows.Forms.Padding(4);
            groupBox2.Name = "groupBox2";
            groupBox2.Padding = new System.Windows.Forms.Padding(4);
            groupBox2.Size = new System.Drawing.Size(779, 88);
            groupBox2.TabIndex = 6;
            groupBox2.TabStop = false;
            groupBox2.Text = "读取|导出";
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Location = new System.Drawing.Point(-1, 0);
            tabControl1.Margin = new System.Windows.Forms.Padding(4);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new System.Drawing.Size(812, 715);
            tabControl1.TabIndex = 7;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(groupBox1);
            tabPage1.Controls.Add(groupBox2);
            tabPage1.Location = new System.Drawing.Point(4, 26);
            tabPage1.Margin = new System.Windows.Forms.Padding(4);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new System.Windows.Forms.Padding(4);
            tabPage1.Size = new System.Drawing.Size(804, 685);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "读取&导出";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            groupBox1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            groupBox1.Controls.Add(excelList);
            groupBox1.Location = new System.Drawing.Point(10, 0);
            groupBox1.Margin = new System.Windows.Forms.Padding(4);
            groupBox1.Name = "groupBox1";
            groupBox1.Padding = new System.Windows.Forms.Padding(4);
            groupBox1.Size = new System.Drawing.Size(779, 549);
            groupBox1.TabIndex = 1;
            groupBox1.TabStop = false;
            groupBox1.Text = "选择文件";
            // 
            // excelList
            // 
            excelList.CheckOnClick = true;
            excelList.FormattingEnabled = true;
            excelList.Location = new System.Drawing.Point(7, 23);
            excelList.Margin = new System.Windows.Forms.Padding(4);
            excelList.Name = "excelList";
            excelList.Size = new System.Drawing.Size(765, 508);
            excelList.TabIndex = 3;
            excelList.ItemCheck += OnExcelListItemCheck;
            // 
            // tabPage2
            // 
            tabPage2.AutoScroll = true;
            tabPage2.Controls.Add(groupBox4);
            tabPage2.Controls.Add(groupBox7);
            tabPage2.Controls.Add(groupBox3);
            tabPage2.Location = new System.Drawing.Point(4, 26);
            tabPage2.Margin = new System.Windows.Forms.Padding(4);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new System.Windows.Forms.Padding(4);
            tabPage2.Size = new System.Drawing.Size(804, 685);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "路径配置";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            groupBox7.BackColor = System.Drawing.Color.Transparent;
            groupBox7.Controls.Add(groupBox9);
            groupBox7.Controls.Add(groupBox8);
            groupBox7.Location = new System.Drawing.Point(8, 321);
            groupBox7.Margin = new System.Windows.Forms.Padding(4);
            groupBox7.Name = "groupBox7";
            groupBox7.Padding = new System.Windows.Forms.Padding(4);
            groupBox7.Size = new System.Drawing.Size(789, 193);
            groupBox7.TabIndex = 5;
            groupBox7.TabStop = false;
            groupBox7.Text = "路径配置";
            // 
            // groupBox5
            // 
            groupBox5.Controls.Add(textBoxAuthorName);
            groupBox5.Location = new System.Drawing.Point(7, 23);
            groupBox5.Margin = new System.Windows.Forms.Padding(4);
            groupBox5.Name = "groupBox5";
            groupBox5.Padding = new System.Windows.Forms.Padding(4);
            groupBox5.Size = new System.Drawing.Size(774, 66);
            groupBox5.TabIndex = 5;
            groupBox5.TabStop = false;
            groupBox5.Text = "作者信息";
            // 
            // textBoxAuthorName
            // 
            textBoxAuthorName.AllowDrop = true;
            textBoxAuthorName.Location = new System.Drawing.Point(7, 26);
            textBoxAuthorName.Margin = new System.Windows.Forms.Padding(4);
            textBoxAuthorName.Name = "textBoxAuthorName";
            textBoxAuthorName.Size = new System.Drawing.Size(762, 23);
            textBoxAuthorName.TabIndex = 0;
            // 
            // groupBox6
            // 
            groupBox6.Controls.Add(btnDelectPathConfig);
            groupBox6.Controls.Add(btnModifyPathConfig);
            groupBox6.Controls.Add(btnAddPathConfig);
            groupBox6.Controls.Add(textBoxConfigName);
            groupBox6.Location = new System.Drawing.Point(7, 97);
            groupBox6.Margin = new System.Windows.Forms.Padding(4);
            groupBox6.Name = "groupBox6";
            groupBox6.Padding = new System.Windows.Forms.Padding(4);
            groupBox6.Size = new System.Drawing.Size(774, 99);
            groupBox6.TabIndex = 4;
            groupBox6.TabStop = false;
            groupBox6.Text = "配置名称";
            // 
            // btnDelectPathConfig
            // 
            btnDelectPathConfig.Location = new System.Drawing.Point(583, 57);
            btnDelectPathConfig.Margin = new System.Windows.Forms.Padding(4);
            btnDelectPathConfig.Name = "btnDelectPathConfig";
            btnDelectPathConfig.Size = new System.Drawing.Size(186, 29);
            btnDelectPathConfig.TabIndex = 3;
            btnDelectPathConfig.Text = "删除配置";
            btnDelectPathConfig.UseVisualStyleBackColor = true;
            btnDelectPathConfig.Click += OnBtnDeletePathConfigClick;
            // 
            // btnModifyPathConfig
            // 
            btnModifyPathConfig.Location = new System.Drawing.Point(583, 21);
            btnModifyPathConfig.Margin = new System.Windows.Forms.Padding(4);
            btnModifyPathConfig.Name = "btnModifyPathConfig";
            btnModifyPathConfig.Size = new System.Drawing.Size(186, 29);
            btnModifyPathConfig.TabIndex = 2;
            btnModifyPathConfig.Text = "修改配置";
            btnModifyPathConfig.UseVisualStyleBackColor = true;
            btnModifyPathConfig.Click += OnBtnModifyPathConfigClick;
            // 
            // btnAddPathConfig
            // 
            btnAddPathConfig.Location = new System.Drawing.Point(583, 20);
            btnAddPathConfig.Margin = new System.Windows.Forms.Padding(4);
            btnAddPathConfig.Name = "btnAddPathConfig";
            btnAddPathConfig.Size = new System.Drawing.Size(186, 29);
            btnAddPathConfig.TabIndex = 1;
            btnAddPathConfig.Text = "添加配置";
            btnAddPathConfig.UseVisualStyleBackColor = true;
            btnAddPathConfig.Click += OnBtnAddPathConfigClick;
            // 
            // textBoxConfigName
            // 
            textBoxConfigName.AllowDrop = true;
            textBoxConfigName.Location = new System.Drawing.Point(7, 24);
            textBoxConfigName.Margin = new System.Windows.Forms.Padding(4);
            textBoxConfigName.Name = "textBoxConfigName";
            textBoxConfigName.Size = new System.Drawing.Size(568, 23);
            textBoxConfigName.TabIndex = 0;
            // 
            // groupBox9
            // 
            groupBox9.Controls.Add(btnSelectExport);
            groupBox9.Controls.Add(textBoxExport);
            groupBox9.Location = new System.Drawing.Point(7, 102);
            groupBox9.Margin = new System.Windows.Forms.Padding(4);
            groupBox9.Name = "groupBox9";
            groupBox9.Padding = new System.Windows.Forms.Padding(4);
            groupBox9.Size = new System.Drawing.Size(773, 66);
            groupBox9.TabIndex = 3;
            groupBox9.TabStop = false;
            groupBox9.Text = "导出路径";
            // 
            // btnSelectExport
            // 
            btnSelectExport.Location = new System.Drawing.Point(582, 23);
            btnSelectExport.Margin = new System.Windows.Forms.Padding(4);
            btnSelectExport.Name = "btnSelectExport";
            btnSelectExport.Size = new System.Drawing.Size(186, 29);
            btnSelectExport.TabIndex = 1;
            btnSelectExport.Text = "选择";
            btnSelectExport.UseVisualStyleBackColor = true;
            btnSelectExport.Click += OnBtnSelectExportClick;
            // 
            // textBoxExport
            // 
            textBoxExport.AllowDrop = true;
            textBoxExport.Location = new System.Drawing.Point(7, 26);
            textBoxExport.Margin = new System.Windows.Forms.Padding(4);
            textBoxExport.Name = "textBoxExport";
            textBoxExport.ReadOnly = true;
            textBoxExport.Size = new System.Drawing.Size(568, 23);
            textBoxExport.TabIndex = 0;
            textBoxExport.DragDrop += OnTextBoxExportDragDrop;
            textBoxExport.DragEnter += OnTextBoxExportDragEnter;
            // 
            // groupBox8
            // 
            groupBox8.Controls.Add(btnSelectExcel);
            groupBox8.Controls.Add(textBoxExcel);
            groupBox8.Location = new System.Drawing.Point(7, 28);
            groupBox8.Margin = new System.Windows.Forms.Padding(4);
            groupBox8.Name = "groupBox8";
            groupBox8.Padding = new System.Windows.Forms.Padding(4);
            groupBox8.Size = new System.Drawing.Size(773, 66);
            groupBox8.TabIndex = 2;
            groupBox8.TabStop = false;
            groupBox8.Text = "Excel路径";
            // 
            // btnSelectExcel
            // 
            btnSelectExcel.Location = new System.Drawing.Point(582, 23);
            btnSelectExcel.Margin = new System.Windows.Forms.Padding(4);
            btnSelectExcel.Name = "btnSelectExcel";
            btnSelectExcel.Size = new System.Drawing.Size(186, 29);
            btnSelectExcel.TabIndex = 1;
            btnSelectExcel.Text = "选择";
            btnSelectExcel.UseVisualStyleBackColor = true;
            btnSelectExcel.Click += OnBtnSelectExcelClick;
            // 
            // textBoxExcel
            // 
            textBoxExcel.AllowDrop = true;
            textBoxExcel.Location = new System.Drawing.Point(7, 26);
            textBoxExcel.Margin = new System.Windows.Forms.Padding(4);
            textBoxExcel.Name = "textBoxExcel";
            textBoxExcel.ReadOnly = true;
            textBoxExcel.Size = new System.Drawing.Size(568, 23);
            textBoxExcel.TabIndex = 0;
            textBoxExcel.DragDrop += OnTextBoxExcelDragDrop;
            textBoxExcel.DragEnter += OnTextBoxExcelDragEnter;
            // 
            // groupBox3
            // 
            groupBox3.BackColor = System.Drawing.Color.Transparent;
            groupBox3.Controls.Add(configListComboBox);
            groupBox3.Location = new System.Drawing.Point(7, 8);
            groupBox3.Margin = new System.Windows.Forms.Padding(4);
            groupBox3.Name = "groupBox3";
            groupBox3.Padding = new System.Windows.Forms.Padding(4);
            groupBox3.Size = new System.Drawing.Size(789, 79);
            groupBox3.TabIndex = 4;
            groupBox3.TabStop = false;
            groupBox3.Text = "当前配置";
            // 
            // configListComboBox
            // 
            configListComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            configListComboBox.FormattingEnabled = true;
            configListComboBox.Location = new System.Drawing.Point(7, 28);
            configListComboBox.Margin = new System.Windows.Forms.Padding(4);
            configListComboBox.Name = "configListComboBox";
            configListComboBox.Size = new System.Drawing.Size(774, 25);
            configListComboBox.TabIndex = 7;
            configListComboBox.SelectedIndexChanged += OnConfigListComboBoxChanged;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(groupBox5);
            groupBox4.Controls.Add(groupBox6);
            groupBox4.Location = new System.Drawing.Point(7, 94);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new System.Drawing.Size(789, 220);
            groupBox4.TabIndex = 6;
            groupBox4.TabStop = false;
            groupBox4.Text = "基本配置";
            // 
            // ExcelExport
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(807, 713);
            Controls.Add(tabControl1);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            Margin = new System.Windows.Forms.Padding(4);
            MaximizeBox = false;
            Name = "ExcelExport";
            Text = "Excel导出工具  WuWu";
            groupBox2.ResumeLayout(false);
            tabControl1.ResumeLayout(false);
            tabPage1.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            tabPage2.ResumeLayout(false);
            groupBox7.ResumeLayout(false);
            groupBox5.ResumeLayout(false);
            groupBox5.PerformLayout();
            groupBox6.ResumeLayout(false);
            groupBox6.PerformLayout();
            groupBox9.ResumeLayout(false);
            groupBox9.PerformLayout();
            groupBox8.ResumeLayout(false);
            groupBox8.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox4.ResumeLayout(false);
            ResumeLayout(false);

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
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.Button btnSelectExport;
        private System.Windows.Forms.TextBox textBoxExport;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.Button btnSelectExcel;
        private System.Windows.Forms.TextBox textBoxExcel;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.Button btnAddPathConfig;
        private System.Windows.Forms.TextBox textBoxConfigName;
        private System.Windows.Forms.ComboBox configListComboBox;
        private System.Windows.Forms.Button btnModifyPathConfig;
        private System.Windows.Forms.Button btnDelectPathConfig;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.TextBox textBoxAuthorName;
        private System.Windows.Forms.GroupBox groupBox4;
    }
}

