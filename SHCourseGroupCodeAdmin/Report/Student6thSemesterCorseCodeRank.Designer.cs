
namespace SHCourseGroupCodeAdmin.Report
{
    partial class Student6thSemesterCorseCodeRank
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.labelX9 = new DevComponents.DotNetBar.LabelX();
            this.labelX7 = new DevComponents.DotNetBar.LabelX();
            this.lnkDefault = new System.Windows.Forms.LinkLabel();
            this.lnkViewMapColumns = new System.Windows.Forms.LinkLabel();
            this.lnkViewTemplate = new System.Windows.Forms.LinkLabel();
            this.lnkChangeTemplate = new System.Windows.Forms.LinkLabel();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.chkAccordingToClass = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.iptSemester = new DevComponents.Editors.IntegerInput();
            this.iptSchoolYear = new DevComponents.Editors.IntegerInput();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).BeginInit();
            this.SuspendLayout();
            // 
            // labelX9
            // 
            this.labelX9.AutoSize = true;
            this.labelX9.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX9.BackgroundStyle.Class = "";
            this.labelX9.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX9.Location = new System.Drawing.Point(142, 8);
            this.labelX9.Name = "labelX9";
            this.labelX9.Size = new System.Drawing.Size(47, 21);
            this.labelX9.TabIndex = 8;
            this.labelX9.Text = "學期：";
            // 
            // labelX7
            // 
            this.labelX7.AutoSize = true;
            this.labelX7.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX7.BackgroundStyle.Class = "";
            this.labelX7.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX7.Location = new System.Drawing.Point(14, 8);
            this.labelX7.Name = "labelX7";
            this.labelX7.Size = new System.Drawing.Size(60, 21);
            this.labelX7.TabIndex = 9;
            this.labelX7.Text = "學年度：";
            // 
            // lnkDefault
            // 
            this.lnkDefault.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lnkDefault.AutoSize = true;
            this.lnkDefault.BackColor = System.Drawing.Color.Transparent;
            this.lnkDefault.Location = new System.Drawing.Point(306, 173);
            this.lnkDefault.Name = "lnkDefault";
            this.lnkDefault.Size = new System.Drawing.Size(86, 17);
            this.lnkDefault.TabIndex = 21;
            this.lnkDefault.TabStop = true;
            this.lnkDefault.Text = "下載預設樣板";
            this.lnkDefault.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkDefault_LinkClicked);
            // 
            // lnkViewMapColumns
            // 
            this.lnkViewMapColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lnkViewMapColumns.AutoSize = true;
            this.lnkViewMapColumns.BackColor = System.Drawing.Color.Transparent;
            this.lnkViewMapColumns.Location = new System.Drawing.Point(188, 173);
            this.lnkViewMapColumns.Name = "lnkViewMapColumns";
            this.lnkViewMapColumns.Size = new System.Drawing.Size(112, 17);
            this.lnkViewMapColumns.TabIndex = 20;
            this.lnkViewMapColumns.TabStop = true;
            this.lnkViewMapColumns.Text = "下載合併欄位總表";
            this.lnkViewMapColumns.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkViewMapColumns_LinkClicked);
            // 
            // lnkViewTemplate
            // 
            this.lnkViewTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lnkViewTemplate.AutoSize = true;
            this.lnkViewTemplate.BackColor = System.Drawing.Color.Transparent;
            this.lnkViewTemplate.Location = new System.Drawing.Point(9, 173);
            this.lnkViewTemplate.Name = "lnkViewTemplate";
            this.lnkViewTemplate.Size = new System.Drawing.Size(86, 17);
            this.lnkViewTemplate.TabIndex = 18;
            this.lnkViewTemplate.TabStop = true;
            this.lnkViewTemplate.Text = "檢視套印樣板";
            this.lnkViewTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkViewTemplate_LinkClicked);
            // 
            // lnkChangeTemplate
            // 
            this.lnkChangeTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lnkChangeTemplate.AutoSize = true;
            this.lnkChangeTemplate.BackColor = System.Drawing.Color.Transparent;
            this.lnkChangeTemplate.Location = new System.Drawing.Point(98, 173);
            this.lnkChangeTemplate.Name = "lnkChangeTemplate";
            this.lnkChangeTemplate.Size = new System.Drawing.Size(86, 17);
            this.lnkChangeTemplate.TabIndex = 19;
            this.lnkChangeTemplate.TabStop = true;
            this.lnkChangeTemplate.Text = "變更套印樣板";
            this.lnkChangeTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkChangeTemplate_LinkClicked);
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnPrint.Enabled = false;
            this.btnPrint.Location = new System.Drawing.Point(455, 167);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(67, 23);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 22;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(528, 167);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(67, 23);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 23;
            this.btnCancel.Text = "離開";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // labelX1
            // 
            this.labelX1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(14, 36);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(579, 102);
            this.labelX1.TabIndex = 26;
            this.labelX1.Text = "1. 列印將同時產生PDF和Word。\r\n2. 學校上傳：請使用PDF檔案。\r\n3. 學生上傳：請使用「學生：其它→電子報表上傳」，選擇「系統編號」選項及Word" +
    "檔案，\r\n　分析後勾選「上傳時Word檔轉為PDF」。\r\n4. 產出條件： 3年級學生且須有110-2學期成績或修課紀錄。";
            // 
            // chkAccordingToClass
            // 
            this.chkAccordingToClass.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkAccordingToClass.AutoSize = true;
            this.chkAccordingToClass.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.chkAccordingToClass.BackgroundStyle.Class = "";
            this.chkAccordingToClass.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkAccordingToClass.Location = new System.Drawing.Point(10, 144);
            this.chkAccordingToClass.Name = "chkAccordingToClass";
            this.chkAccordingToClass.Size = new System.Drawing.Size(174, 21);
            this.chkAccordingToClass.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkAccordingToClass.TabIndex = 27;
            this.chkAccordingToClass.Text = "依據班級名稱建立資料夾";
            // 
            // iptSemester
            // 
            this.iptSemester.AllowEmptyState = false;
            this.iptSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSemester.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSemester.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSemester.Enabled = false;
            this.iptSemester.Location = new System.Drawing.Point(181, 6);
            this.iptSemester.MaxValue = 2;
            this.iptSemester.MinValue = 2;
            this.iptSemester.Name = "iptSemester";
            this.iptSemester.Size = new System.Drawing.Size(34, 25);
            this.iptSemester.TabIndex = 29;
            this.iptSemester.Value = 2;
            // 
            // iptSchoolYear
            // 
            this.iptSchoolYear.AllowEmptyState = false;
            this.iptSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSchoolYear.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSchoolYear.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSchoolYear.Enabled = false;
            this.iptSchoolYear.Location = new System.Drawing.Point(74, 6);
            this.iptSchoolYear.MaxValue = 110;
            this.iptSchoolYear.MinValue = 110;
            this.iptSchoolYear.Name = "iptSchoolYear";
            this.iptSchoolYear.Size = new System.Drawing.Size(46, 25);
            this.iptSchoolYear.TabIndex = 28;
            this.iptSchoolYear.Value = 110;
            // 
            // Student6thSemesterCorseCodeRank
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(605, 199);
            this.Controls.Add(this.iptSemester);
            this.Controls.Add(this.iptSchoolYear);
            this.Controls.Add(this.chkAccordingToClass);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.labelX9);
            this.Controls.Add(this.labelX7);
            this.Controls.Add(this.lnkDefault);
            this.Controls.Add(this.lnkViewMapColumns);
            this.Controls.Add(this.lnkViewTemplate);
            this.Controls.Add(this.lnkChangeTemplate);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.btnCancel);
            this.DoubleBuffered = true;
            this.MaximumSize = new System.Drawing.Size(800, 300);
            this.MinimumSize = new System.Drawing.Size(621, 187);
            this.Name = "Student6thSemesterCorseCodeRank";
            this.Text = "學生第6學期修課紀錄";
            this.Load += new System.EventHandler(this.Student6thSemesterCorseCodeRank_Load);
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private DevComponents.DotNetBar.LabelX labelX9;
        private DevComponents.DotNetBar.LabelX labelX7;
        private System.Windows.Forms.LinkLabel lnkDefault;
        private System.Windows.Forms.LinkLabel lnkViewMapColumns;
        private System.Windows.Forms.LinkLabel lnkViewTemplate;
        private System.Windows.Forms.LinkLabel lnkChangeTemplate;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkAccordingToClass;
        private DevComponents.Editors.IntegerInput iptSemester;
        private DevComponents.Editors.IntegerInput iptSchoolYear;
    }
}