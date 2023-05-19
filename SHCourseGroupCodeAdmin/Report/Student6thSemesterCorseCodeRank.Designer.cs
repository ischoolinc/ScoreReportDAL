
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
            this.chkNCredit = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.chkNScore = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
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
            this.lnkDefault.Location = new System.Drawing.Point(306, 189);
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
            this.lnkViewMapColumns.Location = new System.Drawing.Point(188, 189);
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
            this.lnkViewTemplate.Location = new System.Drawing.Point(9, 189);
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
            this.lnkChangeTemplate.Location = new System.Drawing.Point(98, 189);
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
            this.btnPrint.Location = new System.Drawing.Point(455, 186);
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
            this.btnCancel.Location = new System.Drawing.Point(528, 186);
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
            this.labelX1.Location = new System.Drawing.Point(14, 35);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(579, 120);
            this.labelX1.TabIndex = 26;
            this.labelX1.Text = "1. 列印將同時產生PDF和Word。\r\n2. 學校上傳：請使用PDF檔案。\r\n3. 學生上傳：請使用「學生：其它→電子報表上傳」，選擇「系統編號」選項及Word" +
    "檔案，\r\n　分析後勾選「上傳時Word檔轉為PDF」。\r\n4. 產出條件： 3年級學生且須有第六學期成績或修課紀錄。\r\n5. 預設課程類型:計算學分且需要計分之" +
    "課程(可透過加選其他類型課程)";
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
            this.chkAccordingToClass.Location = new System.Drawing.Point(10, 160);
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
            this.iptSemester.Location = new System.Drawing.Point(181, 7);
            this.iptSemester.MaxValue = 2;
            this.iptSemester.MinValue = 2;
            this.iptSemester.Name = "iptSemester";
            this.iptSemester.Size = new System.Drawing.Size(34, 25);
            this.iptSemester.TabIndex = 29;
            this.iptSemester.Value = 2;
            this.iptSemester.WatermarkEnabled = false;
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
            this.iptSchoolYear.Location = new System.Drawing.Point(67, 7);
            this.iptSchoolYear.MaxValue = 200;
            this.iptSchoolYear.MinValue = 108;
            this.iptSchoolYear.Name = "iptSchoolYear";
            this.iptSchoolYear.Size = new System.Drawing.Size(46, 25);
            this.iptSchoolYear.TabIndex = 28;
            this.iptSchoolYear.Value = 110;
            this.iptSchoolYear.WatermarkEnabled = false;
            // 
            // chkNCredit
            // 
            this.chkNCredit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkNCredit.AutoSize = true;
            this.chkNCredit.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.chkNCredit.BackgroundStyle.Class = "";
            this.chkNCredit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkNCredit.Location = new System.Drawing.Point(310, 160);
            this.chkNCredit.Name = "chkNCredit";
            this.chkNCredit.Size = new System.Drawing.Size(80, 21);
            this.chkNCredit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkNCredit.TabIndex = 30;
            this.chkNCredit.Text = "不計學分";
            // 
            // chkNScore
            // 
            this.chkNScore.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkNScore.AutoSize = true;
            this.chkNScore.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.chkNScore.BackgroundStyle.Class = "";
            this.chkNScore.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkNScore.Location = new System.Drawing.Point(404, 160);
            this.chkNScore.Name = "chkNScore";
            this.chkNScore.Size = new System.Drawing.Size(80, 21);
            this.chkNScore.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkNScore.TabIndex = 31;
            this.chkNScore.Text = "不需計分";
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(195, 160);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(117, 21);
            this.labelX2.TabIndex = 32;
            this.labelX2.Text = "加選其他類型課程:";
            // 
            // labelX3
            // 
            this.labelX3.AutoSize = true;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(383, 160);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(20, 21);
            this.labelX3.TabIndex = 33;
            this.labelX3.Text = "及";
            // 
            // Student6thSemesterCorseCodeRank
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(605, 217);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.chkNScore);
            this.Controls.Add(this.chkNCredit);
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
        private DevComponents.DotNetBar.Controls.CheckBoxX chkNCredit;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkNScore;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.LabelX labelX3;
    }
}