namespace SHCourseGroupCodeAdmin.UIForm
{
    partial class frmCheckSCAttendCourseCode
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
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.btnRun = new DevComponents.DotNetBar.ButtonX();
            this.iptSchoolYear = new DevComponents.Editors.IntegerInput();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.iptSemester = new DevComponents.Editors.IntegerInput();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.cboGradeYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.txtDesc = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.chkPreScoreXls = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.chkPreScoreXlsN = new DevComponents.DotNetBar.Controls.CheckBoxX();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.AutoSize = true;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.Location = new System.Drawing.Point(281, 322);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 30);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 11;
            this.btnCancel.Text = "離開";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnRun
            // 
            this.btnRun.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnRun.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRun.AutoSize = true;
            this.btnRun.BackColor = System.Drawing.Color.Transparent;
            this.btnRun.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnRun.Location = new System.Drawing.Point(181, 322);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 30);
            this.btnRun.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnRun.TabIndex = 10;
            this.btnRun.Text = "執行";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
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
            this.iptSchoolYear.Location = new System.Drawing.Point(71, 17);
            this.iptSchoolYear.MaxValue = 200;
            this.iptSchoolYear.MinValue = 100;
            this.iptSchoolYear.Name = "iptSchoolYear";
            this.iptSchoolYear.Size = new System.Drawing.Size(59, 25);
            this.iptSchoolYear.TabIndex = 14;
            this.iptSchoolYear.Value = 100;
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
            this.labelX2.Location = new System.Drawing.Point(19, 19);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(47, 21);
            this.labelX2.TabIndex = 15;
            this.labelX2.Text = "學年度";
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
            this.iptSemester.Location = new System.Drawing.Point(181, 17);
            this.iptSemester.MaxValue = 2;
            this.iptSemester.MinValue = 1;
            this.iptSemester.Name = "iptSemester";
            this.iptSemester.Size = new System.Drawing.Size(51, 25);
            this.iptSemester.TabIndex = 16;
            this.iptSemester.Value = 1;
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
            this.labelX3.Location = new System.Drawing.Point(143, 19);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(34, 21);
            this.labelX3.TabIndex = 17;
            this.labelX3.Text = "學期";
            // 
            // cboGradeYear
            // 
            this.cboGradeYear.DisplayMember = "Text";
            this.cboGradeYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboGradeYear.FormattingEnabled = true;
            this.cboGradeYear.ItemHeight = 19;
            this.cboGradeYear.Location = new System.Drawing.Point(288, 17);
            this.cboGradeYear.Name = "cboGradeYear";
            this.cboGradeYear.Size = new System.Drawing.Size(68, 25);
            this.cboGradeYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboGradeYear.TabIndex = 18;
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(248, 19);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(34, 21);
            this.labelX1.TabIndex = 19;
            this.labelX1.Text = "年級";
            // 
            // txtDesc
            // 
            // 
            // 
            // 
            this.txtDesc.Border.Class = "TextBoxBorder";
            this.txtDesc.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtDesc.Location = new System.Drawing.Point(19, 58);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.Size = new System.Drawing.Size(344, 241);
            this.txtDesc.TabIndex = 20;
            // 
            // chkPreScoreXls
            // 
            this.chkPreScoreXls.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkPreScoreXls.AutoSize = true;
            this.chkPreScoreXls.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.chkPreScoreXls.BackgroundStyle.Class = "";
            this.chkPreScoreXls.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkPreScoreXls.Location = new System.Drawing.Point(13, 308);
            this.chkPreScoreXls.Name = "chkPreScoreXls";
            this.chkPreScoreXls.Size = new System.Drawing.Size(156, 21);
            this.chkPreScoreXls.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkPreScoreXls.TabIndex = 21;
            this.chkPreScoreXls.Text = "產生預檢成績名冊(日)";
            this.chkPreScoreXls.Click += new System.EventHandler(this.chkPreScoreXls_Click);
            // 
            // chkPreScoreXlsN
            // 
            this.chkPreScoreXlsN.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkPreScoreXlsN.AutoSize = true;
            this.chkPreScoreXlsN.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.chkPreScoreXlsN.BackgroundStyle.Class = "";
            this.chkPreScoreXlsN.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkPreScoreXlsN.Location = new System.Drawing.Point(13, 333);
            this.chkPreScoreXlsN.Name = "chkPreScoreXlsN";
            this.chkPreScoreXlsN.Size = new System.Drawing.Size(156, 21);
            this.chkPreScoreXlsN.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkPreScoreXlsN.TabIndex = 22;
            this.chkPreScoreXlsN.Text = "產生預檢成績名冊(進)";
            this.chkPreScoreXlsN.Click += new System.EventHandler(this.chkPreScoreXlsN_Click);
            // 
            // frmCheckSCAttendCourseCode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(377, 361);
            this.Controls.Add(this.chkPreScoreXlsN);
            this.Controls.Add(this.chkPreScoreXls);
            this.Controls.Add(this.txtDesc);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.cboGradeYear);
            this.Controls.Add(this.iptSemester);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.iptSchoolYear);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRun);
            this.DoubleBuffered = true;
            this.Name = "frmCheckSCAttendCourseCode";
            this.Text = "修課檢核課程代碼";
            this.Load += new System.EventHandler(this.frmCheckSCAttendCourseCode_Load);
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnCancel;
        private DevComponents.DotNetBar.ButtonX btnRun;
        private DevComponents.Editors.IntegerInput iptSchoolYear;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.Editors.IntegerInput iptSemester;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboGradeYear;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.TextBoxX txtDesc;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkPreScoreXls;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkPreScoreXlsN;
    }
}