namespace SHCourseGroupCodeAdmin.UIForm
{
    partial class frmCheckSemsScoreCourseCode
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
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.btnRun = new DevComponents.DotNetBar.ButtonX();
            this.iptSemester = new DevComponents.Editors.IntegerInput();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.iptSchoolYear = new DevComponents.Editors.IntegerInput();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.cboGradeYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX4 = new DevComponents.DotNetBar.LabelX();
            this.labelX5 = new DevComponents.DotNetBar.LabelX();
            this.txtDesc = new DevComponents.DotNetBar.Controls.TextBoxX();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).BeginInit();
            this.SuspendLayout();
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
            this.labelX1.Location = new System.Drawing.Point(30, 20);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(60, 21);
            this.labelX1.TabIndex = 17;
            this.labelX1.Text = "檢核目前";
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.AutoSize = true;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.Location = new System.Drawing.Point(300, 381);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 30);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 15;
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
            this.btnRun.Location = new System.Drawing.Point(206, 381);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 30);
            this.btnRun.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnRun.TabIndex = 14;
            this.btnRun.Text = "執行";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
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
            this.iptSemester.Location = new System.Drawing.Point(183, 58);
            this.iptSemester.MaxValue = 2;
            this.iptSemester.MinValue = 1;
            this.iptSemester.Name = "iptSemester";
            this.iptSemester.Size = new System.Drawing.Size(47, 25);
            this.iptSemester.TabIndex = 20;
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
            this.labelX3.Location = new System.Drawing.Point(236, 60);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(101, 21);
            this.labelX3.TabIndex = 21;
            this.labelX3.Text = "學期的學期成績";
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
            this.iptSchoolYear.Location = new System.Drawing.Point(53, 58);
            this.iptSchoolYear.MaxValue = 200;
            this.iptSchoolYear.MinValue = 108;
            this.iptSchoolYear.Name = "iptSchoolYear";
            this.iptSchoolYear.Size = new System.Drawing.Size(59, 25);
            this.iptSchoolYear.TabIndex = 18;
            this.iptSchoolYear.Value = 108;
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
            this.labelX2.Location = new System.Drawing.Point(117, 60);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(67, 21);
            this.labelX2.TabIndex = 19;
            this.labelX2.Text = "學年度  第";
            // 
            // cboGradeYear
            // 
            this.cboGradeYear.DisplayMember = "Text";
            this.cboGradeYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboGradeYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboGradeYear.FormattingEnabled = true;
            this.cboGradeYear.ItemHeight = 19;
            this.cboGradeYear.Location = new System.Drawing.Point(96, 18);
            this.cboGradeYear.Name = "cboGradeYear";
            this.cboGradeYear.Size = new System.Drawing.Size(68, 25);
            this.cboGradeYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboGradeYear.TabIndex = 22;
            // 
            // labelX4
            // 
            this.labelX4.AutoSize = true;
            this.labelX4.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX4.BackgroundStyle.Class = "";
            this.labelX4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX4.Location = new System.Drawing.Point(170, 20);
            this.labelX4.Name = "labelX4";
            this.labelX4.Size = new System.Drawing.Size(60, 21);
            this.labelX4.TabIndex = 23;
            this.labelX4.Text = "年級學生";
            // 
            // labelX5
            // 
            this.labelX5.AutoSize = true;
            this.labelX5.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX5.BackgroundStyle.Class = "";
            this.labelX5.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX5.Location = new System.Drawing.Point(30, 60);
            this.labelX5.Name = "labelX5";
            this.labelX5.Size = new System.Drawing.Size(20, 21);
            this.labelX5.TabIndex = 24;
            this.labelX5.Text = "於";
            // 
            // txtDesc
            // 
            this.txtDesc.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            // 
            // 
            // 
            this.txtDesc.Border.Class = "TextBoxBorder";
            this.txtDesc.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtDesc.Location = new System.Drawing.Point(12, 93);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.Size = new System.Drawing.Size(361, 272);
            this.txtDesc.TabIndex = 25;
            // 
            // frmCheckSemsScoreCourseCode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(387, 419);
            this.Controls.Add(this.txtDesc);
            this.Controls.Add(this.labelX5);
            this.Controls.Add(this.labelX4);
            this.Controls.Add(this.cboGradeYear);
            this.Controls.Add(this.iptSemester);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.iptSchoolYear);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRun);
            this.DoubleBuffered = true;
            this.Name = "frmCheckSemsScoreCourseCode";
            this.Text = "學期成績檢核課程代碼";
            this.Load += new System.EventHandler(this.frmCheckSemsScoreCourseCode_Load);
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private DevComponents.DotNetBar.ButtonX btnRun;
        private DevComponents.Editors.IntegerInput iptSemester;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.Editors.IntegerInput iptSchoolYear;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboGradeYear;
        private DevComponents.DotNetBar.LabelX labelX4;
        private DevComponents.DotNetBar.LabelX labelX5;
        private DevComponents.DotNetBar.Controls.TextBoxX txtDesc;
    }
}