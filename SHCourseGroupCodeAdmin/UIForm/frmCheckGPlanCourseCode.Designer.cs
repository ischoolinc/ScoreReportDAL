namespace SHCourseGroupCodeAdmin.UIForm
{
    partial class frmCheckGPlanCourseCode
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
            this.cbxSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cbxSemester = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.btnRun = new DevComponents.DotNetBar.ButtonX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.txtDesc = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.SuspendLayout();
            // 
            // cbxSchoolYear
            // 
            this.cbxSchoolYear.DisplayMember = "Text";
            this.cbxSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxSchoolYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxSchoolYear.FormattingEnabled = true;
            this.cbxSchoolYear.ItemHeight = 19;
            this.cbxSchoolYear.Location = new System.Drawing.Point(71, 20);
            this.cbxSchoolYear.Name = "cbxSchoolYear";
            this.cbxSchoolYear.Size = new System.Drawing.Size(58, 25);
            this.cbxSchoolYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxSchoolYear.TabIndex = 0;
            // 
            // cbxSemester
            // 
            this.cbxSemester.DisplayMember = "Text";
            this.cbxSemester.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxSemester.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxSemester.FormattingEnabled = true;
            this.cbxSemester.ItemHeight = 19;
            this.cbxSemester.Location = new System.Drawing.Point(197, 20);
            this.cbxSemester.Name = "cbxSemester";
            this.cbxSemester.Size = new System.Drawing.Size(42, 25);
            this.cbxSemester.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxSemester.TabIndex = 1;
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
            this.labelX1.Location = new System.Drawing.Point(13, 22);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(47, 21);
            this.labelX1.TabIndex = 2;
            this.labelX1.Text = "學年度";
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
            this.labelX2.Location = new System.Drawing.Point(157, 22);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(34, 21);
            this.labelX2.TabIndex = 3;
            this.labelX2.Text = "學期";
            // 
            // btnRun
            // 
            this.btnRun.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnRun.AutoSize = true;
            this.btnRun.BackColor = System.Drawing.Color.Transparent;
            this.btnRun.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnRun.Location = new System.Drawing.Point(116, 251);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 25);
            this.btnRun.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnRun.TabIndex = 4;
            this.btnRun.Text = "執行";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.AutoSize = true;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.Location = new System.Drawing.Point(210, 251);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 25);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "離開";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // txtDesc
            // 
            // 
            // 
            // 
            this.txtDesc.Border.Class = "TextBoxBorder";
            this.txtDesc.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtDesc.Location = new System.Drawing.Point(13, 57);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtDesc.Size = new System.Drawing.Size(272, 179);
            this.txtDesc.TabIndex = 6;
            // 
            // frmCheckGPlanCourseCode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(301, 284);
            this.Controls.Add(this.txtDesc);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.cbxSemester);
            this.Controls.Add(this.cbxSchoolYear);
            this.DoubleBuffered = true;
            this.Name = "frmCheckGPlanCourseCode";
            this.Text = "開課檢核課程代碼";
            this.Load += new System.EventHandler(this.frmCheckGPlanCourseCode_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxSchoolYear;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxSemester;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.ButtonX btnRun;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private DevComponents.DotNetBar.Controls.TextBoxX txtDesc;
    }
}