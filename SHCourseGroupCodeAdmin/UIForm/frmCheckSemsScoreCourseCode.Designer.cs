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
            this.iptGradeYear = new DevComponents.Editors.IntegerInput();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.btnRun = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.iptGradeYear)).BeginInit();
            this.SuspendLayout();
            // 
            // iptGradeYear
            // 
            this.iptGradeYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptGradeYear.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptGradeYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptGradeYear.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptGradeYear.Location = new System.Drawing.Point(108, 17);
            this.iptGradeYear.MaxValue = 6;
            this.iptGradeYear.MinValue = 1;
            this.iptGradeYear.Name = "iptGradeYear";
            this.iptGradeYear.ShowUpDown = true;
            this.iptGradeYear.Size = new System.Drawing.Size(51, 25);
            this.iptGradeYear.TabIndex = 16;
            this.iptGradeYear.Value = 3;
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
            this.labelX1.Location = new System.Drawing.Point(66, 19);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(34, 21);
            this.labelX1.TabIndex = 17;
            this.labelX1.Text = "年級";
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.AutoSize = true;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.Location = new System.Drawing.Point(127, 74);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 25);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 15;
            this.btnCancel.Text = "離開";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnRun
            // 
            this.btnRun.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnRun.AutoSize = true;
            this.btnRun.BackColor = System.Drawing.Color.Transparent;
            this.btnRun.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnRun.Location = new System.Drawing.Point(33, 74);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 25);
            this.btnRun.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnRun.TabIndex = 14;
            this.btnRun.Text = "執行";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // frmCheckSemsScoreCourseCode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(235, 117);
            this.Controls.Add(this.iptGradeYear);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRun);
            this.DoubleBuffered = true;
            this.Name = "frmCheckSemsScoreCourseCode";
            this.Text = "學期成績檢核課程代碼";
            this.Load += new System.EventHandler(this.frmCheckSemsScoreCourseCode_Load);
            ((System.ComponentModel.ISupportInitialize)(this.iptGradeYear)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.Editors.IntegerInput iptGradeYear;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private DevComponents.DotNetBar.ButtonX btnRun;
    }
}