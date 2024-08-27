namespace SHGraduationWarning.UIForm
{
    partial class frmNoGraduationPlanStudent
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
            this.btnAddStudentTemp = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.lblMsg = new DevComponents.DotNetBar.LabelX();
            this.SuspendLayout();
            // 
            // btnAddStudentTemp
            // 
            this.btnAddStudentTemp.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAddStudentTemp.BackColor = System.Drawing.Color.Transparent;
            this.btnAddStudentTemp.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAddStudentTemp.Location = new System.Drawing.Point(12, 91);
            this.btnAddStudentTemp.Name = "btnAddStudentTemp";
            this.btnAddStudentTemp.Size = new System.Drawing.Size(118, 23);
            this.btnAddStudentTemp.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnAddStudentTemp.TabIndex = 0;
            this.btnAddStudentTemp.Text = "加入學生待處理";
            this.btnAddStudentTemp.Click += new System.EventHandler(this.btnAddStudentTemp_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(296, 91);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 1;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // lblMsg
            // 
            this.lblMsg.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblMsg.BackgroundStyle.Class = "";
            this.lblMsg.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblMsg.Location = new System.Drawing.Point(12, 12);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(366, 62);
            this.lblMsg.TabIndex = 2;
            this.lblMsg.Text = "有100位學生沒有設定課程規劃，請設定後再進行查詢作業。";
            this.lblMsg.WordWrap = true;
            // 
            // frmNoGraduationPlanStudent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(390, 128);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnAddStudentTemp);
            this.DoubleBuffered = true;
            this.Name = "frmNoGraduationPlanStudent";
            this.Text = "沒有設定課程規劃學生";
            this.Load += new System.EventHandler(this.frmNoGraduationPlanStudent_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnAddStudentTemp;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.LabelX lblMsg;
    }
}