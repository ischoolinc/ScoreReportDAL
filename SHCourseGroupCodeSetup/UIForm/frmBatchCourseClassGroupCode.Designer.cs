﻿namespace SHCourseGroupCodeSetup.UIForm
{
    partial class frmBatchCourseClassGroupCode
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
            this.cbxCourseGroupCode = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.btnRun = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.SuspendLayout();
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelX1.Location = new System.Drawing.Point(24, 20);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(51, 23);
            this.labelX1.TabIndex = 0;
            this.labelX1.Text = "群科班";
            // 
            // cbxCourseGroupCode
            // 
            this.cbxCourseGroupCode.DisplayMember = "Text";
            this.cbxCourseGroupCode.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxCourseGroupCode.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbxCourseGroupCode.FormattingEnabled = true;
            this.cbxCourseGroupCode.ItemHeight = 19;
            this.cbxCourseGroupCode.Location = new System.Drawing.Point(96, 20);
            this.cbxCourseGroupCode.Name = "cbxCourseGroupCode";
            this.cbxCourseGroupCode.Size = new System.Drawing.Size(416, 25);
            this.cbxCourseGroupCode.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxCourseGroupCode.TabIndex = 1;
            // 
            // btnRun
            // 
            this.btnRun.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnRun.AutoSize = true;
            this.btnRun.BackColor = System.Drawing.Color.Transparent;
            this.btnRun.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnRun.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnRun.Location = new System.Drawing.Point(342, 75);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 28);
            this.btnRun.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnRun.TabIndex = 2;
            this.btnRun.Text = "產生";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnExit.Location = new System.Drawing.Point(437, 75);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 28);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 3;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // frmBatchCourseClassGroupCode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(537, 119);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.cbxCourseGroupCode);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.Name = "frmBatchCourseClassGroupCode";
            this.Text = "批次產生群科班";
            this.Load += new System.EventHandler(this.frmBatchCourseClassGroupCode_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxCourseGroupCode;
        private DevComponents.DotNetBar.ButtonX btnRun;
        private DevComponents.DotNetBar.ButtonX btnExit;
    }
}