
namespace SHCourseGroupCodeAdmin.UIForm
{
    partial class frmCopyCourseGroupSetting
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
            this.cboGraduationPlanName = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.btnCopy = new DevComponents.DotNetBar.ButtonX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
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
            this.labelX1.Location = new System.Drawing.Point(12, 21);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(114, 21);
            this.labelX1.TabIndex = 0;
            this.labelX1.Text = "課程規畫表名稱：";
            // 
            // cboGraduationPlanName
            // 
            this.cboGraduationPlanName.DisplayMember = "Text";
            this.cboGraduationPlanName.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboGraduationPlanName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboGraduationPlanName.FormattingEnabled = true;
            this.cboGraduationPlanName.ItemHeight = 19;
            this.cboGraduationPlanName.Location = new System.Drawing.Point(124, 19);
            this.cboGraduationPlanName.Name = "cboGraduationPlanName";
            this.cboGraduationPlanName.Size = new System.Drawing.Size(298, 25);
            this.cboGraduationPlanName.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboGraduationPlanName.TabIndex = 1;
            // 
            // btnCopy
            // 
            this.btnCopy.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCopy.BackColor = System.Drawing.Color.Transparent;
            this.btnCopy.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCopy.Location = new System.Drawing.Point(266, 58);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(75, 23);
            this.btnCopy.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCopy.TabIndex = 2;
            this.btnCopy.Text = "複製";
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.Location = new System.Drawing.Point(347, 58);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "取消";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // frmCopyCourseGroupSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 91);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.cboGraduationPlanName);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.Name = "frmCopyCourseGroupSetting";
            this.Text = "複製課程群組設定";
            this.Load += new System.EventHandler(this.frmCopyCourseGroupSetting_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboGraduationPlanName;
        private DevComponents.DotNetBar.ButtonX btnCopy;
        private DevComponents.DotNetBar.ButtonX btnCancel;
    }
}