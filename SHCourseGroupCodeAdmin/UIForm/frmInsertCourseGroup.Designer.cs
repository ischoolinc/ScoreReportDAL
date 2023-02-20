
namespace SHCourseGroupCodeAdmin.UIForm
{
    partial class frmInsertCourseGroup
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmInsertCourseGroup));
            this.tbCredit = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.btnInsert = new DevComponents.DotNetBar.ButtonX();
            this.tbCourseGroupName = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.cpColorPicker = new DevComponents.DotNetBar.ColorPickerButton();
            this.SuspendLayout();
            // 
            // tbCredit
            // 
            // 
            // 
            // 
            this.tbCredit.Border.Class = "TextBoxBorder";
            this.tbCredit.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbCredit.Location = new System.Drawing.Point(132, 52);
            this.tbCredit.Name = "tbCredit";
            this.tbCredit.Size = new System.Drawing.Size(171, 25);
            this.tbCredit.TabIndex = 11;
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
            this.labelX2.Location = new System.Drawing.Point(16, 54);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(114, 21);
            this.labelX2.TabIndex = 10;
            this.labelX2.Text = "群組修課學分數：";
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.Location = new System.Drawing.Point(254, 93);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(92, 23);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 13;
            this.btnCancel.Text = "取消";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnInsert
            // 
            this.btnInsert.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnInsert.BackColor = System.Drawing.Color.Transparent;
            this.btnInsert.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnInsert.Location = new System.Drawing.Point(132, 93);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(92, 23);
            this.btnInsert.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnInsert.TabIndex = 12;
            this.btnInsert.Text = "新增";
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // tbCourseGroupName
            // 
            // 
            // 
            // 
            this.tbCourseGroupName.Border.Class = "TextBoxBorder";
            this.tbCourseGroupName.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbCourseGroupName.Location = new System.Drawing.Point(132, 20);
            this.tbCourseGroupName.Name = "tbCourseGroupName";
            this.tbCourseGroupName.Size = new System.Drawing.Size(171, 25);
            this.tbCourseGroupName.TabIndex = 8;
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
            this.labelX1.Location = new System.Drawing.Point(29, 22);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(101, 21);
            this.labelX1.TabIndex = 7;
            this.labelX1.Text = "課程群組名稱：";
            // 
            // cpColorPicker
            // 
            this.cpColorPicker.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.cpColorPicker.BackColor = System.Drawing.Color.Transparent;
            this.cpColorPicker.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.cpColorPicker.Image = ((System.Drawing.Image)(resources.GetObject("cpColorPicker.Image")));
            this.cpColorPicker.Location = new System.Drawing.Point(309, 21);
            this.cpColorPicker.Name = "cpColorPicker";
            this.cpColorPicker.SelectedColorImageRectangle = new System.Drawing.Rectangle(2, 2, 12, 12);
            this.cpColorPicker.Size = new System.Drawing.Size(37, 23);
            this.cpColorPicker.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cpColorPicker.TabIndex = 9;
            // 
            // frmInsertCourseGroup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(372, 133);
            this.Controls.Add(this.tbCredit);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.cpColorPicker);
            this.Controls.Add(this.tbCourseGroupName);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.Name = "frmInsertCourseGroup";
            this.Text = "新增課程群組";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.TextBoxX tbCredit;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private DevComponents.DotNetBar.ButtonX btnInsert;
        private DevComponents.DotNetBar.ColorPickerButton cpColorPicker;
        private DevComponents.DotNetBar.Controls.TextBoxX tbCourseGroupName;
        private DevComponents.DotNetBar.LabelX labelX1;
    }
}