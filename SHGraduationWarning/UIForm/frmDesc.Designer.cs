namespace SHGraduationWarning.UIForm
{
    partial class frmDesc
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
            this.buttonClose = new DevComponents.DotNetBar.ButtonX();
            this.textDesc = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.SuspendLayout();
            // 
            // buttonClose
            // 
            this.buttonClose.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose.BackColor = System.Drawing.Color.Transparent;
            this.buttonClose.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonClose.Location = new System.Drawing.Point(599, 361);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(75, 23);
            this.buttonClose.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonClose.TabIndex = 0;
            this.buttonClose.Text = "關閉";
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // textDesc
            // 
            this.textDesc.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            // 
            // 
            // 
            this.textDesc.Border.Class = "TextBoxBorder";
            this.textDesc.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.textDesc.Location = new System.Drawing.Point(13, 13);
            this.textDesc.Multiline = true;
            this.textDesc.Name = "textDesc";
            this.textDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textDesc.Size = new System.Drawing.Size(661, 337);
            this.textDesc.TabIndex = 1;
            // 
            // frmDesc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(686, 391);
            this.Controls.Add(this.textDesc);
            this.Controls.Add(this.buttonClose);
            this.DoubleBuffered = true;
            this.Name = "frmDesc";
            this.Text = "說明";
            this.Load += new System.EventHandler(this.frmDesc_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX buttonClose;
        private DevComponents.DotNetBar.Controls.TextBoxX textDesc;
    }
}