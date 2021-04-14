namespace SHCourseGroupCodeSetup.DetailContent
{
    partial class UCClassGroupCodeItem
    {
        /// <summary> 
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.cbxCourseGroupCode = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.SuspendLayout();
            // 
            // labelX1
            // 
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(36, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(52, 23);
            this.labelX1.TabIndex = 0;
            this.labelX1.Text = "群 科 班";
            // 
            // cbxCourseGroupCode
            // 
            this.cbxCourseGroupCode.DisplayMember = "Text";
            this.cbxCourseGroupCode.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxCourseGroupCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxCourseGroupCode.FormattingEnabled = true;
            this.cbxCourseGroupCode.ItemHeight = 19;
            this.cbxCourseGroupCode.Location = new System.Drawing.Point(109, 11);
            this.cbxCourseGroupCode.Name = "cbxCourseGroupCode";
            this.cbxCourseGroupCode.Size = new System.Drawing.Size(404, 25);
            this.cbxCourseGroupCode.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxCourseGroupCode.TabIndex = 1;
            this.cbxCourseGroupCode.TextChanged += new System.EventHandler(this.cbxCourseGroupCode_TextChanged);
            this.cbxCourseGroupCode.Validated += new System.EventHandler(this.cbxCourseGroupCode_Validated);
            // 
            // UCClassGroupCodeItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbxCourseGroupCode);
            this.Controls.Add(this.labelX1);
            this.Name = "UCClassGroupCodeItem";
            this.Size = new System.Drawing.Size(550, 50);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxCourseGroupCode;
    }
}
