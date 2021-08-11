namespace SHCourseGroupCodeAdmin.UIForm
{
    partial class frmCreateGPlanQueryAndSetup108
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgData = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.lblDiffCount = new DevComponents.DotNetBar.LabelX();
            this.lblAddCount = new DevComponents.DotNetBar.LabelX();
            this.lblUpdateCount = new DevComponents.DotNetBar.LabelX();
            this.lblDelCount = new DevComponents.DotNetBar.LabelX();
            this.lblNoChangeCount = new DevComponents.DotNetBar.LabelX();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).BeginInit();
            this.SuspendLayout();
            // 
            // dgData
            // 
            this.dgData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgData.BackgroundColor = System.Drawing.Color.White;
            this.dgData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgData.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgData.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgData.Location = new System.Drawing.Point(22, 26);
            this.dgData.Name = "dgData";
            this.dgData.RowTemplate.Height = 24;
            this.dgData.Size = new System.Drawing.Size(925, 416);
            this.dgData.TabIndex = 0;
            // 
            // lblDiffCount
            // 
            this.lblDiffCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblDiffCount.AutoSize = true;
            this.lblDiffCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblDiffCount.BackgroundStyle.Class = "";
            this.lblDiffCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblDiffCount.Location = new System.Drawing.Point(22, 458);
            this.lblDiffCount.Name = "lblDiffCount";
            this.lblDiffCount.Size = new System.Drawing.Size(82, 21);
            this.lblDiffCount.TabIndex = 1;
            this.lblDiffCount.Text = "差異筆數0筆";
            // 
            // lblAddCount
            // 
            this.lblAddCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblAddCount.AutoSize = true;
            this.lblAddCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblAddCount.BackgroundStyle.Class = "";
            this.lblAddCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblAddCount.Location = new System.Drawing.Point(144, 458);
            this.lblAddCount.Name = "lblAddCount";
            this.lblAddCount.Size = new System.Drawing.Size(55, 21);
            this.lblAddCount.TabIndex = 2;
            this.lblAddCount.Text = "新增0筆";
            // 
            // lblUpdateCount
            // 
            this.lblUpdateCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblUpdateCount.AutoSize = true;
            this.lblUpdateCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblUpdateCount.BackgroundStyle.Class = "";
            this.lblUpdateCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblUpdateCount.Location = new System.Drawing.Point(285, 458);
            this.lblUpdateCount.Name = "lblUpdateCount";
            this.lblUpdateCount.Size = new System.Drawing.Size(55, 21);
            this.lblUpdateCount.TabIndex = 3;
            this.lblUpdateCount.Text = "更新0筆";
            // 
            // lblDelCount
            // 
            this.lblDelCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblDelCount.AutoSize = true;
            this.lblDelCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblDelCount.BackgroundStyle.Class = "";
            this.lblDelCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblDelCount.Location = new System.Drawing.Point(415, 458);
            this.lblDelCount.Name = "lblDelCount";
            this.lblDelCount.Size = new System.Drawing.Size(55, 21);
            this.lblDelCount.TabIndex = 4;
            this.lblDelCount.Text = "刪除0筆";
            // 
            // lblNoChangeCount
            // 
            this.lblNoChangeCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblNoChangeCount.AutoSize = true;
            this.lblNoChangeCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblNoChangeCount.BackgroundStyle.Class = "";
            this.lblNoChangeCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblNoChangeCount.Location = new System.Drawing.Point(561, 458);
            this.lblNoChangeCount.Name = "lblNoChangeCount";
            this.lblNoChangeCount.Size = new System.Drawing.Size(55, 21);
            this.lblNoChangeCount.TabIndex = 5;
            this.lblNoChangeCount.Text = "略過0筆";
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.AutoSize = true;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(872, 458);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 25);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 6;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // frmCreateGPlanQueryAndSetup108
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(967, 502);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblNoChangeCount);
            this.Controls.Add(this.lblDelCount);
            this.Controls.Add(this.lblUpdateCount);
            this.Controls.Add(this.lblAddCount);
            this.Controls.Add(this.lblDiffCount);
            this.Controls.Add(this.dgData);
            this.DoubleBuffered = true;
            this.MaximizeBox = true;
            this.Name = "frmCreateGPlanQueryAndSetup108";
            this.Text = "查詢及設定";
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dgData;
        private DevComponents.DotNetBar.LabelX lblDiffCount;
        private DevComponents.DotNetBar.LabelX lblAddCount;
        private DevComponents.DotNetBar.LabelX lblUpdateCount;
        private DevComponents.DotNetBar.LabelX lblDelCount;
        private DevComponents.DotNetBar.LabelX lblNoChangeCount;
        private DevComponents.DotNetBar.ButtonX btnSave;
    }
}