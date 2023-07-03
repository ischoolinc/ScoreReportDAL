namespace SHCourseGroupCodeAdmin.UIForm
{
    partial class frmCreateGPlanItemSetup108
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblGroupName = new DevComponents.DotNetBar.LabelX();
            this.lblDiffCount = new DevComponents.DotNetBar.LabelX();
            this.lblAddCount = new DevComponents.DotNetBar.LabelX();
            this.lblUpdateCount = new DevComponents.DotNetBar.LabelX();
            this.lblDelCount = new DevComponents.DotNetBar.LabelX();
            this.lblNoChangeCount = new DevComponents.DotNetBar.LabelX();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.dgData = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.menuUpdateStatus = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.btnBatchModify = new DevComponents.DotNetBar.ButtonX();
            this.lblUpdateLevelCount = new DevComponents.DotNetBar.LabelX();
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).BeginInit();
            this.SuspendLayout();
            // 
            // lblGroupName
            // 
            this.lblGroupName.AutoSize = true;
            this.lblGroupName.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblGroupName.BackgroundStyle.Class = "";
            this.lblGroupName.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblGroupName.Location = new System.Drawing.Point(22, 21);
            this.lblGroupName.Name = "lblGroupName";
            this.lblGroupName.Size = new System.Drawing.Size(60, 21);
            this.lblGroupName.TabIndex = 0;
            this.lblGroupName.Text = "群科班：";
            // 
            // lblDiffCount
            // 
            this.lblDiffCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblDiffCount.AutoSize = true;
            this.lblDiffCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblDiffCount.BackgroundStyle.Class = "";
            this.lblDiffCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblDiffCount.Location = new System.Drawing.Point(613, 21);
            this.lblDiffCount.Name = "lblDiffCount";
            this.lblDiffCount.Size = new System.Drawing.Size(87, 21);
            this.lblDiffCount.TabIndex = 1;
            this.lblDiffCount.Text = "差異科目數：";
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
            this.lblAddCount.Location = new System.Drawing.Point(27, 363);
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
            this.lblUpdateCount.Location = new System.Drawing.Point(123, 363);
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
            this.lblDelCount.Location = new System.Drawing.Point(342, 363);
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
            this.lblNoChangeCount.Location = new System.Drawing.Point(438, 363);
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
            this.btnSave.Location = new System.Drawing.Point(683, 361);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 25);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 6;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // dgData
            // 
            this.dgData.AllowUserToAddRows = false;
            this.dgData.AllowUserToDeleteRows = false;
            this.dgData.AllowUserToResizeRows = false;
            this.dgData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgData.BackgroundColor = System.Drawing.Color.White;
            this.dgData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgData.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgData.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgData.Location = new System.Drawing.Point(22, 55);
            this.dgData.Name = "dgData";
            this.dgData.RowTemplate.Height = 24;
            this.dgData.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgData.Size = new System.Drawing.Size(736, 290);
            this.dgData.TabIndex = 7;
            this.dgData.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgData_CellContentClick);
            this.dgData.CellMouseEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgData_CellMouseEnter);
            this.dgData.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgData_CellValueChanged);
            // 
            // menuUpdateStatus
            // 
            this.menuUpdateStatus.Name = "menuUpdateStatus";
            this.menuUpdateStatus.Size = new System.Drawing.Size(61, 4);
            this.menuUpdateStatus.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.menuUpdateStatus_ItemClicked);
            // 
            // btnBatchModify
            // 
            this.btnBatchModify.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnBatchModify.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBatchModify.BackColor = System.Drawing.Color.Transparent;
            this.btnBatchModify.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnBatchModify.Location = new System.Drawing.Point(549, 361);
            this.btnBatchModify.Name = "btnBatchModify";
            this.btnBatchModify.Size = new System.Drawing.Size(122, 25);
            this.btnBatchModify.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnBatchModify.TabIndex = 6;
            this.btnBatchModify.Text = "批次修改處理方式";
            this.btnBatchModify.Click += new System.EventHandler(this.btnBatchModify_Click);
            // 
            // lblUpdateLevelCount
            // 
            this.lblUpdateLevelCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblUpdateLevelCount.AutoSize = true;
            this.lblUpdateLevelCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblUpdateLevelCount.BackgroundStyle.Class = "";
            this.lblUpdateLevelCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblUpdateLevelCount.Location = new System.Drawing.Point(219, 363);
            this.lblUpdateLevelCount.Name = "lblUpdateLevelCount";
            this.lblUpdateLevelCount.Size = new System.Drawing.Size(82, 21);
            this.lblUpdateLevelCount.TabIndex = 4;
            this.lblUpdateLevelCount.Text = "級別更新0筆";
            // 
            // frmCreateGPlanItemSetup108
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(790, 398);
            this.Controls.Add(this.lblUpdateLevelCount);
            this.Controls.Add(this.btnBatchModify);
            this.Controls.Add(this.dgData);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblNoChangeCount);
            this.Controls.Add(this.lblDelCount);
            this.Controls.Add(this.lblUpdateCount);
            this.Controls.Add(this.lblAddCount);
            this.Controls.Add(this.lblDiffCount);
            this.Controls.Add(this.lblGroupName);
            this.DoubleBuffered = true;
            this.MaximizeBox = true;
            this.Name = "frmCreateGPlanItemSetup108";
            this.Text = "設定";
            this.Load += new System.EventHandler(this.frmCreateGPlanItemSetup108_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lblGroupName;
        private DevComponents.DotNetBar.LabelX lblDiffCount;
        private DevComponents.DotNetBar.LabelX lblAddCount;
        private DevComponents.DotNetBar.LabelX lblUpdateCount;
        private DevComponents.DotNetBar.LabelX lblDelCount;
        private DevComponents.DotNetBar.LabelX lblNoChangeCount;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgData;
        private System.Windows.Forms.ContextMenuStrip menuUpdateStatus;
        private DevComponents.DotNetBar.ButtonX btnBatchModify;
        private DevComponents.DotNetBar.LabelX lblUpdateLevelCount;
    }
}