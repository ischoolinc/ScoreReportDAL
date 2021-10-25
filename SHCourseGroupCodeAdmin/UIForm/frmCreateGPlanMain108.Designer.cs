namespace SHCourseGroupCodeAdmin.UIForm
{
    partial class frmCreateGPlanMain108
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblTitle = new DevComponents.DotNetBar.LabelX();
            this.lblGroupCount = new DevComponents.DotNetBar.LabelX();
            this.lblAddCount = new DevComponents.DotNetBar.LabelX();
            this.lblUpdateCount = new DevComponents.DotNetBar.LabelX();
            this.lblNoChangeCount = new DevComponents.DotNetBar.LabelX();
            this.dgData = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colEntrySchoolYear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGroupName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colChangeDesc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colUpdateSetup = new System.Windows.Forms.DataGridViewButtonColumn();
            this.btnQueryAndSet = new DevComponents.DotNetBar.ButtonX();
            this.btnCreate = new DevComponents.DotNetBar.ButtonX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).BeginInit();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblTitle.BackgroundStyle.Class = "";
            this.lblTitle.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblTitle.Location = new System.Drawing.Point(27, 12);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(748, 23);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "與「課程代碼總表」比對，變動課程規劃表如下。請選取「查詢及設定」確認更新項目。";
            // 
            // lblGroupCount
            // 
            this.lblGroupCount.AutoSize = true;
            this.lblGroupCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblGroupCount.BackgroundStyle.Class = "";
            this.lblGroupCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblGroupCount.Location = new System.Drawing.Point(27, 46);
            this.lblGroupCount.Name = "lblGroupCount";
            this.lblGroupCount.Size = new System.Drawing.Size(82, 21);
            this.lblGroupCount.TabIndex = 1;
            this.lblGroupCount.Text = "群科班數0筆";
            // 
            // lblAddCount
            // 
            this.lblAddCount.AutoSize = true;
            this.lblAddCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblAddCount.BackgroundStyle.Class = "";
            this.lblAddCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblAddCount.Location = new System.Drawing.Point(217, 46);
            this.lblAddCount.Name = "lblAddCount";
            this.lblAddCount.Size = new System.Drawing.Size(55, 21);
            this.lblAddCount.TabIndex = 2;
            this.lblAddCount.Text = "新增0筆";
            // 
            // lblUpdateCount
            // 
            this.lblUpdateCount.AutoSize = true;
            this.lblUpdateCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblUpdateCount.BackgroundStyle.Class = "";
            this.lblUpdateCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblUpdateCount.Location = new System.Drawing.Point(402, 46);
            this.lblUpdateCount.Name = "lblUpdateCount";
            this.lblUpdateCount.Size = new System.Drawing.Size(55, 21);
            this.lblUpdateCount.TabIndex = 3;
            this.lblUpdateCount.Text = "更新0筆";
            // 
            // lblNoChangeCount
            // 
            this.lblNoChangeCount.AutoSize = true;
            this.lblNoChangeCount.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblNoChangeCount.BackgroundStyle.Class = "";
            this.lblNoChangeCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblNoChangeCount.Location = new System.Drawing.Point(581, 46);
            this.lblNoChangeCount.Name = "lblNoChangeCount";
            this.lblNoChangeCount.Size = new System.Drawing.Size(68, 21);
            this.lblNoChangeCount.TabIndex = 4;
            this.lblNoChangeCount.Text = "無變動0筆";
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
            this.dgData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colEntrySchoolYear,
            this.colGroupName,
            this.colChangeDesc,
            this.colUpdateSetup});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgData.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgData.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgData.Location = new System.Drawing.Point(28, 73);
            this.dgData.Name = "dgData";
            this.dgData.RowTemplate.Height = 24;
            this.dgData.Size = new System.Drawing.Size(832, 450);
            this.dgData.TabIndex = 5;
            this.dgData.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgData_CellContentClick);
            // 
            // colEntrySchoolYear
            // 
            this.colEntrySchoolYear.FillWeight = 80F;
            this.colEntrySchoolYear.HeaderText = "入學年";
            this.colEntrySchoolYear.Name = "colEntrySchoolYear";
            this.colEntrySchoolYear.ReadOnly = true;
            this.colEntrySchoolYear.Width = 80;
            // 
            // colGroupName
            // 
            this.colGroupName.HeaderText = "群科班名稱";
            this.colGroupName.Name = "colGroupName";
            this.colGroupName.ReadOnly = true;
            this.colGroupName.Width = 300;
            // 
            // colChangeDesc
            // 
            this.colChangeDesc.HeaderText = "變更說明";
            this.colChangeDesc.Name = "colChangeDesc";
            this.colChangeDesc.ReadOnly = true;
            this.colChangeDesc.Width = 150;
            // 
            // colUpdateSetup
            // 
            this.colUpdateSetup.HeaderText = "更新設定";
            this.colUpdateSetup.Name = "colUpdateSetup";
            this.colUpdateSetup.ReadOnly = true;
            // 
            // btnQueryAndSet
            // 
            this.btnQueryAndSet.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnQueryAndSet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnQueryAndSet.AutoSize = true;
            this.btnQueryAndSet.BackColor = System.Drawing.Color.Transparent;
            this.btnQueryAndSet.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnQueryAndSet.Location = new System.Drawing.Point(28, 543);
            this.btnQueryAndSet.Name = "btnQueryAndSet";
            this.btnQueryAndSet.Size = new System.Drawing.Size(78, 25);
            this.btnQueryAndSet.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnQueryAndSet.TabIndex = 6;
            this.btnQueryAndSet.Text = "查詢及設定";
            this.btnQueryAndSet.Click += new System.EventHandler(this.btnQueryAndSet_Click);
            // 
            // btnCreate
            // 
            this.btnCreate.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCreate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCreate.AutoSize = true;
            this.btnCreate.BackColor = System.Drawing.Color.Transparent;
            this.btnCreate.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCreate.Location = new System.Drawing.Point(677, 543);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(75, 25);
            this.btnCreate.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCreate.TabIndex = 7;
            this.btnCreate.Text = "產生";
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.AutoSize = true;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.Location = new System.Drawing.Point(781, 543);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 25);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "取消";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // frmCreateGPlanMain108
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(881, 580);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.btnQueryAndSet);
            this.Controls.Add(this.dgData);
            this.Controls.Add(this.lblNoChangeCount);
            this.Controls.Add(this.lblUpdateCount);
            this.Controls.Add(this.lblAddCount);
            this.Controls.Add(this.lblGroupCount);
            this.Controls.Add(this.lblTitle);
            this.DoubleBuffered = true;
            this.Name = "frmCreateGPlanMain108";
            this.Text = "依課程代碼總表產生課程規劃";
            this.Load += new System.EventHandler(this.frmCreateGPlanMain108_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lblTitle;
        private DevComponents.DotNetBar.LabelX lblGroupCount;
        private DevComponents.DotNetBar.LabelX lblAddCount;
        private DevComponents.DotNetBar.LabelX lblUpdateCount;
        private DevComponents.DotNetBar.LabelX lblNoChangeCount;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgData;
        private DevComponents.DotNetBar.ButtonX btnQueryAndSet;
        private DevComponents.DotNetBar.ButtonX btnCreate;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private System.Windows.Forms.DataGridViewTextBoxColumn colEntrySchoolYear;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGroupName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colChangeDesc;
        private System.Windows.Forms.DataGridViewButtonColumn colUpdateSetup;
    }
}