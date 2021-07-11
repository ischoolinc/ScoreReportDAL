namespace SHCourseGroupCodeAdmin.UIForm
{
    partial class frmCreateClassGPlanHasData
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
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.lblGroupName = new DevComponents.DotNetBar.LabelX();
            this.dgData = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.btnCreate = new DevComponents.DotNetBar.ButtonX();
            this.colGPName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colUsedClass = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDiffSujeCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colUpdateSubjCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSet = new System.Windows.Forms.DataGridViewButtonColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).BeginInit();
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
            this.labelX1.Location = new System.Drawing.Point(13, 13);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(60, 21);
            this.labelX1.TabIndex = 0;
            this.labelX1.Text = "群科班：";
            // 
            // lblGroupName
            // 
            this.lblGroupName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblGroupName.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblGroupName.BackgroundStyle.Class = "";
            this.lblGroupName.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblGroupName.Location = new System.Drawing.Point(81, 12);
            this.lblGroupName.Name = "lblGroupName";
            this.lblGroupName.Size = new System.Drawing.Size(755, 23);
            this.lblGroupName.TabIndex = 1;
            // 
            // dgData
            // 
            this.dgData.AllowUserToAddRows = false;
            this.dgData.AllowUserToDeleteRows = false;
            this.dgData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgData.BackgroundColor = System.Drawing.Color.White;
            this.dgData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colGPName,
            this.colUsedClass,
            this.colDiffSujeCount,
            this.colUpdateSubjCount,
            this.colSet});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgData.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgData.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgData.Location = new System.Drawing.Point(13, 43);
            this.dgData.Name = "dgData";
            this.dgData.RowTemplate.Height = 24;
            this.dgData.Size = new System.Drawing.Size(823, 150);
            this.dgData.TabIndex = 2;
            this.dgData.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgData_CellContentClick);
            // 
            // btnCreate
            // 
            this.btnCreate.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCreate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCreate.BackColor = System.Drawing.Color.Transparent;
            this.btnCreate.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCreate.Location = new System.Drawing.Point(761, 210);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(75, 23);
            this.btnCreate.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCreate.TabIndex = 3;
            this.btnCreate.Text = "產生";
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // colGPName
            // 
            this.colGPName.HeaderText = "課程規劃";
            this.colGPName.Name = "colGPName";
            this.colGPName.ReadOnly = true;
            this.colGPName.Width = 150;
            // 
            // colUsedClass
            // 
            this.colUsedClass.HeaderText = "採用班級";
            this.colUsedClass.Name = "colUsedClass";
            this.colUsedClass.ReadOnly = true;
            this.colUsedClass.Width = 200;
            // 
            // colDiffSujeCount
            // 
            this.colDiffSujeCount.HeaderText = "差異科目數";
            this.colDiffSujeCount.Name = "colDiffSujeCount";
            this.colDiffSujeCount.ReadOnly = true;
            // 
            // colUpdateSubjCount
            // 
            this.colUpdateSubjCount.HeaderText = "更新科目數";
            this.colUpdateSubjCount.Name = "colUpdateSubjCount";
            this.colUpdateSubjCount.ReadOnly = true;
            // 
            // colSet
            // 
            this.colSet.HeaderText = "     ";
            this.colSet.Name = "colSet";
            this.colSet.ReadOnly = true;
            // 
            // frmCreateClassGPlanHasData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(848, 255);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.dgData);
            this.Controls.Add(this.lblGroupName);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.Name = "frmCreateClassGPlanHasData";
            this.Text = "關聯課程規劃";
            this.Load += new System.EventHandler(this.frmCreateClassGPlanHasData_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX lblGroupName;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgData;
        private DevComponents.DotNetBar.ButtonX btnCreate;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGPName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colUsedClass;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDiffSujeCount;
        private System.Windows.Forms.DataGridViewTextBoxColumn colUpdateSubjCount;
        private System.Windows.Forms.DataGridViewButtonColumn colSet;
    }
}