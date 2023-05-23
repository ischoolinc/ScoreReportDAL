namespace SHGraduationWarning.UIForm
{
    partial class frmMain
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblMsg = new DevComponents.DotNetBar.LabelX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnQuery = new DevComponents.DotNetBar.ButtonX();
            this.dgDataGW = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.comboClass = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.comboGradeYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX5 = new DevComponents.DotNetBar.LabelX();
            this.comboDept = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.tabControl1 = new DevComponents.DotNetBar.TabControl();
            this.tabControlPanel1 = new DevComponents.DotNetBar.TabControlPanel();
            this.tbItemGW = new DevComponents.DotNetBar.TabItem(this.components);
            this.tabControlPanel2 = new DevComponents.DotNetBar.TabControlPanel();
            this.dgDataChkEdit = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.tbItemChkEdit = new DevComponents.DotNetBar.TabItem(this.components);
            this.tabControlPanel3 = new DevComponents.DotNetBar.TabControlPanel();
            this.dgData2ChkEdit = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.tbItem2ChkEdit = new DevComponents.DotNetBar.TabItem(this.components);
            this.btnUpdate = new DevComponents.DotNetBar.ButtonX();
            this.btnDel = new DevComponents.DotNetBar.ButtonX();
            this.btnReport = new DevComponents.DotNetBar.ButtonX();
            this.buttonUpdateDSubjectName = new DevComponents.DotNetBar.ButtonX();
            this.buttonLoadDesc = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.dgDataGW)).BeginInit();
            this.groupPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tabControl1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabControlPanel1.SuspendLayout();
            this.tabControlPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgDataChkEdit)).BeginInit();
            this.tabControlPanel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgData2ChkEdit)).BeginInit();
            this.SuspendLayout();
            // 
            // lblMsg
            // 
            this.lblMsg.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblMsg.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblMsg.BackgroundStyle.Class = "";
            this.lblMsg.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblMsg.Location = new System.Drawing.Point(17, 752);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(413, 23);
            this.lblMsg.TabIndex = 58;
            this.lblMsg.Text = "共0筆。";
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(971, 752);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 55;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnQuery
            // 
            this.btnQuery.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnQuery.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnQuery.BackColor = System.Drawing.Color.Transparent;
            this.btnQuery.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnQuery.Location = new System.Drawing.Point(887, 752);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(75, 23);
            this.btnQuery.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnQuery.TabIndex = 56;
            this.btnQuery.Text = "檢查";
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // dgDataGW
            // 
            this.dgDataGW.AllowUserToAddRows = false;
            this.dgDataGW.AllowUserToDeleteRows = false;
            this.dgDataGW.BackgroundColor = System.Drawing.Color.White;
            this.dgDataGW.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgDataGW.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgDataGW.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgDataGW.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgDataGW.Location = new System.Drawing.Point(1, 1);
            this.dgDataGW.Name = "dgDataGW";
            this.dgDataGW.RowTemplate.Height = 24;
            this.dgDataGW.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgDataGW.Size = new System.Drawing.Size(1030, 575);
            this.dgDataGW.TabIndex = 57;
            // 
            // labelX2
            // 
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(320, 19);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(33, 23);
            this.labelX2.TabIndex = 48;
            this.labelX2.Text = "班級";
            // 
            // comboClass
            // 
            this.comboClass.DisplayMember = "Text";
            this.comboClass.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboClass.FormattingEnabled = true;
            this.comboClass.ItemHeight = 19;
            this.comboClass.Location = new System.Drawing.Point(359, 18);
            this.comboClass.Name = "comboClass";
            this.comboClass.Size = new System.Drawing.Size(121, 25);
            this.comboClass.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comboClass.TabIndex = 49;
            this.comboClass.SelectedIndexChanged += new System.EventHandler(this.comboClass_SelectedIndexChanged);
            // 
            // groupPanel1
            // 
            this.groupPanel1.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.labelX1);
            this.groupPanel1.Controls.Add(this.comboGradeYear);
            this.groupPanel1.Controls.Add(this.labelX5);
            this.groupPanel1.Controls.Add(this.comboDept);
            this.groupPanel1.Controls.Add(this.labelX2);
            this.groupPanel1.Controls.Add(this.comboClass);
            this.groupPanel1.Location = new System.Drawing.Point(17, 19);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(503, 87);
            // 
            // 
            // 
            this.groupPanel1.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel1.Style.BackColorGradientAngle = 90;
            this.groupPanel1.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel1.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderBottomWidth = 1;
            this.groupPanel1.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel1.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderLeftWidth = 1;
            this.groupPanel1.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderRightWidth = 1;
            this.groupPanel1.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderTopWidth = 1;
            this.groupPanel1.Style.Class = "";
            this.groupPanel1.Style.CornerDiameter = 4;
            this.groupPanel1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel1.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel1.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseDown.Class = "";
            this.groupPanel1.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseOver.Class = "";
            this.groupPanel1.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel1.TabIndex = 62;
            this.groupPanel1.Text = "請選擇學生";
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(13, 19);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(33, 23);
            this.labelX1.TabIndex = 54;
            this.labelX1.Text = "年級";
            // 
            // comboGradeYear
            // 
            this.comboGradeYear.DisplayMember = "Text";
            this.comboGradeYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboGradeYear.FormattingEnabled = true;
            this.comboGradeYear.ItemHeight = 19;
            this.comboGradeYear.Location = new System.Drawing.Point(52, 18);
            this.comboGradeYear.Name = "comboGradeYear";
            this.comboGradeYear.Size = new System.Drawing.Size(75, 25);
            this.comboGradeYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comboGradeYear.TabIndex = 55;
            this.comboGradeYear.SelectedIndexChanged += new System.EventHandler(this.comboGradeYear_SelectedIndexChanged);
            // 
            // labelX5
            // 
            this.labelX5.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX5.BackgroundStyle.Class = "";
            this.labelX5.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX5.Location = new System.Drawing.Point(137, 19);
            this.labelX5.Name = "labelX5";
            this.labelX5.Size = new System.Drawing.Size(33, 23);
            this.labelX5.TabIndex = 52;
            this.labelX5.Text = "科別";
            // 
            // comboDept
            // 
            this.comboDept.DisplayMember = "Text";
            this.comboDept.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboDept.FormattingEnabled = true;
            this.comboDept.ItemHeight = 19;
            this.comboDept.Location = new System.Drawing.Point(176, 18);
            this.comboDept.Name = "comboDept";
            this.comboDept.Size = new System.Drawing.Size(125, 25);
            this.comboDept.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comboDept.TabIndex = 53;
            this.comboDept.SelectedIndexChanged += new System.EventHandler(this.comboDept_SelectedIndexChanged);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.BackColor = System.Drawing.Color.Transparent;
            this.tabControl1.CanReorderTabs = true;
            this.tabControl1.Controls.Add(this.tabControlPanel1);
            this.tabControl1.Controls.Add(this.tabControlPanel2);
            this.tabControl1.Controls.Add(this.tabControlPanel3);
            this.tabControl1.Location = new System.Drawing.Point(17, 122);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedTabFont = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Bold);
            this.tabControl1.SelectedTabIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1032, 606);
            this.tabControl1.TabIndex = 63;
            this.tabControl1.TabLayoutType = DevComponents.DotNetBar.eTabLayoutType.FixedWithNavigationBox;
            this.tabControl1.Tabs.Add(this.tbItemChkEdit);
            this.tabControl1.Tabs.Add(this.tbItem2ChkEdit);
            this.tabControl1.Tabs.Add(this.tbItemGW);
            this.tabControl1.Text = "tabControl1";
            this.tabControl1.SelectedTabChanged += new DevComponents.DotNetBar.TabStrip.SelectedTabChangedEventHandler(this.tabControl1_SelectedTabChanged);
            this.tabControl1.Click += new System.EventHandler(this.tabControl1_Click);
            // 
            // tabControlPanel1
            // 
            this.tabControlPanel1.Controls.Add(this.dgDataGW);
            this.tabControlPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlPanel1.Location = new System.Drawing.Point(0, 29);
            this.tabControlPanel1.Name = "tabControlPanel1";
            this.tabControlPanel1.Padding = new System.Windows.Forms.Padding(1);
            this.tabControlPanel1.Size = new System.Drawing.Size(1032, 577);
            this.tabControlPanel1.Style.BackColor1.Color = System.Drawing.Color.FromArgb(((int)(((byte)(142)))), ((int)(((byte)(179)))), ((int)(((byte)(231)))));
            this.tabControlPanel1.Style.BackColor2.Color = System.Drawing.Color.FromArgb(((int)(((byte)(223)))), ((int)(((byte)(237)))), ((int)(((byte)(254)))));
            this.tabControlPanel1.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine;
            this.tabControlPanel1.Style.BorderColor.Color = System.Drawing.Color.FromArgb(((int)(((byte)(59)))), ((int)(((byte)(97)))), ((int)(((byte)(156)))));
            this.tabControlPanel1.Style.BorderSide = ((DevComponents.DotNetBar.eBorderSide)(((DevComponents.DotNetBar.eBorderSide.Left | DevComponents.DotNetBar.eBorderSide.Right) 
            | DevComponents.DotNetBar.eBorderSide.Bottom)));
            this.tabControlPanel1.Style.GradientAngle = 90;
            this.tabControlPanel1.TabIndex = 1;
            this.tabControlPanel1.TabItem = this.tbItemGW;
            // 
            // tbItemGW
            // 
            this.tbItemGW.AttachedControl = this.tabControlPanel1;
            this.tbItemGW.Name = "tbItemGW";
            this.tbItemGW.Text = "畢業預警";
            this.tbItemGW.Visible = false;
            this.tbItemGW.Click += new System.EventHandler(this.tbItemGW_Click);
            // 
            // tabControlPanel2
            // 
            this.tabControlPanel2.Controls.Add(this.dgDataChkEdit);
            this.tabControlPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlPanel2.Location = new System.Drawing.Point(0, 29);
            this.tabControlPanel2.Name = "tabControlPanel2";
            this.tabControlPanel2.Padding = new System.Windows.Forms.Padding(1);
            this.tabControlPanel2.Size = new System.Drawing.Size(1032, 577);
            this.tabControlPanel2.Style.BackColor1.Color = System.Drawing.Color.FromArgb(((int)(((byte)(142)))), ((int)(((byte)(179)))), ((int)(((byte)(231)))));
            this.tabControlPanel2.Style.BackColor2.Color = System.Drawing.Color.FromArgb(((int)(((byte)(223)))), ((int)(((byte)(237)))), ((int)(((byte)(254)))));
            this.tabControlPanel2.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine;
            this.tabControlPanel2.Style.BorderColor.Color = System.Drawing.Color.FromArgb(((int)(((byte)(59)))), ((int)(((byte)(97)))), ((int)(((byte)(156)))));
            this.tabControlPanel2.Style.BorderSide = ((DevComponents.DotNetBar.eBorderSide)(((DevComponents.DotNetBar.eBorderSide.Left | DevComponents.DotNetBar.eBorderSide.Right) 
            | DevComponents.DotNetBar.eBorderSide.Bottom)));
            this.tabControlPanel2.Style.GradientAngle = 90;
            this.tabControlPanel2.TabIndex = 2;
            this.tabControlPanel2.TabItem = this.tbItemChkEdit;
            // 
            // dgDataChkEdit
            // 
            this.dgDataChkEdit.AllowUserToAddRows = false;
            this.dgDataChkEdit.AllowUserToDeleteRows = false;
            this.dgDataChkEdit.BackgroundColor = System.Drawing.Color.White;
            this.dgDataChkEdit.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgDataChkEdit.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgDataChkEdit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgDataChkEdit.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgDataChkEdit.Location = new System.Drawing.Point(1, 1);
            this.dgDataChkEdit.Name = "dgDataChkEdit";
            this.dgDataChkEdit.RowTemplate.Height = 24;
            this.dgDataChkEdit.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgDataChkEdit.Size = new System.Drawing.Size(1030, 575);
            this.dgDataChkEdit.TabIndex = 58;
            // 
            // tbItemChkEdit
            // 
            this.tbItemChkEdit.AttachedControl = this.tabControlPanel2;
            this.tbItemChkEdit.Name = "tbItemChkEdit";
            this.tbItemChkEdit.Text = "資料合理檢查(科目級別)";
            this.tbItemChkEdit.Click += new System.EventHandler(this.tbItemChkEdit_Click);
            // 
            // tabControlPanel3
            // 
            this.tabControlPanel3.Controls.Add(this.dgData2ChkEdit);
            this.tabControlPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlPanel3.Location = new System.Drawing.Point(0, 29);
            this.tabControlPanel3.Name = "tabControlPanel3";
            this.tabControlPanel3.Padding = new System.Windows.Forms.Padding(1);
            this.tabControlPanel3.Size = new System.Drawing.Size(1032, 577);
            this.tabControlPanel3.Style.BackColor1.Color = System.Drawing.Color.FromArgb(((int)(((byte)(142)))), ((int)(((byte)(179)))), ((int)(((byte)(231)))));
            this.tabControlPanel3.Style.BackColor2.Color = System.Drawing.Color.FromArgb(((int)(((byte)(223)))), ((int)(((byte)(237)))), ((int)(((byte)(254)))));
            this.tabControlPanel3.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine;
            this.tabControlPanel3.Style.BorderColor.Color = System.Drawing.Color.FromArgb(((int)(((byte)(59)))), ((int)(((byte)(97)))), ((int)(((byte)(156)))));
            this.tabControlPanel3.Style.BorderSide = ((DevComponents.DotNetBar.eBorderSide)(((DevComponents.DotNetBar.eBorderSide.Left | DevComponents.DotNetBar.eBorderSide.Right) 
            | DevComponents.DotNetBar.eBorderSide.Bottom)));
            this.tabControlPanel3.Style.GradientAngle = 90;
            this.tabControlPanel3.TabIndex = 3;
            this.tabControlPanel3.TabItem = this.tbItem2ChkEdit;
            // 
            // dgData2ChkEdit
            // 
            this.dgData2ChkEdit.AllowUserToAddRows = false;
            this.dgData2ChkEdit.AllowUserToDeleteRows = false;
            this.dgData2ChkEdit.BackgroundColor = System.Drawing.Color.White;
            this.dgData2ChkEdit.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgData2ChkEdit.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgData2ChkEdit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgData2ChkEdit.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgData2ChkEdit.Location = new System.Drawing.Point(1, 1);
            this.dgData2ChkEdit.Name = "dgData2ChkEdit";
            this.dgData2ChkEdit.RowTemplate.Height = 24;
            this.dgData2ChkEdit.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgData2ChkEdit.Size = new System.Drawing.Size(1030, 575);
            this.dgData2ChkEdit.TabIndex = 59;
            // 
            // tbItem2ChkEdit
            // 
            this.tbItem2ChkEdit.AttachedControl = this.tabControlPanel3;
            this.tbItem2ChkEdit.Name = "tbItem2ChkEdit";
            this.tbItem2ChkEdit.Text = "資料合理檢查(科目屬性)";
            this.tbItem2ChkEdit.Click += new System.EventHandler(this.tbItem2ChkEdit_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnUpdate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnUpdate.BackColor = System.Drawing.Color.Transparent;
            this.btnUpdate.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnUpdate.Location = new System.Drawing.Point(850, 34);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnUpdate.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnUpdate.TabIndex = 64;
            this.btnUpdate.Text = "更新";
            this.btnUpdate.Visible = false;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnDel
            // 
            this.btnDel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnDel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDel.BackColor = System.Drawing.Color.Transparent;
            this.btnDel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnDel.Location = new System.Drawing.Point(932, 34);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(75, 23);
            this.btnDel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnDel.TabIndex = 65;
            this.btnDel.Text = "刪除";
            this.btnDel.Visible = false;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnReport
            // 
            this.btnReport.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnReport.BackColor = System.Drawing.Color.Transparent;
            this.btnReport.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnReport.Location = new System.Drawing.Point(783, 752);
            this.btnReport.Name = "btnReport";
            this.btnReport.Size = new System.Drawing.Size(95, 23);
            this.btnReport.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnReport.TabIndex = 68;
            this.btnReport.Text = "產生報表";
            this.btnReport.Click += new System.EventHandler(this.btnReport_Click);
            // 
            // buttonUpdateDSubjectName
            // 
            this.buttonUpdateDSubjectName.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonUpdateDSubjectName.BackColor = System.Drawing.Color.Transparent;
            this.buttonUpdateDSubjectName.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonUpdateDSubjectName.Location = new System.Drawing.Point(537, 83);
            this.buttonUpdateDSubjectName.Name = "buttonUpdateDSubjectName";
            this.buttonUpdateDSubjectName.Size = new System.Drawing.Size(131, 23);
            this.buttonUpdateDSubjectName.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonUpdateDSubjectName.TabIndex = 69;
            this.buttonUpdateDSubjectName.Text = "更新報部科目名稱";
            this.buttonUpdateDSubjectName.Click += new System.EventHandler(this.buttonUpdateDSubjectName_Click);
            // 
            // buttonLoadDesc
            // 
            this.buttonLoadDesc.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonLoadDesc.BackColor = System.Drawing.Color.Transparent;
            this.buttonLoadDesc.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonLoadDesc.Location = new System.Drawing.Point(537, 34);
            this.buttonLoadDesc.Name = "buttonLoadDesc";
            this.buttonLoadDesc.Size = new System.Drawing.Size(59, 23);
            this.buttonLoadDesc.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonLoadDesc.TabIndex = 70;
            this.buttonLoadDesc.Text = "說明";
            this.buttonLoadDesc.Click += new System.EventHandler(this.buttonLoadDesc_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1061, 787);
            this.Controls.Add(this.buttonLoadDesc);
            this.Controls.Add(this.buttonUpdateDSubjectName);
            this.Controls.Add(this.btnReport);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.groupPanel1);
            this.DoubleBuffered = true;
            this.MaximizeBox = true;
            this.Name = "frmMain";
            this.Text = "畢業預警與成績資料合理性檢查";
            this.Load += new System.EventHandler(this.frmMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgDataGW)).EndInit();
            this.groupPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.tabControl1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabControlPanel1.ResumeLayout(false);
            this.tabControlPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgDataChkEdit)).EndInit();
            this.tabControlPanel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgData2ChkEdit)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lblMsg;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.ButtonX btnQuery;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgDataGW;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comboClass;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
        private DevComponents.DotNetBar.LabelX labelX5;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comboDept;
        private DevComponents.DotNetBar.TabControl tabControl1;
        private DevComponents.DotNetBar.TabControlPanel tabControlPanel1;
        private DevComponents.DotNetBar.TabItem tbItemGW;
        private DevComponents.DotNetBar.TabControlPanel tabControlPanel2;
        private DevComponents.DotNetBar.TabItem tbItemChkEdit;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgDataChkEdit;
        private DevComponents.DotNetBar.ButtonX btnUpdate;
        private DevComponents.DotNetBar.ButtonX btnDel;
        private DevComponents.DotNetBar.TabControlPanel tabControlPanel3;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgData2ChkEdit;
        private DevComponents.DotNetBar.TabItem tbItem2ChkEdit;
        private DevComponents.DotNetBar.ButtonX btnReport;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comboGradeYear;
        private DevComponents.DotNetBar.ButtonX buttonUpdateDSubjectName;
        private DevComponents.DotNetBar.ButtonX buttonLoadDesc;
    }
}