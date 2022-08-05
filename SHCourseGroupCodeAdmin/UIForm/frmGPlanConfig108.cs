﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using SHCourseGroupCodeAdmin.DAO;
using DevComponents.DotNetBar;
using System.Xml.Linq;
using DevComponents.AdvTree;
using FISCA.LogAgent;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmGPlanConfig108 : BaseForm
    {
        BackgroundWorker _bgWorker;
        DataAccess _da;
        List<GPlanInfo108> GP108List;
        GPlanInfo108 SelectInfo = null;
        bool isDgDataChange = false;
        bool isUDDgDataChange = false;
        bool isLoadUDDataFinish = true;

        // 檢查科目是否重複
        List<string> chkSubjectNameList = new List<string>();
        List<string> chkSubjectNameDList = new List<string>();

        int dgColIdx = 0, dgRowIdx = 0, dgUDColIdx = 0, dgUDRowIdx = 0;

        private Node _SelectItem;
        Dictionary<string, bool> _AdvTreeExpandStatus = new Dictionary<string, bool>();

        public frmGPlanConfig108()
        {
            InitializeComponent();

            this.expandablePanel1.TitleText = "課程規劃表(108課綱適用)";
            _AdvTreeExpandStatus.Clear();

            GP108List = new List<GPlanInfo108>();

            _da = new DataAccess();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取完成。");

            LoadData();
        }

        private void LoadData()
        {
            string noSchoolYear = "未分類";
            Dictionary<string, Node> itemNodes = new Dictionary<string, Node>();

            foreach (GPlanInfo108 data in GP108List)
            {
                string schoolYear = data.EntrySchoolYear;

                if (string.IsNullOrEmpty(data.EntrySchoolYear))
                {
                    schoolYear = noSchoolYear;
                }

                if (!itemNodes.ContainsKey(schoolYear))
                {
                    itemNodes.Add(schoolYear, new Node());
                    itemNodes[schoolYear].Text = (schoolYear == noSchoolYear) ? noSchoolYear : schoolYear + "學年度";
                    itemNodes[schoolYear].TagString = (schoolYear == noSchoolYear) ? "" : schoolYear;
                }

                Node childNode = new Node();
                childNode.Tag = data;
                childNode.Text = data.RefGPName;
                childNode.Name = data.RefGPName;

                if (_AdvTreeExpandStatus.ContainsKey(itemNodes[schoolYear].TagString))
                    itemNodes[schoolYear].Expanded = _AdvTreeExpandStatus[itemNodes[schoolYear].TagString];

                itemNodes[schoolYear].Nodes.Add(childNode);
            }

            List<string> sortedKey = itemNodes.Keys.ToList<string>();
            sortedKey.Sort(delegate (string key1, string key2)
            {
                if (key1 == "未分類") return 1;
                if (key2 == "未分類") return -1;

                string sort1 = key1.PadLeft(10, '0');
                string sort2 = key2.PadLeft(10, '0');
                return sort2.CompareTo(sort1);
            });


            // 把結果填入畫面
            #region 把結果填入畫面
            advTree1.BeginUpdate();
            advTree1.Nodes.Clear();
            foreach (string key in sortedKey)
            {
                advTree1.Nodes.Add(itemNodes[key]);
            }

            if (_AdvTreeExpandStatus.Count == 0)
            {
                if (advTree1.Nodes.Count > 0)
                    advTree1.Nodes[0].Expand();
            }

            advTree1.EndUpdate();
            #endregion 把結果填入畫面

            if (_SelectItem != null)
            {
                advTree1.SelectedNode = _SelectItem;
            }

            LoadDataGridViewColumns();
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取課程規劃表...", e.ProgressPercentage);
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            GP108List = _da.GetGPlanInfo108All();
            _bgWorker.ReportProgress(100);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // 不能修改，如果需要修改待開發
            MsgBox.Show("儲存完成");
        }

        private void frmGPlanConfig108_Load(object sender, EventArgs e)
        {
            ReloadData();
        }

        private void ReloadData()
        {
            advTree1.Nodes.Clear();
            btnEditName.Enabled = btnUpdate.Enabled = btnDelete.Enabled = false;
            tabItem1.Visible = tabItem2.Visible = tabItem4.Visible = false;
            _bgWorker.RunWorkerAsync();
        }

        private void LoadDataGridViewColumns()
        {
            dgData.Columns.Clear();
            dgUDData.Columns.Clear();
            try
            {

                DataGridViewTextBoxColumn tbDomain = new DataGridViewTextBoxColumn();
                tbDomain.Name = "領域";
                tbDomain.Width = 80;
                tbDomain.HeaderText = "領域";
                tbDomain.ReadOnly = true;
                tbDomain.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbScoreType = new DataGridViewTextBoxColumn();
                tbScoreType.Name = "分項類別";
                tbScoreType.Width = 80;
                tbScoreType.HeaderText = "分項類別";
                tbScoreType.ReadOnly = true;
                tbScoreType.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbSubjectName = new DataGridViewTextBoxColumn();
                tbSubjectName.Name = "科目名稱";
                tbSubjectName.Width = 160;
                tbSubjectName.HeaderText = "科目名稱";
                tbSubjectName.ReadOnly = true;
                tbSubjectName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbRequiredBy = new DataGridViewTextBoxColumn();
                tbRequiredBy.Name = "校訂部定";
                tbRequiredBy.Width = 40;
                tbRequiredBy.HeaderText = "校訂部定";
                tbRequiredBy.ReadOnly = true;
                tbRequiredBy.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbIsRequired = new DataGridViewTextBoxColumn();
                tbIsRequired.Name = "必選修";
                tbIsRequired.Width = 40;
                tbIsRequired.HeaderText = "必選修";
                tbIsRequired.ReadOnly = true;
                tbIsRequired.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbGS11 = new DataGridViewTextBoxColumn();
                tbGS11.Name = "1上";
                tbGS11.Width = 40;
                tbGS11.HeaderText = "1上";
                tbGS11.ReadOnly = true;
                tbGS11.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbGS12 = new DataGridViewTextBoxColumn();
                tbGS12.Name = "1下";
                tbGS12.Width = 40;
                tbGS12.HeaderText = "1下";
                tbGS12.ReadOnly = true;
                tbGS12.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbGS21 = new DataGridViewTextBoxColumn();
                tbGS21.Name = "2上";
                tbGS21.Width = 40;
                tbGS21.HeaderText = "2上";
                tbGS21.ReadOnly = true;
                tbGS21.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbGS22 = new DataGridViewTextBoxColumn();
                tbGS22.Name = "2下";
                tbGS22.Width = 40;
                tbGS22.HeaderText = "2下";
                tbGS22.ReadOnly = true;
                tbGS22.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbGS31 = new DataGridViewTextBoxColumn();
                tbGS31.Name = "3上";
                tbGS31.Width = 40;
                tbGS31.HeaderText = "3上";
                tbGS31.ReadOnly = true;
                tbGS31.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbGS32 = new DataGridViewTextBoxColumn();
                tbGS32.Name = "3下";
                tbGS32.Width = 40;
                tbGS32.HeaderText = "3下";
                tbGS32.ReadOnly = true;
                tbGS32.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbNotIncludedInCalc = new DataGridViewTextBoxColumn();
                tbNotIncludedInCalc.Name = "不需評分";
                tbNotIncludedInCalc.Width = 40;
                tbNotIncludedInCalc.HeaderText = "不需評分";
                tbNotIncludedInCalc.ReadOnly = true;
                tbNotIncludedInCalc.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbNotIncludedInCredit = new DataGridViewTextBoxColumn();
                tbNotIncludedInCredit.Name = "不計學分";
                tbNotIncludedInCredit.Width = 40;
                tbNotIncludedInCredit.HeaderText = "不計學分";
                tbNotIncludedInCredit.ReadOnly = true;
                tbNotIncludedInCredit.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;


                DataGridViewTextBoxColumn tbOpenStatus = new DataGridViewTextBoxColumn();
                tbOpenStatus.Name = "開課方式";
                tbOpenStatus.Width = 40;
                tbOpenStatus.HeaderText = "開課方式";
                tbOpenStatus.ReadOnly = true;
                tbOpenStatus.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbCourseCode = new DataGridViewTextBoxColumn();
                tbCourseCode.Name = "課程代碼";
                tbCourseCode.Width = 300;
                tbCourseCode.HeaderText = "課程代碼";
                tbCourseCode.ReadOnly = true;
                tbCourseCode.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgData.Columns.Add(tbDomain);
                dgData.Columns.Add(tbScoreType);
                dgData.Columns.Add(tbSubjectName);
                dgData.Columns.Add(tbRequiredBy);
                dgData.Columns.Add(tbIsRequired);
                dgData.Columns.Add(tbGS11);
                dgData.Columns.Add(tbGS12);
                dgData.Columns.Add(tbGS21);
                dgData.Columns.Add(tbGS22);
                dgData.Columns.Add(tbGS31);
                dgData.Columns.Add(tbGS32);
                dgData.Columns.Add(tbNotIncludedInCalc);
                dgData.Columns.Add(tbNotIncludedInCredit);
                dgData.Columns.Add(tbOpenStatus);
                dgData.Columns.Add(tbCourseCode);

                //// 因為自動排序有些問題，先將關閉
                //foreach(DataGridViewColumn col in dgData.Columns )
                //{
                //    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //}


                // --使用者自訂
                DataGridViewTextBoxColumn tbUDDomain = new DataGridViewTextBoxColumn();
                tbUDDomain.Name = "領域";
                tbUDDomain.Width = 80;
                tbUDDomain.HeaderText = "領域";
                tbUDDomain.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewComboBoxColumn cbUDScoreType = new DataGridViewComboBoxColumn();
                cbUDScoreType.Name = "分項類別";
                cbUDScoreType.Width = 80;
                cbUDScoreType.DropDownWidth = 90;
                cbUDScoreType.HeaderText = "分項類別";
                cbUDScoreType.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                List<string> ScoreTypeList = new List<string>();
                ScoreTypeList.Add("學業");
                ScoreTypeList.Add("專業科目");
                ScoreTypeList.Add("實習科目");

                DataTable dtUDScoreType = new DataTable();
                dtUDScoreType.Columns.Add("VALUE");
                dtUDScoreType.Columns.Add("ITEM");

                foreach (string str in ScoreTypeList)
                {
                    DataRow dr = dtUDScoreType.NewRow();
                    dr["VALUE"] = str;
                    dr["ITEM"] = str;
                    dtUDScoreType.Rows.Add(dr);
                }
                cbUDScoreType.DataSource = dtUDScoreType;
                cbUDScoreType.DisplayMember = "ITEM";
                cbUDScoreType.ValueMember = "VALUE";



                DataGridViewTextBoxColumn tbUDSubjectName = new DataGridViewTextBoxColumn();
                tbUDSubjectName.Name = "科目名稱";
                tbUDSubjectName.Width = 160;
                tbUDSubjectName.HeaderText = "科目名稱";
                tbUDSubjectName.ReadOnly = false;
                tbUDSubjectName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewComboBoxColumn cbUDRequiredBy = new DataGridViewComboBoxColumn();
                cbUDRequiredBy.Name = "校訂部定";
                cbUDRequiredBy.Width = 60;
                cbUDRequiredBy.HeaderText = "校訂部定";
                cbUDRequiredBy.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                List<string> RequiredByList = new List<string>();
                RequiredByList.Add("部定");
                RequiredByList.Add("校訂");

                DataTable dtUDRequiredBy = new DataTable();
                dtUDRequiredBy.Columns.Add("VALUE");
                dtUDRequiredBy.Columns.Add("ITEM");

                foreach (string str in RequiredByList)
                {
                    DataRow dr = dtUDRequiredBy.NewRow();
                    dr["VALUE"] = str;
                    dr["ITEM"] = str;
                    dtUDRequiredBy.Rows.Add(dr);
                }

                cbUDRequiredBy.DataSource = dtUDRequiredBy;
                cbUDRequiredBy.DisplayMember = "ITEM";
                cbUDRequiredBy.ValueMember = "VALUE";

                DataGridViewComboBoxColumn cbUDIsRequired = new DataGridViewComboBoxColumn();
                cbUDIsRequired.Name = "必選修";
                cbUDIsRequired.Width = 60;
                cbUDIsRequired.HeaderText = "必選修";
                cbUDIsRequired.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                List<string> RequiredList = new List<string>();
                RequiredList.Add("必修");
                RequiredList.Add("選修");

                DataTable dtUDRequired = new DataTable();
                dtUDRequired.Columns.Add("VALUE");
                dtUDRequired.Columns.Add("ITEM");

                foreach (string str in RequiredList)
                {
                    DataRow dr = dtUDRequired.NewRow();
                    dr["VALUE"] = str;
                    dr["ITEM"] = str;
                    dtUDRequired.Rows.Add(dr);
                }

                cbUDIsRequired.DataSource = dtUDRequired;
                cbUDIsRequired.DisplayMember = "ITEM";
                cbUDIsRequired.ValueMember = "VALUE";


                DataGridViewTextBoxColumn tbUDGS11 = new DataGridViewTextBoxColumn();
                tbUDGS11.Name = "1上";
                tbUDGS11.Width = 40;
                tbUDGS11.HeaderText = "1上";
                tbUDGS11.ReadOnly = false;
                tbUDGS11.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbUDGS12 = new DataGridViewTextBoxColumn();
                tbUDGS12.Name = "1下";
                tbUDGS12.Width = 40;
                tbUDGS12.HeaderText = "1下";
                tbUDGS12.ReadOnly = false;
                tbUDGS12.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbUDGS21 = new DataGridViewTextBoxColumn();
                tbUDGS21.Name = "2上";
                tbUDGS21.Width = 40;
                tbUDGS21.HeaderText = "2上";
                tbUDGS21.ReadOnly = false;
                tbUDGS21.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbUDGS22 = new DataGridViewTextBoxColumn();
                tbUDGS22.Name = "2下";
                tbUDGS22.Width = 40;
                tbUDGS22.HeaderText = "2下";
                tbUDGS22.ReadOnly = false;
                tbUDGS22.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbUDGS31 = new DataGridViewTextBoxColumn();
                tbUDGS31.Name = "3上";
                tbUDGS31.Width = 40;
                tbUDGS31.HeaderText = "3上";
                tbUDGS31.ReadOnly = false;
                tbUDGS31.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbUDGS32 = new DataGridViewTextBoxColumn();
                tbUDGS32.Name = "3下";
                tbUDGS32.Width = 40;
                tbUDGS32.HeaderText = "3下";
                tbUDGS32.ReadOnly = false;
                tbUDGS32.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewComboBoxColumn cbUDNotIncludedInCalc = new DataGridViewComboBoxColumn();
                cbUDNotIncludedInCalc.Name = "不需評分";
                cbUDNotIncludedInCalc.Width = 40;
                cbUDNotIncludedInCalc.HeaderText = "不需評分";
                cbUDNotIncludedInCalc.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                List<string> YesNoList = new List<string>();
                YesNoList.Add("是");
                YesNoList.Add("否");

                DataTable dtUDNotIncludedInCalc = new DataTable();
                dtUDNotIncludedInCalc.Columns.Add("VALUE");
                dtUDNotIncludedInCalc.Columns.Add("ITEM");

                foreach (string str in YesNoList)
                {
                    DataRow dr = dtUDNotIncludedInCalc.NewRow();
                    dr["VALUE"] = str;
                    dr["ITEM"] = str;
                    dtUDNotIncludedInCalc.Rows.Add(dr);
                }

                cbUDNotIncludedInCalc.DataSource = dtUDNotIncludedInCalc;
                cbUDNotIncludedInCalc.DisplayMember = "ITEM";
                cbUDNotIncludedInCalc.ValueMember = "VALUE";

                DataGridViewComboBoxColumn cbUDNotIncludedInCredit = new DataGridViewComboBoxColumn();
                cbUDNotIncludedInCredit.Name = "不計學分";
                cbUDNotIncludedInCredit.Width = 40;
                cbUDNotIncludedInCredit.HeaderText = "不計學分";
                cbUDNotIncludedInCredit.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataTable dtUDNotIncludedInCredit = new DataTable();
                dtUDNotIncludedInCredit.Columns.Add("VALUE");
                dtUDNotIncludedInCredit.Columns.Add("ITEM");

                foreach (string str in YesNoList)
                {
                    DataRow dr = dtUDNotIncludedInCredit.NewRow();
                    dr["VALUE"] = str;
                    dr["ITEM"] = str;
                    dtUDNotIncludedInCredit.Rows.Add(dr);
                }

                cbUDNotIncludedInCredit.DataSource = dtUDNotIncludedInCredit;
                cbUDNotIncludedInCredit.DisplayMember = "ITEM";
                cbUDNotIncludedInCredit.ValueMember = "VALUE";


                DataGridViewComboBoxColumn cbUDOpenStatus = new DataGridViewComboBoxColumn();
                cbUDOpenStatus.Name = "開課方式";
                cbUDOpenStatus.Width = 60;
                cbUDOpenStatus.HeaderText = "開課方式";
                cbUDOpenStatus.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                List<string> OpenList = new List<string>();
                OpenList.Add("原班");
                OpenList.Add("跨班");

                DataTable dtUDOpenStatus = new DataTable();
                dtUDOpenStatus.Columns.Add("VALUE");
                dtUDOpenStatus.Columns.Add("ITEM");

                foreach (string str in OpenList)
                {
                    DataRow dr = dtUDOpenStatus.NewRow();
                    dr["VALUE"] = str;
                    dr["ITEM"] = str;
                    dtUDOpenStatus.Rows.Add(dr);
                }

                cbUDOpenStatus.DataSource = dtUDOpenStatus;
                cbUDOpenStatus.DisplayMember = "ITEM";
                cbUDOpenStatus.ValueMember = "VALUE";


                dgUDData.Columns.Add(tbUDDomain);
                dgUDData.Columns.Add(cbUDScoreType);
                dgUDData.Columns.Add(tbUDSubjectName);
                dgUDData.Columns.Add(cbUDRequiredBy);
                dgUDData.Columns.Add(cbUDIsRequired);
                dgUDData.Columns.Add(tbUDGS11);
                dgUDData.Columns.Add(tbUDGS12);
                dgUDData.Columns.Add(tbUDGS21);
                dgUDData.Columns.Add(tbUDGS22);
                dgUDData.Columns.Add(tbUDGS31);
                dgUDData.Columns.Add(tbUDGS32);
                dgUDData.Columns.Add(cbUDNotIncludedInCalc);
                dgUDData.Columns.Add(cbUDNotIncludedInCredit);
                dgUDData.Columns.Add(cbUDOpenStatus);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void advTree1_Click(object sender, EventArgs e)
        {
            btnEditName.Enabled = btnUpdate.Enabled = btnDelete.Enabled = false;
        }

        private void advTree1_NodeClick(object sender, TreeNodeMouseEventArgs e)
        {
            if (isDgDataChange || isUDDgDataChange)
            {
                if (DialogResult.No == MsgBox.Show("變更尚未儲存，確定離開？", MessageBoxButtons.YesNo))
                {
                    return;
                }
                isDgDataChange = isUDDgDataChange = false;
            }


            isLoadUDDataFinish = false;
            this.lblGroupName.Text = "";
            lblUDGroupName.Text = "";
            dgUDData.Rows.Clear();
            dgData.Rows.Clear();
            chkSubjectNameList.Clear();
            SelectInfo = null;

            if (!(e.Node.Tag is GPlanInfo108))
            {
                _SelectItem = null;
                tabItem1.Visible = tabItem2.Visible = tabItem4.Visible = false;
                return;
            }

            if (_SelectItem != null)
                _SelectItem.Checked = false;

            _SelectItem = e.Node;
            _SelectItem.Checked = true;

            SelectInfo = (GPlanInfo108)_SelectItem.Tag;

            // 判斷功能項目是否顯示
            if (SelectInfo != null)
            {
                if (string.IsNullOrEmpty(SelectInfo.GDCCode))
                {
                    btnEditName.Enabled = btnDelete.Enabled = true;
                    tabItem1.Visible = false;
                    tabItem2.Visible = tabItem4.Visible = true;
                    tabControl1.SelectedTabIndex = 1;

                }
                else
                {
                    tabItem1.Visible = true;
                    tabItem2.Visible = tabItem4.Visible = true;
                    btnEditName.Enabled = btnDelete.Enabled = false;
                    tabControl1.SelectedTab = tabItem4;
                    tabControl1.SelectedTab = tabItem1;
                    //tabControl1.SelectedTabIndex = 0;
                }
            }

            try
            {
                // 解析 XML
                if (SelectInfo.RefGPContentXml == null)
                    SelectInfo.RefGPContentXml = XElement.Parse(SelectInfo.RefGPContent);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            #region  課程規劃表(國教署)
            lblGroupName.Text = SelectInfo.RefGPName;

            // 資料整理
            Dictionary<string, List<XElement>> dataDict = new Dictionary<string, List<XElement>>();
            foreach (XElement elm in SelectInfo.RefGPContentXml.Elements("Subject"))
            {
                string idx = elm.Element("Grouping").Attribute("RowIndex").Value;

                if (!dataDict.ContainsKey(idx))
                    dataDict.Add(idx, new List<XElement>());

                dataDict[idx].Add(elm);
            }


            dgData.Rows.Clear();

            // 填入資料
            foreach (string idx in dataDict.Keys)
            {
                int rowIdx = dgData.Rows.Add();
                XElement firstElm = null;
                if (dataDict[idx].Count > 0)
                {
                    firstElm = dataDict[idx][0];
                }

                // 將資料存入 Tag
                dgData.Rows[rowIdx].Tag = firstElm;

                dgData.Rows[rowIdx].Cells["領域"].Value = firstElm.Attribute("Domain").Value;
                dgData.Rows[rowIdx].Cells["分項類別"].Value = firstElm.Attribute("Entry").Value;
                dgData.Rows[rowIdx].Cells["科目名稱"].Value = firstElm.Attribute("SubjectName").Value;

                chkSubjectNameList.Add(firstElm.Attribute("SubjectName").Value);

                if (firstElm.Attribute("RequiredBy").Value == "部訂")
                {
                    dgData.Rows[rowIdx].Cells["校訂部定"].Value = "部定";
                }
                else
                    dgData.Rows[rowIdx].Cells["校訂部定"].Value = firstElm.Attribute("RequiredBy").Value;

                dgData.Rows[rowIdx].Cells["必選修"].Value = firstElm.Attribute("Required").Value;

                foreach (XElement elmD in dataDict[idx])
                {
                    try
                    {
                        if (elmD.Attribute("GradeYear").Value == "1" && elmD.Attribute("Semester").Value == "1")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgData.Rows[rowIdx].Cells["1上"].Tag = elmD;
                            dgData.Rows[rowIdx].Cells["1上"].Style.BackColor = c1.BackgroundColor;
                            dgData.Rows[rowIdx].Cells["1上"].Value = c1.StringValue;
                            dgData.Rows[rowIdx].Cells["1上"].ToolTipText = c1.ToolTipText;
                        }

                        if (elmD.Attribute("GradeYear").Value == "1" && elmD.Attribute("Semester").Value == "2")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgData.Rows[rowIdx].Cells["1下"].Tag = elmD;
                            dgData.Rows[rowIdx].Cells["1下"].Style.BackColor = c1.BackgroundColor;
                            dgData.Rows[rowIdx].Cells["1下"].Value = c1.StringValue;
                            dgData.Rows[rowIdx].Cells["1下"].ToolTipText = c1.ToolTipText;
                        }

                        if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "1")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgData.Rows[rowIdx].Cells["2上"].Tag = elmD;
                            dgData.Rows[rowIdx].Cells["2上"].Style.BackColor = c1.BackgroundColor;
                            dgData.Rows[rowIdx].Cells["2上"].Value = c1.StringValue;
                            dgData.Rows[rowIdx].Cells["2上"].ToolTipText = c1.ToolTipText;
                        }

                        if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "2")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgData.Rows[rowIdx].Cells["2下"].Tag = elmD;
                            dgData.Rows[rowIdx].Cells["2下"].Style.BackColor = c1.BackgroundColor;
                            dgData.Rows[rowIdx].Cells["2下"].Value = c1.StringValue;
                            dgData.Rows[rowIdx].Cells["2下"].ToolTipText = c1.ToolTipText;

                        }

                        if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "1")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgData.Rows[rowIdx].Cells["3上"].Tag = elmD;
                            dgData.Rows[rowIdx].Cells["3上"].Style.BackColor = c1.BackgroundColor;
                            dgData.Rows[rowIdx].Cells["3上"].Value = c1.StringValue;
                            dgData.Rows[rowIdx].Cells["3上"].ToolTipText = c1.ToolTipText;

                        }

                        if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "2")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgData.Rows[rowIdx].Cells["3下"].Tag = elmD;
                            dgData.Rows[rowIdx].Cells["3下"].Style.BackColor = c1.BackgroundColor;
                            dgData.Rows[rowIdx].Cells["3下"].Value = c1.StringValue;
                            dgData.Rows[rowIdx].Cells["3下"].ToolTipText = c1.ToolTipText;
                        }


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                dgData.Rows[rowIdx].Cells["不需評分"].Value = "否";
                dgData.Rows[rowIdx].Cells["不計學分"].Value = "否";

                if (firstElm.Attribute("NotIncludedInCalc").Value == "True")
                    dgData.Rows[rowIdx].Cells["不需評分"].Value = "是";

                if (firstElm.Attribute("NotIncludedInCredit").Value == "True")
                    dgData.Rows[rowIdx].Cells["不計學分"].Value = "是";

                dgData.Rows[rowIdx].Cells["開課方式"].Value = firstElm.Attribute("開課方式").Value;
                dgData.Rows[rowIdx].Cells["課程代碼"].Value = firstElm.Attribute("課程代碼").Value;

            }

            #endregion

            #region 使用者自訂科目(自訂課程規劃)
            lblUDGroupName.Text = SelectInfo.RefGPName;
            // 取得使用者自訂
            Dictionary<string, List<XElement>> elmUDRoot = SelectInfo.GetUserDefSubjectDict();

            // 資料整理
            Dictionary<string, List<XElement>> dataUDDict = new Dictionary<string, List<XElement>>();
            foreach (string key in elmUDRoot.Keys)
                foreach (XElement elm in elmUDRoot[key])
                {
                    string idx = elm.Element("Grouping").Attribute("RowIndex").Value;

                    if (!dataUDDict.ContainsKey(idx))
                        dataUDDict.Add(idx, new List<XElement>());

                    dataUDDict[idx].Add(elm);
                }

            dgUDData.Rows.Clear();
            // 填入資料
            foreach (string idx in dataUDDict.Keys)
            {
                int rowIdx = dgUDData.Rows.Add();
                XElement firstElm = null;
                if (dataUDDict[idx].Count > 0)
                {
                    firstElm = dataUDDict[idx][0];
                }

                // 將資料存入 Tag
                dgUDData.Rows[rowIdx].Tag = firstElm;

                dgUDData.Rows[rowIdx].Cells["領域"].Value = firstElm.Attribute("Domain").Value;
                dgUDData.Rows[rowIdx].Cells["分項類別"].Value = firstElm.Attribute("Entry").Value;
                dgUDData.Rows[rowIdx].Cells["科目名稱"].Value = firstElm.Attribute("SubjectName").Value;

                if (firstElm.Attribute("RequiredBy").Value == "部訂")
                {
                    dgUDData.Rows[rowIdx].Cells["校訂部定"].Value = "部定";
                }
                else
                    dgUDData.Rows[rowIdx].Cells["校訂部定"].Value = firstElm.Attribute("RequiredBy").Value;

                dgUDData.Rows[rowIdx].Cells["必選修"].Value = firstElm.Attribute("Required").Value;

                foreach (XElement elmD in dataUDDict[idx])
                {
                    try
                    {
                        if (elmD.Attribute("GradeYear").Value == "1" && elmD.Attribute("Semester").Value == "1")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgUDData.Rows[rowIdx].Cells["1上"].Tag = elmD;
                            dgUDData.Rows[rowIdx].Cells["1上"].Style.BackColor = c1.BackgroundColor;
                            dgUDData.Rows[rowIdx].Cells["1上"].Value = elmD.Attribute("Credit").Value;
                            dgUDData.Rows[rowIdx].Cells["1上"].ToolTipText = c1.ToolTipText;
                        }

                        if (elmD.Attribute("GradeYear").Value == "1" && elmD.Attribute("Semester").Value == "2")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgUDData.Rows[rowIdx].Cells["1下"].Tag = elmD;
                            dgUDData.Rows[rowIdx].Cells["1下"].Style.BackColor = c1.BackgroundColor;
                            dgUDData.Rows[rowIdx].Cells["1下"].Value = elmD.Attribute("Credit").Value;
                            dgUDData.Rows[rowIdx].Cells["1下"].ToolTipText = c1.ToolTipText;
                        }

                        if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "1")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgUDData.Rows[rowIdx].Cells["2上"].Tag = elmD;
                            dgUDData.Rows[rowIdx].Cells["2上"].Style.BackColor = c1.BackgroundColor;
                            dgUDData.Rows[rowIdx].Cells["2上"].Value = elmD.Attribute("Credit").Value;
                            dgUDData.Rows[rowIdx].Cells["2上"].ToolTipText = c1.ToolTipText;
                        }

                        if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "2")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgUDData.Rows[rowIdx].Cells["2下"].Tag = elmD;
                            dgUDData.Rows[rowIdx].Cells["2下"].Style.BackColor = c1.BackgroundColor;
                            dgUDData.Rows[rowIdx].Cells["2下"].Value = elmD.Attribute("Credit").Value;
                            dgUDData.Rows[rowIdx].Cells["2下"].ToolTipText = c1.ToolTipText;

                        }

                        if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "1")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgUDData.Rows[rowIdx].Cells["3上"].Tag = elmD;
                            dgUDData.Rows[rowIdx].Cells["3上"].Style.BackColor = c1.BackgroundColor;
                            dgUDData.Rows[rowIdx].Cells["3上"].Value = elmD.Attribute("Credit").Value;
                            dgUDData.Rows[rowIdx].Cells["3上"].ToolTipText = c1.ToolTipText;

                        }

                        if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "2")
                        {
                            CreditInfo c1 = GetCreditAttr(elmD);
                            dgUDData.Rows[rowIdx].Cells["3下"].Tag = elmD;
                            dgUDData.Rows[rowIdx].Cells["3下"].Style.BackColor = c1.BackgroundColor;
                            dgUDData.Rows[rowIdx].Cells["3下"].Value = elmD.Attribute("Credit").Value;
                            dgUDData.Rows[rowIdx].Cells["3下"].ToolTipText = c1.ToolTipText;
                        }


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                dgUDData.Rows[rowIdx].Cells["不需評分"].Value = "否";
                dgUDData.Rows[rowIdx].Cells["不計學分"].Value = "否";

                if (firstElm.Attribute("NotIncludedInCalc").Value == "True")
                    dgUDData.Rows[rowIdx].Cells["不需評分"].Value = "是";

                if (firstElm.Attribute("NotIncludedInCredit").Value == "True")
                    dgUDData.Rows[rowIdx].Cells["不計學分"].Value = "是";

                dgUDData.Rows[rowIdx].Cells["開課方式"].Value = firstElm.Attribute("開課方式").Value;


            }

            #endregion


            #region 處理採用班級
            listViewEx1.SuspendLayout();
            listViewEx1.Items.Clear();
            listViewEx1.Groups.Clear();
            // 取得採用班級
            Dictionary<string, List<DataRow>> classRows = _da.GetGPlanRefClaasByID(SelectInfo.RefGPID);

            foreach (string key in classRows.Keys)
            {
                string groupKey;

                groupKey = key + "　年級";

                foreach (DataRow dr in classRows[key])
                {
                    ListViewGroup group = listViewEx1.Groups[groupKey];
                    if (group == null)
                        group = listViewEx1.Groups.Add(groupKey, groupKey);

                    string c_name = dr["class_name"] + "(" + dr["stud_cot"] + ")";
                    ListViewItem lvi = new ListViewItem(c_name, 0, group);
                    listViewEx1.Items.Add(lvi);
                }
            }
            listViewEx1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listViewEx1.ResumeLayout();
            #endregion



            btnUpdate.Enabled = false;
            isDgDataChange = false;
            isUDDgDataChange = false;
            isLoadUDDataFinish = true;

        }

        private void TabItem2_Click(object sender, EventArgs e)
        {

        }

        private CreditInfo GetCreditAttr(XElement elm)
        {
            CreditInfo value = new CreditInfo();
            if (elm != null)
            {
                int idx = -1;

                if (elm.Attribute("GradeYear").Value == "1" && elm.Attribute("Semester").Value == "1")
                    idx = 0;

                if (elm.Attribute("GradeYear").Value == "1" && elm.Attribute("Semester").Value == "2")
                    idx = 1;

                if (elm.Attribute("GradeYear").Value == "2" && elm.Attribute("Semester").Value == "1")
                    idx = 2;

                if (elm.Attribute("GradeYear").Value == "2" && elm.Attribute("Semester").Value == "2")
                    idx = 3;

                if (elm.Attribute("GradeYear").Value == "3" && elm.Attribute("Semester").Value == "1")
                    idx = 4;

                if (elm.Attribute("GradeYear").Value == "3" && elm.Attribute("Semester").Value == "2")
                    idx = 5;


                // 對開初始
                value.isOpenD = false;
                // null 表示使用者未設定
                value.isSetOpenD = null;
                value.BackgroundColor = Color.White;
                value.ToolTipText = "";

                if (elm.Attribute("設定對開") != null)
                {
                    if (elm.Attribute("設定對開").Value == "是")
                        value.isSetOpenD = true;

                    if (elm.Attribute("設定對開").Value == "否")
                        value.isSetOpenD = false;

                    if (elm.Attribute("設定對開").Value == "")
                        value.isSetOpenD = null;
                }

                if (elm.Attribute("授課學期學分") != null)
                {
                    char[] cc = elm.Attribute("授課學期學分").Value.ToCharArray();

                    if (idx < cc.Length)
                    {
                        value.StringValue = cc[idx].ToString();
                        value.isOpenD = true;

                        int cr;
                        if (int.TryParse(value.StringValue, out cr))
                        {
                            value.isOpenD = false;
                        }
                        else
                            value.isOpenD = true;
                    }

                }

                // 先判斷使用者設定對開
                if (value.isSetOpenD.HasValue)
                {
                    // 設定 是
                    if (value.isSetOpenD.Value)
                    {
                        value.BackgroundColor = Color.LightPink;
                        value.ToolTipText = "使用者設定";
                    }
                    else
                    {
                        // 設定否
                        if (value.isSetOpenD.Value == false)
                        {
                            value.BackgroundColor = Color.White;
                            value.ToolTipText = "";
                        }
                        else
                        {
                            if (value.isOpenD)
                            {
                                value.BackgroundColor = Color.Yellow;
                                value.ToolTipText = "系統判斷";
                            }
                        }
                    }
                }
                else
                {
                    // 使用者未設
                    // 程式判斷對開
                    if (value.isOpenD)
                    {
                        value.BackgroundColor = Color.Yellow;
                        value.ToolTipText = "系統判斷";
                    }
                }
            }

            return value;
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTabIndex == 0)
            {
                if (dgColIdx > 4 && dgColIdx < 11 && dgRowIdx > -1)
                {
                    XElement elm = dgData.Rows[dgRowIdx].Cells[dgColIdx].Tag as XElement;
                    if (elm != null)
                    {
                        elm.SetAttributeValue("設定對開", "是");
                    }
                    CreditInfo c1 = GetCreditAttr(elm);
                    dgData.Rows[dgRowIdx].Cells[dgColIdx].Tag = elm;
                    dgData.Rows[dgRowIdx].Cells[dgColIdx].Style.BackColor = c1.BackgroundColor;
                    dgData.Rows[dgRowIdx].Cells[dgColIdx].ToolTipText = c1.ToolTipText;


                    // 更新資料
                    UpdateGPlaDataSubjectOpen(SelectInfo, elm, "設定對開:是");
                    SetIsDirtyDisplay(true);
                }
            }


            if (tabControl1.SelectedTabIndex == 1)
            {
                if (dgUDColIdx > 4 && dgUDColIdx < 11 && dgUDRowIdx > -1)
                {
                    XElement elm = dgUDData.Rows[dgUDRowIdx].Cells[dgUDColIdx].Tag as XElement;
                    if (elm != null)
                    {
                        elm.SetAttributeValue("設定對開", "是");
                    }
                    CreditInfo c1 = GetCreditAttr(elm);
                    dgUDData.Rows[dgUDRowIdx].Cells[dgUDColIdx].Tag = elm;
                    dgUDData.Rows[dgUDRowIdx].Cells[dgUDColIdx].Style.BackColor = c1.BackgroundColor;
                    dgUDData.Rows[dgUDRowIdx].Cells[dgUDColIdx].ToolTipText = c1.ToolTipText;
                    SetIsDirtyDisplay(true);
                }
            }

        }


        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTabIndex == 0)
            {
                if (dgColIdx > 4 && dgColIdx < 11 && dgRowIdx > -1)
                {
                    XElement elm = dgData.Rows[dgRowIdx].Cells[dgColIdx].Tag as XElement;
                    if (elm != null)
                    {
                        elm.SetAttributeValue("設定對開", "否");
                    }

                    CreditInfo c1 = GetCreditAttr(elm);
                    dgData.Rows[dgRowIdx].Cells[dgColIdx].Tag = elm;
                    dgData.Rows[dgRowIdx].Cells[dgColIdx].Style.BackColor = c1.BackgroundColor;
                    dgData.Rows[dgRowIdx].Cells[dgColIdx].ToolTipText = c1.ToolTipText;
                    // 更新資料
                    UpdateGPlaDataSubjectOpen(SelectInfo, elm, "設定對開:否");
                    SetIsDirtyDisplay(true);
                }
            }

            if (tabControl1.SelectedTabIndex == 1)
            {
                if (dgUDColIdx > 4 && dgUDColIdx < 11 && dgUDRowIdx > -1)
                {
                    XElement elm = dgUDData.Rows[dgUDRowIdx].Cells[dgUDColIdx].Tag as XElement;
                    if (elm != null)
                    {
                        elm.SetAttributeValue("設定對開", "否");
                    }

                    CreditInfo c1 = GetCreditAttr(elm);
                    dgUDData.Rows[dgUDRowIdx].Cells[dgUDColIdx].Tag = elm;
                    dgUDData.Rows[dgUDRowIdx].Cells[dgUDColIdx].Style.BackColor = c1.BackgroundColor;
                    dgUDData.Rows[dgUDRowIdx].Cells[dgUDColIdx].ToolTipText = c1.ToolTipText;
                    SetIsDirtyDisplay(true);
                }
            }
        }

        private void dgData_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                dgColIdx = e.ColumnIndex;
                dgRowIdx = e.RowIndex;
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            if (isDgDataChange || isUDDgDataChange)
            {
                if (DialogResult.No == MsgBox.Show("變更尚未儲存，確定離開？", MessageBoxButtons.YesNo))
                {
                    return;
                }
                isDgDataChange = isUDDgDataChange = false;
            }

            btnCreate.Enabled = false;
            frmAddGPlan fgg = new frmAddGPlan();
            if (fgg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FISCA.Features.Invoke("GraduationPlanSyncAllBackground");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                ReloadData();
            }
            btnCreate.Enabled = true;
        }

        private void btnEditName_Click(object sender, EventArgs e)
        {
            if (SelectInfo != null)
            {
                btnEditName.Enabled = false;

                frmEditGPlanName fdg = new frmEditGPlanName();
                fdg.SetGPlanInfo108(SelectInfo);
                if (fdg.ShowDialog() == DialogResult.OK)
                {
                    ReloadData();
                }
                btnEditName.Enabled = true;
            }
            else
            {
                MsgBox.Show("請選擇課程規劃表");
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (MsgBox.Show("請問要刪除「" + SelectInfo.RefGPName + "」? 選「是」將刪除。", "刪除課程規劃表", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                int c = _da.DeleteGPlanByID(SelectInfo.RefGPID);
                btnDelete.Enabled = false;
                _SelectItem = null;
                try
                {
                    FISCA.Features.Invoke("GraduationPlanSyncAllBackground");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                ReloadData();
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {

                btnUpdate.Enabled = false;

                // 檢查使用這自訂科目資料
                if (CheckUDDataGridData() == false)
                {
                    MsgBox.Show("自訂課程規劃表資料有錯誤，無法儲存。");
                    btnUpdate.Enabled = true;
                    return;
                }

                // 檢查使用者自訂
                if (isUDDgDataChange)
                {
                    SelectInfo.SetUserDefSubjectDict(ConvertDGYDDToXML());
                }
                // 回寫資料
                _da.UpdateGPlanXML(SelectInfo.RefGPID, SelectInfo.RefGPContentXml.ToString());
                SelectInfo.RefGPContent = SelectInfo.RefGPContentXml.ToString();
                _SelectItem.Tag = SelectInfo;
                SetIsDirtyDisplay(false);
                MsgBox.Show("儲存完成");
                //       ApplicationLog.Log("課程規劃表(108適用)", logMsg);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        /// <summary>
        /// 轉換使用者自訂欄位成XML
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, List<XElement>> ConvertDGYDDToXML()
        {
            Dictionary<string, List<XElement>> value = new Dictionary<string, List<XElement>>();
            // 學年期
            List<string> gsList = new List<string>();
            gsList.Add("1上");
            gsList.Add("1下");
            gsList.Add("2上");
            gsList.Add("2下");
            gsList.Add("3上");
            gsList.Add("3下");

            int rowIdx = 1;
            foreach (DataGridViewRow dr in dgUDData.Rows)
            {
                if (dr.IsNewRow)
                    continue;

                List<XElement> dataXList = new List<XElement>();

                foreach (string key in gsList)
                {
                    if (dr.Cells[key].Value != null)
                    {
                        XElement elm = new XElement("Subject");
                        XElement elmG = new XElement("Grouping");
                        elmG.SetAttributeValue("RowIndex", rowIdx);
                        elm.Add(elmG);

                        elm.SetAttributeValue("Domain", GetDRVCellValue(dr, "領域"));
                        elm.SetAttributeValue("Entry", GetDRVCellValue(dr, "分項類別"));

                        if (GetDRVCellValue(dr, "不需評分") == "是")
                            elm.SetAttributeValue("NotIncludedInCalc", "True");
                        else
                            elm.SetAttributeValue("NotIncludedInCalc", "False");

                        if (GetDRVCellValue(dr, "不計學分") == "是")
                            elm.SetAttributeValue("NotIncludedInCredit", "True");
                        else
                            elm.SetAttributeValue("NotIncludedInCredit", "False");

                        if (GetDRVCellValue(dr, "校訂部定") == "部定")
                            elm.SetAttributeValue("RequiredBy", "部訂");
                        else
                            elm.SetAttributeValue("RequiredBy", "校訂");

                        elm.SetAttributeValue("Required", GetDRVCellValue(dr, "必選修"));

                        elm.SetAttributeValue("SubjectName", GetDRVCellValue(dr, "科目名稱"));
                        elm.SetAttributeValue("開課方式", GetDRVCellValue(dr, "開課方式"));

                        elm.SetAttributeValue("Credit", GetDRVCellValue(dr, key));


                        if (key == "1上")
                        {
                            elm.SetAttributeValue("設定對開", GetDDOpenString(dr.Cells[key]));
                            elm.SetAttributeValue("GradeYear", "1");
                            elm.SetAttributeValue("Semester", "1");
                        }
                        else if (key == "1下")
                        {
                            elm.SetAttributeValue("設定對開", GetDDOpenString(dr.Cells[key]));
                            elm.SetAttributeValue("GradeYear", "1");
                            elm.SetAttributeValue("Semester", "2");
                        }
                        else if (key == "2上")
                        {
                            elm.SetAttributeValue("設定對開", GetDDOpenString(dr.Cells[key]));
                            elm.SetAttributeValue("GradeYear", "2");
                            elm.SetAttributeValue("Semester", "1");
                        }
                        else if (key == "2下")
                        {
                            elm.SetAttributeValue("設定對開", GetDDOpenString(dr.Cells[key]));
                            elm.SetAttributeValue("GradeYear", "2");
                            elm.SetAttributeValue("Semester", "2");
                        }
                        else if (key == "3上")
                        {
                            elm.SetAttributeValue("設定對開", GetDDOpenString(dr.Cells[key]));
                            elm.SetAttributeValue("GradeYear", "3");
                            elm.SetAttributeValue("Semester", "1");
                        }
                        else if (key == "3下")
                        {
                            elm.SetAttributeValue("設定對開", GetDDOpenString(dr.Cells[key]));
                            elm.SetAttributeValue("GradeYear", "3");
                            elm.SetAttributeValue("Semester", "2");
                        }
                        else
                        {

                        }

                        dataXList.Add(elm);
                    }
                }

                if (dataXList.Count > 0)
                {
                    XElement elm = dataXList[0];
                    string key = GetAttribute(elm, "Domain") + "_" + GetAttribute(elm, "Entry") + "_" + GetAttribute(elm, "Required") + "_" + GetAttribute(elm, "RequiredBy") + "_" + GetAttribute(elm, "SubjectName");

                    if (!value.ContainsKey(key))
                        value.Add(key, dataXList);
                }

                rowIdx++;
            }
            return value;
        }

        /// <summary>
        ///  取得使用者對開設定
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private string GetDDOpenString(DataGridViewCell cell)
        {
            string value = "";

            if (cell.Tag != null)
            {
                try
                {
                    XElement elm = XElement.Parse(cell.Tag.ToString());
                    if (elm.Attribute("設定對開") != null)
                        value = elm.Attribute("設定對開").Value;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            return value;
        }

        private string GetDRVCellValue(DataGridViewRow row, string cellName)
        {
            string value = "";
            if (row.Cells[cellName].Value != null)
                value = row.Cells[cellName].Value.ToString();

            return value;
        }

        private string UpdateGPlaDataSubjectOpen(GPlanInfo108 item, XElement updateData, string logMsg)
        {
            string value = "";
            if (item != null)
            {
                try
                {
                    // 比對資料並更新設定對開
                    foreach (XElement sElm in item.RefGPContentXml.Elements("Subject"))
                    {
                        if (sElm.Attribute("Semester").Value == updateData.Attribute("Semester").Value &&
                            sElm.Attribute("GradeYear").Value == updateData.Attribute("GradeYear").Value &&
                            sElm.Attribute("Domain").Value == updateData.Attribute("Domain").Value &&
                            sElm.Attribute("課程代碼").Value == updateData.Attribute("課程代碼").Value &&
                            sElm.Attribute("授課學期學分").Value == updateData.Attribute("授課學期學分").Value)
                        {
                            sElm.SetAttributeValue("設定對開", updateData.Attribute("設定對開").Value);
                        }
                    }

                    //// 回寫資料
                    //_da.UpdateGPlanXML(item.RefGPID, item.RefGPContentXml.ToString());
                    //item.RefGPContent = item.RefGPContentXml.ToString();
                    //_SelectItem.Tag = item;

                    //ApplicationLog.Log("課程規劃表(108適用)", logMsg);
                }
                catch (Exception ex)
                {
                    value = ex.Message;
                }
            }
            else
            {
                value = "沒有資料";
            }

            return value;
        }

        private void frmGPlanConfig108_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (isDgDataChange || isUDDgDataChange)
            {
                if (DialogResult.No == MsgBox.Show("變更尚未儲存，確定離開？", MessageBoxButtons.YesNo))
                {
                    e.Cancel = true;
                    return;
                }
                isDgDataChange = isUDDgDataChange = false;
            }
        }

        private void dgUDData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (tabControl1.SelectedTabIndex == 1 && isLoadUDDataFinish == true)
            {
                SetIsDirtyDisplay(true);
                if (e.RowIndex > -1 && e.ColumnIndex > -1)
                {
                    dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";

                    if (e.ColumnIndex >= 5 && e.ColumnIndex <= 10)
                    {
                        if (dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                        {
                            string strValue = dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                            // 檢查填學分數
                            if (e.ColumnIndex >= 5 && e.ColumnIndex <= 10)
                            {
                                decimal d;
                                if (decimal.TryParse(strValue, out d) == false)
                                {
                                    dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "必須填數字";
                                }
                            }
                        }

                    }
                    else
                    {
                        if (dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null)
                            dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "不能空值";
                        else
                        {
                            if (dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
                                dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "不能空值";
                            else
                            {
                                string strValue = dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                                // 檢查校部定 3
                                if (e.ColumnIndex == 3)
                                {
                                    if (strValue != "部定" && strValue != "校訂")
                                    {
                                        dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "必須填部定或校訂";
                                    }
                                }

                                // 檢查必選修 4                        
                                if (e.ColumnIndex == 4)
                                {
                                    if (strValue != "必修" && strValue != "選修")
                                    {
                                        dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "必須填必修或選修";
                                    }
                                }

                                // 檢查不需評分 NotIncludedInCalc 11
                                if (e.ColumnIndex == 11)
                                {
                                    if (strValue != "是" && strValue != "否")
                                    {
                                        dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "必須填是或否";
                                    }
                                }

                                // 檢查不計學分 NotIncludedInCredit 12
                                if (e.ColumnIndex == 12)
                                {
                                    if (strValue != "是" && strValue != "否")
                                    {
                                        dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "必須填是或否";
                                    }
                                }

                                // 檢查開課方式 13
                                if (e.ColumnIndex == 13)
                                {
                                    if (strValue != "原班" && strValue != "跨班")
                                    {
                                        dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "必須填原班或跨班";
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void SetIsDirtyDisplay(bool isD)
        {
            // 原本課程規劃
            if (tabControl1.SelectedTabIndex == 0)
            {
                lblGroupName.Text = SelectInfo.RefGPName + (isD ? " (<font color=\"Chocolate\">已變更</font>)" : "");
                btnUpdate.Enabled = isD;
                isDgDataChange = isD;
            }

            // 自訂課程
            if (tabControl1.SelectedTabIndex == 1)
            {
                lblUDGroupName.Text = SelectInfo.RefGPName + (isD ? " (<font color=\"Chocolate\">已變更</font>)" : "");
                btnUpdate.Enabled = isD;
                isUDDgDataChange = isD;
            }
        }

        private void dgUDData_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            // index 5 ~ 10 輸入學分數
            if (e.ColumnIndex >= 5 && e.ColumnIndex <= 10)
                dgUDData.ImeMode = ImeMode.Off;
            else
                dgUDData.ImeMode = ImeMode.On;

        }

        private void dgUDData_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                dgUDColIdx = e.ColumnIndex;
                dgUDRowIdx = e.RowIndex;
            }

            if (e.ColumnIndex > -1 && e.RowIndex > -1)
            {
                if (e.ColumnIndex == 0)
                {
                    if (dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null)
                    {
                        dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "不分";
                    }
                    else
                    {
                        if (dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "")
                        {
                            dgUDData.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "不分";
                        }
                    }
                }
            }

        }

        private void dgUDData_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dgUDData.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
            dgUDData.Rows[e.RowIndex].Selected = true;
        }

        private void dgUDData_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgUDData.IsCurrentCellDirty)
            {
                dgUDData.CommitEdit(DataGridViewDataErrorContexts.Commit);

            }
        }

        private string GetAttribute(XElement elm, string attrName)
        {
            string value = "";

            if (elm.Attribute(attrName) != null)
                value = elm.Attribute(attrName).Value;

            return value;
        }

        /// <summary>
        ///  檢查使用自訂欄位值
        /// </summary>
        /// <returns></returns>
        private bool CheckUDDataGridData()
        {
            bool value = true;

            // 檢查輸入值
            foreach (DataGridViewRow dr in dgUDData.Rows)
            {
                dr.ErrorText = "";
                foreach (DataGridViewCell cell in dr.Cells)
                {
                    cell.ErrorText = "";
                }
            }

            chkSubjectNameDList.Clear();


            foreach (DataGridViewRow dr in dgUDData.Rows)
            {
                if (dr.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in dr.Cells)
                {

                    // 檢查填學分數
                    if (cell.ColumnIndex >= 5 && cell.ColumnIndex <= 10)
                    {
                        if (cell.Value != null)
                        {
                            string strValue = dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value.ToString();

                            // 檢查填學分數
                            if (cell.ColumnIndex >= 5 && cell.ColumnIndex <= 10)
                            {
                                decimal d;
                                if (decimal.TryParse(strValue, out d) == false)
                                {
                                    dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].ErrorText = "必須填數字";
                                    value = false;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (cell.Value == null)
                        {
                            cell.ErrorText = "不能空值";
                            value = false;
                        }
                        else
                        {
                            if (dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value.ToString() == "")
                            {
                                dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].ErrorText = "不能空值";
                                value = false;
                            }
                            else
                            {
                                string strValue = dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value.ToString();

                                // 檢查科目名稱
                                if (cell.ColumnIndex == 2)
                                {
                                    if (chkSubjectNameList.Contains(strValue))
                                    {
                                        dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].ErrorText = "科目名稱與國教署科目名稱相同";
                                        value = false;
                                    }

                                    if (chkSubjectNameDList.Contains(strValue))
                                    {
                                        dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].ErrorText = "科目名稱重複";
                                        value = false;
                                    }
                                    else
                                        chkSubjectNameDList.Add(strValue);
                                }

                                // 檢查校部定 3
                                if (cell.ColumnIndex == 3)
                                {
                                    if (strValue != "部定" && strValue != "校訂")
                                    {
                                        dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].ErrorText = "必須填部定或校訂";
                                        value = false;
                                    }
                                }

                                // 檢查必選修 4                        
                                if (cell.ColumnIndex == 4)
                                {
                                    if (strValue != "必修" && strValue != "選修")
                                    {
                                        dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].ErrorText = "必須填必修或選修";
                                        value = false;
                                    }
                                }

                                // 檢查不需評分 NotIncludedInCalc 11
                                if (cell.ColumnIndex == 11)
                                {
                                    if (strValue != "是" && strValue != "否")
                                    {
                                        dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].ErrorText = "必須填是或否";
                                        value = false;
                                    }
                                }

                                // 檢查不計學分 NotIncludedInCredit 12
                                if (cell.ColumnIndex == 12)
                                {
                                    if (strValue != "是" && strValue != "否")
                                    {
                                        dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].ErrorText = "必須填是或否";
                                        value = false;
                                    }
                                }

                                // 檢查開課方式 13
                                if (cell.ColumnIndex == 13)
                                {
                                    if (strValue != "原班" && strValue != "跨班")
                                    {
                                        dgUDData.Rows[cell.RowIndex].Cells[cell.ColumnIndex].ErrorText = "必須填原班或跨班";
                                        value = false;
                                    }
                                }
                            }
                        }
                    }

                }
            }


            return value;
        }
    }
}
