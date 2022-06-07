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

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmGPlanConfig108 : BaseForm
    {
        BackgroundWorker _bgWorker;
        DataAccess _da;
        List<GPlanInfo108> GP108List;

        private Node _SelectItem;
        Dictionary<string, bool> _AdvTreeExpandStatus = new Dictionary<string, bool>();

        public frmGPlanConfig108()
        {
            InitializeComponent();

            this.expandablePanel1.TitleText = "課程規劃表(108課綱)";
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
            advTree1.Nodes.Clear();
            _bgWorker.RunWorkerAsync();
        }

        private void LoadDataGridViewColumns()
        {
            dgData.Columns.Clear();
            try
            {

                DataGridViewTextBoxColumn tbDomain = new DataGridViewTextBoxColumn();
                tbDomain.Name = "領域";
                tbDomain.Width = 70;
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
                tbSubjectName.Width = 150;
                tbSubjectName.HeaderText = "科目名稱";
                tbSubjectName.ReadOnly = true;
                tbSubjectName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbRequiredBy = new DataGridViewTextBoxColumn();
                tbRequiredBy.Name = "校訂部定";
                tbRequiredBy.Width = 60;
                tbRequiredBy.HeaderText = "校訂部定";
                tbRequiredBy.ReadOnly = true;
                tbRequiredBy.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbIsRequired = new DataGridViewTextBoxColumn();
                tbIsRequired.Name = "必選修";
                tbIsRequired.Width = 60;
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
                tbGS31.Width = 60;
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
                tbNotIncludedInCalc.Width = 60;
                tbNotIncludedInCalc.HeaderText = "不需評分";
                tbNotIncludedInCalc.ReadOnly = true;
                tbNotIncludedInCalc.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxColumn tbNotIncludedInCredit = new DataGridViewTextBoxColumn();
                tbNotIncludedInCredit.Name = "不計學分";
                tbNotIncludedInCredit.Width = 60;
                tbNotIncludedInCredit.HeaderText = "不計學分";
                tbNotIncludedInCredit.ReadOnly = true;
                tbNotIncludedInCredit.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;


                DataGridViewTextBoxColumn tbOpenStatus = new DataGridViewTextBoxColumn();
                tbOpenStatus.Name = "開課方式";
                tbOpenStatus.Width = 60;
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

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void advTree1_Click(object sender, EventArgs e)
        {
            
        }

        private void advTree1_NodeClick(object sender, TreeNodeMouseEventArgs e)
        {
            this.lblGroupName.Text = "";
            dgData.Rows.Clear();

            if (!(e.Node.Tag is GPlanInfo108))
            {
                _SelectItem = null;
                return;
            }

            if (_SelectItem != null)
                _SelectItem.Checked = false;

            _SelectItem = e.Node;            
            _SelectItem.Checked = true;

            GPlanInfo108 info = (GPlanInfo108)_SelectItem.Tag;

            dgData.Rows.Clear();

            try
            {
                // 解析 XML
                info.RefGPContentXml = XElement.Parse(info.RefGPContent);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
            lblGroupName.Text = info.RefGPName;

            // 資料整理
            Dictionary<string, List<XElement>> dataDict = new Dictionary<string, List<XElement>>();
            foreach (XElement elm in info.RefGPContentXml.Elements("Subject"))
            {
                string idx = elm.Element("Grouping").Attribute("RowIndex").Value;

                if (!dataDict.ContainsKey(idx))
                    dataDict.Add(idx, new List<XElement>());

                dataDict[idx].Add(elm);
            }

            // 取得採用班級
            Dictionary<string, List<DataRow>> classRows = _da.GetGPlanRefClaasByID(info.RefGPID);


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
                            dgData.Rows[rowIdx].Cells["1上"].Value = GetCreditAttr(elmD);
                        }

                        if (elmD.Attribute("GradeYear").Value == "1" && elmD.Attribute("Semester").Value == "2")
                        {
                            dgData.Rows[rowIdx].Cells["1下"].Value = GetCreditAttr(elmD);
                        }

                        if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "1")
                        {
                            dgData.Rows[rowIdx].Cells["2上"].Value = GetCreditAttr(elmD);
                        }

                        if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "2")
                        {
                            dgData.Rows[rowIdx].Cells["2下"].Value = GetCreditAttr(elmD);

                        }

                        if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "1")
                        {
                            dgData.Rows[rowIdx].Cells["3上"].Value = GetCreditAttr(elmD);

                        }

                        if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "2")
                        {
                            dgData.Rows[rowIdx].Cells["3下"].Value = GetCreditAttr(elmD);
                        }


                    }
                    catch (Exception ex) {
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
            

            listViewEx1.SuspendLayout();
            listViewEx1.Items.Clear();
            listViewEx1.Groups.Clear();

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

        }

        private string GetCreditAttr(XElement elm)
        {
            string value = "";

            if (elm.Attribute("學分") != null)
            {
                value = elm.Attribute("學分").Value;
            }else
            {
                if (elm.Attribute("Credit") != null)
                {
                    value = elm.Attribute("Credit").Value;
                }
            }

            return value;
        }

    }
}
