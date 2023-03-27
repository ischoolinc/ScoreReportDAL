using System;
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
using DevComponents.DotNetBar.Controls;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmGPlanConfig108 : BaseForm
    {
        // 讀資料用變數
        BackgroundWorker _bgWorker;
        DataAccess _da;
        List<GPlanInfo108> GP108List;

        // TreeView用變數
        Dictionary<string, bool> _AdvTreeExpandStatus = new Dictionary<string, bool>();
        private Node _SelectItem;
        GPlanInfo108 SelectInfo = null;

        bool isDgDataChange = false;
        bool isUDDgDataChange = false;
        bool isLoadUDDataFinish = true;
        bool isUDRowSelect = false;

        // 檢查科目是否重複
        List<string> chkSubjectNameList = new List<string>();
        List<string> chkSubjectNameDList = new List<string>();

        int dgColIdx = 0, dgRowIdx = 0, dgUDColIdx = 0, dgUDRowIdx = 0;

        // 總表及課程群組通用
        List<GPlanCourseInfo108> _CourseInfoList = new List<GPlanCourseInfo108>();
        TabItem _SelectedTab = null; // TreeView切換需回到null

        // 總表用
        bool _MainIsLoading = false;
        bool _MainIsFirstLoad = false;
        List<DataGridViewRow> _MainRowList = new List<DataGridViewRow>();
        DataGridViewRow _MainSelectedRow = new DataGridViewRow();
        bool _IsMainDataDirty = false;

        // 課程群組用
        bool _CourseGroupIsLoading = false;
        bool _CourseGroupIsFirstLoad = false;
        List<DataGridViewRow> _CourseGroupRowList = new List<DataGridViewRow>();
        List<CourseGroupSetting> _CourseGroupSettingList = new List<CourseGroupSetting>();
        List<DataGridViewCell> _CourseGroupFocusCellList = new List<DataGridViewCell>();
        Dictionary<string, List<DataGridViewCell>> _CourseGroupSettingDic = new Dictionary<string, List<DataGridViewCell>>(); // 群組設定與所屬課程
        bool _IsCourseGroupDataDirty = false;

        // 學分統計
        string _CourseType = ""; // 用來判斷學制的變數
        List<DataGridViewCell> _One1CourseList = new List<DataGridViewCell>(); // 一上
        List<DataGridViewCell> _One2CourseList = new List<DataGridViewCell>(); // 一下
        List<DataGridViewCell> _Two1CourseList = new List<DataGridViewCell>(); // 二上
        List<DataGridViewCell> _Two2CourseList = new List<DataGridViewCell>(); // 二下
        List<DataGridViewCell> _Three1CourseList = new List<DataGridViewCell>(); // 三上
        List<DataGridViewCell> _Three2CourseList = new List<DataGridViewCell>(); // 三下

        //// 技術型高中學分統計TextBox清單
        List<TextBoxX> _TechNormalSubjectRequiredbyDepartList;
        List<TextBoxX> _TechNormalSubjectRequiredbySchoolRequiredList;
        List<TextBoxX> _TechNormalSubjectRequiredbySchoolNonRequiredList;
        List<TextBoxX> _TechProfessionalSubjectRequiredByDepartProfessionalList;
        List<TextBoxX> _TechProfessionalSubjectRequiredByDepartPracticeList;
        List<TextBoxX> _TechProfessionalSubjectRequiredBySchoolProfessionRequiredList;
        List<TextBoxX> _TechProfessionalSubjectRequiredBySchoolProfessionNonRequiredList;
        List<TextBoxX> _TechProfessionalSubjectRequiredBySchoolPracticeRequiredList;
        List<TextBoxX> _TechProfessionalSubjectRequiredBySchoolPracticeNonRequiredList;

        //// 普通型高中學分統計TextBox清單
        List<TextBoxX> _NormalNormalRequiredByDepartList;
        List<TextBoxX> _NormalNormalSubjectRequiredBySchoolRequiredList;
        List<TextBoxX> _NormalNormalSubjectRequiredBySchoolNonRequiredList;

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

            #region 學分統計

            #region 技術型高中學分統計TextBox清單
            _TechNormalSubjectRequiredbyDepartList = new List<TextBoxX>()
            {
                tbNormalSubjectRequiredbyDepart1_1
                , tbNormalSubjectRequiredbyDepart1_2
                , tbNormalSubjectRequiredbyDepart2_1
                , tbNormalSubjectRequiredbyDepart2_2
                , tbNormalSubjectRequiredbyDepart3_1
                , tbNormalSubjectRequiredbyDepart3_2
            };
            _TechNormalSubjectRequiredbySchoolRequiredList = new List<TextBoxX>()
            {
                tbNormalSubjectRequiredbySchoolRequired1_1
                , tbNormalSubjectRequiredbySchoolRequired1_2
                , tbNormalSubjectRequiredbySchoolRequired2_1
                , tbNormalSubjectRequiredbySchoolRequired2_2
                , tbNormalSubjectRequiredbySchoolRequired3_1
                , tbNormalSubjectRequiredbySchoolRequired3_2
            };
            _TechNormalSubjectRequiredbySchoolNonRequiredList = new List<TextBoxX>()
            {
                tbNormalSubjectRequiredbySchoolNonRequired1_1
                , tbNormalSubjectRequiredbySchoolNonRequired1_2
                , tbNormalSubjectRequiredbySchoolNonRequired2_1
                , tbNormalSubjectRequiredbySchoolNonRequired2_2
                , tbNormalSubjectRequiredbySchoolNonRequired3_1
                , tbNormalSubjectRequiredbySchoolNonRequired3_2
            };
            _TechProfessionalSubjectRequiredByDepartProfessionalList = new List<TextBoxX>()
            {
                tbProfessionalSubjectRequiredByDepartProfessional1_1
                , tbProfessionalSubjectRequiredByDepartProfessional1_2
                , tbProfessionalSubjectRequiredByDepartProfessional2_1
                , tbProfessionalSubjectRequiredByDepartProfessional2_2
                , tbProfessionalSubjectRequiredByDepartProfessional3_1
                , tbProfessionalSubjectRequiredByDepartProfessional3_2
            };
            _TechProfessionalSubjectRequiredByDepartPracticeList = new List<TextBoxX>()
            {
                tbProfessionalSubjectRequiredByDepartPractice1_1
                , tbProfessionalSubjectRequiredByDepartPractice1_2
                , tbProfessionalSubjectRequiredByDepartPractice2_1
                , tbProfessionalSubjectRequiredByDepartPractice2_2
                , tbProfessionalSubjectRequiredByDepartPractice3_1
                , tbProfessionalSubjectRequiredByDepartPractice3_2
            };
            _TechProfessionalSubjectRequiredBySchoolProfessionRequiredList = new List<TextBoxX>()
            {
                tbProfessionalSubjectRequiredBySchoolProfessionRequired1_1
                , tbProfessionalSubjectRequiredBySchoolProfessionRequired1_2
                , tbProfessionalSubjectRequiredBySchoolProfessionRequired2_1
                , tbProfessionalSubjectRequiredBySchoolProfessionRequired2_2
                , tbProfessionalSubjectRequiredBySchoolProfessionRequired3_1
                , tbProfessionalSubjectRequiredBySchoolProfessionRequired3_2
            };
            _TechProfessionalSubjectRequiredBySchoolProfessionNonRequiredList = new List<TextBoxX>()
            {
                tbProfessionalSubjectRequiredBySchoolProfessionNonRequired1_1
                , tbProfessionalSubjectRequiredBySchoolProfessionNonRequired1_2
                , tbProfessionalSubjectRequiredBySchoolProfessionNonRequired2_1
                , tbProfessionalSubjectRequiredBySchoolProfessionNonRequired2_2
                , tbProfessionalSubjectRequiredBySchoolProfessionNonRequired3_1
                , tbProfessionalSubjectRequiredBySchoolProfessionNonRequired3_2
            };
            _TechProfessionalSubjectRequiredBySchoolPracticeRequiredList = new List<TextBoxX>()
            {
                tbProfessionalSubjectRequiredBySchoolPracticeRequired1_1
                , tbProfessionalSubjectRequiredBySchoolPracticeRequired1_2
                , tbProfessionalSubjectRequiredBySchoolPracticeRequired2_1
                , tbProfessionalSubjectRequiredBySchoolPracticeRequired2_2
                , tbProfessionalSubjectRequiredBySchoolPracticeRequired3_1
                , tbProfessionalSubjectRequiredBySchoolPracticeRequired3_2
            };
            _TechProfessionalSubjectRequiredBySchoolPracticeNonRequiredList = new List<TextBoxX>()
            {
                tbProfessionalSubjectRequiredBySchoolPracticeNonRequired1_1
                , tbProfessionalSubjectRequiredBySchoolPracticeNonRequired1_2
                , tbProfessionalSubjectRequiredBySchoolPracticeNonRequired2_1
                , tbProfessionalSubjectRequiredBySchoolPracticeNonRequired2_2
                , tbProfessionalSubjectRequiredBySchoolPracticeNonRequired3_1
                , tbProfessionalSubjectRequiredBySchoolPracticeNonRequired3_2
            };
            #endregion

            #region 普通型高中學分統計TextBox清單
            _NormalNormalRequiredByDepartList = new List<TextBoxX>()
            {
                tbNormalNormalRequiredByDepart1_1
                , tbNormalNormalRequiredByDepart1_2
                , tbNormalNormalRequiredByDepart2_1
                , tbNormalNormalRequiredByDepart2_2
                , tbNormalNormalRequiredByDepart3_1
                , tbNormalNormalRequiredByDepart3_2
            };
            _NormalNormalSubjectRequiredBySchoolRequiredList = new List<TextBoxX>()
            {
                tbNormalNormalSubjectRequiredBySchoolRequired1_1
                , tbNormalNormalSubjectRequiredBySchoolRequired1_2
                , tbNormalNormalSubjectRequiredBySchoolRequired2_1
                , tbNormalNormalSubjectRequiredBySchoolRequired2_2
                , tbNormalNormalSubjectRequiredBySchoolRequired3_1
                , tbNormalNormalSubjectRequiredBySchoolRequired3_2
            };
            _NormalNormalSubjectRequiredBySchoolNonRequiredList = new List<TextBoxX>()
            {
                tbNormalNormalSubjectRequiredBySchoolNonRequired1_1
                , tbNormalNormalSubjectRequiredBySchoolNonRequired1_2
                , tbNormalNormalSubjectRequiredBySchoolNonRequired2_1
                , tbNormalNormalSubjectRequiredBySchoolNonRequired2_2
                , tbNormalNormalSubjectRequiredBySchoolNonRequired3_1
                , tbNormalNormalSubjectRequiredBySchoolNonRequired3_2
            };
            #endregion

            #endregion
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
            tabItem1.Visible = tabItem2.Visible = tabItem4.Visible = tbiCourseGroupMain.Visible = tbiSetCourseGroup.Visible = false;
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
            btnEditName.Enabled = btnDelete.Enabled = false;
        }

        private void advTree1_NodeClick(object sender, TreeNodeMouseEventArgs e)
        {
            if (isDgDataChange || isUDDgDataChange || _IsMainDataDirty || _IsCourseGroupDataDirty)
            {
                if (DialogResult.No == MsgBox.Show("變更尚未儲存，確定離開？", MessageBoxButtons.YesNo))
                {
                    return;
                }

                // 將剛剛選取的SelectedInfo重置
                SelectInfo.RefGPContentXml = XElement.Parse(SelectInfo.RefGPContent);

                isDgDataChange = isUDDgDataChange = _IsMainDataDirty = _IsCourseGroupDataDirty = false;
            }

            // 讀取選取node
            if (e != null)
            {
                _SelectItem = e.Node;
                _SelectItem.Checked = true;

                if (!(e.Node.Tag is GPlanInfo108))
                {
                    _SelectItem = null;
                    tabItem1.Visible = tabItem2.Visible = tabItem4.Visible = false;
                    tbiCourseGroupMain.Visible = false;
                    tbiSetCourseGroup.Visible = false;
                    return;
                }

                if (_SelectItem != null)
                    _SelectItem.Checked = false;
            }
            else // e為null，代表是儲存後重新讀取
            {
                _SelectItem = (Node)sender;
            }

            SelectInfo = (GPlanInfo108)_SelectItem.Tag;

            // 初始化
            isUDRowSelect = false;
            isLoadUDDataFinish = false;
            this.lblGroupName.Text = "";
            lblUDGroupName.Text = "";
            dgUDData.Rows.Clear();
            dgData.Rows.Clear();
            chkSubjectNameList.Clear();

            //// 課程群組總表
            _CourseInfoList = new List<GPlanCourseInfo108>();
            lbMainGraduationPlanName.Text = SelectInfo.RefGPName;
            _MainRowList = new List<DataGridViewRow>();
            _MainSelectedRow = new DataGridViewRow();

            //// 設定課程群組
            lbCourseGroupGraduationPlanName.Text = SelectInfo.RefGPName;
            _CourseGroupRowList = new List<DataGridViewRow>();
            _CourseGroupSettingList.Clear();
            _CourseGroupSettingDic = new Dictionary<string, List<DataGridViewCell>>();

            // 學分統計
            _CourseType = "";

            // 判斷功能項目是否顯示
            if (SelectInfo != null)
            {
                if (string.IsNullOrEmpty(SelectInfo.GDCCode))
                {
                    btnEditName.Enabled = btnDelete.Enabled = true;
                    tabItem1.Visible = false;
                    tabItem2.Visible = tabItem4.Visible = true;
                    tbiCourseGroupMain.Visible = false;
                    tbiSetCourseGroup.Visible = false;

                    tabControl1.SelectedTabIndex = 1;

                }
                else
                {
                    tabItem1.Visible = true;
                    tbiCourseGroupMain.Visible = true;
                    tbiSetCourseGroup.Visible = true;
                    tabItem2.Visible = tabItem4.Visible = true;
                    btnEditName.Enabled = btnDelete.Enabled = false;

                    if (e != null)
                    {
                        _SelectedTab = null;
                        tabControl1.SelectedTab = tabItem4;
                        tabControl1.SelectedTab = tabItem1;
                        //tabControl1.SelectedTabIndex = 0; 
                    }
                    else // e為null，代表是儲存後重新讀取
                    {
                        tabControl1.SelectedTab = _SelectedTab;
                    }

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

            // 將選取的課程規畫表用RowIndex切成各個科目
            Dictionary<string, List<XElement>> dataDict = new Dictionary<string, List<XElement>>();
            foreach (XElement elm in SelectInfo.RefGPContentXml.Elements("Subject"))
            {
                string idx = elm.Element("Grouping").Attribute("RowIndex").Value;

                if (!dataDict.ContainsKey(idx))
                    dataDict.Add(idx, new List<XElement>());

                dataDict[idx].Add(elm);
            }

            // 讀取課程群組設定
            if (SelectInfo.RefGPContentXml.Element("CourseGroupSetting") != null)
            {
                foreach (XElement element in SelectInfo.RefGPContentXml.Element("CourseGroupSetting").Elements("CourseGroup"))
                {
                    CourseGroupSetting courseGroupSetting = new CourseGroupSetting();
                    courseGroupSetting.CourseGroupName = element.Attribute("Name").Value;
                    courseGroupSetting.CourseGroupCredit = element.Attribute("Credit").Value;
                    courseGroupSetting.CourseGroupColor = Color.FromArgb(Int32.Parse(element.Attribute("Color").Value));
                    courseGroupSetting.IsSchoolYearCourseGroup = element.Attribute("IsSchoolYearCourseGroup") == null ? false : bool.Parse(element.Attribute("IsSchoolYearCourseGroup").Value);
                    courseGroupSetting.CourseGroupElement = element;
                    _CourseGroupSettingList.Add(courseGroupSetting);

                    if (!_CourseGroupSettingDic.ContainsKey(courseGroupSetting.CourseGroupName))
                    {
                        _CourseGroupSettingDic.Add(courseGroupSetting.CourseGroupName, new List<DataGridViewCell>());
                    }
                }
            }

            #region  課程規劃表(國教署)
            lblGroupName.Text = SelectInfo.RefGPName;
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

            #region 課程群組
            foreach (string index in dataDict.Keys)
            {
                XElement firstElement = null;
                if (dataDict[index].Count > 0)
                {
                    firstElement = dataDict[index][0];
                }

                _CourseType = firstElement.Attribute("課程類別").Value;

                GPlanCourseInfo108 courseInfo = new GPlanCourseInfo108();
                courseInfo.RequiredBy = firstElement.Attribute("RequiredBy").Value == "部訂" ? "部定" : firstElement.Attribute("RequiredBy").Value;
                courseInfo.Required = firstElement.Attribute("Required").Value;
                courseInfo.SpecialCategory = firstElement.Attribute("特殊類別") == null ? "" : firstElement.Attribute("特殊類別").Value;
                courseInfo.SubjectAttribute = firstElement.Attribute("科目屬性") == null ? "" : firstElement.Attribute("科目屬性").Value;
                courseInfo.Entry = firstElement.Attribute("Entry").Value;
                courseInfo.DomainName = firstElement.Attribute("Domain").Value;
                courseInfo.OfficialSubjectName = firstElement.Attribute("OfficialSubjectName") == null ? firstElement.Attribute("SubjectName").Value : firstElement.Attribute("OfficialSubjectName").Value;
                courseInfo.StartLevel = firstElement.Element("Grouping").Attribute("startLevel").Value;
                courseInfo.SubjectName = firstElement.Attribute("SubjectName").Value;
                courseInfo.SchoolYearGroupName = firstElement.Attribute("指定學年科目名稱") == null ? "" : firstElement.Attribute("指定學年科目名稱").Value;
                courseInfo.CourseCode = firstElement.Attribute("課程代碼").Value;
                courseInfo.CourseContentList = dataDict[index];
                _CourseInfoList.Add(courseInfo);

                // 總表表格用
                DataGridViewRow mainRow = new DataGridViewRow();
                mainRow.CreateCells(dgvMain);
                mainRow.Tag = courseInfo;
                mainRow.Cells[0].Value = courseInfo.RequiredBy;
                mainRow.Cells[1].Value = courseInfo.Required;
                mainRow.Cells[2].Value = courseInfo.SpecialCategory;
                mainRow.Cells[3].Value = courseInfo.SubjectAttribute;
                mainRow.Cells[4].Value = courseInfo.Entry;
                mainRow.Cells[5].Value = courseInfo.DomainName;
                mainRow.Cells[6].Value = courseInfo.OfficialSubjectName;
                mainRow.Cells[13].Value = courseInfo.StartLevel;
                mainRow.Cells[14].Value = courseInfo.SubjectName;
                mainRow.Cells[15].Value = courseInfo.SchoolYearGroupName;
                foreach (XElement element in courseInfo.CourseContentList)
                {
                    string credit = element.Attribute("Credit").Value;
                    string courseGroupName = element.Attribute("分組名稱") == null ? "" : element.Attribute("分組名稱").Value;
                    Color courseGroupColor = Color.White;

                    // 讀取群組設定
                    if (_CourseGroupSettingList.Where(x => x.CourseGroupName == courseGroupName).Count() > 0)
                    {
                        CourseGroupSetting courseGroupSetting = _CourseGroupSettingList.Where(x => x.CourseGroupName == courseGroupName).ToList()[0];
                        courseGroupColor = courseGroupSetting.CourseGroupColor;
                    }

                    if (element.Attribute("GradeYear").Value == "1" && element.Attribute("Semester").Value == "1")
                    {
                        mainRow.Cells[7].Tag = element;
                        mainRow.Cells[7].Value = element.Attribute("Credit").Value;
                        mainRow.Cells[7].Style.BackColor = courseGroupColor;
                    }
                    if (element.Attribute("GradeYear").Value == "1" && element.Attribute("Semester").Value == "2")
                    {
                        mainRow.Cells[8].Tag = element;
                        mainRow.Cells[8].Value = element.Attribute("Credit").Value;
                        mainRow.Cells[8].Style.BackColor = courseGroupColor;
                    }
                    if (element.Attribute("GradeYear").Value == "2" && element.Attribute("Semester").Value == "1")
                    {
                        mainRow.Cells[9].Tag = element;
                        mainRow.Cells[9].Value = element.Attribute("Credit").Value;
                        mainRow.Cells[9].Style.BackColor = courseGroupColor;
                    }
                    if (element.Attribute("GradeYear").Value == "2" && element.Attribute("Semester").Value == "2")
                    {
                        mainRow.Cells[10].Tag = element;
                        mainRow.Cells[10].Value = element.Attribute("Credit").Value;
                        mainRow.Cells[10].Style.BackColor = courseGroupColor;
                    }
                    if (element.Attribute("GradeYear").Value == "3" && element.Attribute("Semester").Value == "1")
                    {
                        mainRow.Cells[11].Tag = element;
                        mainRow.Cells[11].Value = element.Attribute("Credit").Value;
                        mainRow.Cells[11].Style.BackColor = courseGroupColor;
                    }
                    if (element.Attribute("GradeYear").Value == "3" && element.Attribute("Semester").Value == "2")
                    {
                        mainRow.Cells[12].Tag = element;
                        mainRow.Cells[12].Value = element.Attribute("Credit").Value;
                        mainRow.Cells[12].Style.BackColor = courseGroupColor;
                    }
                }
                _MainRowList.Add(mainRow);

                // 課程群組表格用
                DataGridViewRow courseGroupRow = new DataGridViewRow();
                courseGroupRow.CreateCells(dgvCourseGroup);
                courseGroupRow.Tag = courseInfo;
                courseGroupRow.Cells[0].Value = courseInfo.RequiredBy;
                courseGroupRow.Cells[1].Value = courseInfo.Required;
                courseGroupRow.Cells[2].Value = courseInfo.SpecialCategory;
                courseGroupRow.Cells[3].Value = courseInfo.SubjectAttribute;
                courseGroupRow.Cells[4].Value = courseInfo.Entry;
                courseGroupRow.Cells[5].Value = courseInfo.DomainName;
                courseGroupRow.Cells[6].Value = courseInfo.OfficialSubjectName;
                foreach (XElement element in courseInfo.CourseContentList)
                {
                    string credit = element.Attribute("Credit").Value;
                    string courseGroupName = element.Attribute("分組名稱") == null ? "" : element.Attribute("分組名稱").Value;
                    Color courseGroupColor = Color.White;

                    // 讀取群組設定
                    if (_CourseGroupSettingList.Where(x => x.CourseGroupName == courseGroupName).Count() > 0)
                    {
                        CourseGroupSetting courseGroupSetting = _CourseGroupSettingList.Where(x => x.CourseGroupName == courseGroupName).ToList()[0];
                        courseGroupColor = courseGroupSetting.CourseGroupColor;
                    }

                    if (element.Attribute("GradeYear").Value == "1" && element.Attribute("Semester").Value == "1")
                    {
                        courseGroupRow.Cells[7].Tag = element;
                        courseGroupRow.Cells[7].Value = credit;
                        courseGroupRow.Cells[7].Style.BackColor = courseGroupColor;

                        // 依據群組設定分類至CourseGroupSettingDic
                        if (_CourseGroupSettingDic.ContainsKey(courseGroupName))
                        {
                            _CourseGroupSettingDic[courseGroupName].Add(courseGroupRow.Cells[7]);
                        }
                    }
                    if (element.Attribute("GradeYear").Value == "1" && element.Attribute("Semester").Value == "2")
                    {
                        courseGroupRow.Cells[8].Tag = element;
                        courseGroupRow.Cells[8].Value = credit;
                        courseGroupRow.Cells[8].Style.BackColor = courseGroupColor;

                        // 依據群組設定分類至CourseGroupSettingDic
                        if (_CourseGroupSettingDic.ContainsKey(courseGroupName))
                        {
                            _CourseGroupSettingDic[courseGroupName].Add(courseGroupRow.Cells[8]);
                        }
                    }
                    if (element.Attribute("GradeYear").Value == "2" && element.Attribute("Semester").Value == "1")
                    {
                        courseGroupRow.Cells[9].Tag = element;
                        courseGroupRow.Cells[9].Value = credit;
                        courseGroupRow.Cells[9].Style.BackColor = courseGroupColor;

                        // 依據群組設定分類至CourseGroupSettingDic
                        if (_CourseGroupSettingDic.ContainsKey(courseGroupName))
                        {
                            _CourseGroupSettingDic[courseGroupName].Add(courseGroupRow.Cells[9]);
                        }
                    }
                    if (element.Attribute("GradeYear").Value == "2" && element.Attribute("Semester").Value == "2")
                    {
                        courseGroupRow.Cells[10].Tag = element;
                        courseGroupRow.Cells[10].Value = credit;
                        courseGroupRow.Cells[10].Style.BackColor = courseGroupColor;

                        // 依據群組設定分類至CourseGroupSettingDic
                        if (_CourseGroupSettingDic.ContainsKey(courseGroupName))
                        {
                            _CourseGroupSettingDic[courseGroupName].Add(courseGroupRow.Cells[10]);
                        }
                    }
                    if (element.Attribute("GradeYear").Value == "3" && element.Attribute("Semester").Value == "1")
                    {
                        courseGroupRow.Cells[11].Tag = element;
                        courseGroupRow.Cells[11].Value = credit;
                        courseGroupRow.Cells[11].Style.BackColor = courseGroupColor;

                        // 依據群組設定分類至CourseGroupSettingDic
                        if (_CourseGroupSettingDic.ContainsKey(courseGroupName))
                        {
                            _CourseGroupSettingDic[courseGroupName].Add(courseGroupRow.Cells[11]);
                        }
                    }
                    if (element.Attribute("GradeYear").Value == "3" && element.Attribute("Semester").Value == "2")
                    {
                        courseGroupRow.Cells[12].Tag = element;
                        courseGroupRow.Cells[12].Value = credit;
                        courseGroupRow.Cells[12].Style.BackColor = courseGroupColor;

                        // 依據群組設定分類至CourseGroupSettingDic
                        if (_CourseGroupSettingDic.ContainsKey(courseGroupName))
                        {
                            _CourseGroupSettingDic[courseGroupName].Add(courseGroupRow.Cells[12]);
                        }
                    }
                }
                _CourseGroupRowList.Add(courseGroupRow);
            }
            #endregion

            _MainIsFirstLoad = true;
            LoadMainComboBoxData(null, null);
            _CourseGroupIsFirstLoad = true;
            LoadCourseGroupComboBoxData(null, null);
            LoadCourseGroupSettingDataGridView();

            lbCacluateCredit_LinkClicked(null, null); // 計算學分

            btnUpdate.Enabled = false;
            isDgDataChange = false;
            isUDDgDataChange = false;
            _IsMainDataDirty = false;
            _IsCourseGroupDataDirty = false;
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
            if (isDgDataChange || isUDDgDataChange || _IsMainDataDirty || _IsCourseGroupDataDirty)
            {
                if (DialogResult.No == MsgBox.Show("變更尚未儲存，確定離開？", MessageBoxButtons.YesNo))
                {
                    return;
                }
                isDgDataChange = isUDDgDataChange = _IsMainDataDirty = _IsCourseGroupDataDirty = false;
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
                advTree1_NodeClick(_SelectItem, null);

                ApplicationLog.Log("班級課程規劃表(108課綱適用)", "修改", $"修改 「{SelectInfo.RefGPName}」 課程規畫表資訊。");
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
            if (isDgDataChange || isUDDgDataChange || _IsMainDataDirty || _IsCourseGroupDataDirty)
            {
                if (DialogResult.No == MsgBox.Show("變更尚未儲存，確定離開？", MessageBoxButtons.YesNo))
                {
                    e.Cancel = true;
                    return;
                }
                isDgDataChange = isUDDgDataChange = _IsMainDataDirty = _IsCourseGroupDataDirty = false;
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
            if (SelectInfo != null)
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

                // 課程群組總表
                if (tabControl1.SelectedTabIndex == 4)
                {
                    lbMainGraduationPlanName.Text = SelectInfo.RefGPName + (isD ? " (<font color=\"Chocolate\">已變更</font>)" : "");
                    btnUpdate.Enabled = isD;
                    _IsMainDataDirty = isD;
                }

                // 設定課程群組
                if (tabControl1.SelectedTabIndex == 5)
                {
                    lbCourseGroupGraduationPlanName.Text = SelectInfo.RefGPName + (isD ? " (<font color=\"Chocolate\">已變更</font>)" : "");
                    btnUpdate.Enabled = isD;
                    _IsMainDataDirty = isD;
                }
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

            isUDRowSelect = true;

        }

        private void dgUDData_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (isUDRowSelect)
            {
                SetIsDirtyDisplay(true);
                btnUpdate.Enabled = true;
                isUDRowSelect = false;
            }
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

        private void tabControl1_SelectedTabChanged(object sender, TabStripTabChangedEventArgs e)
        {
            _SelectedTab = tabControl1.SelectedTab;
        }

        // 總表用事件
        private void LoadMainComboBoxData(object sender, EventArgs args)
        {
            if (_MainIsLoading == true)
            {
                return;
            }
            _MainIsLoading = true;
            this.SuspendLayout();

            if (_MainIsFirstLoad)
            {
                // 第一次載入，所有下拉式清單刷新
                // 初始化下拉式清單
                cboMainRequiredBy.Items.Clear();
                cboMainRequiredBy.Items.Add("全部");
                cboMainRequiredBy.SelectedIndex = 0;
                cboMainRequired.Items.Clear();
                cboMainRequired.Items.Add("全部");
                cboMainRequired.SelectedIndex = 0;
                cboMainSpecialCategory.Items.Clear();
                cboMainSpecialCategory.Items.Add("全部");
                cboMainSpecialCategory.SelectedIndex = 0;
                cboMainSubjectAttribute.Items.Clear();
                cboMainSubjectAttribute.Items.Add("全部");
                cboMainSubjectAttribute.SelectedIndex = 0;
                cboMainEntry.Items.Clear();
                cboMainEntry.Items.Add("全部");
                cboMainEntry.SelectedIndex = 0;
                cboMainDomainName.Items.Clear();
                cboMainDomainName.Items.Add("全部");
                cboMainDomainName.SelectedIndex = 0;
                foreach (GPlanCourseInfo108 courseInfo in _CourseInfoList)
                {
                    if (!cboMainRequiredBy.Items.Contains(courseInfo.RequiredBy))
                    {
                        cboMainRequiredBy.Items.Add(courseInfo.RequiredBy);
                    }
                    if (!cboMainRequired.Items.Contains(courseInfo.Required))
                    {
                        cboMainRequired.Items.Add(courseInfo.Required);
                    }
                    if (!cboMainSpecialCategory.Items.Contains(courseInfo.SpecialCategory))
                    {
                        cboMainSpecialCategory.Items.Add(courseInfo.SpecialCategory);
                    }
                    if (!cboMainSubjectAttribute.Items.Contains(courseInfo.SubjectAttribute))
                    {
                        cboMainSubjectAttribute.Items.Add(courseInfo.SubjectAttribute);
                    }
                    if (!cboMainEntry.Items.Contains(courseInfo.Entry))
                    {
                        cboMainEntry.Items.Add(courseInfo.Entry);
                    }
                    if (!cboMainDomainName.Items.Contains(courseInfo.DomainName))
                    {
                        cboMainDomainName.Items.Add(courseInfo.DomainName);
                    }
                }

                _MainIsFirstLoad = false;
            }
            else
            {
                // 非第一次載入代表是篩選觸發事件
                List<GPlanCourseInfo108> filterList = new List<GPlanCourseInfo108>();

                foreach (GPlanCourseInfo108 courseInfo in _CourseInfoList)
                {
                    bool show = true;
                    if (cboMainRequiredBy.Text != "" && cboMainRequiredBy.Text != "全部" && cboMainRequiredBy.Text != courseInfo.RequiredBy)
                        show = false;
                    if (cboMainRequired.Text != "" && cboMainRequired.Text != "全部" && cboMainRequired.Text != courseInfo.Required)
                        show = false;
                    if (cboMainSpecialCategory.Text != "" && cboMainSpecialCategory.Text != "全部" && cboMainSpecialCategory.Text != courseInfo.SpecialCategory)
                        show = false;
                    if (cboMainSubjectAttribute.Text != "" && cboMainSubjectAttribute.Text != "全部" && cboMainSubjectAttribute.Text != courseInfo.SubjectAttribute)
                        show = false;
                    if (cboMainEntry.Text != "" && cboMainEntry.Text != "全部" && cboMainEntry.Text != courseInfo.Entry)
                        show = false;
                    if (cboMainDomainName.Text != "" && cboMainDomainName.Text != "全部" && cboMainDomainName.Text != courseInfo.DomainName)
                        show = false;
                    if (show)
                        filterList.Add(courseInfo);
                }

                if (cboMainRequiredBy.Text == "全部")
                {
                    cboMainRequiredBy.Items.Clear();
                    cboMainRequiredBy.Items.Add("全部");
                    cboMainRequiredBy.Items.AddRange(filterList.Select(x => x.RequiredBy).Distinct().ToArray());
                    cboMainRequiredBy.SelectedIndex = 0;
                }

                if (cboMainRequired.Text == "全部")
                {
                    cboMainRequired.Items.Clear();
                    cboMainRequired.Items.Add("全部");
                    cboMainRequired.Items.AddRange(filterList.Select(x => x.Required).Distinct().ToArray());
                    cboMainRequired.SelectedIndex = 0;
                }

                if (cboMainSpecialCategory.Text == "全部")
                {
                    cboMainSpecialCategory.Items.Clear();
                    cboMainSpecialCategory.Items.Add("全部");
                    cboMainSpecialCategory.Items.AddRange(filterList.Select(x => x.SpecialCategory).Distinct().ToArray());
                    cboMainSpecialCategory.SelectedIndex = 0;
                }

                if (cboMainSubjectAttribute.Text == "全部")
                {
                    cboMainSubjectAttribute.Items.Clear();
                    cboMainSubjectAttribute.Items.Add("全部");
                    cboMainSubjectAttribute.Items.AddRange(filterList.Select(x => x.SubjectAttribute).Distinct().ToArray());
                    cboMainSubjectAttribute.SelectedIndex = 0;
                }

                if (cboMainEntry.Text == "全部")
                {
                    cboMainEntry.Items.Clear();
                    cboMainEntry.Items.Add("全部");
                    cboMainEntry.Items.AddRange(filterList.Select(x => x.Entry).Distinct().ToArray());
                    cboMainEntry.SelectedIndex = 0;
                }

                if (cboMainDomainName.Text == "全部")
                {
                    cboMainDomainName.Items.Clear();
                    cboMainDomainName.Items.Add("全部");
                    cboMainDomainName.Items.AddRange(filterList.Select(x => x.DomainName).Distinct().ToArray());
                    cboMainDomainName.SelectedIndex = 0;
                }

            }

            this.ResumeLayout();
            _MainIsLoading = false;
            LoadMainDataGridViewData();
        }

        private void LoadMainDataGridViewData()
        {
            _MainIsLoading = true;
            this.SuspendLayout();
            dgvMain.Rows.Clear();

            List<DataGridViewRow> filterRowList = new List<DataGridViewRow>();

            foreach (DataGridViewRow row in _MainRowList)
            {
                bool show = true;
                if (cboMainRequiredBy.Text != "" && cboMainRequiredBy.Text != "全部" && cboMainRequiredBy.Text != row.Cells[0].Value.ToString())
                    show = false;
                if (cboMainRequired.Text != "" && cboMainRequired.Text != "全部" && cboMainRequired.Text != row.Cells[1].Value.ToString())
                    show = false;
                if (cboMainSpecialCategory.Text != "" && cboMainSpecialCategory.Text != "全部" && cboMainSpecialCategory.Text != row.Cells[2].Value.ToString())
                    show = false;
                if (cboMainSubjectAttribute.Text != "" && cboMainSubjectAttribute.Text != "全部" && cboMainSubjectAttribute.Text != row.Cells[3].Value.ToString())
                    show = false;
                if (cboMainEntry.Text != "" && cboMainEntry.Text != "全部" && cboMainEntry.Text != row.Cells[4].Value.ToString())
                    show = false;
                if (cboMainDomainName.Text != "" && cboMainDomainName.Text != "全部" && cboMainDomainName.Text != row.Cells[5].Value.ToString())
                    show = false;
                if (show)
                    filterRowList.Add(row);
            }
            dgvMain.Rows.AddRange(filterRowList.ToArray());
            this.ResumeLayout();
            _MainIsLoading = false;
        }

        private void dgvMain_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvMain.SelectedRows.Count > 0)
            {
                DataGridViewRow row = dgvMain.SelectedRows[0];

                if (row.Tag != null)
                {
                    GPlanCourseInfo108 courseInfo = (GPlanCourseInfo108)row.Tag;
                    tbMainStartLevel.Text = courseInfo.StartLevel;
                    tbMainSubjectName.Text = courseInfo.SubjectName;
                    tbMainSchoolYearGroupName.Text = courseInfo.SchoolYearGroupName;
                    _MainSelectedRow = row;
                }
            }
        }

        private void btnMainSet_Click(object sender, EventArgs e)
        {
            if (_MainSelectedRow.Tag != null)
            {
                GPlanCourseInfo108 courseInfo = (GPlanCourseInfo108)_MainSelectedRow.Tag;

                if (_CourseInfoList.Where(x => x.CourseCode != courseInfo.CourseCode && x.SubjectName == tbMainSubjectName.Text).Count() > 0)
                {
                    MessageBox.Show("科目名稱不可重複");
                    return;
                }

                int index = 0;
                string rowIndex = courseInfo.CourseContentList[0].Element("Grouping").Attribute("RowIndex") == null ? "" : courseInfo.CourseContentList[0].Element("Grouping").Attribute("RowIndex").Value;
                if (Int32.TryParse(rowIndex, out index))
                {
                    index -= 1;
                    int startLevel = 0;
                    if (Int32.TryParse(tbMainStartLevel.Text, out startLevel))
                    {
                        _MainRowList[index].Cells[13].Value = tbMainStartLevel.Text;
                        _MainRowList[index].Cells[14].Value = tbMainSubjectName.Text;
                        _MainRowList[index].Cells[15].Value = tbMainSchoolYearGroupName.Text;
                        courseInfo.StartLevel = tbMainStartLevel.Text;
                        courseInfo.SubjectName = tbMainSubjectName.Text;
                        courseInfo.SchoolYearGroupName = tbMainSchoolYearGroupName.Text;
                        int level = startLevel;
                        foreach (XElement element in courseInfo.CourseContentList)
                        {
                            element.Element("Grouping").SetAttributeValue("startLevel", tbMainStartLevel.Text);
                            element.SetAttributeValue("Level", level);
                            element.SetAttributeValue("SubjectName", tbMainSubjectName.Text);
                            element.SetAttributeValue("指定學年科目名稱", tbMainSchoolYearGroupName.Text);
                            level++;
                        }

                        LoadMainDataGridViewData();

                        dgvMain.Rows[index].Selected = true;
                        SetIsDirtyDisplay(true);
                    }
                }
            }
        }

        private void btnMainCancel_Click(object sender, EventArgs e)
        {
            GPlanCourseInfo108 courseInfo = (GPlanCourseInfo108)_MainSelectedRow.Tag;
            tbMainStartLevel.Text = courseInfo.StartLevel;
            tbMainSubjectName.Text = courseInfo.SubjectName;
            tbMainSchoolYearGroupName.Text = courseInfo.SchoolYearGroupName;
        }

        // 課程群組用事件
        //// 篩選
        private void LoadCourseGroupComboBoxData(object sender, EventArgs e)
        {
            if (_CourseGroupIsLoading == true)
            {
                return;
            }
            _CourseGroupIsLoading = true;
            this.SuspendLayout();

            if (_CourseGroupIsFirstLoad)
            {
                // 第一次載入，所有下拉式清單刷新
                // 初始化下拉式清單
                cboCourseGroupRequiredBy.Items.Clear();
                cboCourseGroupRequiredBy.Items.Add("全部");
                cboCourseGroupRequiredBy.SelectedIndex = 0;
                cboCourseGroupRequired.Items.Clear();
                cboCourseGroupRequired.Items.Add("全部");
                cboCourseGroupRequired.SelectedIndex = 0;
                cboCourseGroupSpecialCategory.Items.Clear();
                cboCourseGroupSpecialCategory.Items.Add("全部");
                cboCourseGroupSpecialCategory.SelectedIndex = 0;
                cboCourseGroupSubjectAttribute.Items.Clear();
                cboCourseGroupSubjectAttribute.Items.Add("全部");
                cboCourseGroupSubjectAttribute.SelectedIndex = 0;
                cboCourseGroupEntry.Items.Clear();
                cboCourseGroupEntry.Items.Add("全部");
                cboCourseGroupEntry.SelectedIndex = 0;
                cboCourseGroupDomainName.Items.Clear();
                cboCourseGroupDomainName.Items.Add("全部");
                cboCourseGroupDomainName.SelectedIndex = 0;
                foreach (GPlanCourseInfo108 courseInfo in _CourseInfoList)
                {
                    if (!cboCourseGroupRequiredBy.Items.Contains(courseInfo.RequiredBy))
                    {
                        cboCourseGroupRequiredBy.Items.Add(courseInfo.RequiredBy);
                    }
                    if (!cboCourseGroupRequired.Items.Contains(courseInfo.Required))
                    {
                        cboCourseGroupRequired.Items.Add(courseInfo.Required);
                    }
                    if (!cboCourseGroupSpecialCategory.Items.Contains(courseInfo.SpecialCategory))
                    {
                        cboCourseGroupSpecialCategory.Items.Add(courseInfo.SpecialCategory);
                    }
                    if (!cboCourseGroupSubjectAttribute.Items.Contains(courseInfo.SubjectAttribute))
                    {
                        cboCourseGroupSubjectAttribute.Items.Add(courseInfo.SubjectAttribute);
                    }
                    if (!cboCourseGroupEntry.Items.Contains(courseInfo.Entry))
                    {
                        cboCourseGroupEntry.Items.Add(courseInfo.Entry);
                    }
                    if (!cboCourseGroupDomainName.Items.Contains(courseInfo.DomainName))
                    {
                        cboCourseGroupDomainName.Items.Add(courseInfo.DomainName);
                    }
                }

                _CourseGroupIsFirstLoad = false;
            }
            else
            {
                // 非第一次載入代表是篩選觸發事件
                List<GPlanCourseInfo108> filterList = new List<GPlanCourseInfo108>();

                foreach (GPlanCourseInfo108 courseInfo in _CourseInfoList)
                {
                    bool show = true;
                    if (cboCourseGroupRequiredBy.Text != "" && cboCourseGroupRequiredBy.Text != "全部" && cboCourseGroupRequiredBy.Text != courseInfo.RequiredBy)
                        show = false;
                    if (cboCourseGroupRequired.Text != "" && cboCourseGroupRequired.Text != "全部" && cboCourseGroupRequired.Text != courseInfo.Required)
                        show = false;
                    if (cboCourseGroupSpecialCategory.Text != "" && cboCourseGroupSpecialCategory.Text != "全部" && cboCourseGroupSpecialCategory.Text != courseInfo.SpecialCategory)
                        show = false;
                    if (cboCourseGroupSubjectAttribute.Text != "" && cboCourseGroupSubjectAttribute.Text != "全部" && cboCourseGroupSubjectAttribute.Text != courseInfo.SubjectAttribute)
                        show = false;
                    if (cboCourseGroupEntry.Text != "" && cboCourseGroupEntry.Text != "全部" && cboCourseGroupEntry.Text != courseInfo.Entry)
                        show = false;
                    if (cboCourseGroupDomainName.Text != "" && cboCourseGroupDomainName.Text != "全部" && cboCourseGroupDomainName.Text != courseInfo.DomainName)
                        show = false;
                    if (show)
                        filterList.Add(courseInfo);
                }

                if (cboCourseGroupRequiredBy.Text == "全部")
                {
                    cboCourseGroupRequiredBy.Items.Clear();
                    cboCourseGroupRequiredBy.Items.Add("全部");
                    cboCourseGroupRequiredBy.Items.AddRange(filterList.Select(x => x.RequiredBy).Distinct().ToArray());
                    cboCourseGroupRequiredBy.SelectedIndex = 0;
                }

                if (cboCourseGroupRequired.Text == "全部")
                {
                    cboCourseGroupRequired.Items.Clear();
                    cboCourseGroupRequired.Items.Add("全部");
                    cboCourseGroupRequired.Items.AddRange(filterList.Select(x => x.Required).Distinct().ToArray());
                    cboCourseGroupRequired.SelectedIndex = 0;
                }

                if (cboCourseGroupSpecialCategory.Text == "全部")
                {
                    cboCourseGroupSpecialCategory.Items.Clear();
                    cboCourseGroupSpecialCategory.Items.Add("全部");
                    cboCourseGroupSpecialCategory.Items.AddRange(filterList.Select(x => x.SpecialCategory).Distinct().ToArray());
                    cboCourseGroupSpecialCategory.SelectedIndex = 0;
                }

                if (cboCourseGroupSubjectAttribute.Text == "全部")
                {
                    cboCourseGroupSubjectAttribute.Items.Clear();
                    cboCourseGroupSubjectAttribute.Items.Add("全部");
                    cboCourseGroupSubjectAttribute.Items.AddRange(filterList.Select(x => x.SubjectAttribute).Distinct().ToArray());
                    cboCourseGroupSubjectAttribute.SelectedIndex = 0;
                }

                if (cboCourseGroupEntry.Text == "全部")
                {
                    cboCourseGroupEntry.Items.Clear();
                    cboCourseGroupEntry.Items.Add("全部");
                    cboCourseGroupEntry.Items.AddRange(filterList.Select(x => x.Entry).Distinct().ToArray());
                    cboCourseGroupEntry.SelectedIndex = 0;
                }

                if (cboCourseGroupDomainName.Text == "全部")
                {
                    cboCourseGroupDomainName.Items.Clear();
                    cboCourseGroupDomainName.Items.Add("全部");
                    cboCourseGroupDomainName.Items.AddRange(filterList.Select(x => x.DomainName).Distinct().ToArray());
                    cboCourseGroupDomainName.SelectedIndex = 0;
                }

            }

            this.ResumeLayout();
            _CourseGroupIsLoading = false;
            LoadCourseGroupDataGridViewData();
        }

        private void LoadCourseGroupDataGridViewData()
        {
            _CourseGroupIsLoading = true;
            this.SuspendLayout();
            dgvCourseGroup.Rows.Clear();

            List<DataGridViewRow> filterRowList = new List<DataGridViewRow>();

            foreach (DataGridViewRow row in _CourseGroupRowList)
            {
                bool show = true;
                if (cboCourseGroupRequiredBy.Text != "" && cboCourseGroupRequiredBy.Text != "全部" && cboCourseGroupRequiredBy.Text != row.Cells[0].Value.ToString())
                    show = false;
                if (cboCourseGroupRequired.Text != "" && cboCourseGroupRequired.Text != "全部" && cboCourseGroupRequired.Text != row.Cells[1].Value.ToString())
                    show = false;
                if (cboCourseGroupSpecialCategory.Text != "" && cboCourseGroupSpecialCategory.Text != "全部" && cboCourseGroupSpecialCategory.Text != row.Cells[2].Value.ToString())
                    show = false;
                if (cboCourseGroupSubjectAttribute.Text != "" && cboCourseGroupSubjectAttribute.Text != "全部" && cboCourseGroupSubjectAttribute.Text != row.Cells[3].Value.ToString())
                    show = false;
                if (cboCourseGroupEntry.Text != "" && cboCourseGroupEntry.Text != "全部" && cboCourseGroupEntry.Text != row.Cells[4].Value.ToString())
                    show = false;
                if (cboCourseGroupDomainName.Text != "" && cboCourseGroupDomainName.Text != "全部" && cboCourseGroupDomainName.Text != row.Cells[5].Value.ToString())
                    show = false;
                if (show)
                    filterRowList.Add(row);
            }
            dgvCourseGroup.Rows.AddRange(filterRowList.ToArray());
            this.ResumeLayout();
            _CourseGroupIsLoading = false;
        }

        //// 群組管理 
        private void LoadCourseGroupSettingDataGridView()
        {
            dgvCourseGroupManageGroup.Rows.Clear();

            foreach (CourseGroupSetting setting in _CourseGroupSettingList)
            {
                // 宣告16*16點陣圖
                Bitmap bitmap = new Bitmap(16, 16);

                using (Graphics gx = Graphics.FromImage(bitmap))
                {
                    Rectangle rect = new Rectangle(2, 2, 13, 13); // 外框
                    Rectangle rectSmall = new Rectangle(3, 3, 12, 12); // 填滿

                    SolidBrush brush = new SolidBrush(setting.CourseGroupColor); // 設定填滿色彩
                    gx.FillRectangle(brush, rectSmall);
                    gx.DrawRectangle(Pens.Black, rect);
                }

                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgvCourseGroupManageGroup);
                row.Tag = setting;
                row.Cells[0].Value = bitmap;
                row.Cells[1].Value = setting.CourseGroupName;
                row.Cells[2].Value = setting.CourseGroupCredit;
                row.Cells[3].Value = setting.IsSchoolYearCourseGroup;
                dgvCourseGroupManageGroup.Rows.Add(row);
            }
        }

        private void lbCopyCourseGroupSetting_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmCopyCourseGroupSetting frm = new frmCopyCourseGroupSetting(GP108List, SelectInfo);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                // 讀取課程群組設定
                _CourseGroupSettingList.Clear();
                if (SelectInfo.RefGPContentXml.Element("CourseGroupSetting") != null)
                {
                    foreach (XElement element in SelectInfo.RefGPContentXml.Element("CourseGroupSetting").Elements("CourseGroup"))
                    {
                        CourseGroupSetting courseGroupSetting = new CourseGroupSetting();
                        courseGroupSetting.CourseGroupName = element.Attribute("Name").Value;
                        courseGroupSetting.CourseGroupCredit = element.Attribute("Credit").Value;
                        courseGroupSetting.CourseGroupColor = Color.FromArgb(Int32.Parse(element.Attribute("Color").Value));
                        courseGroupSetting.IsSchoolYearCourseGroup = element.Attribute("IsSchoolYearCourseGroup") == null ? false : bool.Parse(element.Attribute("IsSchoolYearCourseGroup").Value);
                        courseGroupSetting.CourseGroupElement = element;
                        _CourseGroupSettingList.Add(courseGroupSetting);

                        if (!_CourseGroupSettingDic.ContainsKey(courseGroupSetting.CourseGroupName))
                        {
                            _CourseGroupSettingDic.Add(courseGroupSetting.CourseGroupName, new List<DataGridViewCell>());
                        }
                    }
                }

                LoadCourseGroupSettingDataGridView();
                SetIsDirtyDisplay(true);
            }
        }

        private void dgvCourseGroupManageGroup_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // 結束編輯，一併修改已經群組的課程
            DataGridViewRow row = dgvCourseGroupManageGroup.Rows[e.RowIndex];
            CourseGroupSetting setting = (CourseGroupSetting)row.Tag;

            switch (e.ColumnIndex)
            {
                case 1:
                    // 先將舊的群組另外存起來
                    string oldCourseGroupName = setting.CourseGroupName;
                    string newCourseGroupName = row.Cells[e.ColumnIndex].Value.ToString();

                    if (newCourseGroupName == oldCourseGroupName)
                    {
                        return;
                    }

                    if (_CourseGroupSettingList.Where(x => x.CourseGroupName == newCourseGroupName).Count() > 0)
                    {
                        row.Cells[e.ColumnIndex].Value = oldCourseGroupName;
                        MessageBox.Show("課程群組名稱不可重複");
                        return;
                    }

                    setting.CourseGroupName = newCourseGroupName;
                    setting.CourseGroupElement.SetAttributeValue("Name", newCourseGroupName);
                    List<DataGridViewCell> cellList = new List<DataGridViewCell>();

                    // 處理CourseGroupSetting這個節點
                    foreach (XElement element in SelectInfo.RefGPContentXml.Element("CourseGroupSetting").Elements("CourseGroup").Where(x => x.Attribute("Name").Value == oldCourseGroupName).ToList())
                    {
                        element.SetAttributeValue("Name", newCourseGroupName);
                    }

                    // 處理名稱對應課程的Dictionary
                    if (_CourseGroupSettingDic.ContainsKey(oldCourseGroupName))
                    {
                        cellList = _CourseGroupSettingDic[oldCourseGroupName];

                        // 移除舊的
                        _CourseGroupSettingDic.Remove(oldCourseGroupName);
                    }

                    // 依照新的名稱建立字典並加入原本的課程
                    if (!_CourseGroupSettingDic.ContainsKey(newCourseGroupName))
                    {
                        _CourseGroupSettingDic.Add(newCourseGroupName, cellList);
                    }
                    else
                    {
                        _CourseGroupSettingDic[newCourseGroupName].Clear();
                        _CourseGroupSettingDic[newCourseGroupName].AddRange(cellList);
                    }

                    foreach (DataGridViewCell cell in _CourseGroupSettingDic[newCourseGroupName])
                    {
                        XElement element = (XElement)cell.Tag;
                        element.SetAttributeValue("分組名稱", newCourseGroupName);
                    }
                    SetIsDirtyDisplay(true);
                    break;

                case 2:
                    string credit = row.Cells[e.ColumnIndex].Value.ToString();
                    string courseGroupName = setting.CourseGroupName;

                    if (credit == setting.CourseGroupCredit)
                    {
                        return;
                    }

                    setting.CourseGroupCredit = credit;
                    setting.CourseGroupElement.SetAttributeValue("Credit", credit);

                    // 處理CourseGroupSetting這個節點
                    foreach (XElement element in SelectInfo.RefGPContentXml.Element("CourseGroupSetting").Elements("CourseGroup").Where(x => x.Attribute("Name").Value == courseGroupName).ToList())
                    {
                        element.SetAttributeValue("Credit", credit);
                    }

                    // 處理名稱對應課程的Dictionary
                    foreach (DataGridViewCell cell in _CourseGroupSettingDic[courseGroupName])
                    {
                        XElement element = (XElement)cell.Tag;
                        element.SetAttributeValue("分組修課學分數", credit);
                    }

                    SetIsDirtyDisplay(true);
                    break;

                default:
                    break;
            }
        }

        private void dgvCourseGroupManageGroup_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex == 3)
            {
                dgvCourseGroupManageGroup.EndEdit();
            }
        }

        private void dgvCourseGroupManageGroup_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex == 3)
            {
                DataGridViewRow row = dgvCourseGroupManageGroup.Rows[e.RowIndex];
                CourseGroupSetting setting = (CourseGroupSetting)row.Tag;
                DataGridViewCheckBoxCell cell = (DataGridViewCheckBoxCell)row.Cells[e.ColumnIndex];
                string courseGroupName = setting.CourseGroupName;
                bool isSchoolYearCourseGroup = (bool)cell.Value;

                if (isSchoolYearCourseGroup == setting.IsSchoolYearCourseGroup)
                {
                    return;
                }

                setting.IsSchoolYearCourseGroup = isSchoolYearCourseGroup;
                setting.CourseGroupElement.SetAttributeValue("IsSchoolYearCourseGroup", isSchoolYearCourseGroup);

                // 處理CourseGroupSetting這個節點
                foreach (XElement element in SelectInfo.RefGPContentXml.Element("CourseGroupSetting").Elements("CourseGroup").Where(x => x.Attribute("Name").Value == courseGroupName).ToList())
                {
                    element.SetAttributeValue("IsSchoolYearCourseGroup", isSchoolYearCourseGroup);
                }

                SetIsDirtyDisplay(true);
            }
        }

        private void lbInsertCourseGroup_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmInsertCourseGroup frm = new frmInsertCourseGroup(_CourseGroupSettingList);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (SelectInfo.RefGPContentXml.Element("CourseGroupSetting") == null)
                {
                    XElement element = new XElement("CourseGroupSetting");
                    SelectInfo.RefGPContentXml.Add(element);
                }

                // 編輯SelectInfo中的群組設定
                XElement courseGroupElement = SelectInfo.RefGPContentXml.Element("CourseGroupSetting");

                // 新增的群組，index應為最後一筆
                CourseGroupSetting newCourseGroupSetting = _CourseGroupSettingList.Last();

                XElement newCourseGroupElement = new XElement("CourseGroup");
                newCourseGroupElement.SetAttributeValue("Name", newCourseGroupSetting.CourseGroupName);
                newCourseGroupElement.SetAttributeValue("Credit", newCourseGroupSetting.CourseGroupCredit);
                newCourseGroupElement.SetAttributeValue("Color", newCourseGroupSetting.CourseGroupColor.ToArgb());
                newCourseGroupElement.SetAttributeValue("IsSchoolYearCourseGroup", newCourseGroupSetting.IsSchoolYearCourseGroup);
                courseGroupElement.Add(newCourseGroupElement);

                if (!_CourseGroupSettingDic.ContainsKey(newCourseGroupSetting.CourseGroupName))
                {
                    _CourseGroupSettingDic.Add(newCourseGroupSetting.CourseGroupName, new List<DataGridViewCell>());
                }

                LoadCourseGroupSettingDataGridView();

                // 選取最新一筆
                foreach (DataGridViewRow row in dgvCourseGroupManageGroup.SelectedRows)
                {
                    row.Selected = false;
                }
                dgvCourseGroupManageGroup.Rows[dgvCourseGroupManageGroup.Rows.Count - 1].Selected = true;

                SetIsDirtyDisplay(true);
            }
        }

        private void dgvCourseGroupManageGroup_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvCourseGroupManageGroup.Rows.Count > 0)
            {
                foreach (DataGridViewRow row in dgvCourseGroupManageGroup.SelectedRows)
                {
                    this.SuspendLayout();

                    foreach (DataGridViewCell cell in _CourseGroupFocusCellList)
                    {
                        cell.Style.Font = new Font("Microsoft JhengHei", 10, FontStyle.Regular);
                        cell.Style.ForeColor = Color.Black;
                    }
                    _CourseGroupFocusCellList.Clear();

                    string courseGroupName = row.Cells[1].Value.ToString();
                    if (_CourseGroupSettingDic.ContainsKey(courseGroupName))
                    {
                        foreach (DataGridViewCell cell in _CourseGroupSettingDic[courseGroupName])
                        {
                            cell.Style.Font = new Font("Microsoft JhengHei", 10, FontStyle.Bold);
                            cell.Style.ForeColor = Color.Red;
                            _CourseGroupFocusCellList.Add(cell);
                        }
                    }

                    this.ResumeLayout();
                }
            }
        }

        private void lbDeleteCourseGroup_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            foreach (DataGridViewRow row in dgvCourseGroupManageGroup.SelectedRows)
            {
                CourseGroupSetting setting = _CourseGroupSettingList[row.Index];

                // 刪除設定時，一併清除已經設定的群組
                foreach (DataGridViewCell cell in _CourseGroupSettingDic[setting.CourseGroupName])
                {
                    XElement element = (XElement)cell.Tag;
                    element.SetAttributeValue("分組名稱", "");
                    element.SetAttributeValue("分組修課學分數", "");
                    cell.Style.BackColor = Color.White;
                }
                _CourseGroupSettingDic.Remove(setting.CourseGroupName);

                _CourseGroupSettingList.Remove(setting);
            }

            // 編輯SelectInfo中的群組設定
            XElement courseGroupElement = SelectInfo.RefGPContentXml.Element("CourseGroupSetting");
            courseGroupElement.Elements("CourseGroup").Remove(); // 刪除現有Elements中的資料，重新讀取dataGridView並寫入資料

            foreach (CourseGroupSetting setting in _CourseGroupSettingList)
            {
                XElement element = new XElement("CourseGroup");
                element.SetAttributeValue("Name", setting.CourseGroupName);
                element.SetAttributeValue("Credit", setting.CourseGroupCredit);
                element.SetAttributeValue("Color", setting.CourseGroupColor.ToArgb());
                element.SetAttributeValue("IsSchoolYearCourseGroup", setting.IsSchoolYearCourseGroup);
                courseGroupElement.Add(element);

                if (!_CourseGroupSettingDic.ContainsKey(setting.CourseGroupName))
                {
                    _CourseGroupSettingDic.Add(setting.CourseGroupName, new List<DataGridViewCell>());
                }
            }

            LoadCourseGroupSettingDataGridView();
            SetIsDirtyDisplay(true);
        }

        //// 學分統計
        private void lbCacluateCredit_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 學分統計
            if (!string.IsNullOrEmpty(_CourseType))
            {
                if (_CourseType == "技術型高中" || _CourseType == "進修部" || _CourseType == "實用技能學程(日)" || _CourseType == "實用技能學程(夜)")
                {
                    tcSwitchCreditStatistics.SelectedTab = tbiCreditTechnical;
                    tbiCreditNormal.Visible = false;

                    // 學業部定必修
                    int index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> normalSubjectRequiredbyDepartmentList = _CourseGroupRowList.Where(
                        x => x.Cells[4].Value.ToString() == tbNormalSubject.Text
                        && x.Cells[0].Value.ToString() == tbNormalSubjectRequiredbyDepart.Text).ToList();
                    foreach (TextBoxX tb in _TechNormalSubjectRequiredbyDepartList)
                    {
                        int credit = normalSubjectRequiredbyDepartmentList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = normalSubjectRequiredbyDepartmentList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 學業校訂必修
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> normalSubjectRequiredbySchoolRequiredList = _CourseGroupRowList.Where(
                        x => x.Cells[4].Value.ToString() == tbNormalSubject.Text
                        && x.Cells[0].Value.ToString() == tbNormalSubjectRequiredbySchool.Text
                        && x.Cells[1].Value.ToString() == tbNormalSubjectRequiredbySchoolRequired.Text).ToList();
                    foreach (TextBoxX tb in _TechNormalSubjectRequiredbySchoolRequiredList)
                    {
                        int credit = normalSubjectRequiredbySchoolRequiredList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = normalSubjectRequiredbySchoolRequiredList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 學業校訂選修
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> normalSubjectRequiredbySchoolNonRequiredList = _CourseGroupRowList.Where(
                        x => x.Cells[4].Value.ToString() == tbNormalSubject.Text
                        && x.Cells[0].Value.ToString() == tbNormalSubjectRequiredbySchool.Text
                        && x.Cells[1].Value.ToString() == tbNormalSubjectRequiredbySchoolNonRequired.Text
                        && x.Cells[6].Value.ToString() != "團體活動時間"
                        && x.Cells[6].Value.ToString() != "彈性學習時間").ToList();
                    foreach (TextBoxX tb in _TechNormalSubjectRequiredbySchoolNonRequiredList)
                    {
                        int credit = normalSubjectRequiredbySchoolNonRequiredList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = normalSubjectRequiredbySchoolNonRequiredList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 部定專業科目
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> professionalSubjectRequiredByDepartProfessionalList = _CourseGroupRowList.Where(
                        x => x.Cells[0].Value.ToString() == tbProfessionalSubjectRequiredByDepart.Text
                        && x.Cells[4].Value.ToString() == tbProfessionalSubjectRequiredByDepartProfessional.Text).ToList();
                    foreach (TextBoxX tb in _TechProfessionalSubjectRequiredByDepartProfessionalList)
                    {
                        int credit = professionalSubjectRequiredByDepartProfessionalList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = professionalSubjectRequiredByDepartProfessionalList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 部定實習科目
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> professionalSubjectRequiredByDepartPracticeList = _CourseGroupRowList.Where(
                        x => x.Cells[0].Value.ToString() == tbProfessionalSubjectRequiredByDepart.Text
                        && x.Cells[4].Value.ToString() == tbProfessionalSubjectRequiredByDepartPractice.Text).ToList();
                    foreach (TextBoxX tb in _TechProfessionalSubjectRequiredByDepartPracticeList)
                    {
                        int credit = professionalSubjectRequiredByDepartPracticeList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = professionalSubjectRequiredByDepartPracticeList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 校訂專業科目必修
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> professionalSubjectRequiredBySchoolProfessionRequiredList = _CourseGroupRowList.Where(
                        x => x.Cells[0].Value.ToString() == tbProfessionalSubjectRequiredBySchool.Text
                        && x.Cells[4].Value.ToString() == tbProfessionalSubjectRequiredBySchoolProfession.Text
                        && x.Cells[1].Value.ToString() == tbProfessionalSubjectRequiredBySchoolProfessionRequired.Text).ToList();
                    foreach (TextBoxX tb in _TechProfessionalSubjectRequiredBySchoolProfessionRequiredList)
                    {
                        int credit = professionalSubjectRequiredBySchoolProfessionRequiredList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = professionalSubjectRequiredBySchoolProfessionRequiredList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 校訂專業科目選修
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> professionalSubjectRequiredBySchoolProfessionNonRequiredList = _CourseGroupRowList.Where(
                        x => x.Cells[0].Value.ToString() == tbProfessionalSubjectRequiredBySchool.Text
                        && x.Cells[4].Value.ToString() == tbProfessionalSubjectRequiredBySchoolProfession.Text
                        && x.Cells[1].Value.ToString() == tbProfessionalSubjectRequiredBySchoolProfessionNonRequired.Text).ToList();
                    foreach (TextBoxX tb in _TechProfessionalSubjectRequiredBySchoolProfessionNonRequiredList)
                    {
                        int credit = professionalSubjectRequiredBySchoolProfessionNonRequiredList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = professionalSubjectRequiredBySchoolProfessionNonRequiredList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 校訂實習科目必修
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> professionalSubjectRequiredBySchoolPracticeRequiredList = _CourseGroupRowList.Where(
                        x => x.Cells[0].Value.ToString() == tbProfessionalSubjectRequiredBySchool.Text
                        && x.Cells[4].Value.ToString() == tbProfessionalSubjectRequiredBySchoolPractice.Text
                        && x.Cells[1].Value.ToString() == tbProfessionalSubjectRequiredBySchoolPracticeRequired.Text).ToList();
                    foreach (TextBoxX tb in _TechProfessionalSubjectRequiredBySchoolPracticeRequiredList)
                    {
                        int credit = professionalSubjectRequiredBySchoolPracticeRequiredList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = professionalSubjectRequiredBySchoolPracticeRequiredList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 校訂實習科目選修
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> professionalSubjectRequiredBySchoolPracticeNonRequiredList = _CourseGroupRowList.Where(
                        x => x.Cells[0].Value.ToString() == tbProfessionalSubjectRequiredBySchool.Text
                        && x.Cells[4].Value.ToString() == tbProfessionalSubjectRequiredBySchoolPractice.Text
                        && x.Cells[1].Value.ToString() == tbProfessionalSubjectRequiredBySchoolPracticeNonRequired.Text).ToList();
                    foreach (TextBoxX tb in _TechProfessionalSubjectRequiredBySchoolPracticeNonRequiredList)
                    {
                        int credit = professionalSubjectRequiredBySchoolPracticeNonRequiredList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = professionalSubjectRequiredBySchoolPracticeNonRequiredList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    #region 合計
                    tbNormalSubjectRequiredbyDepartSummary.Text = _TechNormalSubjectRequiredbyDepartList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbNormalSubjectRequiredbySchoolRequiredSummary.Text = _TechNormalSubjectRequiredbySchoolRequiredList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbNormalSubjectRequiredbySchoolNonRequiredSummary.Text = _TechNormalSubjectRequiredbySchoolNonRequiredList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbProfessionalSubjectRequiredByDepartProfessionalSummary.Text = _TechProfessionalSubjectRequiredByDepartProfessionalList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbProfessionalSubjectRequiredByDepartPracticeSummary.Text = _TechProfessionalSubjectRequiredByDepartPracticeList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbProfessionalSubjectRequiredBySchoolProfessionRequiredSummary.Text = _TechProfessionalSubjectRequiredBySchoolProfessionRequiredList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbProfessionalSubjectRequiredBySchoolProfessionNonRequiredSummary.Text = _TechProfessionalSubjectRequiredBySchoolProfessionNonRequiredList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbProfessionalSubjectRequiredBySchoolPracticeRequiredSummary.Text = _TechProfessionalSubjectRequiredBySchoolPracticeRequiredList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbProfessionalSubjectRequiredBySchoolPracticeNonRequiredSummary.Text = _TechProfessionalSubjectRequiredBySchoolPracticeNonRequiredList
                        .Sum(x => int.Parse(x.Text)).ToString();


                    tbCreditSummary1_1.Text = (new List<string>() {
                        tbNormalSubjectRequiredbyDepart1_1.Text,
                        tbNormalSubjectRequiredbySchoolRequired1_1.Text,
                        tbNormalSubjectRequiredbySchoolNonRequired1_1.Text,
                        tbProfessionalSubjectRequiredByDepartProfessional1_1.Text,
                        tbProfessionalSubjectRequiredByDepartPractice1_1.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionRequired1_1.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionNonRequired1_1.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeRequired1_1.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeNonRequired1_1.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbCreditSummary1_2.Text = (new List<string>() {
                        tbNormalSubjectRequiredbyDepart1_2.Text,
                        tbNormalSubjectRequiredbySchoolRequired1_2.Text,
                        tbNormalSubjectRequiredbySchoolNonRequired1_2.Text,
                        tbProfessionalSubjectRequiredByDepartProfessional1_2.Text,
                        tbProfessionalSubjectRequiredByDepartPractice1_2.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionRequired1_2.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionNonRequired1_2.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeRequired1_2.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeNonRequired1_2.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbCreditSummary2_1.Text = (new List<string>() {
                        tbNormalSubjectRequiredbyDepart2_1.Text,
                        tbNormalSubjectRequiredbySchoolRequired2_1.Text,
                        tbNormalSubjectRequiredbySchoolNonRequired2_1.Text,
                        tbProfessionalSubjectRequiredByDepartProfessional2_1.Text,
                        tbProfessionalSubjectRequiredByDepartPractice2_1.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionRequired2_1.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionNonRequired2_1.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeRequired2_1.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeNonRequired2_1.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbCreditSummary2_2.Text = (new List<string>() {
                        tbNormalSubjectRequiredbyDepart2_2.Text,
                        tbNormalSubjectRequiredbySchoolRequired2_2.Text,
                        tbNormalSubjectRequiredbySchoolNonRequired2_2.Text,
                        tbProfessionalSubjectRequiredByDepartProfessional2_2.Text,
                        tbProfessionalSubjectRequiredByDepartPractice2_2.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionRequired2_2.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionNonRequired2_2.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeRequired2_2.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeNonRequired2_2.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbCreditSummary3_1.Text = (new List<string>() {
                        tbNormalSubjectRequiredbyDepart3_1.Text,
                        tbNormalSubjectRequiredbySchoolRequired3_1.Text,
                        tbNormalSubjectRequiredbySchoolNonRequired3_1.Text,
                        tbProfessionalSubjectRequiredByDepartProfessional3_1.Text,
                        tbProfessionalSubjectRequiredByDepartPractice3_1.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionRequired3_1.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionNonRequired3_1.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeRequired3_1.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeNonRequired3_1.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbCreditSummary3_2.Text = (new List<string>() {
                        tbNormalSubjectRequiredbyDepart3_2.Text,
                        tbNormalSubjectRequiredbySchoolRequired3_2.Text,
                        tbNormalSubjectRequiredbySchoolNonRequired3_2.Text,
                        tbProfessionalSubjectRequiredByDepartProfessional3_2.Text,
                        tbProfessionalSubjectRequiredByDepartPractice3_2.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionRequired3_2.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionNonRequired3_2.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeRequired3_2.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeNonRequired3_2.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbCreditSummary.Text = (new List<string>() {
                        tbNormalSubjectRequiredbyDepartSummary.Text,
                        tbNormalSubjectRequiredbySchoolRequiredSummary.Text,
                        tbNormalSubjectRequiredbySchoolNonRequiredSummary.Text,
                        tbProfessionalSubjectRequiredByDepartProfessionalSummary.Text,
                        tbProfessionalSubjectRequiredByDepartPracticeSummary.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionRequiredSummary.Text,
                        tbProfessionalSubjectRequiredBySchoolProfessionNonRequiredSummary.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeRequiredSummary.Text,
                        tbProfessionalSubjectRequiredBySchoolPracticeNonRequiredSummary.Text
                    }).Sum(x => int.Parse(x)).ToString();
                    #endregion
                }
                else if (_CourseType == "普通型高中" || _CourseType == "綜合型高中")
                {
                    tcSwitchCreditStatistics.SelectedTab = tbiCreditNormal;
                    tbiCreditTechnical.Visible = false;

                    // 學業部定必修
                    int index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> normalSubjectRequiredbyDepartmentList = _CourseGroupRowList.Where(
                        x => x.Cells[4].Value.ToString() == tbNormalNormalSubject.Text
                        && x.Cells[0].Value.ToString() == tbNormalNormalRequiredByDepart.Text).ToList();
                    foreach (TextBoxX tb in _NormalNormalRequiredByDepartList)
                    {
                        int credit = normalSubjectRequiredbyDepartmentList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = normalSubjectRequiredbyDepartmentList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 學業校訂必修
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> normalSubjectRequiredbySchoolRequiredList = _CourseGroupRowList.Where(
                        x => x.Cells[4].Value.ToString() == tbNormalNormalSubject.Text
                        && x.Cells[0].Value.ToString() == tbNormalNormalSubjectRequiredBySchool.Text
                        && x.Cells[1].Value.ToString() == tbNormalNormalSubjectRequiredBySchoolRequired.Text).ToList();
                    foreach (TextBoxX tb in _NormalNormalSubjectRequiredBySchoolRequiredList)
                    {
                        int credit = normalSubjectRequiredbySchoolRequiredList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = normalSubjectRequiredbySchoolRequiredList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    // 學業校訂選修
                    index = 7; // 一上從第7欄開始
                    List<DataGridViewRow> normalSubjectRequiredbySchoolNonRequiredList = _CourseGroupRowList.Where(
                        x => x.Cells[4].Value.ToString() == tbNormalNormalSubject.Text
                        && x.Cells[0].Value.ToString() == tbNormalNormalSubjectRequiredBySchool.Text
                        && x.Cells[1].Value.ToString() == tbNormalNormalSubjectRequiredBySchoolNonRequired.Text
                        && x.Cells[6].Value.ToString() != "團體活動時間"
                        && x.Cells[6].Value.ToString() != "彈性學習時間").ToList();
                    foreach (TextBoxX tb in _NormalNormalSubjectRequiredBySchoolNonRequiredList)
                    {
                        int credit = normalSubjectRequiredbySchoolNonRequiredList
                            .Where(x => x.Cells[index].Style.BackColor == Color.White)
                            .Select(x => int.Parse(x.Cells[index].Value == null ? "0" : x.Cells[index].Value.ToString()))
                            .Sum();
                        tb.Text = credit.ToString();

                        int groupCredit = 0;
                        List<string> courseGroupNameList = normalSubjectRequiredbySchoolNonRequiredList
                            .Where(x => x.Cells[index].Style.BackColor != Color.White && x.Cells[index].Tag != null)
                            .Select(x => ((XElement)x.Cells[index].Tag).Attribute("分組名稱") == null ? "" : ((XElement)x.Cells[index].Tag).Attribute("分組名稱").Value.ToString())
                            .Distinct().ToList();
                        foreach (string groupName in courseGroupNameList)
                        {
                            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == groupName).Count() > 0)
                            {
                                string creditString = _CourseGroupSettingList.First(x => x.CourseGroupName == groupName).CourseGroupCredit;
                                if (int.TryParse(creditString, out groupCredit))
                                {
                                    credit += groupCredit;
                                }
                            }
                        }

                        tb.Text = credit.ToString();
                        index++;
                    }

                    #region 合計
                    tbNormalNormalRequiredByDepartSummary.Text = _NormalNormalRequiredByDepartList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbNormalNormalSubjectRequiredBySchoolRequiredSummary.Text = _NormalNormalSubjectRequiredBySchoolRequiredList
                        .Sum(x => int.Parse(x.Text)).ToString();
                    tbNormalNormalSubjectRequiredBySchoolNonRequiredSummary.Text = _NormalNormalSubjectRequiredBySchoolNonRequiredList
                        .Sum(x => int.Parse(x.Text)).ToString();

                    tbNormalSummary1_1.Text = (new List<string>() {
                        tbNormalNormalRequiredByDepart1_1.Text,
                        tbNormalNormalSubjectRequiredBySchoolRequired1_1.Text,
                        tbNormalNormalSubjectRequiredBySchoolNonRequired1_1.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbNormalSummary1_2.Text = (new List<string>() {
                        tbNormalNormalRequiredByDepart1_2.Text,
                        tbNormalNormalSubjectRequiredBySchoolRequired1_2.Text,
                        tbNormalNormalSubjectRequiredBySchoolNonRequired1_2.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbNormalSummary2_1.Text = (new List<string>() {
                        tbNormalNormalRequiredByDepart2_1.Text,
                        tbNormalNormalSubjectRequiredBySchoolRequired2_1.Text,
                        tbNormalNormalSubjectRequiredBySchoolNonRequired2_1.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbNormalSummary2_2.Text = (new List<string>() {
                        tbNormalNormalRequiredByDepart2_2.Text,
                        tbNormalNormalSubjectRequiredBySchoolRequired2_2.Text,
                        tbNormalNormalSubjectRequiredBySchoolNonRequired2_2.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbNormalSummary3_1.Text = (new List<string>() {
                        tbNormalNormalRequiredByDepart3_1.Text,
                        tbNormalNormalSubjectRequiredBySchoolRequired3_1.Text,
                        tbNormalNormalSubjectRequiredBySchoolNonRequired3_1.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbNormalSummary3_2.Text = (new List<string>() {
                        tbNormalNormalRequiredByDepart3_2.Text,
                        tbNormalNormalSubjectRequiredBySchoolRequired3_2.Text,
                        tbNormalNormalSubjectRequiredBySchoolNonRequired3_2.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    tbNormalSummary.Text = (new List<string>() {
                        tbNormalNormalRequiredByDepartSummary.Text,
                        tbNormalNormalSubjectRequiredBySchoolRequiredSummary.Text,
                        tbNormalNormalSubjectRequiredBySchoolNonRequiredSummary.Text,
                    }).Sum(x => int.Parse(x)).ToString();
                    #endregion
                }
                else
                {
                    tcSwitchCreditStatistics.SelectedTab = null;
                    tbiCreditTechnical.Visible = false;
                    tbiCreditNormal.Visible = false;
                }
            }
            #endregion
        }

        //// 設定群組
        private void dgvCourseGroup_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvCourseGroupManageGroup.Rows.Count > 0 && e.RowIndex > -1)
            {
                if (e.ColumnIndex > 6 && e.ColumnIndex < 13)
                {
                    DataGridViewCell cell = dgvCourseGroup[e.ColumnIndex, e.RowIndex];

                    if (cell.Value == null)
                    {
                        return;
                    }

                    int groupIndex = dgvCourseGroupManageGroup.SelectedRows[0].Index;
                    string courseGroupName = _CourseGroupSettingList[groupIndex].CourseGroupName;
                    string courseGroupCredit = _CourseGroupSettingList[groupIndex].CourseGroupCredit;
                    Color color = _CourseGroupSettingList[groupIndex].CourseGroupColor;

                    XElement element = (XElement)cell.Tag;

                    // 格子按下時，如果顏色與原本的相同，會取消群組
                    if (cell.Style.BackColor == color)
                    {
                        cell.Style.BackColor = Color.White;
                        element.SetAttributeValue("分組名稱", "");
                        element.SetAttributeValue("分組修課學分數", "");
                        if (_CourseGroupSettingDic.ContainsKey(courseGroupName))
                        {
                            if (_CourseGroupSettingDic[courseGroupName].Contains(cell))
                            {
                                _CourseGroupSettingDic[courseGroupName].Remove(cell);
                            }
                        }
                    }
                    else
                    {
                        cell.Style.BackColor = color;
                        element.SetAttributeValue("分組名稱", courseGroupName);
                        element.SetAttributeValue("分組修課學分數", courseGroupCredit);
                        if (_CourseGroupSettingDic.ContainsKey(courseGroupName))
                        {
                            _CourseGroupSettingDic[courseGroupName].Add(cell);
                        }
                    }
                    SetIsDirtyDisplay(true);
                }
            }
        }
    }
}
