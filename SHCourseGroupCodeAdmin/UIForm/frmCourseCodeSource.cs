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
using System.Web.Script.Serialization;
using System.IO;
using System.Net;
using FISCA.Authentication;
using Aspose.Cells;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCourseCodeSource : BaseForm
    {
       
        string jsonSource = "";
        string _SchoolCode = "";
        string DSNS = "";
        Dictionary<string, int> _ColIdxDict;
        CoureseCodeChecker ccChecker;
        List<CourseCodeRoot> _dataList = new List<CourseCodeRoot>();
        Dictionary<string, string> _SchoolNDict;
        public frmCourseCodeSource()
        {
            InitializeComponent();
            _ColIdxDict = new Dictionary<string, int>();
        }

        public void SetData(string jsonString, List<CourseCodeRoot> dataList)
        {
            jsonSource = jsonString;
            _dataList = dataList;

   

        }


        private void LoadData()
        {
            dgData.Rows.Clear();
            foreach (var r in _dataList)
            {
                foreach (var rd in r.課程資料)
                {
                    int rowIdx = dgData.Rows.Add();
                    dgData.Rows[rowIdx].Tag = rd;
                    dgData.Rows[rowIdx].Cells["實施學年度"].Value = r.實施學年度;
                    dgData.Rows[rowIdx].Cells["課程代碼"].Value = rd.課程代碼;
                    dgData.Rows[rowIdx].Cells["課程類型"].Value = r.課程類型;
                    dgData.Rows[rowIdx].Cells["課程代碼"].Value = rd.課程代碼;
                    dgData.Rows[rowIdx].Cells["科目名稱"].Value = rd.科目名稱;
                    dgData.Rows[rowIdx].Cells["授課學期學分節數"].Value = rd.授課學期學分節數;
                    dgData.Rows[rowIdx].Cells["授課學期開課方式"].Value = rd.授課學期開課方式;
                    dgData.Rows[rowIdx].Cells["課程屬性"].Value = rd.課程屬性;
                }
            }

            lblCount.Text = "共" + dgData.Rows.Count + "筆";
        }

        private void LoadColumns()
        {
            DataGridViewTextBoxColumn tbSchoolYear = new DataGridViewTextBoxColumn();
            tbSchoolYear.Name = "實施學年度";
            tbSchoolYear.Width = 70;
            tbSchoolYear.HeaderText = "實施學年度";
            tbSchoolYear.ReadOnly = true;
            tbSchoolYear.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewTextBoxColumn tbCourseT = new DataGridViewTextBoxColumn();
            tbCourseT.Name = "課程類型";
            tbCourseT.Width = 20;
            tbCourseT.HeaderText = "課程類型";
            tbCourseT.ReadOnly = true;
            tbCourseT.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewTextBoxColumn tbCourseCode = new DataGridViewTextBoxColumn();
            tbCourseCode.Name = "課程代碼";
            tbCourseCode.Width = 220;
            tbCourseCode.HeaderText = "課程代碼";
            tbCourseCode.ReadOnly = true;
            tbCourseCode.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewTextBoxColumn tbSubjectName = new DataGridViewTextBoxColumn();
            tbSubjectName.Name = "科目名稱";
            tbSubjectName.Width = 250;
            tbSubjectName.HeaderText = "科目名稱";
            tbSubjectName.ReadOnly = true;
            tbSubjectName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewTextBoxColumn tbCredit = new DataGridViewTextBoxColumn();
            tbCredit.Name = "授課學期學分節數";
            tbCredit.Width = 80;
            tbCredit.HeaderText = "授課學期學分節數";
            tbCredit.ReadOnly = true;
            tbCredit.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;


            DataGridViewTextBoxColumn tbOpenType = new DataGridViewTextBoxColumn();
            tbOpenType.Name = "授課學期開課方式";
            tbOpenType.Width = 100;
            tbOpenType.HeaderText = "授課學期開課方式";
            tbOpenType.ReadOnly = true;
            tbOpenType.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewTextBoxColumn tbCourseType = new DataGridViewTextBoxColumn();
            tbCourseType.Name = "課程屬性";
            tbCourseType.Width = 100;
            tbCourseType.HeaderText = "課程屬性";
            tbCourseType.ReadOnly = true;
            tbCourseType.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgData.Columns.Add(tbSchoolYear);
            dgData.Columns.Add(tbCourseT);
            dgData.Columns.Add(tbCourseCode);
            dgData.Columns.Add(tbSubjectName);
            dgData.Columns.Add(tbCredit);
            dgData.Columns.Add(tbOpenType);
            dgData.Columns.Add(tbCourseType);
        }


        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void frmCourseCodeSource_Load(object sender, EventArgs e)
        {
            // 載入資料欄位
            LoadColumns();

            int sy;
            if (int.TryParse(K12.Data.School.DefaultSchoolYear, out sy))
            {
                for (int s = sy - 2; s <= sy + 1; s++)
                    cboSchoolYear.Items.Add(s);
            }
            _SchoolCode = K12.Data.School.Code;
            DSNS = DSAServices.AccessPoint;
            
            // 進校需要轉換日校學校代碼取得資料
            _SchoolNDict = Utility.GetSchoolNMapping();

            // 進校轉換日校
            if (_SchoolNDict.ContainsKey(DSNS))
            {
                _SchoolCode = _SchoolNDict[DSNS];
            }

           
            ccChecker = new CoureseCodeChecker(_SchoolCode, DSNS);

            lblSchoolCode.Text = "學校代碼：" + _SchoolCode;
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
            cboSchoolYear.DropDownStyle = ComboBoxStyle.DropDownList;      

            LoadDataToDataGridView();
        }


        private void LoadDataToDataGridView()
        {
            ControlEnable(false);
            LoadCourseCodeJSONData();
            LoadData();
            ControlEnable(true);
        }

        private void LoadCourseCodeJSONData()
        {
            int sy;

            if (int.TryParse(cboSchoolYear.Text, out sy))
            {
                ccChecker.LoadCouseCodeSourceData(sy, sy);
                Dictionary<int, List<CourseCodeRoot>> courseDataDict = ccChecker.GetCourseCodeRootDict();
                Dictionary<int, string> jsondataDict = ccChecker.GetJSONStringDict();
                if (jsondataDict.Count > 0 && jsondataDict[sy] != "")
                {
                    SetData(jsondataDict[sy], courseDataDict[sy]);
                }
            }
        }

        private void ControlEnable(bool value)
        {
            btnExcel.Enabled = btnJSON.Enabled  = value;
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            ControlEnable(false);
            LoadCourseCodeJSONData();
            // 取得樣板
            try
            {
                Workbook wb = new Workbook(new MemoryStream(Properties.Resources.課程代碼原始資料樣板));
                Worksheet wst = wb.Worksheets[0];
                wst.Name = cboSchoolYear.Text + "實施年度";
                _ColIdxDict.Clear();
                // 讀取欄位與索引            
                for (int co = 0; co <= wst.Cells.MaxDataColumn; co++)
                {
                    _ColIdxDict.Add(wst.Cells[0, co].StringValue, co);
                }

                int rowIdx = 1;
                foreach (var r in _dataList)
                {
                    foreach (var rd in r.課程資料)
                    {
                        wst.Cells[rowIdx, GetColIndex("實施學年度")].PutValue(r.實施學年度);
                        wst.Cells[rowIdx, GetColIndex("課程類型")].PutValue(r.課程類型);
                        wst.Cells[rowIdx, GetColIndex("課程代碼")].PutValue(rd.課程代碼);
                        wst.Cells[rowIdx, GetColIndex("科目名稱")].PutValue(rd.科目名稱);
                        wst.Cells[rowIdx, GetColIndex("授課學期學分節數")].PutValue(rd.授課學期學分節數);
                        wst.Cells[rowIdx, GetColIndex("授課學期開課方式")].PutValue(rd.授課學期開課方式);
                        wst.Cells[rowIdx, GetColIndex("課程屬性")].PutValue(rd.課程屬性);
                        rowIdx++;
                    }
                }

                wst.AutoFitColumns();

                Utility.ExprotXls("課程代碼原始資料", wb);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            ControlEnable(true);
        }

        private void btnJSON_Click(object sender, EventArgs e)
        {
            ControlEnable(false);
            LoadCourseCodeJSONData();
            if (!string.IsNullOrWhiteSpace(jsonSource))
            {
                Utility.ExprotJSON(cboSchoolYear.Text + "課程代碼原始資料", jsonSource);
            }
            else
            {
                MsgBox.Show("沒有資料");
            }

            ControlEnable(true);
        }


        private int GetColIndex(string name)
        {
            int value = 0;
            if (_ColIdxDict.ContainsKey(name))
                value = _ColIdxDict[name];

            return value;
        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            dgData.Rows.Clear();
            lblCount.Text = "共0筆";
            LoadDataToDataGridView();
        }
    }
}
