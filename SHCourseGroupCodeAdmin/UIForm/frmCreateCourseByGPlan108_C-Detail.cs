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

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateCourseByGPlan108_C_Detail : BaseForm
    {
        string _SchoolYear = "", _Semester = "";
        Dictionary<string, SubjectCourseInfo> _SubjectCourseInfoDict;

        public frmCreateCourseByGPlan108_C_Detail()
        {
            InitializeComponent();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void SetSubjectCourseInfoDict(Dictionary<string, SubjectCourseInfo> data)
        {
            _SubjectCourseInfoDict = data;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            // 畫面資料檢查

            // 取得畫面資料
            foreach (DataGridViewRow drv in dgData.Rows)
            {
                SubjectCourseInfo data = drv.Tag as SubjectCourseInfo;
                int co;
                if (int.TryParse(drv.Cells["開課數"].Value.ToString(), out co))
                {
                    data.CourseCount = co;
                    string skey = drv.Cells["開課學期"].Value + "_" + drv.Cells["科目名稱"].Value;
                    if (_SubjectCourseInfoDict.ContainsKey(skey))
                        _SubjectCourseInfoDict[skey].CourseCount = co;
                }
            }

            frmCreateCourseByGPlan108_C_Create fcc = new frmCreateCourseByGPlan108_C_Create();
            fcc.SetSchoolYearSemester(_SchoolYear, _Semester);
            fcc.SetSubjectCourseInfoDict(_SubjectCourseInfoDict);
            if (fcc.ShowDialog() == DialogResult.OK)
            {
                this.DialogResult = DialogResult.OK;
            }
        }

        private void frmCreateCourseByGPlan108_C_Detail_Load(object sender, EventArgs e)
        {
            lblSchoolYear.Text = "開設課程學年度學期：" + _SchoolYear + "學年度 第" + _Semester + "學期";

            LoadColumns();
            LoadData();
        }

        public void SetSchoolYearSemester(string SchoolYear, string Semester)
        {
            _SchoolYear = SchoolYear;
            _Semester = Semester;
        }

        private void LoadColumns()
        {
            DataGridViewTextBoxColumn tbSubject = new DataGridViewTextBoxColumn();
            tbSubject.Name = "科目名稱";
            tbSubject.Width = 180;
            tbSubject.HeaderText = "科目名稱";
            tbSubject.ReadOnly = true;
            tbSubject.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewTextBoxColumn tbSemester = new DataGridViewTextBoxColumn();
            tbSemester.Name = "開課學期";
            tbSemester.Width = 50;
            tbSemester.HeaderText = "開課學期";
            tbSemester.ReadOnly = true;
            tbSemester.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewTextBoxColumn tbClass = new DataGridViewTextBoxColumn();
            tbClass.Name = "採用班級";
            tbClass.Width = 280;
            tbClass.HeaderText = "採用班級";
            tbClass.ReadOnly = true;
            tbClass.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewTextBoxColumn tbCount = new DataGridViewTextBoxColumn();
            tbCount.Name = "開課數";
            tbCount.Width = 60;
            tbCount.HeaderText = "開課數";
            tbCount.ReadOnly = false;
            tbCount.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgData.Columns.Add(tbSubject);
            dgData.Columns.Add(tbSemester);
            dgData.Columns.Add(tbClass);
            dgData.Columns.Add(tbCount);

            // 只允取數字
            dgData.EditMode = DataGridViewEditMode.EditOnEnter;
            dgData.ImeMode = ImeMode.Off;
        }

        private void LoadData()
        {
            dgData.Rows.Clear();
            foreach (string subjKey in _SubjectCourseInfoDict.Keys)
            {
                SubjectCourseInfo data = _SubjectCourseInfoDict[subjKey];
                int rowIdx = dgData.Rows.Add();
                dgData.Rows[rowIdx].Tag = data;
                dgData.Rows[rowIdx].Cells["科目名稱"].Value = data.SubjectName;
                dgData.Rows[rowIdx].Cells["開課學期"].Value = data.OpenSemester;
                dgData.Rows[rowIdx].Cells["採用班級"].Value = string.Join(",", data.ClassNameDict.Keys.ToArray());
                dgData.Rows[rowIdx].Cells["開課數"].Value = data.CourseCount;
            }
        }
    }
}
