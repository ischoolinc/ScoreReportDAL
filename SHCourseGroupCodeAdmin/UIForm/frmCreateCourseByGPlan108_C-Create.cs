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
    public partial class frmCreateCourseByGPlan108_C_Create : BaseForm
    {
        string _SchoolYear = "", _Semester = "";
        Dictionary<string, SubjectCourseInfo> _SubjectCourseInfoDict;
        BackgroundWorker _bwWorker;
        DataAccess _da;
        StringBuilder _errSb = new StringBuilder();
        public frmCreateCourseByGPlan108_C_Create()
        {
            InitializeComponent();
            _bwWorker = new BackgroundWorker();
            _da = new DataAccess();
            _bwWorker.DoWork += _bwWorker_DoWork;
            _bwWorker.ProgressChanged += _bwWorker_ProgressChanged;
            _bwWorker.RunWorkerCompleted += _bwWorker_RunWorkerCompleted;
            _bwWorker.WorkerReportsProgress = true;
        }

        private void _bwWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            if (Global._CreateCourseErrorMsgList.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("開課課程發生問題：");
                sb.AppendLine(string.Join(",", Global._CreateCourseErrorMsgList.ToArray()));
                MsgBox.Show(sb.ToString(), "開課課程發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
            }

            if (_errSb.Length > 2)
            {
                MsgBox.Show("錯誤：" + _errSb.ToString());
            }
            else
            {
                // 呼叫課程同步
                FISCA.Features.Invoke("CourseSyncAllBackground");
                FISCA.Presentation.MotherForm.SetStatusBarMessage("完成");
                MsgBox.Show("產生完成。");

                // 沒問題關閉畫面
                this.DialogResult = DialogResult.OK;
            }
        }

        private void _bwWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("資料處理中 ...", e.ProgressPercentage);
        }

        private void _bwWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _errSb.Clear();
            _bwWorker.ReportProgress(1);
            // 新增課程
            _errSb.AppendLine(_da.AddGPlanCourse_C_BySchoolYearSemester(_SchoolYear, _Semester, _SubjectCourseInfoDict.Values.ToList()));

            _bwWorker.ReportProgress(100);
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void SetSubjectCourseInfoDict(Dictionary<string, SubjectCourseInfo> data)
        {
            _SubjectCourseInfoDict = data;
        }

        private void frmCreateCourseByGPlan108_C_Create_Load(object sender, EventArgs e)
        {
            // 說明文字
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("班級開課清單：");
            foreach (SubjectCourseInfo data in _SubjectCourseInfoDict.Values)
            {
                if (data.CourseCount > 0)
                    sb.AppendLine(data.SubjectName + "(開課學期：" + data.OpenSemester + "):" + data.CourseCount + "門");
            }

            txtMsg.Text = sb.ToString();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            btnCreate.Enabled = false;
            _bwWorker.RunWorkerAsync();
        }

        public void SetSchoolYearSemester(string SchoolYear, string Semester)
        {
            _SchoolYear = SchoolYear;
            _Semester = Semester;
        }

    }
}
