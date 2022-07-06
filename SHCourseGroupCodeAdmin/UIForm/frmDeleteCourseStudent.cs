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
using FISCA.Data;
using FISCA.LogAgent;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmDeleteCourseStudent : BaseForm
    {

        List<string> _SelectIDList;
        Dictionary<string, string> _CourseIDNameDict;
        List<string> _SCAttendIDList;

        BackgroundWorker _bgWorker;
        BackgroundWorker _bgWorkerDel;

        public frmDeleteCourseStudent()
        {
            InitializeComponent();
            _SelectIDList = new List<string>();
            _CourseIDNameDict = new Dictionary<string, string>();
            _SCAttendIDList = new List<string>();
            _bgWorker = new BackgroundWorker();
            _bgWorkerDel = new BackgroundWorker();

            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;

            _bgWorkerDel.DoWork += _bgWorkerDel_DoWork;
            _bgWorkerDel.RunWorkerCompleted += _bgWorkerDel_RunWorkerCompleted;
            _bgWorkerDel.ProgressChanged += _bgWorkerDel_ProgressChanged;
            _bgWorkerDel.WorkerReportsProgress = true;


        }

        private void _bgWorkerDel_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("刪除課程資料中 ...", e.ProgressPercentage);
        }

        private void _bgWorkerDel_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 呼叫課程同步
            FISCA.Features.Invoke("CourseSyncAllBackground");
            FISCA.Presentation.MotherForm.SetStatusBarMessage("刪除課程資料完成.");
            MsgBox.Show("刪除課程資料完成");
            this.Close();

        }

        private void _bgWorkerDel_DoWork(object sender, DoWorkEventArgs e)
        {
            K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();
            _bgWorkerDel.ReportProgress(1);
            // 檢查是否刪除修課相關
            if (_SCAttendIDList.Count > 0)
            {

                // 先刪除評量成績
                try
                {
                    string delSCESQL = "DELETE FROM sce_take  WHERE ref_sc_attend_id IN(" + string.Join(",", _SCAttendIDList.ToArray()) + ")";
                    uh.Execute(delSCESQL);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("刪除學生修課評量成績發生錯誤," + ex.Message);
                }

                _bgWorkerDel.ReportProgress(30);
                // 再刪除修課紀錄
                try
                {
                    string delSCSQL = "DELETE FROM sc_attend WHERE id IN(" + string.Join(",", _SCAttendIDList.ToArray()) + ")";
                    uh.Execute(delSCSQL);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("刪除學生修課紀錄發生錯誤," + ex.Message);
                }
            }

            _bgWorkerDel.ReportProgress(60);
            // 刪除課程資料
            try
            {
                string delCOSQL = "DELETE FROM course WHERE id IN(" + string.Join(",", _SelectIDList.ToArray()) + ")";
                uh.Execute(delCOSQL);

                // log 
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("刪除課程資料：");
                foreach(string name in _CourseIDNameDict.Values)
                {
                    sb.AppendLine(name);
                }

                ApplicationLog.Log("課程.刪除課程與修課學生", sb.ToString());

            }
            catch (Exception ex)
            {
                Console.WriteLine("刪除課程紀錄發生錯誤," + ex.Message);
            }

            _bgWorkerDel.ReportProgress(100);

        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取資料中 ...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnDel.Enabled = true;

            FISCA.Presentation.MotherForm.SetStatusBarMessage("課程資料讀取完成.");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("刪除課程時會一併刪除【修課紀錄】及【評量成績】，成績刪除後將無法恢復。請確認是否要刪除 " + _CourseIDNameDict.Keys.Count + " 筆課程 ? ");
            //if (_SCAttendIDList.Count > 0)
            //{
            //    sb.AppendLine("這些課程包含 " + _SCAttendIDList.Count + " 筆修課學生也會一起刪除。");
            //}
            lblMsg.Text = sb.ToString();

        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _CourseIDNameDict.Clear();
            _SCAttendIDList.Clear();
            _bgWorker.ReportProgress(1);

            if (_SelectIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();

                // 取得課程ID與名稱
                string coIDNameSQL = "SELECT id,school_year,semester,course_name FROM course WHERE id IN(" + string.Join(",", _SelectIDList.ToArray()) + ") ORDER BY course_name;";

                DataTable dtco = qh.Select(coIDNameSQL);
                if (dtco != null && dtco.Rows.Count > 0)
                    foreach (DataRow dr in dtco.Rows)
                    {
                        string coName = dr["school_year"] + "學年度第" + dr["semester"] + "學期 " + dr["course_name"];
                        _CourseIDNameDict.Add(dr["id"] + "", coName);
                    }

                _bgWorker.ReportProgress(50);

                // 取得修科紀錄     
                string sc_attSQL = "SELECT sc_attend.id FROM sc_attend WHERE ref_course_id IN(" + string.Join(",", _SelectIDList.ToArray()) + ");";

                DataTable dtsc = qh.Select(sc_attSQL);
                if (dtsc != null && dtsc.Rows.Count > 0)
                    foreach (DataRow dr in dtsc.Rows)
                    {
                        _SCAttendIDList.Add(dr["id"] + "");
                    }

            }

            _bgWorker.ReportProgress(100);
        }

        public void SetCourseIDs(List<string> CourseIDs)
        {
            _SelectIDList = CourseIDs;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            PasswordForm pf = new PasswordForm();
            if(pf.ShowDialog() == DialogResult.OK)
            {
                if (pf.GetPass())
                {
                    btnDel.Enabled = false;
                    _bgWorkerDel.RunWorkerAsync();
                }               
            }            
        }

        private void frmDeleteCourseStudent_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;

            btnDel.Enabled = false;

            // 載入資料檢查
            _bgWorker.RunWorkerAsync();
        }
    }
}
