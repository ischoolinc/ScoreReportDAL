using FISCA;
using FISCA.Authentication;
using FISCA.Data;
using FISCA.LogAgent;
using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmDeleteCourseStudent : BaseForm
    {

        List<string> _SelectIDList;
        Dictionary<string, string> _CourseIDNameDict;
        List<string> _SCAttendIDList;

        BackgroundWorker _bgWorker;
        BackgroundWorker _bgWorkerDel;

        /// <summary>
        /// 認證狀態
        /// </summary>
        private enum AuthStatus { NotVerified, Success, Fail }

        private AuthStatus _authStatus = AuthStatus.NotVerified;

        /// <summary>
        /// 課程資料是否讀取完成
        /// </summary>
        private bool _dataLoaded = false;

        /// <summary>
        /// 刪除按鈕閃爍用計時器
        /// </summary>
        private Timer _blinkTimer;

        /// <summary>
        /// 剩餘閃爍次數(每次切換顏色算一次)
        /// </summary>
        private int _blinkCount = 0;

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

            _blinkTimer = new Timer();
            _blinkTimer.Interval = 250;
            _blinkTimer.Tick += _blinkTimer_Tick;
            this.FormClosed += (s, e) => _blinkTimer.Dispose();
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
                foreach (string name in _CourseIDNameDict.Values)
                {
                    sb.AppendLine(name);
                }

                ApplicationLog.Log("刪除課程與修課學生", "刪除", sb.ToString());

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
            _dataLoaded = true;
            UpdateDelButton();

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
            if (_authStatus != AuthStatus.Success)
            {
                MsgBox.Show("請先完成身分認證");
                return;
            }

            if (MsgBox.Show(_CourseIDNameDict.Count + "筆課程在系統內將永久刪除無法回復，確定要刪除?", "永久刪除課程", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                return;

            StopBlink();
            btnDel.Enabled = false;
            buttonX1.Enabled = false;
            _bgWorkerDel.RunWorkerAsync();
        }

        private void frmDeleteCourseStudent_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            lblUserName.Text = DSAServices.UserAccount;

            if (DSAServices.AccountType == AccountType.Local)
            {
                MsgBox.Show("請使用Google Account登入\n才可以進行身分認證並刪除資料!");
                this.Close();
                return;
            }

            // 畫面開始時為未認證
            SetAuthStatus(AuthStatus.NotVerified);

            // 載入資料檢查
            _bgWorker.RunWorkerAsync();
        }

        /// <summary>
        /// 認證成功且資料讀取完成，才可執行刪除
        /// </summary>
        private void UpdateDelButton()
        {
            bool canDel = _dataLoaded && _authStatus == AuthStatus.Success;

            // 由不可刪除變成可刪除時，閃爍按鈕提示使用者
            if (canDel && !btnDel.Enabled)
            {
                btnDel.Enabled = true;
                StartBlink();
            }
            else if (!canDel)
            {
                StopBlink();
                btnDel.Enabled = false;
            }
        }

        /// <summary>
        /// 開始閃爍刪除按鈕
        /// </summary>
        private void StartBlink()
        {
            _blinkCount = 6;
            _blinkTimer.Start();
        }

        /// <summary>
        /// 停止閃爍並還原按鈕顏色
        /// </summary>
        private void StopBlink()
        {
            _blinkTimer.Stop();
            _blinkCount = 0;
            btnDel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
        }

        private void _blinkTimer_Tick(object sender, EventArgs e)
        {
            if (_blinkCount <= 0)
            {
                StopBlink();
                return;
            }

            btnDel.ColorTable = (_blinkCount % 2 == 0)
                ? DevComponents.DotNetBar.eButtonColor.BlueWithBackground
                : DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            _blinkCount--;
        }

        /// <summary>
        /// 設定認證狀態，並更新圖示與刪除按鈕
        /// </summary>
        private void SetAuthStatus(AuthStatus status)
        {
            _authStatus = status;
            UpdateDelButton();

            Image old = pictureBox1.Image;
            pictureBox1.Image = CreateStatusIcon(status, pictureBox1.Width, pictureBox1.Height);
            if (old != null)
                old.Dispose();
        }

        /// <summary>
        /// 繪製狀態圖示：未認證(驚嘆號)、成功(打勾)、失敗(X)
        /// </summary>
        private static Image CreateStatusIcon(AuthStatus status, int width, int height)
        {
            int size = Math.Min(width, height) - 2;
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                Color color;
                if (status == AuthStatus.Success)
                    color = Color.FromArgb(46, 160, 67);
                else if (status == AuthStatus.Fail)
                    color = Color.FromArgb(215, 58, 73);
                else
                    color = Color.FromArgb(240, 160, 0);

                float x = (width - size) / 2f, y = (height - size) / 2f;
                using (SolidBrush brush = new SolidBrush(color))
                    g.FillEllipse(brush, x, y, size, size);

                using (Pen pen = new Pen(Color.White, size / 8f))
                {
                    pen.StartCap = pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                    float cx = x + size / 2f, cy = y + size / 2f, r = size / 4.5f;

                    if (status == AuthStatus.Success)
                    {
                        g.DrawLines(pen, new PointF[] {
                            new PointF(cx - r, cy),
                            new PointF(cx - r / 3f, cy + r * 0.7f),
                            new PointF(cx + r, cy - r * 0.6f) });
                    }
                    else if (status == AuthStatus.Fail)
                    {
                        g.DrawLine(pen, cx - r, cy - r, cx + r, cy + r);
                        g.DrawLine(pen, cx + r, cy - r, cx - r, cy + r);
                    }
                    else
                    {
                        g.DrawLine(pen, cx, cy - r * 1.2f, cx, cy + r * 0.3f);
                        float dot = pen.Width;
                        using (SolidBrush white = new SolidBrush(Color.White))
                            g.FillEllipse(white, cx - dot / 2f, cy + r * 0.9f - dot / 2f, dot, dot);
                    }
                }
            }
            return bmp;
        }

        private void btCensorClick(object sender, EventArgs e)
        {
            buttonX1.Enabled = false;
            try
            {
                CensorMessage cm = FISCA.CensorToken.CheckIdentity();
                //認證成功才可刪除，其餘狀態皆視為失敗
                bool success = (cm.censor == CensorToken.Censor.Success);
                SetAuthStatus(success ? AuthStatus.Success : AuthStatus.Fail);

                // 認證失敗時，可重新進行認證
                if (!success)
                    buttonX1.Enabled = true;

                MsgBox.Show(cm.Message);
            }
            catch (Exception ex)
            {
                buttonX1.Enabled = true;
                SetAuthStatus(AuthStatus.Fail);
                MsgBox.Show("認證身分發生錯誤:\n" + ex.Message);
            }
        }
    }
}
