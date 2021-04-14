using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Campus.Windows;
using SHCourseGroupCodeSetup.DAO;


namespace SHCourseGroupCodeSetup.DetailContent
{
    public partial class UCStudentGroupCodeItem : FISCA.Presentation.DetailContent
    {
        // 這功能在設定班級科班群代碼使用
        BackgroundWorker _bgWorker = new BackgroundWorker();
        ChangeListener _ChangeListener = new ChangeListener();
        bool _isBusy = false;
        string StudentGroupCode = "";
        string StudentGroupName = "";
        List<string> GroupNameList = new List<string>();
        DataAccess da = new DataAccess();

        public UCStudentGroupCodeItem()
        {
            InitializeComponent();
            this.Group = "學生班級資訊";

            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _ChangeListener.StatusChanged += _ChangeListener_StatusChanged;
            ComboBox cb = cbxCourseGroupCode as ComboBox;
            _ChangeListener.Add(new ComboBoxSource(cb, ComboBoxSource.ListenAttribute.Text));
        }

        private void _ChangeListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            CancelButtonVisible = (e.Status == ValueStatus.Dirty);
            SaveButtonVisible = (e.Status == ValueStatus.Dirty);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_isBusy)
            {
                _isBusy = false;
                _bgWorker.RunWorkerAsync();
                return;
            }
            LoadData();
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            StudentGroupCode = da.GetStudentCodeByStudentID(PrimaryKey);
            da.LoadMOEGroupCodeDict();
            StudentGroupName = da.GetGroupNameByCode(StudentGroupCode);
            GroupNameList = da.GetGroupNameList();
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            _BGRun();
        }

        private void LoadData()
        {
            _ChangeListener.SuspendListen();
            cbxCourseGroupCode.Text = "";
            cbxCourseGroupCode.Items.Clear();
            cbxCourseGroupCode.Items.Add("");
            foreach (string name in GroupNameList)
                cbxCourseGroupCode.Items.Add(name);

            cbxCourseGroupCode.Text = StudentGroupName;
            _ChangeListener.Reset();
            _ChangeListener.ResumeListen();
            this.Loading = false;
        }

        private void SetData()
        {
            da.SetStudentGroupCodeByStudentID(PrimaryKey, da.GetGroupCodeByName(cbxCourseGroupCode.Text));
        }

        protected override void OnSaveButtonClick(EventArgs e)
        {
            SetData();

            this.CancelButtonVisible = this.SaveButtonVisible = false;
            _BGRun();
        }

        protected override void OnCancelButtonClick(EventArgs e)
        {
            this.CancelButtonVisible = this.SaveButtonVisible = false;
            _BGRun();
        }

        /// <summary>
        /// 呼叫重新讀取資料
        /// </summary>
        private void _BGRun()
        {
            if (_bgWorker.IsBusy)
                _isBusy = true;
            else
            {
                this.Loading = true;
                _bgWorker.RunWorkerAsync();
            }
        }
    }
}
