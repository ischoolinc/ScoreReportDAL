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
    public partial class UCClassGroupCodeItem : FISCA.Presentation.DetailContent
    {
        // 這功能在設定班級科班群代碼使用
        BackgroundWorker _bgWorker = new BackgroundWorker();
        ChangeListener _ChangeListener = new ChangeListener();
        bool _isBusy = false;
        string ClassGroupCode = "";
        string ClassGroupName = "";
        List<string> GroupNameList = new List<string>();
        DataAccess da = new DataAccess();

        public UCClassGroupCodeItem()
        {
            InitializeComponent();
            this.Group = "班級基本資料";

            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _ChangeListener.StatusChanged += _ChangeListener_StatusChanged;
            ComboBox cb = cbxCourseGroupCode as ComboBox;
            _ChangeListener.Add(new ComboBoxSource(cb,ComboBoxSource.ListenAttribute.Text));
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
            ClassGroupCode = da.GetClassGroupCodeByClassID(PrimaryKey);            
            da.LoadMOEGroupCodeDict();
            ClassGroupName = da.GetGroupNameByCode(ClassGroupCode);
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

            cbxCourseGroupCode.Text = ClassGroupName;
            _ChangeListener.Reset();
            _ChangeListener.ResumeListen();
            this.Loading = false;
        }

        private void SetData()
        {
            da.SetClassGroupCodeByClassID(PrimaryKey, da.GetGroupCodeByName(cbxCourseGroupCode.Text));
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

        private void cbxCourseGroupCode_TextChanged(object sender, EventArgs e)
        {

        }

        private void cbxCourseGroupCode_Validated(object sender, EventArgs e)
        {

        }
    }
}
