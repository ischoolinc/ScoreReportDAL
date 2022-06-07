using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using SHCourseGroupCodeAdmin.DAO;
using System.IO;
using FISCA.Data;
using System.Data;

namespace SHCourseGroupCodeAdmin.Report
{
    public class rptDBGPlanXML
    {
        BackgroundWorker _bgWorker;
        StringBuilder sb;

        public rptDBGPlanXML()
        {
            _bgWorker = new BackgroundWorker();
            sb = new StringBuilder();

            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;


        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("系統內課程規劃表XML 產生中...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (sb.Length > 1)
            {
                Utility.ExprotText("系統內課程規劃表XML", sb.ToString());
            }
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            sb.Clear();
            sb.Append("id");
            sb.Append(",");
            sb.Append("name");
            sb.Append(",");
            sb.Append("content");
            sb.Append(",");
            sb.Append("moe_group_code");
            sb.AppendLine();

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select("SELECT id,name,content,moe_group_code FROM graduation_plan ORDER BY ID");

            foreach (DataRow dr in dt.Rows)
            {
                sb.Append(dr["id"] + "");
                sb.Append(",");
                sb.Append(dr["name"] + "");
                sb.Append(",");
                sb.Append(dr["content"] + "");
                sb.Append(",");
                sb.Append(dr["moe_group_code"] + "");
                sb.AppendLine();
            }

        }

        public void Run()
        {
            _bgWorker.RunWorkerAsync();
        }
    }
}
