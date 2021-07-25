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
using System.Net;
using System.IO;
using FISCA.Authentication;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmSyncCourseCodeAPI : BaseForm
    {

        Dictionary<string, string> _SchoolNDict;
        string _SchoolCode = "";
        string DSNS = "";

        public frmSyncCourseCodeAPI()
        {
            InitializeComponent();
            _SchoolNDict = new Dictionary<string, string>();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmSyncCourseCodeAPI_Load(object sender, EventArgs e)
        {
            int sy;
            if (int.TryParse(K12.Data.School.DefaultSchoolYear, out sy))
            {
                for (int s = sy - 3; s <= sy + 3; s++)
                    cboSchoolYear.Items.Add(s);
            }
            _SchoolCode = K12.Data.School.Code;
            DSNS = DSAServices.AccessPoint;

            //// test 
            //DSNS = "n.hzsh.tc.edu.tw";

            _SchoolNDict = Utility.GetSchoolNMapping();

            // 進校轉換日校
            if(_SchoolNDict.ContainsKey(DSNS))
            {
                _SchoolCode = _SchoolNDict[DSNS];
            }
            

            lblSchoolCode.Text = "學校代碼：" + _SchoolCode;
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
            this.MaximumSize = this.MinimumSize = this.Size;
            cboSchoolYear.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // 檢查是否有學校代碼
            if (string.IsNullOrWhiteSpace(_SchoolCode))
            {
                MsgBox.Show("請輸入學校代碼。");
                return;
            }

            btnRun.Enabled = false;
            cboSchoolYear.Enabled = false;

            try
            {
                int sy;
                if (int.TryParse(cboSchoolYear.Text, out sy))
                {
                    for (int s = (sy - 2); s <= sy; s++)
                    {
                        CallServiceBySchoolYear(s);
                    }
                }

                MsgBox.Show("同步完成");
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
            }


            btnRun.Enabled = true;
            cboSchoolYear.Enabled = true;
        }

        private string CallServiceBySchoolYear(int SchoolYear)
        {
            string value = "";
            string content = "";
            try
            {
                //    string DSNS = DSAServices.AccessPoint;
                string school_code = _SchoolCode;


                String targetUrl = @"https://console.1campus.net/api/moeproxy/sync/" + DSNS + "?school_code=" + school_code + "&year=" + SchoolYear + "&rspcmds=true&school_name=手動呼叫";
                HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(targetUrl);
                req.Method = "POST";
                req.ContentType = "application/json";

                if (content != "" && content != null)
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(content);
                    req.ContentLength = bytes.Length;
                    req.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

                    Stream oStreamOut = req.GetRequestStream();
                    oStreamOut.Write(bytes, 0, bytes.Length);
                }

                var response = req.GetResponse();
                Stream receiveStream = response.GetResponseStream();
                StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);
                string data = readStream.ReadToEnd();

            }
            catch (Exception ex)
            {
                return ex.Message;
            }


            return value;
        }

    }
}
