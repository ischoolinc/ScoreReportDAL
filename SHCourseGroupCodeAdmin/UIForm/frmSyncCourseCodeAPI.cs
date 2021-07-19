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
        

        public frmSyncCourseCodeAPI()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmSyncCourseCodeAPI_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;

            try
            {
                int sy;
                if (int.TryParse(K12.Data.School.DefaultSchoolYear, out sy))
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
        }

        private bool CallServiceBySchoolYear(int SchoolYear)
        {
            bool value = false;
            string content = "";
            try
            {
                string DSNS = DSAServices.AccessPoint;
                string school_code = K12.Data.School.Code;
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
                value = true;
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
            }


            return value;
        }

    }
}
