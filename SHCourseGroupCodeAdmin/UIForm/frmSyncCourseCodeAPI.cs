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
using SHCourseGroupCodeAdmin.DAO;
using FISCA.Data;

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
            if (_SchoolNDict.ContainsKey(DSNS))
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

                //try
                //{
                //    // 處理  class_type 問題
                //    // 取得課程規劃表大表
                //    string query1 = "SELECT uid,class_type,group_code FROM $moe.subjectcode";
                //    Dictionary<string, string> mapDict = new Dictionary<string, string>();
                //    mapDict.Add("1", "建教合作-輪調式");
                //    mapDict.Add("2", "建教合作-階梯式");
                //    mapDict.Add("3", "建教合作-實習式");
                //    mapDict.Add("4", "建教合作-其他式");
                //    mapDict.Add("5", "產學訓專班");
                //    mapDict.Add("6", "雙軌旗艦計畫");
                //    mapDict.Add("7", "產攜專班");

                //    List<string> updateSQLList = new List<string>();


                //    QueryHelper qh = new QueryHelper();
                //    DataTable dt = qh.Select(query1);
                //    foreach (DataRow dr in dt.Rows)
                //    {
                //        string class_type = dr["class_type"] + "";
                //        string uid = dr["uid"] + "";
                //        if (class_type == "" || class_type == "undefined")
                //        {
                //            // 取群科班最後一碼
                //            string gcode = dr["group_code"] + "";
                //            if (gcode.Length > 15)
                //            {
                //                string newType = "不分班群";
                //                string code = gcode.Substring(15, 1);
                //                if (mapDict.ContainsKey(code))
                //                {
                //                    newType = mapDict[code];
                //                }

                //                if (!string.IsNullOrEmpty(uid))
                //                {
                //                    string UpdateStr = "UPDATE $moe.subjectcode SET class_type = '" + newType + "' WHERE uid = " + uid + ";";
                //                    updateSQLList.Add(UpdateStr);
                //                }

                //            }
                //        }
                //    }

                //    // 更新處理
                //    if (updateSQLList.Count > 0)
                //    {
                //        K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();
                //        int cot = uh.Execute(updateSQLList);
                //        //  Console.WriteLine("更新" + cot + "筆");
                //    }
                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine(ex.Message);
                //}

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
                //   school_code = "181307"; //"070406";

                              

                // 調整新主機位置 2022/10/19
                String targetUrl = @"https://moe-inte-service-4twhrljvua-de.a.run.app/api/moeproxy/sync/" + DSNS + "?school_code=" + school_code + "&year=" + SchoolYear + "&rspcmds=true&school_name=手動呼叫";
                HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(targetUrl);
                req.Method = "POST";
                req.ContentType = "application/json";
                req.ContentLength = 0;

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

        private void btnGetCourseCodeSource_Click(object sender, EventArgs e)
        {


            //解析資料
        }


    }
}
