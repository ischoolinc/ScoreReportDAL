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
    public partial class frmCourseCodeTest : BaseForm
    {
        string DSNS = "";


        public frmCourseCodeTest()
        {
            InitializeComponent();
            DSNS = DSAServices.AccessPoint;
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            btnRun.Enabled = false;

            string value = "";
            string content = "";

            try
            {
                List<string> SchoolCodeList = GetSchoolCodeList();

                foreach (string school_code in SchoolCodeList)
                {
                    for (int SchoolYear = 108; SchoolYear <= 111; SchoolYear++)
                    {

                        // 取得各校
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

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            btnRun.Enabled = true;
        }

        private void frmCourseCodeTest_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
        }

        private List<string> GetSchoolCodeList()
        {
            List<string> value = new List<string>();
            value.Add("064342");  //臺中市中港高中,cgsh.tc.edu.tw
            value.Add("041305");  //新竹縣忠信高中,chhs.hcc.edu.tw
            value.Add("200401");  //嘉義華南高商,hnvs.cy.edu.tw
            value.Add("070410");  //國立員林家商,ylhcvs.chc.edu.tw
            value.Add("070406");  //國立彰化高商,chsc.chc.edu.tw
            value.Add("193316");  //臺中市立惠文高中,h.hwsh.tc.edu.tw
            value.Add("064308");  //臺中市立后綜高級中學 高中部,h.hzsh.tc.edu.tw
            value.Add("193303");  //臺中市立忠明高級中學高中部,h.cmsh.tc.edu.tw
            value.Add("193313");  //臺中市立西苑高級中學,h.sysh.tc.edu.tw
            value.Add("070401");  //國立彰師附工,sivs.chc.edu.tw
            value.Add("061316");  //臺中市私立青年高級中學(進修部),n.youth.tc.edu.tw
            value.Add("180309");  //國立新竹高級中學,hchs.hc.edu.tw
            value.Add("193315");  //臺中市立東山高級中學,h.tsjh.tc.edu.tw
            value.Add("180301");  //國立新竹科學園區實驗中學高中部,h.nehs.hc.edu.tw
            value.Add("011330");  //康橋國際學校 林口校區 高中,h.kcislk.ntpc.edu.tw
            value.Add("074313");  //彰化縣立二林高中高中部,h.elsh.chc.edu.tw
            value.Add("064324");  //臺中市立大里高中,dljh.tcc.edu.tw
            value.Add("191302");  //臺中市私立葳格高級中學,h.wagor.tc.edu.tw
            value.Add("061317");  //臺中市私立弘文高級中學高中部,h.hwhs.tc.edu.tw
            value.Add("191305");  //臺中市私立新民高中,shinmin.tc.edu.tw
            value.Add("191309");  //臺中市私立明德中學,mdhs.tc.edu.tw
            value.Add("120401");  //國立旗山高級農工職業學校,csvs.khc.edu.tw
            value.Add("121417");  //高雄市私立高苑工商,kyvs.ks.edu.tw
            value.Add("171405");  //基隆市私立光隆家商-日校,klhcvs.kl.edu.tw
            value.Add("130302");  //國立屏東女子高級中學,ptgsh.ptc.edu.tw
            value.Add("061301");  //臺中市私立常春藤高中,h.ivyjhs.tc.edu.tw
            value.Add("044311");  //新竹縣立六家高級中學高中部,h.ljsh.hcc.edu.tw
            value.Add("070409");  //國立員林崇實高級工業職業學校,csvs.chc.edu.tw
            value.Add("194315");  //臺中市立文華高級中等學校,whsh.tc.edu.tw
            value.Add("104326");  //嘉義永慶高中,h.ycsh.cyc.edu.tw
            value.Add("181307");  //新竹市私立磐石高級中學進修部,n.sphs.hc.edu.tw
            value.Add("074323");  //彰化縣立和美高中,h.hmjh.chc.edu.tw
            value.Add("011302");  //康橋國際學校 秀岡校區 高中,h.kcbs.ntpc.edu.tw
            value.Add("100301");  //國立東石高級中學,tssh.cyc.edu.tw
            value.Add("540301");  //國立中山大學附屬國光高級中學高中部,h.kksh.kh.edu.tw
            value.Add("610405");  //國立高餐大附屬中學,h.nkuht.edu.tw
            value.Add("104319");  //嘉義竹崎高中,h.ccjh.cyc.edu.tw
            value.Add("050401");  //國立大湖農工,thvs.mlc.edu.tw
            value.Add("193407");  //臺中市立臺中工業高級中等學校,tcivs.tc.edu.tw

            return value;
        }

    }
}
