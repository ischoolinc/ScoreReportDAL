using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Permission;
using FISCA.Presentation.Controls;
using FISCA;
using FISCA.Presentation;
using SHCourseCodeCheckAndUpdate.UIForm;

namespace SHCourseCodeCheckAndUpdate
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {
            Catalog ribbon1 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon1.Add(new RibbonFeature("B2DD028D-AD4E-42B5-8256-2FFC5927DA09", "修課課程代碼檢查與更新"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["修課課程代碼檢查與更新"].Enable = UserAcl.Current["B2DD028D-AD4E-42B5-8256-2FFC5927DA09"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["修課課程代碼檢查與更新"].Click += delegate
            {
                frmSCAttendChkUpdate fsc = new frmSCAttendChkUpdate();
                fsc.ShowDialog();
            };


            Catalog ribbon2 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon2.Add(new RibbonFeature("D3FC7874-F264-49DF-B3BB-89490FFAB917", "學期成績課程代碼檢查與更新"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["學期成績課程代碼檢查與更新"].Enable = UserAcl.Current["D3FC7874-F264-49DF-B3BB-89490FFAB917"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["學期成績課程代碼檢查與更新"].Click += delegate
            {
                frmSemsScoreChkUpdate fsc = new frmSemsScoreChkUpdate();
                fsc.ShowDialog();
            };


        }
    }
}
