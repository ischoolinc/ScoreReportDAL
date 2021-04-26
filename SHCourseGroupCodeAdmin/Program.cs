using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Permission;
using FISCA.Presentation.Controls;
using K12.Data;
using FISCA;
using FISCA.Presentation;
using SHCourseGroupCodeAdmin.Report;
using SHCourseGroupCodeAdmin.DataCheck;

namespace SHCourseGroupCodeAdmin
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {

            // 權限註冊
            Catalog ribbon1 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon1.Add(new RibbonFeature("C26EDE96-1018-45E5-9280-61D4B0986F80", "課程代碼表"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["課程代碼表"].Enable = UserAcl.Current["C26EDE96-1018-45E5-9280-61D4B0986F80"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["課程代碼表"].Click += delegate
            {
                //Console.WriteLine("課程代碼表");
                rptMOECourseCode moe = new rptMOECourseCode();
                moe.Run(); 
            };

            //Catalog ribbon2 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            //ribbon2.Add(new RibbonFeature("88918C1A-E0BD-48C5-A689-9F778AC776EC", "上傳實際開課"));


            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["上傳實際開課"].Enable = UserAcl.Current["88918C1A-E0BD-48C5-A689-9F778AC776EC"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["上傳實際開課"].Click += delegate
            //{
            //    Console.WriteLine("上傳實際開課");

            //};

            //Catalog ribbon3 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            //ribbon3.Add(new RibbonFeature("C1C2859E-104D-4ABA-8394-F1CA25E1EB95", "手動同步課程代碼"));


            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["設定"]["手動同步課程代碼"].Enable = UserAcl.Current["C1C2859E-104D-4ABA-8394-F1CA25E1EB95"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["設定"]["手動同步課程代碼"].Click += delegate
            //{
            //    Console.WriteLine("手動同步課程代碼");
            //};

            Catalog ribbon4 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon4.Add(new RibbonFeature("90C0A273-6387-49B9-BCCE-EFDC7F5A3931", "檢查班級群科班設定"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["檢查班級群科班設定"].Enable = UserAcl.Current["90C0A273-6387-49B9-BCCE-EFDC7F5A3931"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["檢查班級群科班設定"].Click += delegate
            {
                rptCheckClassGroupCode gCode = new rptCheckClassGroupCode();
                gCode.Run();
            };


            Catalog ribbon5 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon5.Add(new RibbonFeature("A1CE769E-3AB2-404D-B743-1B3DD3E2598E", " 課程代碼開課檢查"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"][" 課程代碼開課檢查"].Enable = UserAcl.Current["A1CE769E-3AB2-404D-B743-1B3DD3E2598E"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"][" 課程代碼開課檢查"].Click += delegate
            {
                chkCheckCourseCode fCode = new chkCheckCourseCode();
                fCode.ShowDialog();
            };
        }
    }
}
