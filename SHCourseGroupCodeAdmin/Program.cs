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

            //Catalog ribbon2 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            //ribbon2.Add(new RibbonFeature("88918C1A-E0BD-48C5-A689-9F778AC776EC", "上傳實際開課"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["課程代碼表"].Enable = UserAcl.Current["C26EDE96-1018-45E5-9280-61D4B0986F80"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["課程代碼表"].Click += delegate
            {
                //Console.WriteLine("課程代碼表");
                rptMOECourseCode moe = new rptMOECourseCode();
                moe.Run(); 
            };


            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["上傳實際開課"].Enable = UserAcl.Current["88918C1A-E0BD-48C5-A689-9F778AC776EC"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["上傳實際開課"].Click += delegate
            //{
            //    Console.WriteLine("上傳實際開課");

            //};
        }
    }
}
