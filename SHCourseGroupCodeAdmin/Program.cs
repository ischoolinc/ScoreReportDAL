﻿using System;
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
using SHCourseGroupCodeAdmin.UIForm;

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
            ribbon5.Add(new RibbonFeature("A1CE769E-3AB2-404D-B743-1B3DD3E2598E", "開課檢核課程代碼"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["開課檢核課程代碼"].Enable = UserAcl.Current["A1CE769E-3AB2-404D-B743-1B3DD3E2598E"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["開課檢核課程代碼"].Click += delegate
            {
                frmCheckGPlanCourseCode fpc = new frmCheckGPlanCourseCode();
                fpc.ShowDialog();
            };

     
            //  被 更新班級課程規劃表 取代
            Catalog ribbon6 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon6.Add(new RibbonFeature("C6063F22-FF7C-4AB5-82F8-4FC82B205B1A", "產生課程規劃"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Enable = UserAcl.Current["C6063F22-FF7C-4AB5-82F8-4FC82B205B1A"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Image = Properties.Resources.update;
            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Click += delegate
            {

                //更新班級課程規劃
                frmCreateClassGPlanMain ccg = new frmCreateClassGPlanMain();
                ccg.ShowDialog();
            };

            //Catalog ribbon6 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            //ribbon6.Add(new RibbonFeature("C6063F22-FF7C-4AB5-82F8-4FC82B205B1A", "更新班級課程規劃"));

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["更新班級課程規劃"].Enable = UserAcl.Current["C6063F22-FF7C-4AB5-82F8-4FC82B205B1A"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["更新班級課程規劃"].Image = Properties.Resources.update;
            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["更新班級課程規劃"].Size = RibbonBarButton.MenuButtonSize.Medium;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["更新班級課程規劃"].Click += delegate
            //{

            //    //更新班級課程規劃
            //      frmCreateGPlanMainBatch ccg = new frmCreateGPlanMainBatch();
            //    ccg.ShowDialog();
            //};



            Catalog ribbon7 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon7.Add(new RibbonFeature("1E66275D-2040-4D1D-8C7D-F5163D230A22", "修課檢核課程代碼"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["修課檢核課程代碼"].Enable = UserAcl.Current["1E66275D-2040-4D1D-8C7D-F5163D230A22"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["修課檢核課程代碼"].Click += delegate
            {
                frmCheckSCAttendCourseCode fcs = new frmCheckSCAttendCourseCode();
                fcs.ShowDialog();
            };


            Catalog ribbon8 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon8.Add(new RibbonFeature("83FB7115-548B-4C63-840C-5967A0725189", "手動同步課程代碼API"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["手動同步課程代碼API"].Enable = UserAcl.Current["83FB7115-548B-4C63-840C-5967A0725189"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["手動同步課程代碼API"].Image = Properties.Resources.update_1;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["手動同步課程代碼API"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["手動同步課程代碼API"].Click += delegate
            {
                frmSyncCourseCodeAPI scAPI = new frmSyncCourseCodeAPI();
                scAPI.ShowDialog();
            };

            Catalog ribbon9 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon9.Add(new RibbonFeature("604F7D79-4B25-41DC-9E45-FCC328AF61C7", "學期成績檢核課程代碼"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["學期成績檢核課程代碼"].Enable = UserAcl.Current["604F7D79-4B25-41DC-9E45-FCC328AF61C7"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"]["學期成績檢核課程代碼"].Click += delegate
            {
                frmCheckSemsScoreCourseCode fcs = new frmCheckSemsScoreCourseCode();
                fcs.ShowDialog();
            };

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"].Image = Properties.Resources.approve;
            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["資料檢查"].Size = RibbonBarButton.MenuButtonSize.Large;

            // --- 開發中
            // 108 版本 -----
            //Catalog ribbon10 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            //ribbon10.Add(new RibbonFeature("84066116-8124-41F4-8149-0877dA75a417", "產生課程規劃"));

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Enable = UserAcl.Current["84066116-8124-41F4-8149-0877dA75a417"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Image = Properties.Resources.online_library;
            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Size = RibbonBarButton.MenuButtonSize.Medium;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Click += delegate
            //{
            //    frmCreateGPlanMain108 fMain = new frmCreateGPlanMain108();
            //    fMain.ShowDialog();

            //};


            //Catalog ribbon11 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            //ribbon11.Add(new RibbonFeature("B1E6A402-F173-42B5-8785-7250CF6D46BD", "課程規劃(108適用)"));

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程規劃(108適用)"].Enable = UserAcl.Current["B1E6A402-F173-42B5-8785-7250CF6D46BD"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程規劃(108適用)"].Image = Properties.Resources.schedule;
            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程規劃(108適用)"].Size = RibbonBarButton.MenuButtonSize.Medium;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程規劃(108適用)"].Click += delegate
            //{
            //    frmGPlanConfig108 fgc = new frmGPlanConfig108();
            //    fgc.ShowDialog();

            //};


            //Catalog ribbon12 = RoleAclSource.Instance["班級"]["教務"]["班級開課"];
            //ribbon12.Add(new RibbonFeature("375D9EC7-E4C2-46FB-AA5D-7ADE4A5F7F03", "依課程規劃表開課(108適用)"));

            //MotherForm.RibbonBarItems["班級", "教務"]["班級開課"]["依課程規劃表開課(108適用)"].Enable = UserAcl.Current["375D9EC7-E4C2-46FB-AA5D-7ADE4A5F7F03"].Executable;

            //MotherForm.RibbonBarItems["班級", "教務"]["班級開課"]["依課程規劃表開課(108適用)"].Click += delegate
            //{
            //    if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
            //    {
            //        frmCreateCourseByGPlan108 fcc = new frmCreateCourseByGPlan108(K12.Presentation.NLDPanels.Class.SelectedSource);
            //        fcc.ShowDialog();
            //    }
            //};

            //Catalog ribbon13 = RoleAclSource.Instance["班級"]["教務"]["班級開課"];
            //ribbon13.Add(new RibbonFeature("E5FD6FD7-7641-4E04-81B7-C2A5A1A346F8", "依課程規劃表開課-跨班(108適用)"));

            //MotherForm.RibbonBarItems["班級", "教務"]["班級開課"]["依課程規劃表開課-跨班(108適用)"].Enable = UserAcl.Current["E5FD6FD7-7641-4E04-81B7-C2A5A1A346F8"].Executable;

            //MotherForm.RibbonBarItems["班級", "教務"]["班級開課"]["依課程規劃表開課-跨班(108適用)"].Click += delegate
            //{
            //    if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
            //    {
            //        frmCreateCourseByGPlan108_C fcc = new frmCreateCourseByGPlan108_C(K12.Presentation.NLDPanels.Class.SelectedSource);
            //        fcc.ShowDialog();
            //    }
            //};

        }
    }
}
