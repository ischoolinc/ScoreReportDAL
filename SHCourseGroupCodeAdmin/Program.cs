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
using Campus.Message;

namespace SHCourseGroupCodeAdmin
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {
            #region 學生第6學期修課紀錄
            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["學生", "資料統計"];
            rbItem1["報表"]["成績相關報表"]["學生第6學期修課紀錄"].Enable = UserAcl.Current["SH_6thSemesterCorseCodeRank"].Executable;
            rbItem1["報表"]["成績相關報表"]["學生第6學期修課紀錄"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Student6thSemesterCorseCodeRank srf = new Student6thSemesterCorseCodeRank();
                    srf.SetStudentIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
                    srf.ShowDialog();

                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選學生。");
                    return;
                }
            };

            // 權限
            Catalog catalog1 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog1.Add(new RibbonFeature("SH_6thSemesterCorseCodeRank", "學生第6學期修課紀錄"));

            #endregion
            // 權限註冊
            Catalog ribbon1 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon1.Add(new RibbonFeature("C26EDE96-1018-45E5-9280-61D4B0986F80", "課程代碼表"));
            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"].Size = RibbonBarButton.MenuButtonSize.Large;
            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"].Image = Properties.Resources.Report;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["課程代碼表"].Enable = UserAcl.Current["C26EDE96-1018-45E5-9280-61D4B0986F80"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["課程代碼表"].Click += delegate
            {
                //Console.WriteLine("課程代碼表");
                rptMOECourseCode moe = new rptMOECourseCode();
                moe.Run(); 
            };

            //Catalog ribbon1_p1 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            //ribbon1_p1.Add(new RibbonFeature("531EC923-C442-4B55-8117-773AD4FE6548", "產生全部課程規劃XML"));

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["產生全部課程規劃XML"].Enable = UserAcl.Current["531EC923-C442-4B55-8117-773AD4FE6548"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["產生全部課程規劃XML"].Click += delegate
            //{
            //    //Console.WriteLine("產生全部課程規劃XML");
            //    rptDBGPlanXML dgp = new rptDBGPlanXML();
            //    dgp.Run();
            //};


            //Catalog ribbon1a = RoleAclSource.Instance["教務作業"]["課程代碼"];
            //ribbon1a.Add(new RibbonFeature("0FA6AC77-7657-41C6-B065-73FDD191F9ED", "群科班學分數總計"));

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["群科班學分數總計"].Enable = UserAcl.Current["0FA6AC77-7657-41C6-B065-73FDD191F9ED"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["報表"]["群科班學分數總計"].Click += delegate
            //{
            //    rptMOECourseCodeSumCredit moe = new rptMOECourseCodeSumCredit();
            //    moe.Run();
            //};


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


            ////  被 更新班級課程規劃表 取代
            //Catalog ribbon6 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            //ribbon6.Add(new RibbonFeature("C6063F22-FF7C-4AB5-82F8-4FC82B205B1A", "產生課程規劃"));

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Enable = UserAcl.Current["C6063F22-FF7C-4AB5-82F8-4FC82B205B1A"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Image = Properties.Resources.update;
            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Size = RibbonBarButton.MenuButtonSize.Medium;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["產生課程規劃"].Click += delegate
            //{

            //    //更新班級課程規劃
            //    frmCreateClassGPlanMain ccg = new frmCreateClassGPlanMain();
            //    ccg.ShowDialog();
            //};

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

            Catalog ribbon8_1 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon8_1.Add(new RibbonFeature("010627DB-C14E-43BB-9407-BF5550391FBA", "課程計畫平台原始課程代碼"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程計畫"]["課程計畫平台原始課程代碼"].Enable = UserAcl.Current["010627DB-C14E-43BB-9407-BF5550391FBA"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程計畫平台原始課程代碼"].Image = Properties.Resources.update_1;
            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程計畫平台原始課程代碼"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程計畫"]["課程計畫平台原始課程代碼"].Click += delegate
            {
                frmCourseCodeSource fcc = new frmCourseCodeSource();
                fcc.ShowDialog();
            };

            Catalog ribbon8 = RoleAclSource.Instance["教務作業"]["課程代碼"];
            ribbon8.Add(new RibbonFeature("83FB7115-548B-4C63-840C-5967A0725189", "手動同步課程代碼API"));

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程計畫"]["手動同步課程代碼API"].Enable = UserAcl.Current["83FB7115-548B-4C63-840C-5967A0725189"].Executable;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程計畫"].Size = RibbonBarButton.MenuButtonSize.Large;
            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程計畫"].Image = Properties.Resources.update_1;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["手動同步課程代碼API"].Image = Properties.Resources.update_1;

            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["手動同步課程代碼API"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程計畫"]["手動同步課程代碼API"].Click += delegate
            {
                frmSyncCourseCodeAPI scAPI = new frmSyncCourseCodeAPI();
                scAPI.ShowDialog();
            };

            //// 測試用，將學校們課程代碼寫入某台主機
            //MotherForm.RibbonBarItems["教務作業", "課程代碼"]["課程計畫"]["手動同步課程代碼API全"].Click += delegate
            //{
            //    frmCourseCodeTest fcc = new frmCourseCodeTest();
            //    fcc.ShowDialog();
            //};


            // 108 版本 -----
            Catalog ribbon10 = RoleAclSource.Instance["教務作業"]["基本設定"];
            ribbon10.Add(new RibbonFeature("C6063F22-FF7C-4AB5-82F8-4FC82B205B1A", "產生課程規劃"));

            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["產生課程規劃"].Enable = UserAcl.Current["C6063F22-FF7C-4AB5-82F8-4FC82B205B1A"].Executable;
            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["產生課程規劃"].BeginGroup = true;

            //MotherForm.RibbonBarItems["教務作業", "基本設定"]["產生課程規劃"].Image = Properties.Resources.online_library;
            //MotherForm.RibbonBarItems["教務作業", "基本設定"]["產生課程規劃"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["產生課程規劃"].Click += delegate
            {
                frmCreateGPlanMain108 fMain = new frmCreateGPlanMain108();
                fMain.ShowDialog();

            };


            Catalog ribbon11 = RoleAclSource.Instance["教務作業"]["基本設定"];
            ribbon11.Add(new RibbonFeature("B1E6A402-F173-42B5-8785-7250CF6D46BD", "班級課程規劃表(108課綱適用)"));

            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["班級課程規劃表(108課綱適用)"].Enable = UserAcl.Current["B1E6A402-F173-42B5-8785-7250CF6D46BD"].Executable;

            //MotherForm.RibbonBarItems["教務作業", "基本設定"]["班級課程規劃表(108課綱適用)"].Image = Properties.Resources.schedule;
            //MotherForm.RibbonBarItems["教務作業", "基本設定"]["班級課程規劃表(108課綱適用)"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["班級課程規劃表(108課綱適用)"].Click += delegate
            {
                frmGPlanConfig108 fgc = new frmGPlanConfig108();
                fgc.ShowDialog();

            };


            Catalog ribbon12 = RoleAclSource.Instance["班級"]["教務"]["班級開課"];
            ribbon12.Add(new RibbonFeature("375D9EC7-E4C2-46FB-AA5D-7ADE4A5F7F03", "依課程規劃表開課(108課綱適用)"));

            MotherForm.RibbonBarItems["班級", "教務"]["班級開課"]["依課程規劃表開課(108課綱適用)"].Enable = UserAcl.Current["375D9EC7-E4C2-46FB-AA5D-7ADE4A5F7F03"].Executable;

            MotherForm.RibbonBarItems["班級", "教務"]["班級開課"]["依課程規劃表開課(108課綱適用)"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {
                    frmCreateCourseByGPlan108 fcc = new frmCreateCourseByGPlan108(K12.Presentation.NLDPanels.Class.SelectedSource);
                    fcc.ShowDialog();
                }
            };

            Catalog ribbon13 = RoleAclSource.Instance["班級"]["教務"]["班級開課"];
            ribbon13.Add(new RibbonFeature("E5FD6FD7-7641-4E04-81B7-C2A5A1A346F8", "依課程規劃表開課-跨班(108課綱適用)"));

            MotherForm.RibbonBarItems["班級", "教務"]["班級開課"]["依課程規劃表開課-跨班(108課綱適用)"].Enable = UserAcl.Current["E5FD6FD7-7641-4E04-81B7-C2A5A1A346F8"].Executable;

            MotherForm.RibbonBarItems["班級", "教務"]["班級開課"]["依課程規劃表開課-跨班(108課綱適用)"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {
                    frmCreateCourseByGPlan108_C fcc = new frmCreateCourseByGPlan108_C(K12.Presentation.NLDPanels.Class.SelectedSource);
                    fcc.ShowDialog();
                }
            };


            Catalog ribbon_d1 = RoleAclSource.Instance["課程"]["功能按鈕"];
            ribbon_d1.Add(new RibbonFeature("358887CE-CCFA-4B7B-9F36-70CFDEAE88F9", "刪除課程與修課學生"));

            K12.Presentation.NLDPanels.Course.ListPaneContexMenu["刪除課程與修課學生"].Enable = FISCA.Permission.UserAcl.Current["358887CE-CCFA-4B7B-9F36-70CFDEAE88F9"].Executable;
            K12.Presentation.NLDPanels.Course.ListPaneContexMenu["刪除課程與修課學生"].Click += delegate {
                if (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0)
                {
                    frmDeleteCourseStudent dss = new frmDeleteCourseStudent();
                    dss.SetCourseIDs(K12.Presentation.NLDPanels.Course.SelectedSource);
                    dss.ShowDialog();                    
                }
            };
        }
    }
}
