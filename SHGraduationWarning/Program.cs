using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using SHGraduationWarning.DAO;
using SHGraduationWarning.UIForm;
using Campus.DocumentValidator;
using static System.Windows.Forms.LinkLabel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SHGraduationWarning
{
    public class Program
    {
        static BackgroundWorker bgLoadUpdateSubjectLevelVal = new BackgroundWorker();

        [FISCA.MainMethod()]
        public static void main()
        {
            bgLoadUpdateSubjectLevelVal.DoWork += delegate (object sender, DoWorkEventArgs e)
            {
                bgLoadUpdateSubjectLevelVal.ReportProgress(30);
                // 取得學生學期成績科目與級別，驗證使用
                Utility._StudentSemesScoreSubjectLevelTemp = Utility.GetStudentSemsScoreSubjectLevelDict();
                bgLoadUpdateSubjectLevelVal.ReportProgress(100);
            };

            bgLoadUpdateSubjectLevelVal.RunWorkerCompleted += delegate (object sender, RunWorkerCompletedEventArgs e)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("");

                ImportExport.ImportUpdateSubjectLevel importUpdateSubjectLevel = new ImportExport.ImportUpdateSubjectLevel();
                importUpdateSubjectLevel.Execute();
                

            };
            bgLoadUpdateSubjectLevelVal.WorkerReportsProgress = true;
            bgLoadUpdateSubjectLevelVal.ProgressChanged += delegate (object sender, ProgressChangedEventArgs e)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("匯入更新學期科目級別驗證資料載入中...", e.ProgressPercentage);
            };

            // 教務作業>批次作業/檢視>成績作業>畢業預警
            Catalog ribbon1 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            ribbon1.Add(new RibbonFeature("6460BE9A-3E82-43C7-ACE7-64E0C6100473", "畢業預警與成績資料合理性檢查"));

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["畢業預警與成績資料合理性檢查"].Enable = UserAcl.Current["6460BE9A-3E82-43C7-ACE7-64E0C6100473"].Executable;

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["畢業預警與成績資料合理性檢查"].Click += delegate
            {
                // 畢業預警與成績資料合理性檢查
                frmMain fm = new frmMain();
                fm.ShowDialog();
            };

            Catalog ribbon2 = RoleAclSource.Instance["學生"]["功能按鈕"];
            ribbon2.Add(new RibbonFeature("3DE8EA91-A176-4100-8401-1FBA9F44786A", "匯入更新學期科目級別"));


            MotherForm.RibbonBarItems["學生", "資料統計"]["匯入"]["成績相關匯入"]["匯入更新學期科目級別"].Enable = FISCA.Permission.UserAcl.Current["3DE8EA91-A176-4100-8401-1FBA9F44786A"].Executable;
            MotherForm.RibbonBarItems["學生", "資料統計"]["匯入"]["成績相關匯入"]["匯入更新學期科目級別"].Click += delegate
            {
                // 取得學生學期成績科目與級別，驗證使用
                //   Utility._StudentSemesScoreSubjectLevelTemp = Utility.GetStudentSemsScoreSubjectLevelDict();

                //ImportExport.ImportUpdateSubjectLevel importUpdateSubjectLevel = new ImportExport.ImportUpdateSubjectLevel();
                //importUpdateSubjectLevel.Execute();
                //frmLoadUpdateSubjectVal fl = new frmLoadUpdateSubjectVal();
                //fl.ShowDialog();
                bgLoadUpdateSubjectLevelVal.RunWorkerAsync();
            };



            Catalog ribbons1 = RoleAclSource.Instance["學生"]["功能按鈕"];
            ribbons1.Add(new RibbonFeature("F2963209-300B-4C57-862F-A7E0374A5C3A", "檢查學生學期科目與課規課程代碼差異"));

            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["檢查學生學期科目與課規課程代碼差異"].Enable = FISCA.Permission.UserAcl.Current["F2963209-300B-4C57-862F-A7E0374A5C3A"].Executable;
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["檢查學生學期科目與課規課程代碼差異"].Click += delegate {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    UIForm.frmCheckStudGPlanSemsScoreCourseCode fc = new frmCheckStudGPlanSemsScoreCourseCode(K12.Presentation.NLDPanels.Student.SelectedSource);
                    fc.ShowDialog();
                }
            };


            // 載入自訂驗證規則
            #region 自訂驗證規則
            FactoryProvider.RowFactory.Add(new ValidationRule.SemsScoreRowValidatorFactory());
            #endregion

        }
    }
}
