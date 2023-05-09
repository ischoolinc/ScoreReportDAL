using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Campus.DocumentValidator;
using Campus.Import2014;
using K12.Data;
using SHGraduationWarning.DAO;

namespace SHGraduationWarning.ImportExport
{
    public class ImportUpdateSubjectLevel : ImportWizard
    {
        // 設定檔
        private ImportOption _Option;

        public override ImportAction GetSupportActions()
        {
            //更新
            return ImportAction.Update;
        }


        public override string GetValidateRule()
        {
            // 資料驗證規格 XML 
            return Properties.Resources.ImportUpdateSubjectLevelVal;
        }

        public override string Import(List<Campus.DocumentValidator.IRowStream> Rows)
        {
            // 實作匯入資料
            // 比對待匯入Excel檔案內資料轉換取得 學生ID List

            Dictionary<string, List<StudUpdateSubjectLevelInfo>> StudUpdateSubjectLevelDict = new Dictionary<string, List<StudUpdateSubjectLevelInfo>>();
            foreach (IRowStream ir in Rows)
            {
                if (ir.Contains("學號") && ir.Contains("狀態"))
                {
                    string key = ir.GetValue("學號") + "_" + ir.GetValue("狀態");
                    if (Utility._AllStudentNumberStatusIDTemp.ContainsKey(key))
                    {
                        string StudentID = Utility._AllStudentNumberStatusIDTemp[key];
                        if (!StudUpdateSubjectLevelDict.ContainsKey(StudentID))
                            StudUpdateSubjectLevelDict.Add(StudentID, new List<StudUpdateSubjectLevelInfo>());

                        StudUpdateSubjectLevelInfo ss = new StudUpdateSubjectLevelInfo();
                        ss.StudentID = StudentID;
                        ss.StudentNumber = ir.GetValue("學號");
                        ss.StudentName = ir.GetValue("姓名");
                        ss.ClassName = ir.GetValue("班級");
                        ss.SeatNo = ir.GetValue("座號");
                        ss.SchoolYear = ir.GetValue("學年度");
                        ss.Semester = ir.GetValue("學期");
                        ss.GradeYear = ir.GetValue("成績年級");
                        ss.SubjectName = ir.GetValue("科目名稱");
                        ss.SubjectNameNew = ir.GetValue("新科目名稱");
                        ss.SubjectLevel = ir.GetValue("科目級別");
                        ss.SubjectLevelNew = ir.GetValue("新科目級別");

                        StudUpdateSubjectLevelDict[StudentID].Add(ss);
                    }

                }
            }


            int TotalCount = 0;


            // 需要修改資料


            // 資料整理
            foreach (IRowStream ir in Rows)
            {
                TotalCount++;
                this.ImportProgress = TotalCount;
                if (ir.Contains("學號") && ir.Contains("狀態"))
                {
                    // 判斷需要新增或修該
                    string key = ir.GetValue("學號") + "_" + ir.GetValue("狀態");


                    // 比對解析 StudentID 
                    if (Utility._AllStudentNumberStatusIDTemp.ContainsKey(key))
                    {
                        //    sid = Utility._AllStudentNumberStatusIDTemp[key];
                    }


                }
            }

            // 資料回寫             

            return "";
        }

        public override void Prepare(ImportOption Option)
        {
            // 畫面初始化後需要載入
            _Option = Option;

        }
    }
}
