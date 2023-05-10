using Campus.DocumentValidator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHGraduationWarning.ValidationRule.RowValidator
{
    public class SemsScoreCheckSubjectSubjectLevelVal : IRowVaildator
    {
        public SemsScoreCheckSubjectSubjectLevelVal()
        {

        }

        #region IRowVaildator 成員

        public string Correct(IRowStream Value)
        {
            return string.Empty;
        }

        public string ToString(string template)
        {
            return template;
        }

        public bool Validate(IRowStream Value)
        {
            bool retVal = true;
            if (Value.Contains("學生系統編號") && Value.Contains("學年度") && Value.Contains("學期") && Value.Contains("成績年級") && Value.Contains("科目名稱") && Value.Contains("科目級別"))
            {
                string key = Value.GetValue("學生系統編號") + "_" + Value.GetValue("學年度") + "_" + Value.GetValue("學期") + "_" + Value.GetValue("成績年級") + "_" + Value.GetValue("科目名稱") + "_" + Value.GetValue("科目級別");

                if (Utility._StudentSemesScoreSubjectLevelTemp.ContainsKey(key))
                {
                    // 有重覆 科目名稱+級別
                    if (Utility._StudentSemesScoreSubjectLevelTemp[key] > 1)
                        retVal = false;
                }

            }

            return retVal;
        }

        #endregion
    }
}
