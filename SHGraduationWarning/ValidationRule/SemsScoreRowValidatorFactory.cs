using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Campus.DocumentValidator;


namespace SHGraduationWarning.ValidationRule
{
    public class SemsScoreRowValidatorFactory : IRowValidatorFactory
    {
        public IRowVaildator CreateRowValidator(string typeName, XmlElement validatorDescription)
        {
            switch (typeName.ToUpper())
            {
                case "SEMSSCORECHECKSTUDENTNUMBERSTATUSVAL":
                    return new RowValidator.SemsScoreCheckStudentNumberStatusVal();

                    // 科目名稱+級別，與原成績比對
                case "SEMSSCORECHECKSUBJECTSUBJECTLEVELVAL":
                    return new RowValidator.SemsScoreCheckSubjectSubjectLevelVal();

                // 新科目名稱+新級別，與原成績比對
                case "SEMSSCORECHECKSUBJECTSUBJECTLEVELVALNEW":
                    return new RowValidator.SemsScoreCheckSubjectSubjectLevelValNew();

                default:
                    return null;
            }
        }
    }
}
