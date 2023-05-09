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

                default:
                    return null;
            }
        }
    }
}
