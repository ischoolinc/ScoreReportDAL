using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace SHGraduationWarning
{
    [FISCA.UDT.TableName("ischool.SHGraduationWarning_student.Configure")]
    public class Configure : FISCA.UDT.ActiveRecord
    {
        public Configure()
        {
          
        }
        /// <summary>
        /// 設定檔名稱
        /// </summary>
        [FISCA.UDT.Field]
        public string Name { get; set; }

        /// <summary>
        /// 列印樣板
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream { get; set; }
        public Aspose.Words.Document Template { get; set; }


        /// <summary>
        /// 目前使用        /// 
        /// </summary>
        [FISCA.UDT.Field]
        public bool Active { get; set; }

        /// <summary>
        /// 在儲存前，把資料填入儲存欄位中
        /// </summary>
        public void Encode()
        {
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            this.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
            this.TemplateStream = Convert.ToBase64String(stream.ToArray());
        }
        /// <summary>
        /// 在資料取出後，把資料從儲存欄位轉換至資料欄位
        /// </summary>
        public void Decode()
        {
            this.Template = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream)));
        }
    }
}
