using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using System.Data;
using System.Xml.Linq;

namespace JHScoreReportDAL
{
    public class Config
    {
        private List<ConfigItem> _SubjectItemList = new List<ConfigItem>();
        private List<ConfigItem> _DomainItemList = new List<ConfigItem>();

        public Config()
        {

        }

        /// <summary>
        /// 取得設定資料
        /// </summary>
        public void GetConfigData()
        {
            string strSQL = "SELECT content FROM list WHERE  name = 'JHEvaluation_Subject_Ordinal' ";
            QueryHelper qh = new QueryHelper();
            try
            {
                DataTable dt = qh.Select(strSQL);
                if (dt.Rows.Count > 0)
                {
                    this._SubjectItemList.Clear();
                    this._DomainItemList.Clear();
                    char c = '"';
                    string rootXML = dt.Rows[0]["content"].ToString().Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;",c.ToString()).Replace("&apos;","'").Replace("&amp;","&");
                    XElement elmRoot = XElement.Parse(rootXML);
                    foreach(XElement elm1 in elmRoot.Elements("Configuration"))
                    {
                        if (elm1.Attribute("Name") != null)
                        {
                            // 科目
                            if (elm1.Attribute("Name").Value == "SubjectOrdinal")
                            {
                                foreach(XElement elmItem in elm1.Element("Subjects").Elements("Subject"))
                                {
                                    ConfigItem item = new ConfigItem();

                                    if (elmItem.Attribute("Group") != null)
                                    {
                                        item.Group = elmItem.Attribute("Group").Value;
                                    }

                                    if (elmItem.Attribute("Name") != null)
                                    {
                                        item.Name = elmItem.Attribute("Name").Value;
                                    }

                                    if (elmItem.Attribute("EnglishName") != null)
                                    {
                                        item.EnglishName = elmItem.Attribute("EnglishName").Value;
                                    }

                                    this._SubjectItemList.Add(item);
                                }

                            }

                            // 領域
                            if (elm1.Attribute("Name").Value == "DomainOrdinal")
                            {
                                foreach (XElement elmItem in elm1.Element("Domains").Elements("Domain"))
                                {
                                    ConfigItem item = new ConfigItem();

                                    if (elmItem.Attribute("Group") != null)
                                    {
                                        item.Group = elmItem.Attribute("Group").Value;
                                    }

                                    if (elmItem.Attribute("Name") != null)
                                    {
                                        item.Name = elmItem.Attribute("Name").Value;
                                    }

                                    if (elmItem.Attribute("EnglishName") != null)
                                    {
                                        item.EnglishName = elmItem.Attribute("EnglishName").Value;
                                    }

                                    this._DomainItemList.Add(item);
                                }
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<ConfigItem> GetSubjectItemList()
        {
            return _SubjectItemList;
        }

        public List<ConfigItem> GetDomainItemList()
        {
            return _DomainItemList;
        }
    }
}
