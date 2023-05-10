using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;   

namespace SHGraduationWarning.UIForm
{
    public partial class frmLoadUpdateSubjectVal : BaseForm
    {
        public frmLoadUpdateSubjectVal()
        {
            InitializeComponent();
        }

        private void frmLoadUpdateSubjectVal_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            
        }
    }
}
