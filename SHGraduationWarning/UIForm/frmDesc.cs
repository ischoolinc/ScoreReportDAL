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
    public partial class frmDesc :BaseForm
    {
        string _Desc = "";
        public frmDesc()
        {
            InitializeComponent();
        }
        
        public void SetDesc(string desc)
        {
            _Desc = desc;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void frmDesc_Load(object sender, EventArgs e)
        {            
            textDesc.Text = _Desc;
            textDesc.ReadOnly = true;
            textDesc.BackColor = Color.White;
        }
    }
}
