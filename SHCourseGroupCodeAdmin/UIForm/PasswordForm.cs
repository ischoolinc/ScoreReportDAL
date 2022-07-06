using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Authentication;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class PasswordForm : BaseForm
    {
        private bool passwordPass = false;

        public PasswordForm()
        {
            InitializeComponent();
            passwordPass = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void PasswordForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            lblUserName.Text = DSAServices.UserAccount;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPassword.Text))
            {
                MsgBox.Show("請輸入密碼!");
                return;
            }
            
            try
            {
                passwordPass = DSAServices.ConfirmPassword(txtPassword.Text, null);
            }
            catch (Exception ex)
            {
                MsgBox.Show(FISCA.ErrorReport.Generate(ex));
                return;
            }

            if (passwordPass)
            {
                passwordPass = true;
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MsgBox.Show("密碼錯誤");
                return;
            }            
        }

        /// <summary>
        ///   取得密碼驗證是否通過
        /// </summary>
        /// <returns></returns>
        public bool GetPass()
        {
            return passwordPass;
        }
    }
}
