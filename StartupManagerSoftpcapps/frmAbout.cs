using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace StartupManagerSoftpcapps
{
    public partial class frmAbout : CustomForm
    {        
        public frmAbout()
        {
            InitializeComponent();
        }
                     
        private void frmAbout_Load(object sender, EventArgs e)
        {
            lblAbout.Text = Module.ApplicationTitle + "\n\n" +
            "Developed by Alexander Triantafyllou\n" +
            "Copyright © 2024 - softpcapps Software\n";                       
            
            ullProductWebpage.Text = Module.ProductWebpageURL;                        
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }                        
    }
}