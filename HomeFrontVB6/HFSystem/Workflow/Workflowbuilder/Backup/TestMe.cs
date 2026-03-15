using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoNotice
{
    public partial class TestMe : Form
    {
        public TestMe()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Sets/Gets Recipient Query
        /// </summary>
        public string Recipient
        {
            get
            {
                return txtTestingRecipient.Text;
            }
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            txtTestingRecipient.Text = "";
            this.Close();
        }

       
    }
}
