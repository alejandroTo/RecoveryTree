using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using RecoveryTree.Controllers;
using System.Windows.Forms;

namespace RecoveryTree
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
           
            InitializeComponent();
            ControllerMain control = new ControllerMain(this);
        }

        private void buttonCreateTextBox_Click(object sender, EventArgs e)
        {

        }

        private void buttonCreateTree_Click(object sender, EventArgs e)
        {

        }
    }
}
