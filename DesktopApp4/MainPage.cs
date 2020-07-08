using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
namespace DesktopApp4
{
    public partial class MainPage : Form
    {
        Thread th;
        public MainPage()
        {
            InitializeComponent();
        }

      
        private void ToolStripButton2_Click_1(object sender, EventArgs e)
        {
           
            this.Close();
            th = new Thread(opennewfor);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();

        }

        private void ToolStripButton1_Click_1(object sender, EventArgs e)
        {
           
            this.Close();
            th = new Thread(opennewform);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();

        }
        private void opennewform(object obj)
        {
            Application.Run(new Form1());

        }
        private void opennewfor(object obj)
        {
            Application.Run(new Form2());

        }

        private void ToolStripButton3_Click(object sender, EventArgs e)
        {
            this.Close();
            th = new Thread(opennewfo);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
        }
        private void opennewfo(object obj)
        {
            Application.Run(new ReadMe());

        }

        private void MainPage_Load(object sender, EventArgs e)
        {

        }
    }
}
