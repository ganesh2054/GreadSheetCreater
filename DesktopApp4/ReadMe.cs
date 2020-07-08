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
    public partial class ReadMe : Form
    {
        Thread th;
        public ReadMe()
        {
            InitializeComponent();
        }

        private void ToolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
           
                this.Close();
                th = new Thread(opennewfo);
                th.SetApartmentState(ApartmentState.STA);
                th.Start();
            
         
        }
        private void opennewfo(object obj)
        {
            Application.Run(new MainPage());

        }
    }
}
