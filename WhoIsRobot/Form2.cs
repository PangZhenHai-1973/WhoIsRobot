using System;
using System.Windows.Forms;

namespace WhoIsRobot
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Shown(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            timer2.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Close();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (Opacity == 1)
            {
                timer2.Enabled = false;
            }
            else
            {
                Opacity += 0.1;
            }
        }
    }
}
