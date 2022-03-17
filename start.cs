using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace project
{
    public partial class start : Form
    {
        public start(int x, List<string> bg, List<string> dr, List<string> no_dr, List<string> end)
        {
            InitializeComponent();
            label2.Text = time_edit(x);

            foreach (var a in bg)
            {
                richTextBox1.AppendText(a + "\r\n");
            }

            
            foreach (var b in dr)
            {
                richTextBox1.SelectionColor = Color.Green;
                richTextBox1.AppendText(b + "\r\n");
                richTextBox1.SelectionColor = richTextBox1.ForeColor;
            }
            

            
            foreach (var c in no_dr)
            {
                richTextBox1.SelectionColor = Color.Red;
                richTextBox1.AppendText(c + "\r\n");
                richTextBox1.SelectionColor = richTextBox1.ForeColor;
            }
            

            foreach (var d in end)
            {
                richTextBox1.SelectionColor = Color.BlueViolet;
                richTextBox1.AppendText(d + "\r\n");
                richTextBox1.SelectionColor = richTextBox1.ForeColor;
            }

        }

        public string time_edit(int t)
        {

            if(t>100)
            {
                double m = Math.Round((double)t / 60);
                double s = t - m*60;
                return m + "m " + s + "s ";
            }


            return t.ToString()+"s ";
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/tpem-a/collect_to_pdf");
        }

        private void start_Load(object sender, EventArgs e)
        {

        }
    }
}
