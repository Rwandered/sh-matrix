using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SHMatrix
{
    public partial class AbUG : Form
    {
        List<string> new_list = new List<string>();
        public AbUG(List<string> List3)
        {
            InitializeComponent();
            
            this.Location = new Point(DataR.AbugX, DataR.AbugY);
             textBox1.Visible = false;
            new_list = List3;

            foreach (string listEl in new_list)
            {
                textBox1.Text += listEl + Environment.NewLine;
            }

        }

        private void AbUG_Load(object sender, EventArgs e)
        {
            this.Location = new Point(DataR.AbugX, DataR.AbugY);
            textBox1.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new_list.Clear();
            textBox1.Clear();
            this.Close();
        }
    }
}
