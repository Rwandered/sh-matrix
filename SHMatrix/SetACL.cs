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
    public partial class SetACL : Form
    {
        public SetACL()
        {
            InitializeComponent();
            this.Location = new Point(DataR.AbugX, DataR.AbugY);
            checkBox30.Checked = true;
            checkBox31.Checked = true;
        }

        private void SetACL_Load(object sender, EventArgs e)
        {
            this.Location = new Point(DataR.AbugX, DataR.AbugY);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void checkBox30_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox30.Checked ==true)
            {
                checkBox1.Enabled = false; 
                checkBox2.Enabled = false; 
                checkBox3.Enabled = false;
                checkBox4.Enabled = false;
                checkBox5.Enabled = false;
                checkBox6.Enabled = false;
                checkBox7.Enabled = false;
                checkBox8.Enabled = false;
                checkBox9.Enabled = false;
                checkBox10.Enabled = false;
                checkBox11.Enabled = false;
                checkBox12.Enabled = false;
                checkBox13.Enabled = false;
                checkBox14.Enabled = false;
                checkBox15.Enabled = false;
                checkBox16.Enabled = false;
                checkBox17.Enabled = false;
                checkBox18.Enabled = false;
                checkBox19.Enabled = false;
                checkBox18.Enabled = false;
                checkBox19.Enabled = false;
                checkBox20.Enabled = false;
                checkBox21.Enabled = false;
                checkBox22.Enabled = false;
                checkBox23.Enabled = false;
                checkBox24.Enabled = false;
                checkBox25.Enabled = false;
                checkBox26.Enabled = false;
                checkBox27.Enabled = false;
                checkBox28.Enabled = false;
             


                Clearbtn.Enabled = false;
                SetOKbtn.Enabled = false;
                cancelBtn.Enabled = false;
                checkBox29.Enabled = false;
                comboBox1.Enabled = false;
            }
            if (checkBox30.Checked == false)
            {
                checkBox1.Enabled = true;
                checkBox2.Enabled = true;
                checkBox3.Enabled = true;
                checkBox4.Enabled = true;
                checkBox5.Enabled = true;
                checkBox6.Enabled = true;
                checkBox7.Enabled = true;
                checkBox8.Enabled = true;
                checkBox9.Enabled = true;
                checkBox10.Enabled = true;
                checkBox11.Enabled = true;
                checkBox12.Enabled = true;
                checkBox13.Enabled = true;
                checkBox14.Enabled = true;
                checkBox15.Enabled = true;
                checkBox16.Enabled = true;
                checkBox17.Enabled = true;
                checkBox18.Enabled = true;
                checkBox19.Enabled = true;
                checkBox18.Enabled = true;
                checkBox19.Enabled = true;
                checkBox20.Enabled = true;
                checkBox21.Enabled = true;
                checkBox22.Enabled = true;
                checkBox23.Enabled = true;
                checkBox24.Enabled = true;
                checkBox25.Enabled = true;
                checkBox26.Enabled = true;
                checkBox27.Enabled = true;
                checkBox28.Enabled = true;


                Clearbtn.Enabled = true;
                SetOKbtn.Enabled = true;
                cancelBtn.Enabled = true;
                checkBox29.Enabled = true;
                comboBox1.Enabled = true; 
            }

        }

        private void checkBox31_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox31.Checked == true)
            {
                checkBox60.Enabled = false;
                checkBox59.Enabled = false;
                checkBox58.Enabled = false;
                checkBox48.Enabled = false;
                checkBox49.Enabled = false;
                checkBox50.Enabled = false;
                checkBox47.Enabled = false;
                checkBox46.Enabled = false;
                checkBox45.Enabled = false;
                checkBox35.Enabled = false;
                checkBox36.Enabled = false;
                checkBox37.Enabled = false;
                



                Clearbtn2.Enabled = false;
                Setbtn2.Enabled = false;
               Cancelbtn2.Enabled = false;
                checkBox32.Enabled = false;
                comboBox2.Enabled = false;
            }
            if (checkBox31.Checked == false)
            {
                checkBox60.Enabled = true;
                checkBox59.Enabled = true;
                checkBox58.Enabled = true;
                checkBox48.Enabled = true;
                checkBox49.Enabled = true;
                checkBox50.Enabled = true;
                checkBox47.Enabled = true;
                checkBox46.Enabled = true;
                checkBox45.Enabled = true;
                checkBox35.Enabled = true;
                checkBox36.Enabled = true;
                checkBox37.Enabled = true;



                Clearbtn2.Enabled = true;
                Setbtn2.Enabled = true;
                Cancelbtn2.Enabled = true;
                checkBox32.Enabled = true;
                comboBox2.Enabled = true;
            }
        }

        private void checkBox60_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox60.Checked)
            {
                if (checkBox47.Checked)
                {
                    #region Первый столбец
                   // checkBox60.Checked = true;
                    checkBox59.Checked = true;
                    checkBox58.Checked = true;
                    checkBox48.Checked = true;
                    checkBox49.Checked = true;
                    checkBox50.Checked = true;
                    #endregion
                    #region Второй  столбец
                    checkBox47.Checked = false;
                    checkBox46.Checked = false;
                    checkBox45.Checked = false;
                    checkBox35.Checked = false;
                    checkBox36.Checked = false;
                    checkBox37.Checked = false;
                    #endregion
                }

                #region Первый столбец
                // checkBox60.Checked = true;
                checkBox59.Checked = true;
                checkBox58.Checked = true;
                checkBox48.Checked = true;
                checkBox49.Checked = true;
                checkBox50.Checked = true;
                #endregion
                #region Второй  столбец
                checkBox47.Checked = false;
                checkBox46.Checked = false;
                checkBox45.Checked = false;
                checkBox35.Checked = false;
                checkBox36.Checked = false;
                checkBox37.Checked = false;
                #endregion
            }
        }

        private void checkBox47_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox47.Checked)
            {
                if (checkBox60.Checked)
                {
                    #region Первый столбец
                     checkBox60.Checked = false;
                    checkBox59.Checked = false;
                    checkBox58.Checked = false;
                    checkBox48.Checked = false;
                    checkBox49.Checked = false;
                    checkBox50.Checked = false;
                    #endregion
                    #region Второй  столбец
                    checkBox47.Checked = true;
                    checkBox46.Checked = true;
                    checkBox45.Checked = true;
                    checkBox35.Checked = true;
                    checkBox36.Checked = true;
                    checkBox37.Checked = true;
                    #endregion
                }

                    #region Первый столбец
                    checkBox60.Checked = false;
                    checkBox59.Checked = false;
                    checkBox58.Checked = false;
                    checkBox48.Checked = false;
                    checkBox49.Checked = false;
                    checkBox50.Checked = false;
                    #endregion
                    #region Второй  столбец
                    checkBox47.Checked = true;
                    checkBox46.Checked = true;
                    checkBox45.Checked = true;
                    checkBox35.Checked = true;
                    checkBox36.Checked = true;
                    checkBox37.Checked = true;
                    #endregion
            }
        }
    }
}
