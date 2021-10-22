using MaterialSkin;
using MaterialSkin.Controls;
using ModernGUI_V3;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AID2
{
    public partial class Form2 : Form //MaterialForm
    {
        comboBox9.Items.Clear();
                HashSet<string> distinct = new HashSet<string>();


                foreach (_Excel.Range cell in range.Cells)
                {
                    string value = (cell.Value2).ToString();

                    if (distinct.Add(value))
                        comboBox9.Items.Add(value);
                }
        public Form2()
        {
            InitializeComponent();

            //var materialSkinManager = MaterialSkinManager.Instance;
            //materialSkinManager.AddFormToManage(this);
            //materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;
            //materialSkinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);

            DateTime d = new DateTime();
            d = DateTime.Now;

            textBox2.Text = d.ToString("dd.MM.yyyy");
           
        }
        

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 form1 = new Form1();
            form1.ShowDialog();
            this.Close();
            

        }

        private void button4_MouseHover(object sender, EventArgs e)
        {
            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(this.button4, "General Service Control Valve 3-way");
        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(this.button1, "General Service Control Valve 2-way");
        }

        private void button2_MouseHover(object sender, EventArgs e)
        {
            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(this.button2, "Food & Pharma Sanitary Control Valve 2-way and 3-way");
        }

        private void button3_MouseHover(object sender, EventArgs e)
        {
            ToolTip ToolTip1 = new ToolTip();
            ToolTip1.SetToolTip(this.button3, "Self Actuated Direct Pressure Reducing Valve 2-way");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void CONTR_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();

            FormPrincipal formprincipal = new FormPrincipal();
            formprincipal.ShowDialog();
            this.Close();
            this.Close();
        }
    }
}
