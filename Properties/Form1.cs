using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using _Excel = Microsoft.Office.Interop.Excel;
using FormSerialisation;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using Spire.Xls;
using _PrintExcel = Spire.Xls;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;
using Workbook = Spire.Xls.Workbook;
using SharpFluids;
using UnitsNet;

namespace GetResultFormulas
{

    
    public partial class Form1 : Form
    {
        public static class Util
        {
            public enum Effect { Roll, Slide, Center, Blend }

            public static void Animate(Control ctl, Effect effect, int msec, int angle)
            {
                int flags = effmap[(int)effect];
                if (ctl.Visible) { flags |= 0x10000; angle += 180; }
                else
                {
                    if (ctl.TopLevelControl == ctl) flags |= 0x20000;
                    else if (effect == Effect.Blend) throw new ArgumentException();
                }
                flags |= dirmap[(angle % 360) / 45];
                bool ok = AnimateWindow(ctl.Handle, msec, flags);
                if (!ok) throw new Exception("Animation failed");
                ctl.Visible = !ctl.Visible;
            }

            private static int[] dirmap = { 1, 5, 4, 6, 2, 10, 8, 9 };
            private static int[] effmap = { 0, 0x40000, 0x10, 0x80000 };

            [DllImport("user32.dll")]
            private static extern bool AnimateWindow(IntPtr handle, int msec, int flags);
        }

        public class Variables
        {
            public static double sound;
            public static double m3h;
            public static double kgh;
            public static int power;
            public static double kelvin1;
            public static double nm3h1;
            public static double p1;
            public static double p2;
            public static double sigma;
            public static double soundNorm;
            public static double soundMin;
            public static double maxvelocity;
            public static int NoiseAttenut;


            //public static double minp2;
            //public static double normp2;
            //public static double mint1;
            //public static double normt1;
            //public static double normnm3h;
            //public static double minnm3h;
            //public static double normm3h;
            //public static double minm3h;

        }
            public Form1()
        {
            InitializeComponent();
            filename = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            
            cmbFluid.Items.Add("Acetylene");
            cmbFluid.Items.Add("AIR");
            cmbFluid.Items.Add("AMMONIA");
            cmbFluid.Items.Add("ARGON");
            cmbFluid.Items.Add("BENZENE");
            cmbFluid.Items.Add("BUTANE");
            cmbFluid.Items.Add("CARBON DIOXIDE");
            cmbFluid.Items.Add("CARBON MONOXIDE");
            cmbFluid.Items.Add("CHLORINE");
            cmbFluid.Items.Add("DOWTHERM-A");
            cmbFluid.Items.Add("ETHANE");
            cmbFluid.Items.Add("ETHYLENE");
            cmbFluid.Items.Add("FLUORINE");
            cmbFluid.Items.Add("GLYCOL");
            cmbFluid.Items.Add("HELIUM");
            cmbFluid.Items.Add("HYDROGEN CHLORIDE");
            cmbFluid.Items.Add("Hydrogen Sulphide");
            cmbFluid.Items.Add("ISOBUTANE");
            cmbFluid.Items.Add("ISOBUTYLENE");
            cmbFluid.Items.Add("METHANE");
            cmbFluid.Items.Add("METHANOL");
            cmbFluid.Items.Add("NATURAL GAS");
            cmbFluid.Items.Add("Neon , Krypton");
            cmbFluid.Items.Add("NITROGEN");
            cmbFluid.Items.Add("Nitrogen (Nitric) Oxide");
            cmbFluid.Items.Add("NITROUS OXIDE");
            cmbFluid.Items.Add("OXYGEN");
            cmbFluid.Items.Add("PHOSGENE");
            cmbFluid.Items.Add("PROPANE");
            cmbFluid.Items.Add("PROPYLENE");
            cmbFluid.Items.Add("STEAM Saturated");
            cmbFluid.Items.Add("STEAM Superheated");
            cmbFluid.Items.Add("Sulphur Dioxide");
            cmbFluid.Items.Add("WATER");
            cmbFluid.Items.Add("Other :");
            cmbFluid.Items.Add("2-Phased Flow :");

            cmbState.Items.Add("Liquid");
            cmbState.Items.Add("Gas");
            cmbState.Items.Add("Steam Saturated");
            cmbState.Items.Add("Steam Superheated");










            cmbVBFD.Items.Add("Open");
            cmbVBFD.Items.Add("Close");





            cmbTC.Items.Add("Equal percentage");
            cmbTC.Items.Add("Linear");
            cmbTC.Items.Add("On-Off");
            cmbTC.Items.Add("Bi-Linear");
            cmbTC.Items.Add("Tri-Linear");
            cmbTC.Items.Add("Soecial");

            
            cmbTBU.Items.Add("RPTFE V-Rings");
            //cmbTBU.Items.Add("PTFE Rings [G]");
            cmbTBU.Items.Add("Graphite rings");
            cmbTBU.Items.Add("Special rings");
            //cmbTBU.Items.Add("Steel Rings [L+Sat.Steam]");
            //cmbTBU.Items.Add("Steel Rings [G]");
            cmbTBU.Items.Add("Unbalanced");






            cmbIPUnit1.Items.Add("kPa");
            cmbIPUnit1.Items.Add("Bar");
            cmbIPUnit1.Items.Add("psi");
            cmbIPUnit1.Items.Add("kg/cm2");
            cmbIPUnit1.Items.Add("mmH2O (Spec.)");
            cmbIPUnit1.Items.Add("mmHg (Spec.)");

            cmbIPUnit2.Items.Add("(g)");
            cmbIPUnit2.Items.Add("(a)");

            cmbITUnit.Items.Add("°F");
            cmbITUnit.Items.Add("°C");
            cmbITUnit.Items.Add("°K (Spec.)");

            cmbITSUnit.Items.Add("°F");
            cmbITSUnit.Items.Add("°C");
            cmbITSUnit.Items.Add("°K (Spec.)");

            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox5.SelectedIndex = 0;
            comboBox6.SelectedIndex = 0;




            cmbITSUnit.SelectedIndex = 1;
            cmbITUnit.SelectedIndex = 1;
            cmbIPUnit2.SelectedIndex = 1;
            cmbIPUnit1.SelectedIndex = 1;
            cmbFluid.SelectedIndex = -1;
            AcceptableNoise.SelectedIndex = 3;

            cmbState.SelectedIndex = 0;
            cmbTBU.SelectedIndex = 3;
            cmbTC.SelectedIndex = 0;

            //cmbTType.SelectedIndex = 0;
            cmbUnit.SelectedIndex = 0;
            //cmbVBFD.SelectedIndex = 1;



        }

        private void Button2_Click(object sender, EventArgs e)
        {
            
            Cursor.Current = Cursors.WaitCursor;

            Fluid fluid;


            if (cmbFluid.SelectedItem.ToString() == "AMMONIA")
            {
                fluid = new Fluid(FluidList.Ammonia);
            }
            else if (cmbFluid.SelectedItem.ToString() == "OXYGEN")
            {
                fluid = new Fluid(FluidList.Oxygen);
            }
            else if (cmbFluid.SelectedItem.ToString() == "DOWTHERM-A")
            {
                fluid = new Fluid(FluidList.InCompDowthermJ);
            }
            else if (cmbFluid.SelectedItem.ToString() == "METHANOL")
            {
                fluid = new Fluid(FluidList.Methanol);
            }
            else if (cmbFluid.SelectedItem.ToString() == "BENZENE")
            {
                fluid = new Fluid(FluidList.Benzene);
            }
            else if (cmbFluid.SelectedItem.ToString() == "BUTANE")
            {
                fluid = new Fluid(FluidList.IsoButane);
            }
            else if (cmbFluid.SelectedItem.ToString() == "ETHANE")
            {
                fluid = new Fluid(FluidList.Ethane);
            }
            else if (cmbFluid.SelectedItem.ToString() == "FLOURINE")
            {
                fluid = new Fluid(FluidList.Fluorine);
            }
            else if (cmbFluid.SelectedItem.ToString() == "GLYCOL")
            {
                fluid = new Fluid(FluidList.MixEthyleneGlycolAQ);
            }
            else if (cmbFluid.SelectedItem.ToString() == "HYDROGEN SULPHIDE")
            {
                fluid = new Fluid(FluidList.HydrogenSulfide);
            }
            else if (cmbFluid.SelectedItem.ToString() == "ISOBUTANE")
            {
                fluid = new Fluid(FluidList.IsoButane);
            }
            else if (cmbFluid.SelectedItem.ToString() == "ISOBUTYLENE")
            {
                fluid = new Fluid(FluidList.IsoButene);
            }
            else if (cmbFluid.SelectedItem.ToString() == "PROPYLENE")
            {
                fluid = new Fluid(FluidList.Propylene);
            }
            else if (cmbFluid.SelectedItem.ToString() == "HYDROGEN")
            {
                fluid = new Fluid(FluidList.Hydrogen);
            }
            else if (cmbFluid.SelectedItem.ToString() == "CARBON DIOXIDE")
            {
                fluid = new Fluid(FluidList.CO2);
            }
            else if (cmbFluid.SelectedItem.ToString() == "METHANE")
            {
                fluid = new Fluid(FluidList.Methane);
            }
            else if (cmbFluid.SelectedItem.ToString() == "NATURAL GAS")
            {
                fluid = new Fluid(FluidList.Methane);
            }
            else if (cmbFluid.SelectedItem.ToString() == "AIR")
            {
                fluid = new Fluid(FluidList.Air);
            }
            else if (cmbFluid.SelectedItem.ToString() == "ARGON")
            {
                fluid = new Fluid(FluidList.Argon);
            }
            else if (cmbFluid.SelectedItem.ToString() == "BUTANE")
            {
                fluid = new Fluid(FluidList.nButane);
            }
            else if (cmbFluid.SelectedItem.ToString() == "CARBON MONOXIDE")
            {
                fluid = new Fluid(FluidList.CO);
            }
            else if (cmbFluid.SelectedItem.ToString() == "HELIUM")
            {
                fluid = new Fluid(FluidList.Helium);
            }

            else if (cmbFluid.SelectedItem.ToString() == "Neon")
            {
                fluid = new Fluid(FluidList.Neon);
            }
            else if (cmbFluid.SelectedItem.ToString() == "NITROGEN")
            {
                fluid = new Fluid(FluidList.Nitrogen);
            }
            else if (cmbFluid.SelectedItem.ToString() == "Nitrogen (Nitric) Oxide")
            {
                fluid = new Fluid(FluidList.NitrousOxide);
            }
            

            else if (cmbFluid.SelectedItem.ToString() == "PROPANE")
            {
                fluid = new Fluid(FluidList.nPropane);
            }
            else if (cmbFluid.SelectedItem.ToString() == "PROPYLENE")
            {
                fluid = new Fluid(FluidList.Propylene);
            }
            else if (cmbFluid.SelectedItem.ToString() == "Sulphur Dioxide")
            {
                fluid = new Fluid(FluidList.SulfurDioxide);
            }

            else if (cmbFluid.SelectedItem.ToString() == "WATER"|| cmbFluid.SelectedItem.ToString() == "STEAM Saturated"|| cmbFluid.SelectedItem.ToString() == "STEAM Superheated")
            {
                fluid = new Fluid(FluidList.Water);
            }


            else
            {
                fluid = new Fluid(FluidList.Ethane);
            }


            fluid.UpdatePT(Pressure.FromBars(int.Parse(txtIPMaxF.Text)), Temperature.FromDegreesCelsius(250));
            MessageBox.Show("Density of " +cmbFluid.SelectedItem.ToString()  + fluid.Density +" "+ fluid.Viscosity);

            //Console.WriteLine("Density of water at 13°C: " + Water.Density);

            //Console.ReadLine();
            
            txtCPress.Text = "";
            textBox6.Text = "";
            txtAPSMaxF.Text = "";
            txtAPSNF.Text = "";
            txtAPSMinF.Text = "";
            //txtFRNF.Text = "";
            //txtSWMWNF.Text = "";
            //txtVSHRNF.Text = "";
            //txtIPNF.Text = "";
            //txtRFCMaxF.Text = "";
            //txtRFCNF.Text = "";
            //txtRFCMinF.Text = "";
            //txtVVNF.Text = "";
            //txtVVMinF.Text = "";
            //txtVVMaxF.Text = "";
            //txtLPD.Text = "";
            //txtAMM.Text = "";
            //txtOPNF.Text = "";
            //txtITNF.Text = "";
            //txtFRMinF.Text = "";
            //txtIPMinF.Text = "";
            //txtOPMinF.Text = "";
            //txtITMinF.Text = "";
            //txtRFCMinF.Text = "";
            //txtRFCMinF.Text = "";
            //txtRFCMaxF.Text = "";




            Application excel = new Application();
            

            _Excel.Workbook workbook = excel.Workbooks.Open(filename + "\\Demo.xlsx", ReadOnly: false, Editable: true);
            _Excel.Worksheet worksheet = workbook.Worksheets["SELECTION"] as _Excel.Worksheet;
            _Excel.Range range;
            _Excel.Range rangecv;
            _Excel.Range rangebonnet;
            _Excel.Range rangepacking;
            _Excel.Range rangeseat;
            //_Excel.Range pneumatic;
            //_Excel.Range electric;
            //_Excel.Range balancing;




            if (worksheet == null)
                return;


            try
            {

                Range row1 = worksheet.Rows.Cells[10, 5];
                row1.Value = cmbFluid.SelectedItem.ToString();

                Range row2 = worksheet.Rows.Cells[10, 20];
                row2.Value = cmbState.SelectedItem.ToString();

                Range row3 = worksheet.Rows.Cells[12, 12];
                row3.Value = cmbUnit.SelectedItem.ToString();

                Range row4 = worksheet.Rows.Cells[12, 18];
                row4.Value = txtFRMaxF.Text;

                Range row5 = worksheet.Rows.Cells[12, 24];
                row5.Value = txtFRNF.Text;

                Range row6 = worksheet.Rows.Cells[12, 30];
                row6.Value = txtFRMinF.Text;

                Range row7 = worksheet.Rows.Cells[12, 36];
                row7.Value = txtFRSO.Text;

                Range row8 = worksheet.Rows.Cells[14, 18];
                row8.Value = txtIPMaxF.Text;

                Range row9 = worksheet.Rows.Cells[14, 24];
                row9.Value = txtIPNF.Text;

                Range row10 = worksheet.Rows.Cells[14, 30];
                row10.Value = txtIPMinF.Text;

                Range row11 = worksheet.Rows.Cells[15, 18];
                row11.Value = txtOPMaxF.Text;

                Range row12 = worksheet.Rows.Cells[15, 24];
                row12.Value = txtOPNF.Text;

                Range row13 = worksheet.Rows.Cells[15, 30];
                row13.Value = txtOPMinF.Text;

                Range row14 = worksheet.Rows.Cells[16, 18];
                row14.Value = txtITMaxF.Text;

                Range row15 = worksheet.Rows.Cells[16, 24];
                row15.Value = txtITNF.Text;

                Range row16 = worksheet.Rows.Cells[16, 30];
                row16.Value = txtITMinF.Text;


                Range row30 = worksheet.Rows.Cells[47, 8];
                row30.Value = cmbTC.SelectedItem.ToString();

                //Range row31 = worksheet.Rows.Cells[48, 10];
                //row31.Value = cmbTBU.SelectedItem.ToString();

                //Range row25 = worksheet.Rows.Cells[38, 8];
                //row25.Value = cmbVBFD.SelectedItem.ToString();

                //Range row28 = worksheet.Rows.Cells[45, 6];
                //row28.Value = cmbTType.SelectedItem.ToString();

                //Range row29 = worksheet.Rows.Cells[46, 6];
                //row29.Value = textBox10.Text;

                





                Range row38 = worksheet.Rows.Cells[14, 12];
                row38.Value = cmbIPUnit1.SelectedItem.ToString();

                Range row39 = worksheet.Rows.Cells[14, 16];
                row39.Value = cmbIPUnit2.SelectedItem.ToString();

                Range row40 = worksheet.Rows.Cells[16, 12];
                row40.Value = cmbITUnit.SelectedItem.ToString();

                Range row41 = worksheet.Rows.Cells[17, 12];
                row41.Value = cmbITSUnit.SelectedItem.ToString();

                Range row43 = worksheet.Rows.Cells[17, 18];
                Range row45 = worksheet.Rows.Cells[17, 30];
                Range row44 = worksheet.Rows.Cells[17, 24];
                if (cmbState.SelectedIndex ==2)
                {
                    txtITSMaxF.Text = isValidS(worksheet.Rows.Cells[109, 43]);
                    txtITSNF.Text = isValidS(worksheet.Rows.Cells[110, 43]);
                    txtITSMinF.Text = isValidS(worksheet.Rows.Cells[111, 43]);
                }
                else
                {
                    
                    
                }
                row43.Value = txtITSMaxF.Text;

                
                row44.Value = txtITSNF.Text;

                
                row45.Value = txtITSMinF.Text;

                txtCPress.Text = isValidS(worksheet.Rows.Cells[10, 30]);
               //textBox10.Text = (Math.Round(double.Parse(txtCPress.Text), 2)*2).ToString();
                txtSWMWNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[18, 24])), 2).ToString();
                txtVSHRNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[19, 24])), 2).ToString();
                txtRFCMaxF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[21, 18])),2).ToString();
                txtRFCNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[21, 24])), 2).ToString();
                txtRFCMinF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[21, 30])), 3).ToString();
                //cmbTSize.Text = isValidS(worksheet.Rows.Cells[46, 6]);
                txtVVMaxF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[24, 18])), 2).ToString();
                txtVVNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[24, 24])), 2).ToString();
                txtVVMinF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[24, 30])), 3).ToString();
                txtAPSNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[23, 27])), 2).ToString();
                txtAPSMaxF.Text = Math.Round(double.Parse((Variables.sound).ToString()), 2).ToString();
                Variables.soundNorm = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[23, 27])), 2);
                txtAPSMinF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[23, 33])), 3).ToString();
                Variables.soundMin = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[23, 33])), 2);
                txtLPD.Text = isValidS(worksheet.Rows.Cells[26, 13]);
                txtAMM.Text = worksheet.Rows.Cells[27, 26].Value.ToString();
                textBox6.Text = worksheet.Rows.Cells[29, 16].Value.ToString();
                textBox5.Text = worksheet.Rows.Cells[26, 28].Value.ToString();
                textBox1.Text = isValidS(worksheet.Rows.Cells[75, 43]);
                textBox2.Text = isValidS(worksheet.Rows.Cells[76, 43]);
                shutoff.Text = worksheet.Rows.Cells[113, 43].Value.ToString();
                if (int.Parse(shutoff.Text)>100)
                {
                    MessageBox.Show("Pressure is too high \nIt's Recommended that you contact the factory", "Pressure Warning");
                }


                Variables.maxvelocity = worksheet.Rows.Cells[71, 43].value;
                //textBox4.Text = worksheet.Rows.Cells[72, 43].Value.ToString();
                textBox8.Text = worksheet.Rows.Cells[78, 43].Value.ToString();
                textBox9.Text = worksheet.Rows.Cells[79, 43].Value.ToString();
                textBox7.Text = worksheet.Rows.Cells[81, 43].Value.ToString();
                string f = worksheet.Rows.Cells[82, 43].Value.ToString();
                string g = worksheet.Rows.Cells[83, 43].Value.ToString();
                string h = worksheet.Rows.Cells[84, 43].Value.ToString();
                
                Variables.kelvin1 = worksheet.Rows.Cells[88, 43].Value;
                Variables.nm3h1 = worksheet.Rows.Cells[89, 43].Value;
                Variables.p1 = worksheet.Rows.Cells[90, 43].Value;
                Variables.p2 = worksheet.Rows.Cells[91, 43].Value;
                Variables.sound = worksheet.Rows.Cells[23, 21].Value;
                Variables.m3h = worksheet.Rows.Cells[86, 43].Value;
                Variables.kgh = worksheet.Rows.Cells[87, 43].Value;
                Variables.sigma = worksheet.Rows.Cells[92, 43].Value;
                textBox14.Text= worksheet.Rows.Cells[92, 43].Value.ToString();




                range = worksheet.get_Range("AQ105", "AQ108") as _Excel.Range;
                rangecv = worksheet.get_Range("AQ96", "AQ99") as _Excel.Range;
                rangepacking = worksheet.get_Range("AU72", "AU77") as _Excel.Range;
                rangeseat = worksheet.get_Range("AR96", "AR99") as _Excel.Range;
                rangebonnet = worksheet.get_Range("AU81", "AU85") as _Excel.Range;
                //pneumatic = worksheet.get_Range("AS102", "AS105") as _Excel.Range;
                //electric = worksheet.get_Range("AS107", "AS110") as _Excel.Range;
                //balancing = worksheet.get_Range("AU102", "AU105") as _Excel.Range;

                // Loop through the cells in the Range and add their values to
                // the combo box
                comboBox9.Items.Clear();
                HashSet<string> distinct = new HashSet<string>();


                foreach (_Excel.Range cell in range.Cells)
                {
                    string value = (cell.Value2).ToString();

                    if (distinct.Add(value))
                        comboBox9.Items.Add(value);
                }
                comboBox9.SelectedIndex = 2;


                //ELECTRIC RANGE
                //if (int.Parse(shutoff.Text) > 111)
                //{
                //    comboBox17.Items.Clear();
                //    comboBox17.Items.Add("RegadaST2 + Balancing");

                //}
                //else
                //{
                //    comboBox17.Items.Clear();
                //    foreach (_Excel.Range cell in electric.Cells)
                //    {
                //        string value = (cell.Value2).ToString();

                //        if (distinct.Add(value))
                //            comboBox17.Items.Add(value);

                //        //comboBox17.Items.Add((cell.Value2).ToString() as string);
                //    }
                //}
                //comboBox17.SelectedIndex = 0;


                //////BALANCING RANGE
                //comboBox18.Items.Clear();
                //foreach (_Excel.Range cell in balancing.Cells)
                //{

                //    comboBox18.Items.Add((cell.Value2).ToString() as string);
                //}

                ////PNEUMATIC RANGE
                //if (comboBox18.Items[0].ToString() == "NAA" /*&& int.Parse(comboBox9.SelectedItem.ToString()) > 49*/)
                //{
                //    comboBox16.Items.Clear();
                //    comboBox16.Items.Add("S500B-1250 (195) 0,8-2,0 (12-30)");
                //    checkBox1.Checked = true;
                //    comboBox16.SelectedIndex = 0;
                //    if (Variables.kelvin1 > 73.15 && Variables.kelvin1 <= 473.15/* && checkBox1.Checked == true comboBox18.Items[0].ToString() == "NAA"*/)
                //    {
                //        cmbTBU.SelectedIndex = 0;
                //    }

                //    else if (Variables.kelvin1 > 73.15 && Variables.kelvin1 <= 673.15/* && checkBox1.Checked == true comboBox18.Items[0].ToString() == "NAA"*/)
                //    {
                //        cmbTBU.SelectedIndex = 1;
                //    }
                //    else
                //    {
                //        cmbTBU.SelectedIndex = 3;
                //    }

                //}
                //else if (comboBox18.Items[0].ToString() == "NAA" && int.Parse(comboBox9.SelectedItem.ToString()) < 49)
                //{
                //    comboBox16.Items.Clear();
                //    comboBox16.Items.Add("Consult Factory");
                //    checkBox1.Checked = false;
                //    comboBox16.SelectedIndex = 0;
                //}
                //else
                //{
                //    comboBox16.Items.Clear();
                //    checkBox1.Checked = false;
                //    foreach (_Excel.Range cell in pneumatic.Cells)
                //    {
                //        comboBox16.Items.Add((cell.Value2).ToString() as string);
                //    }
                //    comboBox16.SelectedIndex = 0;
                //}
                //comboBox16.Items.Clear();
                //foreach (_Excel.Range cell in pneumatic.Cells)
                //{

                //    comboBox16.Items.Add((cell.Value2).ToString() as string);
                //}

                //RANGE SEATBORE
                comboBox15.Items.Clear();
                foreach (_Excel.Range cell in rangeseat.Cells)
                {

                    comboBox15.Items.Add((cell.Value2).ToString() as string);
                }


                //comboBox11.Items.Clear();
                //foreach (_Excel.Range cell in rangecv.Cells)
                //{
                    
                //    comboBox11.Items.Add((cell.Value2).ToString() as string);
                //}

                comboBox12.Items.Clear();
                foreach (_Excel.Range cell in rangepacking.Cells)
                {
                    string value = (cell.Value2).ToString();
                    if (distinct.Add(value))
                        comboBox12.Items.Add(value);
                }

                comboBox8.Items.Clear();
                foreach (_Excel.Range cell in rangebonnet.Cells)
                {
                    string value = (cell.Value2).ToString();
                    if (distinct.Add(value))
                        comboBox8.Items.Add(value);
                }

                //comboBox9.SelectedItem = (textBox1.Text);
                
                


                //richTextBox1.Text = (f + "\nPARTIAL CAV f = " + g + "\nMAX CAV f = " + h).ToString();


                if (h == "MAXIMUM CAVITATION")
                {
                    cmbVBFD.SelectedIndex = 1;
                    richTextBox1.Text = (f + "\n " + h).ToString();
                }
                else
                {
                    cmbVBFD.SelectedIndex = 0;
                    if (g=="PARTIAL CAVITATION")
                    {
                        richTextBox1.Text = (f + "\n" + g).ToString();
                    }
                    else
                    {
                        richTextBox1.Text = (f).ToString();
                    }
                    
                }
                if (checkBox2.Checked == true && Variables.power < 50)
                {
                    Variables.sound -= 5;
                    Variables.NoiseAttenut = 5;
                }
                else if (checkBox2.Checked == true && Variables.power > 50 && Variables.power <= 100)
                {
                    Variables.sound -= 8;
                    Variables.NoiseAttenut = 8;
                }
                else if (checkBox2.Checked == true && Variables.power > 100)
                {
                    Variables.sound -= 12;
                    Variables.NoiseAttenut = 12;
                }
                if (int.Parse(AcceptableNoise.SelectedItem.ToString())< Variables.sound)
                {
                    richTextBox1.Select(richTextBox1.TextLength, 0);
                    richTextBox1.SelectionColor = Color.Red;
                    richTextBox1.AppendText("\nNOISE TOO HIGH");
                }
                else
                {
                    richTextBox1.Select(richTextBox1.TextLength, 0);
                    richTextBox1.SelectionColor = Color.Green;
                    richTextBox1.AppendText("\nNOISE IS OK");
                }
                textBox13.Text = f;
                if (textBox13.Text == "CHOCKED FLOW ! ! ! !")
                {
                    //textBox13.Text = "Chocked Flow";
                    textBox13.BackColor = Color.Red;
                }
                else
                {
                    //textBox13.Text = "NOT Chocked";
                    textBox13.BackColor = Color.Green;
                }




                /*              double d = double.Parse(textBox1.Text);
                                double Qs = double.Parse(txtFRMaxF.Text); //Flow rate unit:m3/h
                                double W = 0;//Qs * .04475 * 44.01; //Mass flow rate W=Qs*N8/N9*M unit:Kg/h

                                double P1 = double.Parse(txtIPMaxF.Text); //Inlet Pressure unit:kPa;
                                double P2 = double.Parse(txtOPMaxF.Text); //Outlet pressure unit:kPa;
                                double T1 = double.Parse(txtITMaxF.Text); //Inlet temperature, unit:K;
                                double D1 = double.Parse(textBox1.Text); //Inlet pipe size, unit:mm
                                double D2 = double.Parse(textBox1.Text); //Outlet pipe size, unit:mm

                                //string Fluid; //fluid type, incompressible or compressible
                                              //double Density = 8.389; //Fluid density, unit kg/m3
                                              //double DensitySTD = 1978; //fluid density at standard condition unit kg/m3
                                double Z1 = .991; //Compressibility, unit kPa;
                                double Zs = 0.994; //Standard compressibility, unit kPa;
                                double Ts = 273; // unit K
                                double Ps = 101.325; // unit kPa
                                double v = .000002526; //Kinematic viscosity, unit: m2/s;
                                double M = 17.38; //Molecular mass;
                                double Gama = 1.31; //Specific heat ration


                                double Fl = .85; //Liquid pressure recovery factor
                                double Fd = .72; //Valve style modifier
                                double Xt = .619;
                                double Kv = double.Parse(textBox2.Text); // Valve flow rate
                                double Cv;

                                double N2 = 0.0016;
                                double N4 = 0.0707;
                                double N5 = 0.0018;
                                double N8 = 1.1;
                                double N9 = 24.6;
                                double N18 = 0.865;
                                double n;//trim factor for non-turbulent flow;
                                double N32 = 140;
                                double N22 = 17.3;
                                double C = Kv;//Valve rated flow rate

                                #region Piping geometry factor, Fp
                                double Fp; //Piping geometry factor
                                double K1;
                                double K2; //Head loss coefficient:
                                double KB1;
                                double KB2;
                                double K;
                                K1 = 0.5 * (1 - d * d / (D1 * D1)) * (1 - d * d / (D1 * D1)); //
                                K2 = (1 - d * d / D2 / D2) * (1 - d * d / D2 / D2);
                                KB1 = 1 - d * d * d * d / D1 / D1 / D1 / D1;
                                KB2 = 1 - d * d * d * d / D2 / D2 / D2 / D2;
                                K = K1 + K2 + KB1 - KB2;
                                //Console.WriteLine("K1 = {0}\nK2 = {1}\nKB1 = {2}\nKB2 = {3}", K1, K2, KB1, KB2);
                                Fp = 1 / Math.Sqrt(1 + (K / N2) * (C / d / d) * (C / d / d));
                                richTextBox1.Text = ("Fp = \n" + Fp).ToString();
                                //Console.WriteLine("Fp = " + Fp);
                                #endregion

                                #region combined liquid pressure recovery factor Flp
                                double Flp = Fl / Math.Sqrt(1 + Fl * Fl / N2 * (K1 + KB1) * (C / d / d) * (C / d / d));
                                richTextBox1.Text = ("Fp = {0}\nFlp = {1}", Fp, Flp).ToString();
                                //Console.WriteLine("Fp = {0}\nFlp = {1}", Fp, Flp);
                                #endregion

                                #region Estimated pressure differential ratio factor with attached fittings, Xtp
                                double Xtp = (Xt / Fp / Fp) / (1 + Xt * (K1 + KB1) / N5 * (C / d / d) * (C / d / d));
                                //Console.WriteLine("Xtp = {0}", Xtp);
                                //Console.WriteLine("Xt = " + Xt);
                                #endregion

                                #region Specific heat ratio factor, Fgama
                                double Fgama = Gama / 1.4;
                                //Console.WriteLine("Fgama = {0}", Fgama);
                                #endregion

                                #region choked or not?
                                double Xchocked;
                                Xchocked = Fgama * Xtp;
                                double DeltaP;
                                DeltaP = P1 - P2;
                                double X;
                                X = DeltaP / P1;
                                double Xsizing;

                                //Console.WriteLine("X = {0} kPa\nXChoked = {1} kPa", X, Xchocked);
                                if (X < Xchocked)
                                {
                                    Xsizing = X;
                                    richTextBox1.Text = ("flow is not choked").ToString();
                                    //Console.WriteLine("flow is not choked");
                                }
                                else
                                {
                                    Xsizing = Xchocked;
                                    richTextBox1.Text = ("Alarm!! flow choked\n").ToString();
                                    //Console.WriteLine("Alarm!! flow choked");
                                }
                                //Console.WriteLine("Xsizing = {0} kPa", Xsizing);
                                #endregion

                                #region Expansion factor, Y
                                double Y1 = 1 - Xsizing / (3 * Xchocked);
                                double Y2 = Xchocked * 2 / 3;
                                double Y;
                                if (Xsizing == Xchocked)
                                {
                                    if (Y1 > Y2)
                                    {
                                        Y = Y1;
                                    }
                                    else
                                    {
                                        Y = Y2;
                                    }
                                }
                                else
                                {
                                    Y = Y1;
                                }
                                //Console.WriteLine("Y = {0}", Y);
                                #endregion

                                #region C calculation
                                double Ccal;

                                if (W == 0)
                                {
                                    Ccal = Qs / N9 / Fp / P1 / Y * Math.Sqrt(M * T1 * Z1 / Xsizing);
                                }
                                else
                                {
                                    Ccal = W / N8 / Fp / P1 / Y * Math.Sqrt(T1 * Z1 / M / Xsizing);
                                }
                                textBox10.Text = ("" + Ccal);
                                //Console.WriteLine("Kv = " + Ccal + " m3/h");

                                double Q;
                                Q = Qs * (Ps / Zs / Ts) * (Z1 * T1 / P1);

                                //Console.WriteLine("Q = " + Q + "m3/h");
                                if (Ccal < C)
                                {
                                    //richTextBox1.Text = ("The calculated Cv is less than the Valve max Cv");
                                    //Console.WriteLine("The calculated Cv is less than the Valve max Cv");
                                }
                                else
                                {
                                    richTextBox1.Text = ("Error! The calculated Cv is larger than the Valve max Cv, this valve should not be considered\n");
                                    //Console.WriteLine("Error! The calculated Cv is larger than the Valve max Cv, this valve should not be considered");
                                }
                                #endregion

                                #region verify the result of the scop
                                double result = Ccal / N18 / d / d;

                                if (result < 0.047)
                                {
                                    Console.WriteLine("result ={0} is in scope", result);
                                }
                                else
                                {
                                    Console.WriteLine("Error! result={0} is not in scope", result);
                                }
                                #endregion
                                #region Rev Reynolds Number
                                double Rev = (N4 * Fd * Q) / (v * Math.Sqrt(Ccal * Fl)) * Math.Sqrt(Math.Sqrt(Fl * Fl * Ccal * Ccal / N2 / d / d / d / d + 1));
                                //Console.WriteLine("Rev = " + Rev);
                                if (Rev >= 10000)
                                {
                                    richTextBox1.Text = ("Flow is turbulent\n");
                                    //Console.WriteLine("Flow is turbulent");
                                    richTextBox1.Text = ("Flow Rate Kv = {0} \nFlow Rate Cv = {1} ", Ccal, Ccal * 1.156).ToString();
                                }
                                else if (Rev >= 10)
                                {
                                    richTextBox1.Text = ("Alarm! Flow is transitional");

                                    #region calculate n
                                    if (C / d / d / N18 >= 0.016)
                                    {
                                        n = N2 / (C / d / d) / (C / d / d);
                                    }
                                    else
                                    {
                                        n = 1 + N32 * Math.Pow(C / d / d, 2 / 3);
                                    }
                                    #endregion

                                    #region calculate Fr
                                    double Fr1;
                                    double Fr2;
                                    Fr1 = 1 + (0.33 * Math.Sqrt(Fl) / Math.Pow(n, 1 / 4)) * Math.Log10(Rev / 10000);
                                    Fr2 = 0.026 / Fl * Math.Sqrt(n * Rev);
                                    double Fr;

                                    if (Fr1 < Fr2)
                                    {
                                        Fr = Fr1;
                                    }
                                    else
                                    {
                                        Fr = Fr2;
                                    }
                                    if (Fr > 1)
                                    {
                                        Fr = 1;
                                    }
                                    else { }
                                    #endregion

                                    #region calculate Y
                                    if (Rev >= 1000 && Rev < 10000)
                                    {
                                        Y = (Rev - 1000) / 9000 * (1 - Xsizing / 3 / Xchocked - Math.Sqrt(1 - X / 2)) + Math.Sqrt(1 - X / 2);
                                    }
                                    else
                                    {
                                        Y = Math.Sqrt(1 - X / 2);
                                    }
                                    #endregion

                                    #region Calculated valve flow rate, CcalÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¾ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¯ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â¦ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¼ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¦ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¾ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢unit : m3/h
                                    Ccal = Qs / N22 / Fr / Y * Math.Sqrt(M * T1 / (P1 + P2) / DeltaP);
                                    richTextBox1.Text =("Kv = " + Ccal);
                                    if (Ccal < C)
                                    {
                                        richTextBox1.Text =("The calculated flow rate is less than the Valve rated flow rate");
                                    }
                                    else
                                    {
                                        richTextBox1.Text =("Error! The calculated flow rate is larger than the Valve rated flow rate, this valve should not be considered");
                                    }
                                    #endregion
                                    richTextBox1.Text =("Flow Rate Kv = {0} m3/h\nFlow Rate Cv = {1} us gal/min", Ccal, Ccal * 1.156).ToString();
                                }
                                else
                                {
                                    richTextBox1.Text =("Alarm! Flow is laminar");

                                    #region calculate n
                                    if (C / d / d / N18 >= 0.016)
                                    {
                                        n = N2 / (C / d / d) / (C / d / d);
                                    }
                                    else
                                    {
                                        n = 1 + N32 * Math.Pow(C / d / d, 2 / 3);
                                    }
                                    richTextBox1.Text =("n = " + n);
                                    #endregion

                                    #region calculate Fr
                                    double Fr2;
                                    Fr2 = 0.026 / Fl * Math.Sqrt(n * Rev);
                                    double Fr;

                                    if (Fr2 < 1)
                                    {
                                        Fr = Fr2;
                                    }
                                    else
                                    {
                                        Fr = 1;
                                    }
                                    richTextBox1.Text =("Fr = " + Fr);
                                    #endregion

                                    #region calculate Y
                                    if (Rev >= 1000 && Rev < 10000)
                                    {
                                        Y = (Rev - 1000) / 9000 * (1 - Xsizing / 3 / Xchocked - Math.Sqrt(1 - X / 2)) + Math.Sqrt(1 - X / 2);
                                    }
                                    else
                                    {
                                        Y = Math.Sqrt(1 - X / 2);
                                    }
                                    #endregion

                                    #region Calculated valve flow rate, unit : m3/h
                                    Ccal = Qs / N22 / Fr / Y * Math.Sqrt(M * T1 / (P1 + P2) / DeltaP);
                                    textBox10.Text =("Kv = " + Ccal);
                                    if (Ccal < C)
                                    {
                                        richTextBox1.Text =("The calculated flow rate is less than the Valve rated flow rate");
                                    }
                                    else
                                    {
                                        richTextBox1.Text =("Error! The calculated flow rate is larger than the Valve rated flow rate, this valve should not be considered");
                                    }
                                    #endregion
                                    richTextBox1.Text =("Flow Rate Kv = {0} m3/h\nFlow Rate Cv = {1} us gal/min", Ccal, Ccal * 1.156).ToString();
                                }
                                #endregion
                */
                //if (Variables.sigma > 1.5 && Variables.sigma <= 2.0)
                //{
                //    textBox11.Text=("Incipient cavitation");
                //    textBox11.BackColor = Color.Yellow;


                //}
                //else if (Variables.sigma > 1.3 && Variables.sigma <= 1.5)
                //{
                //    textBox11.Text = ("Medium cavitation");
                //    textBox11.BackColor = Color.Orange;
                //}
                //else if (Variables.sigma > 1.0 && Variables.sigma <= 1.3)
                //{
                //    textBox11.Text = ("Full cavitation");
                //    textBox11.BackColor = Color.Red;
                //}
                //else
                //{
                //    textBox11.Text = ("No cavitation");
                //    textBox11.BackColor = Color.Green;
                //}

                if ((string)cmbState.SelectedItem == "Liquid")
                {
                    if (Variables.sigma > 1.5)
                    {
                        cmbTType.SelectedIndex = 0;


                    }
                    else if (Variables.sigma > 1.3 && Variables.sigma <= 1.5)
                    {
                        cmbTType.SelectedIndex = 1;
                    }
                    else if (Variables.sigma > 1 && Variables.sigma <= 1.3)
                    {
                        cmbTType.SelectedIndex = 2;
                    }
                    else
                    {
                        cmbTType.SelectedIndex = 2;
                        richTextBox1.AppendText("Consult Factory for Cav trim");
                    }

                }
                else
                {
                    if (Variables.sound > 85 && Variables.sound <= 100)
                    {
                        cmbTType.SelectedIndex = 1;
                    }
                    else if (Variables.sound > 100 && Variables.sound <= 110)
                    {
                        cmbTType.SelectedIndex = 2;
                    }
                    else if (Variables.sound < 85)
                    {
                        cmbTType.SelectedIndex = 0;
                    }
                    else
                    {
                        cmbTType.SelectedIndex = 2;
                        richTextBox1.AppendText("Consult Factory Noise Trim");
                    }
                }
                //if (textBox3.Text == ("  NO"))
                //{
                //    textBox3.BackColor = Color.Red;
                //}
                //else
                //{
                //    textBox3.BackColor = Color.Green;
                //}
                //if (textBox4.Text == ("  NO"))
                //{
                //    textBox4.BackColor = Color.Red;
                //}
                //else
                //{
                //    textBox4.BackColor = Color.Green;
                //}
                
                Util.Animate(groupBox9, Util.Effect.Slide, 1000, 360);
                Cursor.Current = Cursors.Default;
                groupBox9.Visible = false;
                button2.Visible = false;
                groupBox5.Visible = false;
                groupBox10.Visible = true;
                button3.Visible = false;
                button1.Visible = true;
                button4.Visible = true;

            }


            catch (Exception ee)
            {

                _ = MessageBox.Show(/*"One or more conditions is Out of Limit / State"*/ee.ToString());
            }
            finally
            {
                
                excel.DisplayAlerts = false;
                excel.ActiveWorkbook.Save();
                excel.Application.Quit();
                excel.Quit();
                
                /// = MessageBox.Show("You are done");
            }

        }

        

        string filename = "";
        public string isValidS(Range rng)
        {
            if (rng.Value == null)
            {
                return "";
            }
            string str = rng.Value.ToString();
            try
            {
                double f = 0;
                Double.TryParse(str, out f);
                if (f > 0)
                {
                    return str;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return "#";
        }


    

        private void groupBox11_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox15_Enter(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbState_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbFluid.Items.Clear();
            cmbUnit.Items.Clear();

            if ((string)cmbState.SelectedItem == "Liquid")
            {
                label40.Text = ("Viscosity");
                //groupBox14.Text = ("Velocity m/s");
                cmbFluid.Items.Add("WATER");
                cmbFluid.Items.Add("OXYGEN");
                cmbFluid.Items.Add("DOWTHERM-A");
                cmbFluid.Items.Add("METHANOL");
                cmbFluid.Items.Add("AMMONIA");
                cmbFluid.Items.Add("BENZENE");
                cmbFluid.Items.Add("BUTANE");
                cmbFluid.Items.Add("CHLORINE");
                cmbFluid.Items.Add("ETHANE");
                cmbFluid.Items.Add("FLOURINE");
                cmbFluid.Items.Add("GLYCOL");
                cmbFluid.Items.Add("HYDROGEN SULPHIDE");
                cmbFluid.Items.Add("ISOBUTANE");
                cmbFluid.Items.Add("ISOBUTYLENE");
                cmbFluid.Items.Add("PROPYLENE");
                cmbFluid.Items.Add("HYDROGEN");
                cmbFluid.Items.Add("CARBON DIOXIDE");
                cmbFluid.Items.Add("METHANE");
                cmbFluid.Items.Add("NATURAL GAS");
                cmbFluid.SelectedItem = ("WATER");

                cmbUnit.Items.Add("Kg/h");
                cmbUnit.Items.Add("m3/h");
                cmbUnit.Items.Add("L/h");
                cmbUnit.Items.Add("Lb/h (Spec.)");

                //ANTICAV TRIMS FOR LIQUIDS
                cmbTType.Items.Clear();
                cmbTType.Items.Add("Standard Parabolic");
                cmbTType.Items.Add("Anticav 1");
                cmbTType.Items.Add("Anticav 2");
                cmbTType.Items.Add(" ");

                //Grey Steam Temp
                cmbITSUnit.Enabled = false;
                txtITSMaxF.Enabled = false;
                txtITSNF.Enabled = false;
                txtITSMinF.Enabled = false;





            }
            else if ((string)cmbState.SelectedItem == "Steam Saturated")
            {
                cmbFluid.Items.Add("STEAM Saturated");
                cmbFluid.SelectedItem = ("STEAM Saturated");
                cmbUnit.Items.Add("Kg/h");
                cmbUnit.Items.Add("t/h");
                cmbUnit.Items.Add("L/h");
                //groupBox14.Text = ("Velocity mach");
                label40.Text = ("Spec Heat Ratio");
                
                //NOISE TRIM FOR GASES
                cmbTType.Items.Clear();
                cmbTType.Items.Add("Standard Parabolic");
                cmbTType.Items.Add("low noise 1");
                cmbTType.Items.Add("Low noise 2");
                cmbTType.Items.Add(" ");

                cmbITSUnit.Enabled = false;
                txtITSMaxF.Enabled = false;
                txtITSNF.Enabled = false;
                txtITSMinF.Enabled = false;

            }
            else if ((string)cmbState.SelectedItem == "Gas")
            {
                cmbFluid.Items.Add("Acetylene");
                cmbFluid.Items.Add("AIR");
                cmbFluid.Items.Add("AMMONIA");
                cmbFluid.Items.Add("ARGON");
                //groupBox14.Text = ("Velocity mach");
                cmbFluid.Items.Add("BUTANE");
                cmbFluid.Items.Add("CARBON DIOXIDE");
                cmbFluid.Items.Add("CARBON MONOXIDE");
                cmbFluid.Items.Add("CHLORINE");
                cmbFluid.Items.Add("ETHANE");
                cmbFluid.Items.Add("ETHYLENE");
                cmbFluid.Items.Add("FLUORINE");
                cmbFluid.Items.Add("GLYCOL");
                cmbFluid.Items.Add("HELIUM");
                cmbFluid.Items.Add("HYDROGEN CHLORIDE");
                cmbFluid.Items.Add("Hydrogen Sulphide");
                cmbFluid.Items.Add("ISOBUTANE");
                cmbFluid.Items.Add("ISOBUTYLENE");
                cmbFluid.Items.Add("METHANE");
                cmbFluid.Items.Add("NATURAL GAS");
                cmbFluid.Items.Add("Neon , Krypton");
                cmbFluid.Items.Add("NITROGEN");
                cmbFluid.Items.Add("Nitrogen (Nitric) Oxide");
                cmbFluid.Items.Add("NITROUS OXIDE");
                cmbFluid.Items.Add("OXYGEN");
                cmbFluid.Items.Add("PHOSGENE");
                cmbFluid.Items.Add("PROPANE");
                cmbFluid.Items.Add("PROPYLENE");
                cmbFluid.Items.Add("Sulphur Dioxide");
                cmbFluid.SelectedItem = ("AIR");
                cmbUnit.Items.Add("Kg/h");
                cmbUnit.Items.Add("m3/h");
                cmbUnit.Items.Add("Nm3/h");
                cmbUnit.Items.Add("L/h");
                cmbUnit.Items.Add("Lb/h (Spec.)");
                cmbUnit.Items.Add("Sft3/h (Spec.)");
                cmbUnit.Items.Add("SCFH (Spec.)");
                label40.Text = ("Spec Heat Ratio");

                //NOISE TRIM FOR GASES
                cmbTType.Items.Clear();
                cmbTType.Items.Add("Standard Parabolic");
                cmbTType.Items.Add("low noise 1");
                cmbTType.Items.Add("Low noise 2");
                cmbTType.Items.Add(" ");

                cmbITSUnit.Enabled = false;
                txtITSMaxF.Enabled = false;
                txtITSNF.Enabled = false;
                txtITSMinF.Enabled = false;


            }


            else if ((string)cmbState.SelectedItem == "Steam Superheated")
            {
                cmbFluid.Items.Add("STEAM Superheated");
                cmbFluid.SelectedItem = ("STEAM Superheated");
                cmbUnit.Items.Add("Kg/h");
                cmbUnit.Items.Add("Kg/s");
                cmbUnit.Items.Add("t/h");
                cmbUnit.Items.Add("L/h");
                //groupBox14.Text = ("Velocity mach");
                label40.Text = ("Spec Heat Ratio");

                //NOISE TRIM FOR GASES
                cmbTType.Items.Clear();
                cmbTType.Items.Add("Standard Parabolic");
                cmbTType.Items.Add("low noise 1");
                cmbTType.Items.Add("Low noise 2");
                cmbTType.Items.Add(" ");

                cmbITSUnit.Enabled = true;
                txtITSMaxF.Enabled = true;
                txtITSNF.Enabled = true;
                txtITSMinF.Enabled = true;

            }
            if (cmbUnit.SelectedItem == null)
            {
                cmbUnit.SelectedItem = ("Kg/h");
                
            }


        }

        private void cmbFluid_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void txtVVMaxF_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void cmbTSize_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            groupBox5.Visible = false;
            groupBox10.Visible = false;
            groupBox9.Visible = true;
            button1.Visible = false;
            button2.Visible = true;
            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = true;
            button6.Visible = false;
        }


        private void groupBox9_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            groupBox5.Visible = false;
            groupBox10.Visible = true;
            groupBox9.Visible = false;
            button1.Visible = true;
            button2.Visible = false;
            button4.Visible = true;
            button5.Visible = false;
            button6.Visible = false;



        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox4.BackColor==Color.Green && textBox3.BackColor == Color.Green && comboBox12.SelectedIndex != -1 && comboBox8.SelectedIndex != -1 && comboBox11.SelectedIndex != -1 && comboBox9.SelectedIndex != -1 && cmbTType.SelectedIndex != -1 && double.Parse(textBox2.Text)<=double.Parse(comboBox11.SelectedItem.ToString()))
            {
                groupBox5.Visible = true;
                groupBox10.Visible = false;
                groupBox9.Visible = false;
                button1.Visible = false;
                button2.Visible = false;
                button3.Visible = true;
                button4.Visible = false;
                button5.Visible = false;
                button6.Visible = true;
            }
            else
            {
                MessageBox.Show("Please complete selection / Eliminate Red warnings to continue" , "Attention",
                                 MessageBoxButtons.RetryCancel,
                                 MessageBoxIcon.Error);
            }
            if (double.Parse(textBox2.Text) <= double.Parse(comboBox11.SelectedItem.ToString()))
            {
                textBox2.BackColor = Color.Green;
                comboBox11.BackColor = Color.Green;
            }
            else
            {
                textBox2.BackColor = Color.Red;
                comboBox11.BackColor = Color.Red;
            }

            Application excel = new Application();


            _Excel.Workbook workbook = excel.Workbooks.Open(filename + "\\Demo.xlsx", ReadOnly: false, Editable: true);
            _Excel.Worksheet worksheet = workbook.Worksheets["SELECTION"] as _Excel.Worksheet;

            _Excel.Range rangeseat;
            _Excel.Range pneumatic;
            _Excel.Range electric;
            _Excel.Range balancing;






            if (worksheet == null)
                return;


            try
            {
                Range valvesiz = worksheet.Rows.Cells[90, 2];
                valvesiz.Value = comboBox9.SelectedItem.ToString();


                HashSet<string> distinct = new HashSet<string>();


                //range = worksheet.get_Range("AQ105", "AQ108") as _Excel.Range;

                //rangecv = worksheet.get_Range("AQ96", "AQ99") as _Excel.Range;
                //rangepacking = worksheet.get_Range("AU72", "AU77") as _Excel.Range;
                rangeseat = worksheet.get_Range("AR96", "AR99") as _Excel.Range;
                //rangebonnet = worksheet.get_Range("AU81", "AU85") as _Excel.Range;
                pneumatic = worksheet.get_Range("AS102", "AS105") as _Excel.Range;
                electric = worksheet.get_Range("AS107", "AS110") as _Excel.Range;
                balancing = worksheet.get_Range("AU102", "AU105") as _Excel.Range;

                //ELECTRIC RANGE
                if (int.Parse(shutoff.Text) > 111)
                {
                    comboBox17.Items.Clear();
                    comboBox17.Items.Add("RegadaST2 + Balancing");

                }
                else
                {
                    comboBox17.Items.Clear();
                    foreach (_Excel.Range cell in electric.Cells)
                    {
                        string value = (cell.Value2).ToString();

                        if (distinct.Add(value))
                            comboBox17.Items.Add(value);

                        //comboBox17.Items.Add((cell.Value2).ToString() as string);
                    }
                }
                comboBox17.SelectedIndex = 0;


                //BALANCING RANGE
                comboBox18.Items.Clear();
                foreach (_Excel.Range cell in balancing.Cells)
                {

                    comboBox18.Items.Add((cell.Value2).ToString() as string);
                }

                //PNEUMATIC RANGE
                if (comboBox18.Items[0].ToString() == "NAA" /*&& int.Parse(comboBox9.SelectedItem.ToString()) > 49*/)
                {
                    comboBox16.Items.Clear();
                    comboBox16.Items.Add("S500B-1250 (195) 0,8-2,0 (12-30)");
                    checkBox1.Checked = true;
                    comboBox16.SelectedIndex = 0;
                    if (Variables.kelvin1 > 73.15 && Variables.kelvin1 <= 473.15/* && checkBox1.Checked == true comboBox18.Items[0].ToString() == "NAA"*/)
                    {
                        cmbTBU.SelectedIndex = 0;
                    }

                    else if (Variables.kelvin1 > 73.15 && Variables.kelvin1 <= 673.15/* && checkBox1.Checked == true comboBox18.Items[0].ToString() == "NAA"*/)
                    {
                        cmbTBU.SelectedIndex = 1;
                    }
                    else
                    {
                        cmbTBU.SelectedIndex = 3;
                    }

                }
                else if (comboBox18.Items[0].ToString() == "NAA" && int.Parse(comboBox9.SelectedItem.ToString()) < 49)
                {
                    comboBox16.Items.Clear();
                    comboBox16.Items.Add("Consult Factory");
                    checkBox1.Checked = false;
                    comboBox16.SelectedIndex = 0;
                }
                else
                {
                    comboBox16.Items.Clear();
                    checkBox1.Checked = false;
                    foreach (_Excel.Range cell in pneumatic.Cells)
                    {
                        comboBox16.Items.Add((cell.Value2).ToString() as string);
                    }
                    comboBox16.SelectedIndex = 0;
                }





                //Range row1 = worksheet.Rows.Cells[10, 5];
                //row1.Value = cmbFluid.SelectedItem.ToString();
            }


            catch (Exception ee)
            {

                _ = MessageBox.Show(/*"One or more conditions is Out of Limit / State"*/ee.ToString());
            }
            finally
            {

                excel.DisplayAlerts = false;
                excel.ActiveWorkbook.Save();
                excel.Application.Quit();
                excel.Quit();

                /// = MessageBox.Show("You are done");
            }




        }

        private void groupBox19_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox7.SelectedIndex == 1)
            {
                comboBox16.Visible = false;
                comboBox17.Visible = true;
                //comboBox8.Visible = false;
                //label19.Visible = false;
                txtAMM.Visible = true;
                label53.Visible = true;
                label48.Visible = false;
                textBox5.Visible = false;
                comboBox1.Items.Clear();
                //comboBox1.Items.Add("Pillar");
                //comboBox1.Items.Add("Handwheel");
                //comboBox1.Items.Add("Voltage/supply");
                comboBox1.Items.Add("none");
                comboBox1.Items.Add("2-3 Points input");
                comboBox1.Items.Add("4-20 mA input");
                comboBox1.Items.Add("20-4 mA input");
                comboBox1.Items.Add("0-10 V input");
                comboBox1.SelectedIndex = 0;
                //comboBox1.Items.Add("No Output");
                //comboBox1.Items.Add("4-20 mA Output ");
                //comboBox1.Items.Add("0-10 V Output ");
                checkedListBox2.Items.Clear();
                //checkedListBox2.Items.Add("Pillar");
                checkedListBox2.Items.Add("Handwheel");
                //checkedListBox2.Items.Add("Voltage/supply");
                //checkedListBox2.Items.Add("2-3 Points input");
                //checkedListBox2.Items.Add("4-20 mA input");
                //checkedListBox2.Items.Add("20-4 mA input");
                //checkedListBox2.Items.Add("0-10 V input");
                //checkedListBox2.Items.Add("No Output");
                //checkedListBox2.Items.Add("4-20 mA Output ");
                //checkedListBox2.Items.Add("0-10 V Output ");

                groupBox18.Visible = false;
                groupBox12.Location = new System.Drawing.Point(472,17);

                comboBox13.Items.Clear();
                comboBox13.Items.Add("No Output");
                comboBox13.Items.Add("4-20 mA Output");
                comboBox13.Items.Add("0-10 V Output");
                comboBox13.SelectedIndex = 0;
                comboBox14.Items.Clear();
                comboBox14.Items.Add("None");
                comboBox14.Items.Add("Voltage/supply");
                comboBox14.SelectedIndex = 0;
                checkBox6.Visible = false;
                checkBox6.Checked = false;
                groupBox20.Visible = true;
                comboBox22.Visible = false;
                comboBox22.Items.Clear();
                comboBox22.Items.Add("None");
                comboBox22.SelectedIndex = 0;
                label71.Visible = false;
            }


            else//not electric
            {
                comboBox22.Visible = true;
                comboBox22.Items.Clear();
                comboBox22.Items.Add("Standard -6mm Rilsan/Comp-Brass");
                comboBox22.Items.Add("Stainless Steel -1/4 / \"Double\" ferro"); 
                comboBox22.Items.Add("None");
                comboBox22.SelectedIndex = 2;
                label71.Visible = true;
                //comboBox8.Visible = true;
                //label19.Visible = true;
                groupBox20.Visible = false;
                checkBox6.Visible = true;
                comboBox16.Visible = true;
                comboBox17.Visible = false;
                txtAMM.Visible = false;
                label53.Visible = false;
                label48.Visible = true;
                textBox5.Visible = true;
                comboBox1.Items.Clear();
                comboBox1.Items.Add("None");
                comboBox1.SelectedIndex = 0;
                comboBox13.Items.Clear();
                comboBox13.Items.Add("None");
                comboBox13.SelectedIndex = 0; 
                comboBox14.Items.Clear();
                comboBox14.Items.Add("None");
                comboBox14.SelectedIndex = 0;
                //comboBox1.Items.Add("Pillar Yoke");
                //comboBox1.Items.Add("Stainless Steel Actuator");
                //comboBox1.Items.Add("Handwheel");
                checkedListBox2.Items.Clear();
                //checkedListBox2.Items.Add("Pillar Yoke 210mm Zinc");
                checkedListBox2.Items.Add("Handwheel");
                checkedListBox2.Items.Add("Pillar Yoke");
                checkedListBox2.Items.Add("Stainless Steel Actuator");

                groupBox18.Visible = true;
                groupBox12.Location = new System.Drawing.Point(617, 191);

                
                


            }
            if (comboBox7.SelectedIndex == 1)
            {
                if (comboBox17.FindString("Sauter AVF234SF132/232") == comboBox17.SelectedIndex)
                {
                    comboBox21.Items.Clear();
                    comboBox21.Items.Add("Fail Close");
                    comboBox21.Items.Add("Fail Open");

                }
                else
                {
                    comboBox21.Items.Clear();
                    comboBox21.Items.Add("Fail in Place");

                }
            }
            else
            {
                comboBox21.Items.Clear();
                comboBox21.Items.Add("Fail Close");
                comboBox21.Items.Add("Fail Open");
            }
            comboBox21.SelectedIndex = 0;




        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (Double.Parse(textBox10.Text) < Variables.maxvelocity)
            {
                textBox4.BackColor = Color.Green;
                comboBox9.BackColor = Color.Green;
                //comboBox9.ForeColor = Color.Green;
                textBox4.Text = ("VELOCITY OK");
                
                richTextBox1.Select(richTextBox1.TextLength, 0);
                richTextBox1.SelectionColor = Color.Green;
                richTextBox1.AppendText("\nVELOCITY OK");
            }
            else
            {
                textBox4.BackColor = Color.Red;
                comboBox9.BackColor = Color.Red;
                //comboBox9.ForeColor = Color.Red;
                textBox4.Text = ("Velocity too high. Bigger valve size to be selected");
                

                richTextBox1.Select(richTextBox1.TextLength, 0);
                richTextBox1.SelectionColor = Color.Red;
                richTextBox1.AppendText("\nVELOCITY TOO HIGH. Bigger valve size to be selected");
            }

        }
        string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);


        
        private void button5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    FormSerialisor.Deserialise(this, dialog.FileName);
                }
            }

    
        }
        
    

        private void button7_Click(object sender, EventArgs e)
        {
            Application excel2 = new Application();


            _Excel.Workbook workbook = excel2.Workbooks.Open(filename + "\\Demo.xlsx");
            _Excel.Worksheet worksheet = workbook.Worksheets["CV"] as _Excel.Worksheet;






            if (worksheet == null)
                return;


            try
            {
                Range allo = worksheet.Rows.Cells[72, 2];
                allo.Value = AcceptableNoise.SelectedItem.ToString();

                Range maxvel = worksheet.Rows.Cells[73, 2];
                maxvel.Value = textBox10.Text;

                Range normvel = worksheet.Rows.Cells[74, 2];
                normvel.Value = txtVVNF.Text;

                Range minvel = worksheet.Rows.Cells[75, 2];
                minvel.Value = txtVVMinF.Text;

                Range maxnoise = worksheet.Rows.Cells[76, 2];
                maxnoise.Value = txtAPSMaxF.Text;

                Range normnoise = worksheet.Rows.Cells[77, 2];
                normnoise.Value = txtAPSNF.Text;

                Range minnoise = worksheet.Rows.Cells[78, 2];
                minnoise.Value = txtAPSMinF.Text;

                Range rating = worksheet.Rows.Cells[79, 2];
                rating.Value = comboBox10.SelectedItem.ToString();

                Range flowdire = worksheet.Rows.Cells[80, 2];
                flowdire.Value = cmbVBFD.SelectedItem.ToString();

                Range bodymat = worksheet.Rows.Cells[81, 2];
                bodymat.Value = textBox7.Text;

                Range packingmat = worksheet.Rows.Cells[82, 2];
                packingmat.Value = comboBox12.SelectedItem.ToString();

                Range bonnet = worksheet.Rows.Cells[83, 2];
                bonnet.Value = comboBox8.SelectedItem.ToString();

                Range trimtype = worksheet.Rows.Cells[84, 2];
                trimtype.Value = cmbTType.SelectedItem.ToString();

                Range seatb = worksheet.Rows.Cells[85, 2];
                seatb.Value = cmbTSize.Text;

                Range trimmat = worksheet.Rows.Cells[86, 2];
                trimmat.Value = comboBox19.SelectedItem.ToString();

                Range leakage = worksheet.Rows.Cells[87, 2];
                leakage.Value = comboBox20.SelectedItem.ToString();

                Range xtic = worksheet.Rows.Cells[88, 2];
                xtic.Value = cmbTC.SelectedItem.ToString();

                Range balancedplug = worksheet.Rows.Cells[89, 2];
                balancedplug.Value = cmbTBU.SelectedItem.ToString();

                //Range valvesiz = worksheet.Rows.Cells[90, 2];
                //valvesiz.Value = comboBox9.SelectedItem.ToString();

                Range valvecvv = worksheet.Rows.Cells[91, 2];
                valvecvv.Value = comboBox11.SelectedItem.ToString();

                Range acttype = worksheet.Rows.Cells[93, 2];
                acttype.Value = comboBox7.SelectedItem.ToString();




                Range pne = worksheet.Rows.Cells[94, 2];
                pne.Value = comboBox16.SelectedItem.ToString();

                Range ele = worksheet.Rows.Cells[95, 2];
                ele.Value = comboBox17.SelectedItem.ToString();

                Range action = worksheet.Rows.Cells[96, 2];
                action.Value = comboBox21.SelectedItem.ToString();

                Range inpop = worksheet.Rows.Cells[97, 2];
                inpop.Value = comboBox1.SelectedItem.ToString();

                Range outpop = worksheet.Rows.Cells[98, 2];
                outpop.Value = comboBox13.SelectedItem.ToString();

                Range supply = worksheet.Rows.Cells[99, 2];
                supply.Value = comboBox14.SelectedItem.ToString();



                Range pos = worksheet.Rows.Cells[100, 2];
                pos.Value = comboBox2.SelectedItem.ToString();

                Range fil = worksheet.Rows.Cells[101, 2];
                fil.Value = comboBox3.SelectedItem.ToString();

                Range sole = worksheet.Rows.Cells[102, 2];
                sole.Value = comboBox4.SelectedItem.ToString();

                Range swit = worksheet.Rows.Cells[103, 2];
                swit.Value = comboBox5.SelectedItem.ToString();

                Range bos = worksheet.Rows.Cells[104, 2];
                bos.Value = comboBox6.SelectedItem.ToString();

                Range note = worksheet.Rows.Cells[104, 2];
                note.Value = richTextBox2.Text;

                Range nace = worksheet.Rows.Cells[106, 2];
                Range area = worksheet.Rows.Cells[107, 2];
                Range ped = worksheet.Rows.Cells[108, 2];
                Range leaktest = worksheet.Rows.Cells[109, 2];
                Range tubing = worksheet.Rows.Cells[113, 2];

                Range worktest = worksheet.Rows.Cells[126, 2];
                if (checkedListBox1.GetItemCheckState(1) == System.Windows.Forms.CheckState.Checked)
                {
                    worktest.Value = "yes";
                }
                else
                {
                    worktest.Value = "no";
                }
                Range mattest = worksheet.Rows.Cells[127, 2];
                if (checkedListBox1.GetItemCheckState(2) == System.Windows.Forms.CheckState.Checked)
                {
                    mattest.Value = "yes";
                }
                else
                {
                    mattest.Value = "no";
                }
                Range milltest = worksheet.Rows.Cells[128, 2];
                if (checkedListBox1.GetItemCheckState(3) == System.Windows.Forms.CheckState.Checked)
                {
                    milltest.Value = "yes";
                }
                else
                {
                    milltest.Value = "no";
                }


                Range compliance = worksheet.Rows.Cells[116, 2];
                if (checkedListBox1.GetItemCheckState(0) == System.Windows.Forms.CheckState.Checked)
                {
                    compliance.Value = "yes";
                }
                else
                {
                    compliance.Value = "no";
                }
                Range seatleakagewater = worksheet.Rows.Cells[117, 2];
                if (checkedListBox1.GetItemCheckState(9) == System.Windows.Forms.CheckState.Checked)
                {
                    seatleakagewater.Value = "yes";
                }
                else
                {
                    seatleakagewater.Value = "no";
                }
                Range seatleakageair = worksheet.Rows.Cells[118, 2];
                if (checkedListBox1.GetItemCheckState(10) == System.Windows.Forms.CheckState.Checked)
                {
                    seatleakageair.Value = "yes";
                }
                else
                {
                    seatleakageair.Value = "no";
                }
                Range pmi = worksheet.Rows.Cells[119, 2];
                if (checkedListBox1.GetItemCheckState(11) == System.Windows.Forms.CheckState.Checked)
                {
                    pmi.Value = "yes";
                }
                else
                {
                    pmi.Value = "no";
                }
                Range lp = worksheet.Rows.Cells[120, 2];
                if (checkedListBox1.GetItemCheckState(12) == System.Windows.Forms.CheckState.Checked)
                {
                    lp.Value = "yes";
                }
                else
                {
                    lp.Value = "no";
                }
                Range rt = worksheet.Rows.Cells[121, 2];
                if (checkedListBox1.GetItemCheckState(13) == System.Windows.Forms.CheckState.Checked)
                {
                    rt.Value = "yes";
                }
                else
                {
                    rt.Value = "no";
                }
                Range it = worksheet.Rows.Cells[122, 2];
                if (checkedListBox1.GetItemCheckState(14) == System.Windows.Forms.CheckState.Checked)
                {
                    it.Value = "yes";
                }
                else
                {
                    it.Value = "no";
                }
                Range mpi = worksheet.Rows.Cells[123, 2];
                if (checkedListBox1.GetItemCheckState(15) == System.Windows.Forms.CheckState.Checked)
                {
                    mpi.Value = "yes";
                }
                else
                {
                    mpi.Value = "no";
                }
                Range specialpaint = worksheet.Rows.Cells[124, 2];
                if (checkedListBox1.GetItemCheckState(16) == System.Windows.Forms.CheckState.Checked)
                {
                    specialpaint.Value = "yes";
                }
                else
                {
                    specialpaint.Value = "no";
                }
                Range itp = worksheet.Rows.Cells[125, 2];
                if (checkedListBox1.GetItemCheckState(17) == System.Windows.Forms.CheckState.Checked)
                {
                    itp.Value = "yes";
                }
                else
                {
                    itp.Value = "no";
                }

                //insulation

                Range insulation = worksheet.Rows.Cells[111, 2];
                if (checkBox2.Checked == true)
                {
                    insulation.Value = "YES";
                }
                else
                {
                    insulation.Value = "NO";
                }

                //Attenuation


                Range NoiseAttenu = worksheet.Rows.Cells[110, 2];
                NoiseAttenu.Value = Variables.NoiseAttenut;

                //NACE TEST CHECK

                if (checkedListBox1.GetItemCheckState(7) == System.Windows.Forms.CheckState.Checked)
                {
                    nace.Value = "nace";
                }
                else
                {
                    nace.Value = "no";
                }


                //AREA TEST CHECK
                if (checkedListBox1.GetItemCheckState(5) == System.Windows.Forms.CheckState.Checked)
                {
                    area.Value = "ATEX";
                    if (checkedListBox1.GetItemCheckState(6) == System.Windows.Forms.CheckState.Checked)
                    {
                        area.Value = "ATEX + CU-TR";
                    }
                }
                else if (checkedListBox1.GetItemCheckState(6) == System.Windows.Forms.CheckState.Checked)
                {
                    area.Value = "CU-TR";
                }
                else
                {
                    area.Value = "SAFE";
                }


                //PED TEST CHECK
                if (checkedListBox1.GetItemCheckState(4) == System.Windows.Forms.CheckState.Checked)
                {
                    ped.Value = "yes";
                }
                else
                {
                    ped.Value = "no";
                }


                //Hydro LEAKAGE TEST
                if (checkedListBox1.GetItemCheckState(8) == System.Windows.Forms.CheckState.Checked)
                {
                    leaktest.Value = "YES";
                }
                else
                {

                    leaktest.Value = "NO";
                }

                //guiding
                Range guiding = worksheet.Rows.Cells[112, 2];

                if (cmbTType.SelectedIndex > 1 || cmbTBU.SelectedIndex < 4)
                {
                    guiding.Value = "Top + cage";
                }
                else
                {
                    if (int.Parse(comboBox9.SelectedItem.ToString()) < 51)
                    {
                        guiding.Value = "Top Guided Stem";
                    }
                    else
                    {
                        guiding.Value = "Top Guided Shaft";
                    }
                }

                //tubing

                if (comboBox22.SelectedIndex == 1)
                {
                    tubing.Value = "Stainless Steel";
                }
                else
                {
                    tubing.Value = "Std";
                }


                Range steel = worksheet.Rows.Cells[114, 2];
                //stainless steel act
                if (checkBox6.Checked == true)
                {
                    steel.Value = "YES";
                }
                else
                {
                    steel.Value = "NO";
                }

                Range yoke = worksheet.Rows.Cells[115, 2];
                yoke.Value = comboBox23.SelectedItem.ToString();
            }


            catch (Exception ee)
            {

                _ = MessageBox.Show(/*"One or more conditions is Out of Limit / State"*/ee.ToString());
            }
            finally
            {

                excel2.DisplayAlerts = false;
                excel2.ActiveWorkbook.Save();
                excel2.Application.Quit();
                excel2.Quit();

                /// = MessageBox.Show("You are done");
            }

            //Marshal.ReleaseComObject(worksheet);
            //excel2.DisplayAlerts = false;
            //excel2.ActiveWorkbook.Save();
            //excel2.Application.Quit();
            //excel2.Quit();


            //Workbook workbook1 = new Workbook();
            //workbook1.LoadFromFile(filename + "\\Printout.xlsx");

            //PrintDialog dialog = new PrintDialog();
            //dialog.AllowPrintToFile = true;
            //dialog.AllowCurrentPage = true;
            //dialog.AllowSomePages = true;
            //dialog.AllowSelection = true;
            //dialog.UseEXDialog = true;
            //dialog.PrinterSettings.Duplex = Duplex.Simplex;
            //dialog.PrinterSettings.FromPage = 0;
            //dialog.PrinterSettings.ToPage = 8;
            //dialog.PrinterSettings.PrintRange = PrintRange.SomePages;
            //workbook1.PrintDialog = dialog;
            ////workbook1.SaveToFile(path + "Sizing Printout.pdf");

            //PrintDocument pd = workbook1.PrintDocument;

            //if (dialog.ShowDialog() == DialogResult.OK)
            //{/* pd.Print();*/}

            //_PrintExcel.Workbook workbook1 = new _PrintExcel.Workbook();
            //workbook1.LoadFromFile(filename + "\\Printout.xlsx");

            //_PrintExcel.Worksheet sheet = workbook1.Worksheets["printout"];
            //worksheet1.SaveAs(filename + "TEST.xlsx");
            //sheet.SaveToPdf(path + "Sizing Printout.pdf");
            //System.Diagnostics.Process.Start("explorer.exe", path + "Sizing Printout.pdf");
            //workbook1.SaveCopyAs(path + "Sizing Printout.xlsx");






            //PrintDialog dialog = new PrintDialog();
            //dialog.AllowPrintToFile = true;
            //dialog.AllowCurrentPage = true;
            //dialog.AllowSomePages = true;
            //dialog.AllowSelection = true;
            //dialog.UseEXDialog = true;
            //dialog.PrinterSettings.Duplex = Duplex.Simplex;
            //dialog.PrinterSettings.PrintRange = PrintRange.SomePages;
            //workbook.PrintDialog = dialog;
            //PrintDocument pd = workbook.PrintDocument;
            //if (dialog.ShowDialog() == DialogResult.OK)
            //{ pd.Print(); }

            //Application excel = new Application();


            //_Excel.Workbook workbook = excel.Workbooks.Open(filename, ReadOnly: false, Editable: true);
            //_Excel.Worksheet worksheet = workbook.Worksheets["printout"] as _Excel.Worksheet;



            //worksheet.PrintPreview();
            //worksheet.PrintOut(From:1,To:2,Copies:1, Preview:true, Type.Missing, PrintToFile:true, Type.Missing, Type.Missing);
            //worksheet.PrintPreview();
            //PrintDialog printDlg = new PrintDialog();
            /*PrintDocument printDoc = new PrintDocument();
            printDoc.DocumentName = "Print Document";
            printDlg.Document = printDoc;
            printDlg.AllowSelection = true;
            printDlg.AllowSomePages = true;
            if (printDlg.ShowDialog() == DialogResult.OK) printDoc.Print();*/
            //    worksheet.PrintOut(
            //1, 1, 1, Type.Missing,
            //Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //    // Cleanup:
            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();

            //    Marshal.FinalReleaseComObject(worksheet);

            //    workbook.Close(false, Type.Missing, Type.Missing);
            //    Marshal.FinalReleaseComObject(workbook);

            //    excel.Quit();
            //    Marshal.FinalReleaseComObject(excel);

            /*workbook.DisplayAlerts = false;
            excel.ActiveWorkbook.Save(); 
            excel.Application.Quit();
            excel.Quit();*/
        }
        private void button6_Click(object sender, EventArgs e)
        {

            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
            dialog.FilterIndex = 1;
            dialog.FileName = "sizing";

            if (dialog.ShowDialog() == DialogResult.OK)
            {                
                FormSerialisor.Serialise(this, dialog.FileName + "");
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //comboBox9.SelectedItem = textBox1.Text;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            comboBox10.SelectedItem = textBox6.Text;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
        
        

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            Variables.power = int.Parse(comboBox9.SelectedItem.ToString());
            if (cmbState.SelectedIndex==0)
            {
                
                textBox10.Text = Math.Round(double.Parse(((278*Variables.m3h)/(Math.Pow(Variables.power /2, 2)*3.14)).ToString()),2).ToString();
                //textBox10.Text = Math.Round(double.Parse(((278 * Variables.m3h) / (Math.Pow(Variables.power / 2, 2) * 3.14)).ToString()), 2).ToString();
                //textBox10.Text = Math.Round(double.Parse(((278 * Variables.m3h) / (Math.Pow(Variables.power / 2, 2) * 3.14)).ToString()), 2).ToString();

            }
            else
            {
                textBox10.Text = Math.Round(double.Parse((((1.296 * Variables.nm3h1) * Variables.kelvin1 / (Math.Pow(Variables.power, 2) * (Variables.p2)) / 340).ToString())),3).ToString();
                //textBox10.Text = Math.Round(double.Parse((((1.296 * Variables.nm3h1) * Variables.kelvin1 / (Math.Pow(Variables.power, 2) * (Variables.p2)) / 340).ToString())), 3).ToString();
                //textBox10.Text = Math.Round(double.Parse((((1.296 * Variables.nm3h1) * Variables.kelvin1 / (Math.Pow(Variables.power, 2) * (Variables.p2)) / 340).ToString())), 3).ToString();
            }
            //try
            //{
            //    comboBox11.SelectedItem = (textBox2.Text);
            //}
            //catch (Exception k)
            //{
            //    comboBox11.SelectedItem = double.Parse(textBox2.Text);

            //}
            double[] collection = { 0.059, 0.165, 0.33, 0.55, 0.825, 1.1, 1.43, 1.65, 2.2, 2.53, 3.3, 4.95 };
            double[] collection20 = { 0.059, 0.165, 0.33, 0.55, 0.825, 1.1, 1.43, 1.65, 2.2, 2.53, 3.3, 4.95, 6.93 };
            double[] collection25 = { 0.059, 0.165, 0.33, 0.55, 0.825, 1.1, 1.43, 1.65, 2.2, 2.53, 3.3, 4.95, 6.93, 12.1 };
            double[] collection32 = { 4.95, 6.93, 12.1, 19.8 };
            double[] collection40 = { 4.95, 6.93, 12.1, 19.8, 33 };
            double[] collection50 = { 4.95, 6.93, 12.1, 19.8, 33, 51.26 };
            double[] collection65 = { 12.1, 19.8, 33, 51.26, 79.97 };
            double[] collection80 = { 19.8, 33, 51.26, 79.97, 115.5 };
            double[] collection100 = { 33, 51.26, 79.97, 115.5, 176 };
            double[] collection125 = { 51.26, 79.97, 115.5, 176, 293.7 };
            double[] collection150 = { 79.97, 115.5, 176, 293.7, 404.8 };
            double[] collection200 = { 115.5, 176, 293.7, 404.8, 711.7 };
            
            int cvs = int.Parse(comboBox9.SelectedItem.ToString());
            switch (cvs)
            {
                
                case 15:
                    comboBox11.Items.Clear();
                    foreach (double item in collection)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        } 
                    }
                    
                    //comboBox11.Items.Add(0.059);
                    //comboBox11.Items.Add(0.165);
                    //comboBox11.Items.Add(0.33);
                    //comboBox11.Items.Add(0.55);
                    //comboBox11.Items.Add(0.825);
                    //comboBox11.Items.Add(1.1);
                    //comboBox11.Items.Add(1.43);
                    //comboBox11.Items.Add(1.65);
                    //comboBox11.Items.Add(2.2);
                    //comboBox11.Items.Add(2.53);
                    //comboBox11.Items.Add(3.3);
                    //comboBox11.Items.Add(4.95);

                    comboBox11.SelectedItem = double.Parse(textBox2.Text); 
                    break;
                case 20:
                    comboBox11.Items.Clear();
                    
                    foreach (double item in collection20)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }
                    //comboBox11.Items.Add(0.059);
                    //comboBox11.Items.Add(0.165);
                    //comboBox11.Items.Add(0.33);
                    //comboBox11.Items.Add(0.55);
                    //comboBox11.Items.Add(0.825);
                    //comboBox11.Items.Add(1.1);
                    //comboBox11.Items.Add(1.43);
                    //comboBox11.Items.Add(1.65);
                    //comboBox11.Items.Add(2.2);
                    //comboBox11.Items.Add(2.53);
                    //comboBox11.Items.Add(3.3);
                    //comboBox11.Items.Add(4.95);
                    //comboBox11.Items.Add(6.93);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text); 
                    break;
                case 25:
                    comboBox11.Items.Clear();
                    foreach (double item in collection25)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }

                    //comboBox11.Items.Add(0.059);
                    //comboBox11.Items.Add(0.165);
                    //comboBox11.Items.Add(0.33);
                    //comboBox11.Items.Add(0.55);
                    //comboBox11.Items.Add(0.825);
                    //comboBox11.Items.Add(1.1);
                    //comboBox11.Items.Add(1.43);
                    //comboBox11.Items.Add(1.65);
                    //comboBox11.Items.Add(2.2);
                    //comboBox11.Items.Add(2.53);
                    //comboBox11.Items.Add(3.3);
                    //comboBox11.Items.Add(4.95);
                    //comboBox11.Items.Add(6.93);
                    //comboBox11.Items.Add(12.1);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text);

                    break;
                case 32:
                    comboBox11.Items.Clear();
                    foreach (double item in collection32)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }
                    //comboBox11.Items.Add(4.95);
                    //comboBox11.Items.Add(6.93);
                    //comboBox11.Items.Add(12.1);
                    //comboBox11.Items.Add(19.8);

                    comboBox11.SelectedItem = double.Parse(textBox2.Text); 
                    break;
                case 40:
                    comboBox11.Items.Clear();
                    foreach (double item in collection40)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }
                    //comboBox11.Items.Add(4.95);
                    //comboBox11.Items.Add(6.93);
                    //comboBox11.Items.Add(12.1);
                    //comboBox11.Items.Add(19.8);
                    //comboBox11.Items.Add(33);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text); 
                    break;
                case 50:
                    comboBox11.Items.Clear();
                    foreach (double item in collection50)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }
                    //comboBox11.Items.Add(4.95);
                    //comboBox11.Items.Add(6.93);
                    //comboBox11.Items.Add(12.1);
                    //comboBox11.Items.Add(19.8);
                    //comboBox11.Items.Add(33);
                    //comboBox11.Items.Add(51.26);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text); 
                    break;
                case 65:
                    comboBox11.Items.Clear();
                    foreach (double item in collection65)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }
                    //comboBox11.Items.Add(12.1);
                    //comboBox11.Items.Add(19.8);
                    //comboBox11.Items.Add(33);
                    //comboBox11.Items.Add(51.26);
                    //comboBox11.Items.Add(79.97);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text); 
                    break;
                case 80:
                    comboBox11.Items.Clear();

                    foreach (double item in collection80)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }

                    //comboBox11.Items.Add(19.8);
                    //comboBox11.Items.Add(33);
                    //comboBox11.Items.Add(51.26);
                    //comboBox11.Items.Add(79.97);
                    //comboBox11.Items.Add(115.5);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text);
                    break;
                case 100:
                    comboBox11.Items.Clear();
                    foreach (double item in collection100)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }
                    //comboBox11.Items.Add(33);
                    //comboBox11.Items.Add(51.26);
                    //comboBox11.Items.Add(79.97);
                    //comboBox11.Items.Add(115.5);
                    //comboBox11.Items.Add(176);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text); 
                    break;
                case 125:
                    comboBox11.Items.Clear();
                    foreach (double item in collection125)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }
                    //comboBox11.Items.Add(51.26);
                    //comboBox11.Items.Add(79.97);
                    //comboBox11.Items.Add(115.5);
                    //comboBox11.Items.Add(176);
                    //comboBox11.Items.Add(293.7);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text); 
                    break;
                case 150:
                    comboBox11.Items.Clear();
                    foreach (double item in collection150)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }
                    //comboBox11.Items.Add(79.97);
                    //comboBox11.Items.Add(115.5);
                    //comboBox11.Items.Add(176);
                    //comboBox11.Items.Add(293.7);
                    //comboBox11.Items.Add(404.8);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text); 
                    break;
                case 200:
                    comboBox11.Items.Clear();
                    foreach (double item in collection200)
                    {
                        if (item >= double.Parse(textBox2.Text))
                        {
                            comboBox11.Items.Add(item);
                        }
                    }
                    //comboBox11.Items.Add(115);
                    //comboBox11.Items.Add(176);
                    //comboBox11.Items.Add(293.7);
                    //comboBox11.Items.Add(404.8);
                    //comboBox11.Items.Add(711.7);
                    comboBox11.SelectedItem = double.Parse(textBox2.Text);
                    break;
            
            }
            

        }


        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBox10.SelectedItem="PN 16")
            //{
            //    comboBox24.Items.Clear();
            //    comboBox24.Items.Add("");

            //}
            string pn = comboBox10.SelectedItem.ToString();
            switch (pn)
            {
                case "PN16":
                    comboBox24.Items.Clear();
                    comboBox24.Items.Add("FLG EN1092-1 PN16 B1");
                    break;
                case "PN25":
                    comboBox24.Items.Clear();
                    comboBox24.Items.Add("FLG EN1092-1 PN16 B1");
                    comboBox24.Items.Add("ASME B16.5-150 RF");
                    comboBox24.Items.Add("FLG EN1092-1 PN25 B1");
                    
                    break;
                case "#150":
                    comboBox24.Items.Clear();
                    comboBox24.Items.Add("FLG EN1092-1 PN16 B1");
                    comboBox24.Items.Add("ASME B16.5-150 RF");

                    break;
                case "PN40":
                    comboBox24.Items.Clear();
                    comboBox24.Items.Add("FLG EN1092-1 PN16 B1");
                    comboBox24.Items.Add("ASME B16.5-150 RF");
                    comboBox24.Items.Add("FLG EN1092-1 PN25 B1");
                    comboBox24.Items.Add("FLG EN1092-1 PN40 B1");

                    break;
                case "PN63":
                    comboBox24.Items.Clear();
                    comboBox24.Items.Add("FLG EN1092-1 PN16 B1");
                    comboBox24.Items.Add("ASME B16.5-150 RF");
                    comboBox24.Items.Add("FLG EN1092-1 PN25 B1");
                    comboBox24.Items.Add("FLG EN1092-1 PN40 B1");
                    comboBox24.Items.Add("ASME B16.5-300 RF");

                    break;
                case "PN100":
                    comboBox24.Items.Clear();
                    comboBox24.Items.Add("FLG EN1092-1 PN16 B1");
                    comboBox24.Items.Add("ASME B16.5-150 RF");
                    comboBox24.Items.Add("FLG EN1092-1 PN25 B1");
                    comboBox24.Items.Add("FLG EN1092-1 PN40 B1");
                    comboBox24.Items.Add("ASME B16.5-300 RF");
                    comboBox24.Items.Add("EN 10226 BSPT");
                    comboBox24.Items.Add("ASME B1.20.1 NPT");
                    comboBox24.Items.Add("SW ASME B16.11");
                    comboBox24.Items.Add("BW ASME B16.25");

                    break;
                case "#300":
                    comboBox24.Items.Clear();
                    comboBox24.Items.Add("FLG EN1092-1 PN16 B1");
                    comboBox24.Items.Add("ASME B16.5-150 RF");
                    comboBox24.Items.Add("FLG EN1092-1 PN25 B1");
                    comboBox24.Items.Add("FLG EN1092-1 PN40 B1");
                    comboBox24.Items.Add("ASME B16.5-300 RF");
                    break;
                case "#600":
                    comboBox24.Items.Clear();
                    comboBox24.Items.Add("FLG EN1092-1 PN16 B1");
                    comboBox24.Items.Add("ASME B16.5-150 RF");
                    comboBox24.Items.Add("FLG EN1092-1 PN25 B1");
                    comboBox24.Items.Add("FLG EN1092-1 PN40 B1");
                    comboBox24.Items.Add("ASME B16.5-300 RF");
                    comboBox24.Items.Add("EN 10226 BSPT");
                    comboBox24.Items.Add("ASME B1.20.1 NPT");
                    comboBox24.Items.Add("SW ASME B16.11");
                    comboBox24.Items.Add("BW ASME B16.25");
                    break;




            }
            comboBox24.SelectedIndex = 0;
        }

        private void txtAPSMaxF_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtAPSMaxF.Text))
            {

            }
            else
            {
                if (int.Parse(AcceptableNoise.SelectedItem.ToString()) > Double.Parse(txtAPSMaxF.Text))

                {
                    textBox3.Text = ("Noise OK");
                    textBox3.BackColor = Color.Green;
                    richTextBox1.Select(richTextBox1.TextLength, 0);
                    richTextBox1.SelectionColor = Color.Green;
                    richTextBox1.AppendText("\nNOISE LEVELS ARE OK");
                    //richTextBox1.Text=("\nNOISE LEVELS ARE OK");


                }
                else
                {
                    textBox3.Text = ("Too high. Low noise trim 1/2 to be selected");
                    textBox3.BackColor = Color.Red;
                    
                    richTextBox1.Select(richTextBox1.TextLength, 0);
                    richTextBox1.SelectionColor = Color.Red;
                    richTextBox1.AppendText("\nNOISE TOO HIGH. SELECT TRIM TO LOWER");
                    //richTextBox1.Text=("\nNOISE TOO HIGH.SELECT TRIM TO LOWER");
                }
            }
            


        }
        private void cmbTType_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (cmbTType.SelectedIndex == 1)
            {

                txtAPSMaxF.Text = Math.Round(double.Parse((Variables.sound - 15).ToString()),2).ToString();
                txtAPSNF.Text = (Variables.soundNorm - 15).ToString();
                txtAPSMinF.Text = (Variables.soundMin - 15).ToString();
                textBox14.Text = (Variables.sigma + 0.2).ToString();



            }
            else if (cmbTType.SelectedIndex == 2)
            {
                txtAPSMaxF.Text = Math.Round(double.Parse((Variables.sound - 25).ToString()), 2).ToString();
                txtAPSNF.Text = (Variables.soundNorm - 25).ToString();
                txtAPSMinF.Text = (Variables.soundMin - 25).ToString();
                textBox14.Text = (Variables.sigma + 0.3).ToString();
            }
            else
            {
                txtAPSMaxF.Text = Math.Round(double.Parse((Variables.sound).ToString()), 2).ToString();
                txtAPSNF.Text = (Variables.soundNorm).ToString();
                txtAPSMinF.Text = (Variables.soundMin).ToString();
                textBox14.Text = Variables.sigma.ToString();


            }
            if (cmbTType.SelectedItem.ToString() == "Standard Parabolic")
            {
                comboBox19.Items.Clear();
                comboBox19.Items.Add("316L SS");
            }
            else if (cmbTType.SelectedItem.ToString() == "Anticav 1")
            {
                comboBox19.Items.Clear();
                comboBox19.Items.Add("316L SS + Stellite");
            }
            else if (cmbTType.SelectedItem.ToString() == "Anticav 2")
            {
                comboBox19.Items.Clear();
                comboBox19.Items.Add("316L SS + Stellite");
            }
            else if (cmbTType.SelectedItem.ToString() == "low noise 1")
            {
                comboBox19.Items.Clear();
                comboBox19.Items.Add("316L SS + Stellite");
            }
            else if (cmbTType.SelectedItem.ToString() == "Low noise 2")
            {
                comboBox19.Items.Clear();
                comboBox19.Items.Add("316L SS + Stellite");
            }
            else
            {
                comboBox19.Items.Clear();
                comboBox19.Items.Add("316L SS + Stellite");
            }
            comboBox19.SelectedIndex = 0;


        }


        private void txtRFCMaxF_TextChanged(object sender, EventArgs e)
        {
            //if (Convert.ToInt32(txtRFCMaxF.Text) <= 2.86)
            //{
            //    comboBox9.Items.Clear();
            //    comboBox9.Items.Add(15);
            //    comboBox9.Items.Add(20);
            //    comboBox9.Items.Add(25);
            //}
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked==false)
            {
                groupBox1.Visible = false;
                groupBox14.Visible = false;
            }
            else
            {
                groupBox1.Visible = true;
                groupBox14.Visible = true;

            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == false)
            {
                groupBox2.Visible = false;
                groupBox17.Visible = false;

            }
            else
            {
                groupBox2.Visible = true;
                groupBox17.Visible = true;

            }
        }

        private void cmbTBU_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked==true)
            {
                txtSWMWNF.ReadOnly=false;
                txtVSHRNF.ReadOnly = false;
            }
            else
            {
                txtSWMWNF.ReadOnly = true;
                txtVSHRNF.ReadOnly = true;
            }
        }

        private void cmbIPUnit2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbIPUnit1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbIPUnit1.SelectedIndex==1)
            {
                cmbIPUnit2.Visible = true;
            }
            else
            {
                cmbIPUnit2.Visible = false;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox12.Text == "Bellow seal + Graphite" || comboBox12.Text == "Bellow seal + PTFE")
            {
                comboBox8.SelectedItem = "Bellow seal";
            }
            else
            {
                comboBox8.SelectedIndex = -1;
            }
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            //comboBox15.SelectedIndex = comboBox11.SelectedIndex;
            //cmbTSize.Text = comboBox15.SelectedItem.ToString();

            double seatt = double.Parse(comboBox11.SelectedItem.ToString());
            switch (seatt)
            {
                case 0.059:
                    cmbTSize.Text = "3";
                    break;
                case 0.165:
                case 0.33:
                case 0.55:
                case 0.825:
                case 1.1:
                    cmbTSize.Text = "6";
                    break;
                case 1.43:
                    cmbTSize.Text = "9";
                    break;
                case 1.65:
                    cmbTSize.Text = "10";
                    break;
                case 2.2:
                case 2.53:
                case 3.3:
                    cmbTSize.Text = "12";
                    break;

                case 4.95:
                    cmbTSize.Text = "15";
                    break;
                case 6.93:
                    cmbTSize.Text = "19";
                    break;
                case 12.1:
                    cmbTSize.Text = "25";
                    break;
                case 19.8:
                    cmbTSize.Text = "32";
                    break;
                case 33:
                    cmbTSize.Text = "40";
                    break;
                case 51.26:
                    cmbTSize.Text = "50";
                    break;
                case 79.97:
                    cmbTSize.Text = "64";
                    break;
                case 115.5:
                    cmbTSize.Text = "76";
                    break;
                case 176:
                    cmbTSize.Text = "100";
                    break;
                case 293.7:
                    cmbTSize.Text = "126";
                    break;
                case 404.8:
                    cmbTSize.Text = "151";
                    break;
                case 711.7:
                    cmbTSize.Text = "201";
                    break;
                    

            }
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox18.SelectedIndex = comboBox16.SelectedIndex;
            if (comboBox18.SelectedItem.ToString() == "NAA" /*&& selectedsize >= 50*/ )
            {
                checkBox1.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
            }
        }

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox19.SelectedItem.ToString() == "316L SS")
            {
                comboBox20.Items.Clear();
                comboBox20.Items.Add("Metal Class IV");
            }
            else if (comboBox19.SelectedItem.ToString() == "316L SS + Stellite")
            {
                comboBox20.Items.Clear();
                comboBox20.Items.Add("Metal-Hardened Class V");
            }
            else
            {
                    comboBox20.Items.Clear();
                    comboBox20.Items.Add("Soft Class VI");
            }
            comboBox20.SelectedIndex = 0;

        }
        private void txtITMaxF_TextChanged(object sender, EventArgs e)
        {
            //if (Variables.kelvin1 > 73.15 && Variables.kelvin1 <= 473.15 && checkBox1.Checked == true /*comboBox18.Items[0].ToString() == "NAA"*/)
            //{
            //    cmbTBU.SelectedIndex = 0;
            //}

            //else if (Variables.kelvin1 > 73.15 && Variables.kelvin1 <= 673.15 && checkBox1.Checked == true /*comboBox18.Items[0].ToString() == "NAA"*/)
            //{
            //    cmbTBU.SelectedIndex = 1;
            //}
            //else
            //{
            //    cmbTBU.SelectedIndex = 3;
            //}
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            if (double.Parse(textBox14.Text)> 1.5 && double.Parse(textBox14.Text) <= 2.0)
            {
                textBox11.Text = ("Incipient cavitation");
                textBox11.BackColor = Color.Yellow;
                


            }
            else if (double.Parse(textBox14.Text) > 1.3 && double.Parse(textBox14.Text) <= 1.5)
            {
                textBox11.Text = ("Medium cavitation");
                textBox11.BackColor = Color.Orange;
            }

            else if (double.Parse(textBox14.Text) > 1.0 && double.Parse(textBox14.Text) <= 1.3)
            {
                textBox11.Text = ("Full cavitation");
                textBox11.BackColor = Color.Red;
               
            }
            else
            {
                textBox11.Text = ("No cavitation");
                textBox11.BackColor = Color.Green;
                

            }
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //new method

            //Application excel3 = new Application();


            //_Excel.Workbook workbook2 = excel3.Workbooks.Open(filename + "\\Demo.xlsx", ReadOnly: false, Editable: true);
            //_Excel.Worksheet worksheet = workbook.Worksheets["printout"] as _Excel.Worksheet;


            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "taskkill /f /im excel.exe";
            process.StartInfo = startInfo;
            process.Start();
            //process.WaitForExit();

            //added
            process.Dispose();
            process.Close();




            _PrintExcel.Workbook workbook3 = new _PrintExcel.Workbook();
            workbook3.LoadFromFile(filename + "\\Demo.xlsx");

            _PrintExcel.Worksheet sheet1 = workbook3.Worksheets["printout"];


            //PrintDialog dialog = new PrintDialog();
            //dialog.AllowPrintToFile = true;
            //dialog.AllowCurrentPage = true;
            //dialog.AllowSomePages = true;
            //dialog.AllowSelection = true;
            //dialog.UseEXDialog = true;
            //dialog.PrinterSettings.Duplex = Duplex.Simplex;
            //dialog.PrinterSettings.PrintRange = PrintRange.SomePages;
            //workbook2.PrintDialog = dialog;
            PrintDocument pd = workbook3.PrintDocument;
            //if (dialog.ShowDialog() == DialogResult.OK)
            //{ pd.Print(); }



            sheet1.SaveToPdf(path + "Sizing Printout.pdf");
            System.Diagnostics.Process.Start("explorer.exe", path + "Sizing Printout.pdf");
           
            sheet1.Dispose();
            workbook3.Dispose();
            
            
            //workbook3.Application.Quit();
            //workbook3.Quit();



        }

        private void groupBox20_Enter(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog;
            saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Export Excel File";
            saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.FileName = "Valve printout";
            saveFileDialog.InitialDirectory = "Documents";
            saveFileDialog.CheckPathExists = true;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //source file
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wbSource = excel.Workbooks.Open(filename + "\\Demo.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet wsSource = wbSource.Worksheets["printout"];
                //copy all the data from source
                wsSource.UsedRange.Copy(Type.Missing);
                //destination file
                Microsoft.Office.Interop.Excel.Workbook wbDest = excel.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet wsDest = wbDest.ActiveSheet;
                //paste the copy data to destination
                wsSource.Copy(wsDest);
                //set name of the excel worksheet
                wsDest.Name = "1";
                // save to savedialogbox location
                wbDest.SaveAs(saveFileDialog.FileName);

                //close all the workbook and excel application handler
                wbSource.Close();
                wbDest.Close();
                //excel.ActiveWorkbook.Save();
                excel.DisplayAlerts = false;
                excel.Quit();
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (checkedListBox1.GetItemCheckState(9) == System.Windows.Forms.CheckState.Checked)
            //{
            //    checkedListBox1.SetItemChecked(10, false);
            //}
            //else if (checkedListBox1.GetItemCheckState(10) == System.Windows.Forms.CheckState.Checked)
            //{
            //    checkedListBox1.SetItemChecked(9, false);
            //}
        }

        private void checkedListBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (checkedListBox1.GetItemCheckState(9) == System.Windows.Forms.CheckState.Checked)
            {
                checkedListBox1.SetItemChecked(10, false);
            }
            else if (checkedListBox1.GetItemCheckState(10) == System.Windows.Forms.CheckState.Checked)
            {
                checkedListBox1.SetItemChecked(9, false);
            }
            
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox7.SelectedIndex == 1)
            {
                if (comboBox17.FindString("Sauter AVF234SF132/232") == comboBox17.SelectedIndex)
                {
                    comboBox21.Items.Clear();
                    comboBox21.Items.Add("Fail Close");
                    comboBox21.Items.Add("Fail Open");

                }
                else
                {
                    comboBox21.Items.Clear();
                    comboBox21.Items.Add("Fail in Place");

                }
            }
            else
            {
                comboBox21.Items.Clear();
                comboBox21.Items.Add("Fail Close");
                comboBox21.Items.Add("Fail Open");
            }
            comboBox21.SelectedIndex = 0;

        }
     

        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }
    }


}


