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

namespace GetResultFormulas
{

    
    public partial class Form1 : Form
    {

        public class Variables
        {
            public static double sound;
            public static double m3h;
            public static double kgh;
            public static int power;
            public static double kelvin1;
            public static double nm3h1;
            public static double p1;

        }
            public Form1()
        {
            InitializeComponent();
            filename = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Demo.xlsx";
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



            cmbTType.Items.Add("No Perforated Plug");
            cmbTType.Items.Add("low noise 1");
            cmbTType.Items.Add("Low noise 2");
            cmbTType.Items.Add("Multi-Cage");


            cmbTC.Items.Add("Equal percentage");
            cmbTC.Items.Add("Linear");
            cmbTC.Items.Add("On-Off");
            cmbTC.Items.Add("Bi-Linear");
            cmbTC.Items.Add("Tri-Linear");
            cmbTC.Items.Add("Soecial");

            //cmbTBU.Items.Add("Unbalanced");
            cmbTBU.Items.Add("PTFE Rings [L]");
            cmbTBU.Items.Add("PTFE Rings [G]");
            cmbTBU.Items.Add("Graphite rings [L+Sat.Steam]");
            cmbTBU.Items.Add("Graphite rings [G]");
            cmbTBU.Items.Add("Steel Rings [L+Sat.Steam]");
            cmbTBU.Items.Add("Steel Rings [G]");






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




            cmbITSUnit.SelectedIndex = 1;
            cmbITUnit.SelectedIndex = 1;
            cmbIPUnit2.SelectedIndex = 1;
            cmbIPUnit1.SelectedIndex = 1;
            cmbFluid.SelectedIndex = -1;

            cmbState.SelectedIndex = 0;
            cmbTBU.SelectedIndex = 1;
            cmbTC.SelectedIndex = 0;

            cmbTType.SelectedIndex = 0;
            cmbUnit.SelectedIndex = 0;
            //cmbVBFD.SelectedIndex = 1;



        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;


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

            _Excel.Workbook workbook = excel.Workbooks.Open(filename, ReadOnly: false, Editable: true);
            _Excel.Worksheet worksheet = workbook.Worksheets["SELECTION"] as _Excel.Worksheet;
            if (worksheet == null)
                return;
                
                comboBox9.Items.Clear();
                HashSet<string> distinct = new HashSet<string>();


                foreach (_Excel.Range cell in range.Cells)
                {
                    string value = (cell.Value2).ToString();

                    if (distinct.Add(value))
                        comboBox9.Items.Add(value);
                }
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

                Range row31 = worksheet.Rows.Cells[48, 10];
                row31.Value = cmbTBU.SelectedItem.ToString();

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
                row43.Value = txtITSMaxF.Text;

                Range row44 = worksheet.Rows.Cells[17, 24];
                row44.Value = txtITSNF.Text;

                Range row45 = worksheet.Rows.Cells[17, 30];
                row45.Value = txtITSMinF.Text;

                txtCPress.Text = isValidS(worksheet.Rows.Cells[10, 30]);
               // textBox10.Text = (Math.Round(double.Parse(txtCPress.Text), 2)*2).ToString();
                txtSWMWNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[18, 24])), 2).ToString();
                txtVSHRNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[19, 24])), 2).ToString();
                txtRFCMaxF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[21, 18])),2).ToString();
                txtRFCNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[21, 24])), 2).ToString();
                txtRFCMinF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[21, 30])), 3).ToString();
                cmbTSize.Text = isValidS(worksheet.Rows.Cells[46, 6]);
                txtVVMaxF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[24, 18])), 2).ToString();
                txtVVNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[24, 24])), 2).ToString();
                txtVVMinF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[24, 30])), 3).ToString();
                
                txtAPSNF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[23, 27])), 2).ToString();
                txtAPSMinF.Text = Math.Round(double.Parse(isValidS(worksheet.Rows.Cells[23, 33])), 3).ToString();
                txtLPD.Text = isValidS(worksheet.Rows.Cells[26, 13]);
                txtAMM.Text = worksheet.Rows.Cells[27, 26].Value.ToString();
                textBox6.Text = worksheet.Rows.Cells[29, 16].Value.ToString();
                textBox5.Text = worksheet.Rows.Cells[26, 28].Value.ToString();
                textBox1.Text = isValidS(worksheet.Rows.Cells[75, 43]);
                textBox2.Text = isValidS(worksheet.Rows.Cells[76, 43]);
                textBox3.Text = worksheet.Rows.Cells[73, 43].Value.ToString();
                textBox4.Text = worksheet.Rows.Cells[72, 43].Value.ToString();
                textBox8.Text = worksheet.Rows.Cells[78, 43].Value.ToString();
                textBox9.Text = worksheet.Rows.Cells[79, 43].Value.ToString();
                textBox7.Text = worksheet.Rows.Cells[81, 43].Value.ToString();
                string f = worksheet.Rows.Cells[82, 43].Value.ToString();
                string g = worksheet.Rows.Cells[83, 43].Value.ToString();
                string h = worksheet.Rows.Cells[84, 43].Value.ToString();
                
                Variables.kelvin1 = worksheet.Rows.Cells[88, 43].Value;
                Variables.nm3h1 = worksheet.Rows.Cells[89, 43].Value;
                Variables.p1 = worksheet.Rows.Cells[90, 43].Value;
                Variables.sound = worksheet.Rows.Cells[23, 21].Value;
                Variables.m3h = worksheet.Rows.Cells[86, 43].Value;
                Variables.kgh = worksheet.Rows.Cells[87, 43].Value;

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
                /*                double d = double.Parse(textBox1.Text);
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

                _ = MessageBox.Show(ee.ToString());
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
                groupBox14.Text = ("Velocity m/s");
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


            }
            else if ((string)cmbState.SelectedItem == "Steam Saturated")
            {
                cmbFluid.Items.Add("STEAM Saturated");
                cmbFluid.SelectedItem = ("STEAM Saturated");
                cmbUnit.Items.Add("Kg/h");
                cmbUnit.Items.Add("t/h");
                cmbUnit.Items.Add("L/h");
                groupBox14.Text = ("Velocity mach");
                label40.Text = ("Spec Heat Ratio");

            }
            else if ((string)cmbState.SelectedItem == "Gas")
            {
                cmbFluid.Items.Add("Acetylene");
                cmbFluid.Items.Add("AIR");
                cmbFluid.Items.Add("AMMONIA");
                cmbFluid.Items.Add("ARGON");
                groupBox14.Text = ("Velocity mach");
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



            }


            else if ((string)cmbState.SelectedItem == "Steam Superheated")
            {
                cmbFluid.Items.Add("STEAM Superheated");
                cmbFluid.SelectedItem = ("STEAM Superheated");
                cmbUnit.Items.Add("Kg/h");
                cmbUnit.Items.Add("Kg/s");
                cmbUnit.Items.Add("t/h");
                cmbUnit.Items.Add("L/h");
                groupBox14.Text = ("Velocity mach");
                label40.Text = ("Spec Heat Ratio");

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

        private void groupBox19_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox7.SelectedIndex == 1)
            {
                comboBox8.Visible = false;
                label19.Visible = false;
                txtAMM.Visible = true;
                label53.Visible = true;
                label48.Visible = false;
                textBox5.Visible = false;
                comboBox1.Items.Clear();
                comboBox1.Items.Add("Pillar");
                comboBox1.Items.Add("Handwheel");
                comboBox1.Items.Add("Voltage/supply");
                comboBox1.Items.Add("2-3 Points input");
                comboBox1.Items.Add("4-20 mA input");
                comboBox1.Items.Add("20-4 mA input");
                comboBox1.Items.Add("0-10 V input");
                comboBox1.Items.Add("No Output");
                comboBox1.Items.Add("4-20 mA Output ");
                comboBox1.Items.Add("0-10 V Output ");
                checkedListBox2.Items.Clear();
                checkedListBox2.Items.Add("Pillar");
                checkedListBox2.Items.Add("Handwheel");
                checkedListBox2.Items.Add("Voltage/supply");
                checkedListBox2.Items.Add("2-3 Points input");
                checkedListBox2.Items.Add("4-20 mA input");
                checkedListBox2.Items.Add("20-4 mA input");
                checkedListBox2.Items.Add("0-10 V input");
                checkedListBox2.Items.Add("No Output");
                checkedListBox2.Items.Add("4-20 mA Output ");
                checkedListBox2.Items.Add("0-10 V Output ");
            }


            else//not electric
            {
                comboBox8.Visible = true;
                label19.Visible = true;
                txtAMM.Visible = false;
                label53.Visible = false;
                label48.Visible = true;
                textBox5.Visible = true;
                comboBox1.Items.Clear();
                comboBox1.Items.Add("Pillar Yoke 210mm Zinc");
                comboBox1.Items.Add("Pillar Yoke 245mm Zinc");
                comboBox1.Items.Add("Stainless Steel Actuator");
                comboBox1.Items.Add("Handwheel");
                checkedListBox2.Items.Clear();
                checkedListBox2.Items.Add("Pillar Yoke 210mm Zinc");
                checkedListBox2.Items.Add("Handwheel");
                checkedListBox2.Items.Add("Pillar Yoke 245mm Zinc");
                checkedListBox2.Items.Add("Stainless Steel Actuator");


            }
            
            


        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

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

            _PrintExcel.Workbook workbook = new _PrintExcel.Workbook();
            workbook.LoadFromFile(filename);

            _PrintExcel.Worksheet sheet = workbook.Worksheets["printout"];
            sheet.SaveToPdf(path + "Sizing Printout.pdf");
            System.Diagnostics.Process.Start("explorer.exe", path+"Sizing Printout.pdf");
            workbook.Dispose();



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
            comboBox9.SelectedItem = textBox1.Text;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            comboBox10.SelectedItem = textBox6.Text;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
        
        private void cmbTType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbTType.SelectedIndex == 1)
            {
                txtAPSMaxF.Text = (Variables.sound - 15).ToString();
            }
            else if (cmbTType.SelectedIndex == 2)
            {
                txtAPSMaxF.Text = (Variables.sound - 25).ToString();
            }
            else
            {
                txtAPSMaxF.Text = (Variables.sound).ToString();
            }
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            Variables.power = int.Parse(comboBox9.SelectedItem.ToString());
            if (cmbState.SelectedIndex==0)
            {

                textBox10.Text = ((278*Variables.m3h)/(Math.Pow(Variables.power /2, 2)*3.14)).ToString();//(278 * Variables.m3h / Math.Pow(power,2) *3.14).ToString();//(278 * Variables.m3h / (power) * (power) * 3.14).ToString();
            }
            else
            {
                textBox10.Text = (((1.296 * Variables.nm3h1) * Variables.kelvin1 / (Math.Pow(Variables.power, 2) * (Variables.p1)) / 340).ToString());
            }
        }
    }
}


