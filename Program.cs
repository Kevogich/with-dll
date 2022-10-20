using System;
using System.Windows.Forms;
using SharpFluids;
using UnitsNet;
using FLUIDS;


namespace AID2
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.Hide unhide forms 
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
    

}
