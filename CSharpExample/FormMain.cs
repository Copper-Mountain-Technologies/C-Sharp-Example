// Copyright ©2016 Copper Mountain Technologies
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR
// ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH
// THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

using System;
using System.Threading;
using System.Windows.Forms;

namespace CSharpExample
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            // --------------------------------------------------------------------------------------------------------
            /*
            R54Lib.RVNA app = null;

            // get type
            Type type = Type.GetTypeFromProgID("RVNA.Application", "localhost");
            if (type != null)
            {
                // create com interface object
                app = (R54Lib.RVNA)Activator.CreateInstance(type);
            }
            */
            // --------------------------------------------------------------------------------------------------------
            /*
            TR1300Lib.TRVNA app = null;

            // get type
            Type type = Type.GetTypeFromProgID("TRVNA.Application", "localhost");
            if (type != null)
            {
                // create com interface object
                app = (TR1300Lib.TRVNA)Activator.CreateInstance(type);
            }
            */
            // --------------------------------------------------------------------------------------------------------

            S2VNALib.S2VNA app = null;

            // get type
            Type type = Type.GetTypeFromProgID("S2VNA.Application", "localhost");
            if (type != null)
            {
                // create com interface object
                app = (S2VNALib.S2VNA)Activator.CreateInstance(type);
            }

            // --------------------------------------------------------------------------------------------------------
            /*
            S4VNALib.S4VNA app = null;

            // get type
            Type type = Type.GetTypeFromProgID("S4VNA.Application", "localhost");
            if (type != null)
            {
                // create com interface object
                app = (S4VNALib.S4VNA)Activator.CreateInstance(type);
            }
            */
            // --------------------------------------------------------------------------------------------------------

            textBox.Clear();
            textBox.AppendText("Waiting for VNA\n");
            textBox.AppendText("\n");

            if (app == null)
            {
                return;
            }

            /*
            int n = 0;
            while ((app.READy != true) && (n < 100))
            {
                Thread.Sleep(100);
                n++;
                textBox.AppendText(".");
            }
            if (app.READy == false)
            {
                return;
            }
            */

            int n = 0;
            while ((app.Ready != true) && (n < 100))
            {
                Thread.Sleep(100);
                n++;
                textBox.AppendText(".");
            }
            if (app.Ready == false)
            {
                return;
            }

            int channelNumber = 1;
            int traceNumber = 1;

            textBox.AppendText("VNA Ready\n");
            textBox.AppendText("\n");

            textBox.AppendText("Sending System Preset\n");
            app.SCPI.SYSTem.PRESet();
            textBox.AppendText("\n");

            // ========================================================================================================

            // ========================================================================================================

            textBox.AppendText("Setting Trace 1 of Channel 1 to S21\n");
            app.SCPI.CALCulate[channelNumber].PARameter[traceNumber].DEFine = "S21";
            textBox.AppendText("\n");

            textBox.AppendText("Retrieving Frequency Data\n");
            double[] F = app.SCPI.SENSe.FREQuency.DATA;
            textBox.AppendText("\n");

            textBox.AppendText("Retrieving S21 Parameter Data from Trace 1 of Channel 1\n");
            double[] S21 = app.SCPI.CALCulate[channelNumber].TRACe[traceNumber].DATA.SDATa;
            textBox.AppendText("\n");

            // loop thru all frequency points
            for (int i = 0; i < F.Length; i++)
            {
                // print frequency
                textBox.AppendText("Freq: " + string.Format("{0:0.###}", F[i] / 1000000) + " MHz\n");

                // print real part
                textBox.AppendText("Real: " + S21[i * 2] + "\n");

                // print imaginary part
                textBox.AppendText("Imag: " + S21[i * 2 + 1] + "\n");

                textBox.AppendText("\n");
            }

            app.SCPI.SYSTem.TERMinate();

            /*
    object err = app.SCPI.CALCulate[1].SELected.FORMat = "DSWR";

    textBox.AppendText("Retrieving Frequency Data\n");
    double[] F = app.SCPI.SENSe.FREQuency.DATA;
    textBox.AppendText("\n");

    textBox.AppendText("Retrieving DRLOss Data from Trace 1 of Channel 1\n");
    double[] DRLOss = app.SCPI.CALCulate[channelNumber].TRACe[traceNumber].DATA.SDATa;
    textBox.AppendText("\n");
            */
        }
    }
}