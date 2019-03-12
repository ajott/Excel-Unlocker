using System;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace ExcelUnlockerVisual {
    public partial class EUVisual : Form {

        bool hasDisplayedVBAWarning = false;
        bool isAlreadyRunning = false;

        public EUVisual() {
            InitializeComponent();
        }

        private void btnChooseFile_Click(object sender, EventArgs e) {
            ChooseFile();
        }


        public void ChooseFile() {
            string workingFilePath = "c:\\";

            if (Path.GetExtension(tbFilePath.Text) == ".xlsm" || Path.GetExtension(tbFilePath.Text) == ".xlsx" || Path.GetExtension(tbFilePath.Text) == ".xlam") {
                workingFilePath = Path.GetDirectoryName(tbFilePath.Text);
            }


            OpenFileDialog ofdFilePath = new OpenFileDialog {
                InitialDirectory = workingFilePath,
                Filter = "Excel workbook/add-in (*.xlsx; *.xlsm, *.xlam)|*.xlsx; *.xlsm; *.xlam",
                RestoreDirectory = true
            };

            if (ofdFilePath.ShowDialog() == DialogResult.OK) {
                tbFilePath.Text = ofdFilePath.FileName;
                if (Path.GetExtension(tbFilePath.Text) == ".xlsm" || Path.GetExtension(tbFilePath.Text) == ".xlam") {
                    cbUnlockVBA.Enabled = true;
                }
            }
        }

        private void UpdateConsole(int code) {
            if (code == 0) {
                rtbConsole.AppendText("File unlocked successfully", Color.DarkGreen);
                rtbConsole.NewLine();

                if (cbOverwrite.Checked) {
                    rtbConsole.AppendText("Saved as \"", Color.Black);
                    rtbConsole.AppendText(Path.GetFileName(tbFilePath.Text), Color.DarkGoldenrod);
                    rtbConsole.AppendText("\"", Color.Black);
                    rtbConsole.NewLine();
                } else {
                    rtbConsole.AppendText("Saved as \"", Color.Black);
                    rtbConsole.AppendText("unlocked_" + Path.GetFileName(tbFilePath.Text), Color.DarkGoldenrod);
                    rtbConsole.AppendText("\"", Color.Black);
                    rtbConsole.NewLine();
                }

                rtbConsole.AppendText("In directory \"", Color.Black);
                rtbConsole.AppendText(Path.GetDirectoryName(tbFilePath.Text), Color.DarkGoldenrod);
                rtbConsole.AppendText("\"", Color.Black);
                rtbConsole.NewLine();
            } else {
                // Sets progress bar to red
                ModifyProgressBarColor.SetState(pbUnlocker, 2);
                rtbConsole.AppendText("File unlock failed", Color.DarkRed);
                rtbConsole.NewLine();
                rtbConsole.AppendText("Could not write to directory.", Color.Black);
                rtbConsole.NewLine();
                rtbConsole.AppendText("User does not have permission, file is currently open, or file\n\"", Color.Black);
                rtbConsole.AppendText("unlocked_"+Path.GetFileName(tbFilePath.Text), Color.DarkGoldenrod);
                rtbConsole.AppendText("\" already exists \nand overwrite directive not supplied", Color.Black);
                rtbConsole.NewLine();
            }
        }
        

        private void btnUnlock_Click(object sender, EventArgs e) {
            // Sets progress bar to green
            ModifyProgressBarColor.SetState(pbUnlocker, 1);

            if (isAlreadyRunning) {
                return;
            } else {
                isAlreadyRunning = true;

                pbUnlocker.Value = 0;

                if (tbFilePath.Text == "No file selected") {
                    return;
                };


                Progress<int> progressBar = new Progress<int>();
                Progress<int> consoleProg = new Progress<int>();
                progressBar.ProgressChanged += (p, value) => pbUnlocker.Value = value;
                consoleProg.ProgressChanged += (p, value) => UpdateConsole(value);
                Task.Run(() => (UnlockClass.Unlock(tbFilePath.Text, cbOverwrite.Checked, cbUnlockVBA.Checked, progressBar, consoleProg)));


                rtbConsole.Text = "";
                btnChooseFile.Enabled = false;
                cbOverwrite.Enabled = false;
                bwProgress.RunWorkerAsync();
                bwProgress.WorkerReportsProgress = true;

                btnChooseFile.Enabled = true;
                cbOverwrite.Enabled = true;
                ScrollToBottomOfMessages();
                isAlreadyRunning = false;
            }
        }
        private void ScrollToBottomOfMessages() {
            rtbConsole.SelectionStart = rtbConsole.Text.Length;
            rtbConsole.ScrollToCaret();
        }

        private void cbUnlockVBA_CheckedChanged(object sender, EventArgs e) {
            if (cbUnlockVBA.Checked == true && !hasDisplayedVBAWarning) {
                hasDisplayedVBAWarning = true;
                MessageBox.Show("Warning: Unlocking VBA projects is dangerous. Always make a backup copy of your file " +
                    "before attempting. \n" +
                    "When opening the file after unlocking, accept all errors (click \"Yes\" or \"OK\"), " +
                    "then save and re-load the file.", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }
    }

    public static class RichTextBoxExtensions {
        public static void AppendText(this RichTextBox box, string text, Color color) {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;

            box.SelectionColor = color;
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
        }

        public static void NewLine(this RichTextBox box) {
            box.AppendText(Environment.NewLine);
        }
    }

    public static class ModifyProgressBarColor {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr w, IntPtr l);
        public static void SetState(this ProgressBar pBar, int state) {
            // States: 
            // 1: Green
            // 2: Red
            // 3: Yellow
            SendMessage(pBar.Handle, 1040, (IntPtr)state, IntPtr.Zero);
        }
    }
}
