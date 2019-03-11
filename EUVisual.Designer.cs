using System.Drawing;
using System.Windows.Forms;

namespace ExcelUnlockerVisual {
    partial class EUVisual {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EUVisual));
            this.btnChooseFile = new System.Windows.Forms.Button();
            this.tbFilePath = new System.Windows.Forms.TextBox();
            this.ofdFilePath = new System.Windows.Forms.OpenFileDialog();
            this.btnUnlock = new System.Windows.Forms.Button();
            this.cbOverwrite = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.rtbConsole = new System.Windows.Forms.RichTextBox();
            this.cbUnlockVBA = new System.Windows.Forms.CheckBox();
            this.pbUnlocker = new System.Windows.Forms.ProgressBar();
            this.bwProgress = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // btnChooseFile
            // 
            this.btnChooseFile.Location = new System.Drawing.Point(24, 34);
            this.btnChooseFile.Name = "btnChooseFile";
            this.btnChooseFile.Size = new System.Drawing.Size(110, 23);
            this.btnChooseFile.TabIndex = 0;
            this.btnChooseFile.Text = "Choose File...";
            this.btnChooseFile.UseVisualStyleBackColor = true;
            this.btnChooseFile.Click += new System.EventHandler(this.btnChooseFile_Click);
            // 
            // tbFilePath
            // 
            this.tbFilePath.Enabled = false;
            this.tbFilePath.Location = new System.Drawing.Point(156, 36);
            this.tbFilePath.Name = "tbFilePath";
            this.tbFilePath.ReadOnly = true;
            this.tbFilePath.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.tbFilePath.Size = new System.Drawing.Size(292, 20);
            this.tbFilePath.TabIndex = 1;
            this.tbFilePath.Text = "No file selected";
            // 
            // btnUnlock
            // 
            this.btnUnlock.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnUnlock.Location = new System.Drawing.Point(24, 129);
            this.btnUnlock.Name = "btnUnlock";
            this.btnUnlock.Size = new System.Drawing.Size(75, 23);
            this.btnUnlock.TabIndex = 2;
            this.btnUnlock.Text = "Unlock";
            this.btnUnlock.UseVisualStyleBackColor = false;
            this.btnUnlock.Click += new System.EventHandler(this.btnUnlock_Click);
            // 
            // cbOverwrite
            // 
            this.cbOverwrite.AutoSize = true;
            this.cbOverwrite.Location = new System.Drawing.Point(24, 72);
            this.cbOverwrite.Name = "cbOverwrite";
            this.cbOverwrite.Size = new System.Drawing.Size(120, 17);
            this.cbOverwrite.TabIndex = 3;
            this.cbOverwrite.Text = "Overwite original file";
            this.cbOverwrite.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(132, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Select Excel file to unlock:";
            // 
            // rtbConsole
            // 
            this.rtbConsole.BackColor = System.Drawing.SystemColors.Window;
            this.rtbConsole.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.rtbConsole.Enabled = false;
            this.rtbConsole.ForeColor = System.Drawing.Color.Black;
            this.rtbConsole.Location = new System.Drawing.Point(156, 68);
            this.rtbConsole.Name = "rtbConsole";
            this.rtbConsole.ReadOnly = true;
            this.rtbConsole.Size = new System.Drawing.Size(292, 96);
            this.rtbConsole.TabIndex = 5;
            this.rtbConsole.Text = "";
            // 
            // cbUnlockVBA
            // 
            this.cbUnlockVBA.AutoSize = true;
            this.cbUnlockVBA.Enabled = false;
            this.cbUnlockVBA.Location = new System.Drawing.Point(24, 95);
            this.cbUnlockVBA.Name = "cbUnlockVBA";
            this.cbUnlockVBA.Size = new System.Drawing.Size(84, 17);
            this.cbUnlockVBA.TabIndex = 6;
            this.cbUnlockVBA.Text = "Unlock VBA";
            this.cbUnlockVBA.UseVisualStyleBackColor = true;
            this.cbUnlockVBA.CheckedChanged += new System.EventHandler(this.cbUnlockVBA_CheckedChanged);
            // 
            // pbUnlocker
            // 
            this.pbUnlocker.Location = new System.Drawing.Point(24, 172);
            this.pbUnlocker.Name = "pbUnlocker";
            this.pbUnlocker.Size = new System.Drawing.Size(424, 23);
            this.pbUnlocker.TabIndex = 7;
            // 
            // EUVisual
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(465, 207);
            this.Controls.Add(this.pbUnlocker);
            this.Controls.Add(this.cbUnlockVBA);
            this.Controls.Add(this.rtbConsole);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbOverwrite);
            this.Controls.Add(this.btnUnlock);
            this.Controls.Add(this.tbFilePath);
            this.Controls.Add(this.btnChooseFile);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "EUVisual";
            this.Text = "ExcelUnlocker";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnChooseFile;
        private System.Windows.Forms.TextBox tbFilePath;
        private System.Windows.Forms.OpenFileDialog ofdFilePath;
        private System.Windows.Forms.Button btnUnlock;
        private System.Windows.Forms.CheckBox cbOverwrite;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox rtbConsole;
        private System.Windows.Forms.CheckBox cbUnlockVBA;
        private System.Windows.Forms.ProgressBar pbUnlocker;
        private System.ComponentModel.BackgroundWorker bwProgress;
    }

    public class NewProgressBar : ProgressBar {
        public NewProgressBar() {
            this.SetStyle(ControlStyles.UserPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e) {
            Rectangle rec = e.ClipRectangle;

            rec.Width = (int)(rec.Width * ((double)Value / Maximum)) - 4;
            if (ProgressBarRenderer.IsSupported)
                ProgressBarRenderer.DrawHorizontalBar(e.Graphics, e.ClipRectangle);
            rec.Height = rec.Height - 4;
            e.Graphics.FillRectangle(Brushes.Red, 2, 2, rec.Width, rec.Height);
        }
    }
}

