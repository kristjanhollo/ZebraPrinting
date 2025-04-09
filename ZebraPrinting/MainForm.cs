using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;

namespace BulkZPLPrinter
{
    public partial class MainForm : Form
    {
        private List<string> zplFilesQueue = new List<string>();
        private List<string> zplContentQueue = new List<string>();
        private bool isNetworkPrinter = false;
        private PrinterSettings printerSettings = new PrinterSettings();

        public MainForm()
        {
            InitializeComponent();
            LoadPrinters();
        }

        private void LoadPrinters()
        {
            // Load available printers to the printer combobox
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                cmbPrinter.Items.Add(printer);
            }

            if (cmbPrinter.Items.Count > 0)
            {
                cmbPrinter.SelectedIndex = 0;
            }

            rbLocalPrinter.Checked = true;
            UpdatePrinterPanel();
        }

        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.gbPrinterSettings = new System.Windows.Forms.GroupBox();
            this.rbNetworkPrinter = new System.Windows.Forms.RadioButton();
            this.rbLocalPrinter = new System.Windows.Forms.RadioButton();
            this.panelLocalPrinter = new System.Windows.Forms.Panel();
            this.cmbPrinter = new System.Windows.Forms.ComboBox();
            this.lblLocalPrinter = new System.Windows.Forms.Label();
            this.panelNetworkPrinter = new System.Windows.Forms.Panel();
            this.nudPort = new System.Windows.Forms.NumericUpDown();
            this.lblIPPort = new System.Windows.Forms.Label();
            this.txtIPAddress = new System.Windows.Forms.TextBox();
            this.lblIPAddress = new System.Windows.Forms.Label();
            this.gbFiles = new System.Windows.Forms.GroupBox();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.btnRemoveSelected = new System.Windows.Forms.Button();
            this.lstZPLFiles = new System.Windows.Forms.ListBox();
            this.btnAddFiles = new System.Windows.Forms.Button();
            this.btnAddDirectory = new System.Windows.Forms.Button();
            this.gbPrintSettings = new System.Windows.Forms.GroupBox();
            this.lblCopies = new System.Windows.Forms.Label();
            this.nudCopies = new System.Windows.Forms.NumericUpDown();
            this.chkPreview = new System.Windows.Forms.CheckBox();
            this.btnPrint = new System.Windows.Forms.Button();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.progressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.rtbPreview = new System.Windows.Forms.RichTextBox();
            this.gbPreview = new System.Windows.Forms.GroupBox();
            this.gbPrinterSettings.SuspendLayout();
            this.panelLocalPrinter.SuspendLayout();
            this.panelNetworkPrinter.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudPort)).BeginInit();
            this.gbFiles.SuspendLayout();
            this.gbPrintSettings.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudCopies)).BeginInit();
            this.statusStrip.SuspendLayout();
            this.gbPreview.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(223, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "Bulk ZPL Print Utility";
            // 
            // gbPrinterSettings
            // 
            this.gbPrinterSettings.Controls.Add(this.rbNetworkPrinter);
            this.gbPrinterSettings.Controls.Add(this.rbLocalPrinter);
            this.gbPrinterSettings.Controls.Add(this.panelLocalPrinter);
            this.gbPrinterSettings.Controls.Add(this.panelNetworkPrinter);
            this.gbPrinterSettings.Location = new System.Drawing.Point(12, 36);
            this.gbPrinterSettings.Name = "gbPrinterSettings";
            this.gbPrinterSettings.Size = new System.Drawing.Size(386, 172);
            this.gbPrinterSettings.TabIndex = 1;
            this.gbPrinterSettings.TabStop = false;
            this.gbPrinterSettings.Text = "Printer Settings";
            // 
            // rbNetworkPrinter
            // 
            this.rbNetworkPrinter.AutoSize = true;
            this.rbNetworkPrinter.Location = new System.Drawing.Point(191, 22);
            this.rbNetworkPrinter.Name = "rbNetworkPrinter";
            this.rbNetworkPrinter.Size = new System.Drawing.Size(102, 17);
            this.rbNetworkPrinter.TabIndex = 3;
            this.rbNetworkPrinter.Text = "Network Printer";
            this.rbNetworkPrinter.UseVisualStyleBackColor = true;
            this.rbNetworkPrinter.CheckedChanged += new System.EventHandler(this.rbNetworkPrinter_CheckedChanged);
            // 
            // rbLocalPrinter
            // 
            this.rbLocalPrinter.AutoSize = true;
            this.rbLocalPrinter.Checked = true;
            this.rbLocalPrinter.Location = new System.Drawing.Point(16, 22);
            this.rbLocalPrinter.Name = "rbLocalPrinter";
            this.rbLocalPrinter.Size = new System.Drawing.Size(85, 17);
            this.rbLocalPrinter.TabIndex = 2;
            this.rbLocalPrinter.TabStop = true;
            this.rbLocalPrinter.Text = "Local Printer";
            this.rbLocalPrinter.UseVisualStyleBackColor = true;
            this.rbLocalPrinter.CheckedChanged += new System.EventHandler(this.rbLocalPrinter_CheckedChanged);
            // 
            // panelLocalPrinter
            // 
            this.panelLocalPrinter.Controls.Add(this.cmbPrinter);
            this.panelLocalPrinter.Controls.Add(this.lblLocalPrinter);
            this.panelLocalPrinter.Location = new System.Drawing.Point(6, 45);
            this.panelLocalPrinter.Name = "panelLocalPrinter";
            this.panelLocalPrinter.Size = new System.Drawing.Size(374, 54);
            this.panelLocalPrinter.TabIndex = 0;
            // 
            // cmbPrinter
            // 
            this.cmbPrinter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPrinter.FormattingEnabled = true;
            this.cmbPrinter.Location = new System.Drawing.Point(85, 17);
            this.cmbPrinter.Name = "cmbPrinter";
            this.cmbPrinter.Size = new System.Drawing.Size(278, 21);
            this.cmbPrinter.TabIndex = 1;
            // 
            // lblLocalPrinter
            // 
            this.lblLocalPrinter.AutoSize = true;
            this.lblLocalPrinter.Location = new System.Drawing.Point(7, 20);
            this.lblLocalPrinter.Name = "lblLocalPrinter";
            this.lblLocalPrinter.Size = new System.Drawing.Size(40, 13);
            this.lblLocalPrinter.TabIndex = 0;
            this.lblLocalPrinter.Text = "Printer:";
            // 
            // panelNetworkPrinter
            // 
            this.panelNetworkPrinter.Controls.Add(this.nudPort);
            this.panelNetworkPrinter.Controls.Add(this.lblIPPort);
            this.panelNetworkPrinter.Controls.Add(this.txtIPAddress);
            this.panelNetworkPrinter.Controls.Add(this.lblIPAddress);
            this.panelNetworkPrinter.Location = new System.Drawing.Point(6, 99);
            this.panelNetworkPrinter.Name = "panelNetworkPrinter";
            this.panelNetworkPrinter.Size = new System.Drawing.Size(374, 55);
            this.panelNetworkPrinter.TabIndex = 1;
            // 
            // nudPort
            // 
            this.nudPort.Location = new System.Drawing.Point(296, 17);
            this.nudPort.Maximum = new decimal(new int[] {
            65535,
            0,
            0,
            0});
            this.nudPort.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudPort.Name = "nudPort";
            this.nudPort.Size = new System.Drawing.Size(67, 20);
            this.nudPort.TabIndex = 3;
            this.nudPort.Value = new decimal(new int[] {
            9100,
            0,
            0,
            0});
            // 
            // lblIPPort
            // 
            this.lblIPPort.AutoSize = true;
            this.lblIPPort.Location = new System.Drawing.Point(261, 20);
            this.lblIPPort.Name = "lblIPPort";
            this.lblIPPort.Size = new System.Drawing.Size(29, 13);
            this.lblIPPort.TabIndex = 2;
            this.lblIPPort.Text = "Port:";
            // 
            // txtIPAddress
            // 
            this.txtIPAddress.Location = new System.Drawing.Point(85, 17);
            this.txtIPAddress.Name = "txtIPAddress";
            this.txtIPAddress.Size = new System.Drawing.Size(170, 20);
            this.txtIPAddress.TabIndex = 1;
            this.txtIPAddress.Text = "192.168.1.1";
            // 
            // lblIPAddress
            // 
            this.lblIPAddress.AutoSize = true;
            this.lblIPAddress.Location = new System.Drawing.Point(7, 20);
            this.lblIPAddress.Name = "lblIPAddress";
            this.lblIPAddress.Size = new System.Drawing.Size(61, 13);
            this.lblIPAddress.TabIndex = 0;
            this.lblIPAddress.Text = "IP Address:";
            // 
            // gbFiles
            // 
            this.gbFiles.Controls.Add(this.btnClearAll);
            this.gbFiles.Controls.Add(this.btnRemoveSelected);
            this.gbFiles.Controls.Add(this.lstZPLFiles);
            this.gbFiles.Controls.Add(this.btnAddFiles);
            this.gbFiles.Controls.Add(this.btnAddDirectory);
            this.gbFiles.Location = new System.Drawing.Point(12, 214);
            this.gbFiles.Name = "gbFiles";
            this.gbFiles.Size = new System.Drawing.Size(386, 274);
            this.gbFiles.TabIndex = 2;
            this.gbFiles.TabStop = false;
            this.gbFiles.Text = "ZPL Files";
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(191, 243);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(75, 23);
            this.btnClearAll.TabIndex = 4;
            this.btnClearAll.Text = "Clear All";
            this.btnClearAll.UseVisualStyleBackColor = true;
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // btnRemoveSelected
            // 
            this.btnRemoveSelected.Location = new System.Drawing.Point(85, 243);
            this.btnRemoveSelected.Name = "btnRemoveSelected";
            this.btnRemoveSelected.Size = new System.Drawing.Size(100, 23);
            this.btnRemoveSelected.TabIndex = 3;
            this.btnRemoveSelected.Text = "Remove Selected";
            this.btnRemoveSelected.UseVisualStyleBackColor = true;
            this.btnRemoveSelected.Click += new System.EventHandler(this.btnRemoveSelected_Click);
            // 
            // lstZPLFiles
            // 
            this.lstZPLFiles.FormattingEnabled = true;
            this.lstZPLFiles.Location = new System.Drawing.Point(16, 77);
            this.lstZPLFiles.Name = "lstZPLFiles";
            this.lstZPLFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstZPLFiles.Size = new System.Drawing.Size(352, 160);
            this.lstZPLFiles.TabIndex = 2;
            this.lstZPLFiles.SelectedIndexChanged += new System.EventHandler(this.lstZPLFiles_SelectedIndexChanged);
            // 
            // btnAddFiles
            // 
            this.btnAddFiles.Location = new System.Drawing.Point(16, 30);
            this.btnAddFiles.Name = "btnAddFiles";
            this.btnAddFiles.Size = new System.Drawing.Size(110, 30);
            this.btnAddFiles.TabIndex = 0;
            this.btnAddFiles.Text = "Add Files...";
            this.btnAddFiles.UseVisualStyleBackColor = true;
            this.btnAddFiles.Click += new System.EventHandler(this.btnAddFiles_Click);
            // 
            // btnAddDirectory
            // 
            this.btnAddDirectory.Location = new System.Drawing.Point(156, 30);
            this.btnAddDirectory.Name = "btnAddDirectory";
            this.btnAddDirectory.Size = new System.Drawing.Size(110, 30);
            this.btnAddDirectory.TabIndex = 1;
            this.btnAddDirectory.Text = "Add Directory...";
            this.btnAddDirectory.UseVisualStyleBackColor = true;
            this.btnAddDirectory.Click += new System.EventHandler(this.btnAddDirectory_Click);
            // 
            // gbPrintSettings
            // 
            this.gbPrintSettings.Controls.Add(this.lblCopies);
            this.gbPrintSettings.Controls.Add(this.nudCopies);
            this.gbPrintSettings.Controls.Add(this.chkPreview);
            this.gbPrintSettings.Controls.Add(this.btnPrint);
            this.gbPrintSettings.Location = new System.Drawing.Point(12, 494);
            this.gbPrintSettings.Name = "gbPrintSettings";
            this.gbPrintSettings.Size = new System.Drawing.Size(386, 105);
            this.gbPrintSettings.TabIndex = 3;
            this.gbPrintSettings.TabStop = false;
            this.gbPrintSettings.Text = "Print Settings";
            // 
            // lblCopies
            // 
            this.lblCopies.AutoSize = true;
            this.lblCopies.Location = new System.Drawing.Point(16, 33);
            this.lblCopies.Name = "lblCopies";
            this.lblCopies.Size = new System.Drawing.Size(42, 13);
            this.lblCopies.TabIndex = 3;
            this.lblCopies.Text = "Copies:";
            // 
            // nudCopies
            // 
            this.nudCopies.Location = new System.Drawing.Point(77, 31);
            this.nudCopies.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudCopies.Name = "nudCopies";
            this.nudCopies.Size = new System.Drawing.Size(61, 20);
            this.nudCopies.TabIndex = 2;
            this.nudCopies.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // chkPreview
            // 
            this.chkPreview.AutoSize = true;
            this.chkPreview.Location = new System.Drawing.Point(156, 32);
            this.chkPreview.Name = "chkPreview";
            this.chkPreview.Size = new System.Drawing.Size(85, 17);
            this.chkPreview.TabIndex = 1;
            this.chkPreview.Text = "Show ZPL";
            this.chkPreview.UseVisualStyleBackColor = true;
            this.chkPreview.CheckedChanged += new System.EventHandler(this.chkPreview_CheckedChanged);
            // 
            // btnPrint
            // 
            this.btnPrint.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPrint.Location = new System.Drawing.Point(106, 60);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(170, 35);
            this.btnPrint.TabIndex = 0;
            this.btnPrint.Text = "Print ZPL Files";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatus,
            this.progressBar});
            this.statusStrip.Location = new System.Drawing.Point(0, 612);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(784, 22);
            this.statusStrip.TabIndex = 4;
            this.statusStrip.Text = "statusStrip1";
            // 
            // lblStatus
            // 
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(42, 17);
            this.lblStatus.Text = "Ready";
            // 
            // progressBar
            // 
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(200, 16);
            this.progressBar.Visible = false;
            // 
            // rtbPreview
            // 
            this.rtbPreview.Location = new System.Drawing.Point(16, 19);
            this.rtbPreview.Name = "rtbPreview";
            this.rtbPreview.ReadOnly = true;
            this.rtbPreview.Size = new System.Drawing.Size(352, 522);
            this.rtbPreview.TabIndex = 5;
            this.rtbPreview.Text = "";
            // 
            // gbPreview
            // 
            this.gbPreview.Controls.Add(this.rtbPreview);
            this.gbPreview.Location = new System.Drawing.Point(404, 36);
            this.gbPreview.Name = "gbPreview";
            this.gbPreview.Size = new System.Drawing.Size(375, 563);
            this.gbPreview.TabIndex = 6;
            this.gbPreview.TabStop = false;
            this.gbPreview.Text = "ZPL Preview";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 634);
            this.Controls.Add(this.gbPreview);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.gbPrintSettings);
            this.Controls.Add(this.gbFiles);
            this.Controls.Add(this.gbPrinterSettings);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Bulk ZPL Printer";
            this.gbPrinterSettings.ResumeLayout(false);
            this.gbPrinterSettings.PerformLayout();
            this.panelLocalPrinter.ResumeLayout(false);
            this.panelLocalPrinter.PerformLayout();
            this.panelNetworkPrinter.ResumeLayout(false);
            this.panelNetworkPrinter.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudPort)).EndInit();
            this.gbFiles.ResumeLayout(false);
            this.gbPrintSettings.ResumeLayout(false);
            this.gbPrintSettings.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudCopies)).EndInit();
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.gbPreview.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox gbPrinterSettings;
        private System.Windows.Forms.Panel panelLocalPrinter;
        private System.Windows.Forms.Panel panelNetworkPrinter;
        private System.Windows.Forms.Label lblLocalPrinter;
        private System.Windows.Forms.ComboBox cmbPrinter;
        private System.Windows.Forms.Label lblIPAddress;
        private System.Windows.Forms.TextBox txtIPAddress;
        private System.Windows.Forms.Label lblIPPort;
        private System.Windows.Forms.NumericUpDown nudPort;
        private System.Windows.Forms.RadioButton rbNetworkPrinter;
        private System.Windows.Forms.RadioButton rbLocalPrinter;
        private System.Windows.Forms.GroupBox gbFiles;
        private System.Windows.Forms.Button btnAddFiles;
        private System.Windows.Forms.Button btnAddDirectory;
        private System.Windows.Forms.ListBox lstZPLFiles;
        private System.Windows.Forms.Button btnRemoveSelected;
        private System.Windows.Forms.Button btnClearAll;
        private System.Windows.Forms.GroupBox gbPrintSettings;
        private System.Windows.Forms.Button btnPrint;
        private System.Windows.Forms.CheckBox chkPreview;
        private System.Windows.Forms.Label lblCopies;
        private System.Windows.Forms.NumericUpDown nudCopies;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.ToolStripProgressBar progressBar;
        private System.Windows.Forms.RichTextBox rtbPreview;
        private System.Windows.Forms.GroupBox gbPreview;

        private void rbLocalPrinter_CheckedChanged(object sender, EventArgs e)
        {
            UpdatePrinterPanel();
        }

        private void rbNetworkPrinter_CheckedChanged(object sender, EventArgs e)
        {
            UpdatePrinterPanel();
        }

        private void UpdatePrinterPanel()
        {
            isNetworkPrinter = rbNetworkPrinter.Checked;
            panelLocalPrinter.Visible = !isNetworkPrinter;
            panelNetworkPrinter.Visible = isNetworkPrinter;
        }

        private void btnAddFiles_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "ZPL Files (*.zpl)|*.zpl|Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Select ZPL Files";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string file in openFileDialog.FileNames)
                    {
                        if (!zplFilesQueue.Contains(file))
                        {
                            zplFilesQueue.Add(file);
                            lstZPLFiles.Items.Add(Path.GetFileName(file));
                            // Read content for preview
                            zplContentQueue.Add(File.ReadAllText(file));
                        }
                    }

                    UpdateStatus($"Added {openFileDialog.FileNames.Length} file(s)");
                }
            }
        }

        private void btnAddDirectory_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select a folder containing ZPL files";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get all ZPL and TXT files in the directory
                    string[] zplFiles = Directory.GetFiles(folderDialog.SelectedPath, "*.zpl");
                    string[] txtFiles = Directory.GetFiles(folderDialog.SelectedPath, "*.txt");
                    string[] allFiles = zplFiles.Concat(txtFiles).ToArray();

                    int addedCount = 0;
                    foreach (string file in allFiles)
                    {
                        if (!zplFilesQueue.Contains(file))
                        {
                            zplFilesQueue.Add(file);
                            lstZPLFiles.Items.Add(Path.GetFileName(file));
                            // Read content for preview
                            zplContentQueue.Add(File.ReadAllText(file));
                            addedCount++;
                        }
                    }

                    UpdateStatus($"Added {addedCount} file(s) from directory");
                }
            }
        }

        private void btnRemoveSelected_Click(object sender, EventArgs e)
        {
            if (lstZPLFiles.SelectedIndices.Count > 0)
            {
                // We need to remove items in reverse order to maintain correct indices
                for (int i = lstZPLFiles.SelectedIndices.Count - 1; i >= 0; i--)
                {
                    int index = lstZPLFiles.SelectedIndices[i];
                    zplFilesQueue.RemoveAt(index);
                    zplContentQueue.RemoveAt(index);
                    lstZPLFiles.Items.RemoveAt(index);
                }

                UpdateStatus("Removed selected file(s)");

                // Clear preview if nothing selected
                if (lstZPLFiles.SelectedIndex == -1)
                {
                    rtbPreview.Clear();
                }
            }
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            zplFilesQueue.Clear();
            zplContentQueue.Clear();
            lstZPLFiles.Items.Clear();
            rtbPreview.Clear();
            UpdateStatus("Cleared all files");
        }

        private void lstZPLFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstZPLFiles.SelectedIndex >= 0 && chkPreview.Checked)
            {
                rtbPreview.Text = zplContentQueue[lstZPLFiles.SelectedIndex];
            }
        }

        private void chkPreview_CheckedChanged(object sender, EventArgs e)
        {
            gbPreview.Visible = chkPreview.Checked;
            if (this.Width < 800 && chkPreview.Checked)
            {
                this.Width = 800;
            }
            else if (!chkPreview.Checked)
            {
                this.Width = 414; // Original width without preview
            }

            // Show preview if item selected
            if (chkPreview.Checked && lstZPLFiles.SelectedIndex >= 0)
            {
                rtbPreview.Text = zplContentQueue[lstZPLFiles.SelectedIndex];
            }
        }

        private async void btnPrint_Click(object sender, EventArgs e)
        {
            if (zplFilesQueue.Count == 0)
            {
                MessageBox.Show("Please add ZPL files to print.", "No Files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int copies = (int)nudCopies.Value;

            // Configure progress bar
            progressBar.Visible = true;
            progressBar.Minimum = 0;
            progressBar.Maximum = zplFilesQueue.Count * copies;
            progressBar.Value = 0;

            bool success = true;

            // Disable controls during printing
            SetControlsEnabled(false);

            try
            {
                // Print each file the specified number of copies
                for (int i = 0; i < zplFilesQueue.Count; i++)
                {
                    string zplFile = zplFilesQueue[i];
                    string zplContent = zplContentQueue[i];

                    for (int c = 0; c < copies; c++)
                    {
                        UpdateStatus($"Printing {Path.GetFileName(zplFile)} (Copy {c + 1}/{copies})");

                        if (isNetworkPrinter)
                        {
                            // Send to network printer
                            string ipAddress = txtIPAddress.Text;
                            int port = (int)nudPort.Value;

                            if (!await SendToNetworkPrinter(ipAddress, port, zplContent))
                            {
                                success = false;
                                break;
                            }
                        }
                        else
                        {
                            // Send to local printer
                            string printerName = cmbPrinter.SelectedItem.ToString();
                            if (!SendToLocalPrinter(printerName, zplContent))
                            {
                                success = false;
                                break;
                            }
                        }

                        progressBar.Value++;
                        await Task.Delay(100); // Small delay to avoid overwhelming the printer
                    }

                    if (!success)
                    {
                        break;
                    }
                }

                if (success)
                {
                    UpdateStatus("Printing completed successfully");
                    MessageBox.Show("All ZPL files have been printed successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Error during printing");
                MessageBox.Show($"An error occurred while printing: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Re-enable controls
                SetControlsEnabled(true);
                progressBar.Visible = false;
            }
        }

        private void SetControlsEnabled(bool enabled)
        {
            gbPrinterSettings.Enabled = enabled;
            gbFiles.Enabled = enabled;
            gbPrintSettings.Enabled = enabled;
        }

        private void UpdateStatus(string message)
        {
            lblStatus.Text = message;
            Application.DoEvents();
        }

        private async Task<bool> SendToNetworkPrinter(string ipAddress, int port, string zplContent)
        {
            try
            {
                using (TcpClient client = new TcpClient())
                {
                    await client.ConnectAsync(ipAddress, port);

                    using (NetworkStream stream = client.GetStream())
                    {
                        // Convert ZPL content to bytes
                        byte[] data = Encoding.ASCII.GetBytes(zplContent);

                        // Send data to printer
                        await stream.WriteAsync(data, 0, data.Length);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                UpdateStatus("Error: " + ex.Message);
                MessageBox.Show($"Failed to send to network printer: {ex.Message}", "Network Printer Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool SendToLocalPrinter(string printerName, string zplContent)
        {
            try
            {
                // Create a raw print job
                using (PrintDocument pd = new PrintDocument())
                {
                    pd.PrinterSettings.PrinterName = printerName;

                    // Check if printer exists
                    if (!pd.PrinterSettings.IsValid)
                    {
                        throw new Exception($"Printer '{printerName}' is not valid.");
                    }

                    // Handle PrintPage event
                    pd.PrintPage += (sender, e) =>
                    {
                        // Convert ZPL content to bytes and send raw data to printer
                        byte[] data = Encoding.ASCII.GetBytes(zplContent);
                        e.Graphics.DrawString(zplContent, new Font("Courier New", 10), Brushes.Black, 0, 0);
                        e.HasMorePages = false;
                    };

                    pd.Print();
                }

                return true;
            }
            catch (Exception ex)
            {
                UpdateStatus("Error: " + ex.Message);
                MessageBox.Show($"Failed to send to local printer: {ex.Message}", "Local Printer Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
    }
}

