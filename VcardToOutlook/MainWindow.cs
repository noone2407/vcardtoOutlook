using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using VcardToOutlook.Properties;
using Outlook = NetOffice.OutlookApi;

namespace VcardToOutlook
{
    public partial class MainWindow : Form
    {
        string ContactFolder = "C:\\Contacts\\";
        public MainWindow()
        {
            InitializeComponent();
        }
        private void MainWindow_Load(object sender, EventArgs e)
        {
            textBoxOutput.Text = ContactFolder;
            label6.Text = "";
            if (!Directory.Exists(ContactFolder))
                Directory.CreateDirectory(ContactFolder);
        }

        private void buttonSelectSource_Click(object sender, EventArgs e)
        {
            var ofdSource = new OpenFileDialog { Filter = Resources.VcfFilter };
            if (ofdSource.ShowDialog() == DialogResult.OK)
            {
                textBoxInput.Text = ofdSource.FileName;
            }
        }

        private void buttonSelectTarget_Click(object sender, EventArgs e)
        {
            var fbdTarget = new FolderBrowserDialog();
            if (fbdTarget.ShowDialog() == DialogResult.OK)
            {
                textBoxOutput.Text = fbdTarget.SelectedPath;
            }
        }
        private void buttonCut_Click(object sender, EventArgs e)
        {
            string inputFile = textBoxInput.Text;
            string outputFolder = textBoxOutput.Text;
            if (!File.Exists(inputFile)) return;
            if (!Directory.Exists(outputFolder)) return;
            bool clearOldVcfFiles = checkBoxClearOldVcf.Checked;
            bool removeVietnameseSign = checkBoxRemoveVietnameseSign.Checked;

            ResetProgressbar();
            var backgroundWorker = new BackgroundWorker()
            {
                WorkerReportsProgress = true
            };
            backgroundWorker.DoWork += (o, args) =>
            {
                string[] lines = Utils.CleanInputFile(inputFile);
                if (clearOldVcfFiles)
                    Utils.ClearOldVcfFiles(backgroundWorker, outputFolder);
                int counter = Utils.CutVcf(backgroundWorker, lines, outputFolder, removeVietnameseSign);
                args.Result = counter;
            };
            backgroundWorker.ProgressChanged += (o, args) =>
            {
                progressBar.Value = args.ProgressPercentage;
            };
            backgroundWorker.RunWorkerCompleted += (o, args) =>
            {
                progressBar.Value = 100;
                MessageBox.Show("Your VCard was split into " + args.Result + " files.", "Success", MessageBoxButtons.OK);
                progressBar.Visible = false;
            };
            backgroundWorker.RunWorkerAsync();
        }

        private void buttonImport_Click(object sender, EventArgs e)
        {
            string outputFolder = textBoxOutput.Text;
            bool clearOldContact = checkBoxClearOldContact.Checked;
            ResetProgressbar();
            var backgroundWorker = new BackgroundWorker()
            {
                WorkerReportsProgress = true
            };
            backgroundWorker.DoWork += (o, args) =>
            {
                var outlookApplication = new Outlook.Application();
                if (clearOldContact)
                    Utils.ClearOldContact(backgroundWorker, outlookApplication);
                int counter = Utils.ImportContacts(backgroundWorker, outlookApplication, outputFolder);
                args.Result = counter;
            };
            backgroundWorker.ProgressChanged += (o, args) =>
            {
                progressBar.Value = args.ProgressPercentage;
            };
            backgroundWorker.RunWorkerCompleted += (o, args) =>
            {
                progressBar.Value = 100;
                MessageBox.Show(string.Format("Imported {0}contact(s) to outlook.", args.Result), "Success", MessageBoxButtons.OK);
                progressBar.Visible = false;
            };
            backgroundWorker.RunWorkerAsync();
        }
        private void buttonAbout_Click(object sender, EventArgs e)
        {
            var about = new About();
            about.ShowDialog();
        }

        private void linkLabelWebsite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://bbhcm.vn");
        }
    
        private void ResetProgressbar()
        {
            progressBar.Visible = true;
            progressBar.Value = 0;
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
        }
    }
}
