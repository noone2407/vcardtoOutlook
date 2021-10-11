using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using VcardToOutlook.Properties;
using Outlook = NetOffice.OutlookApi;

namespace VcardToOutlook
{
    public partial class MainWindow : Form
    {
        string ContactFolder = "C:\\Contacts\\";
        Version appVersion = new Version(Application.ProductVersion);
        public MainWindow()
        {
            InitializeComponent();
        }
        private void MainWindow_Load(object sender, EventArgs e)
        {
            textBoxOutput.Text = ContactFolder;
            label6.Text = "";
            labelTitle.Text = $"VCardToOutlook {appVersion.Major}.{appVersion.Minor}";
            this.BringToFront();
            if (!Directory.Exists(ContactFolder))
                Directory.CreateDirectory(ContactFolder);
            RunCheckForUpdateBackground();
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

        private void RunCheckForUpdateBackground()
        {
            (new Thread(new ThreadStart(this.CheckForUpdate))
            {
                IsBackground = true,
                Priority = ThreadPriority.Normal
            }).Start();
        }
        private void CheckForUpdate()
        {
            var autoUpdate = new AutoUpdateHelper();
            var checkUpdateresult = autoUpdate.CheckUpdate();
            if (!checkUpdateresult.Success)
                return;
            if (!(checkUpdateresult.Version > appVersion))
                return;
            if (!checkUpdateresult.Mandatory)
            {
                if (MessageBox.Show($"Do you want to update to version {checkUpdateresult.Version.ToString()}?", "There is a new version!", MessageBoxButtons.YesNo) == DialogResult.No)
                    return;
            }
            var downloadUpdateResult = autoUpdate.DownloadUpdate(checkUpdateresult.Url);
            if (!downloadUpdateResult.Success)
            {
                MessageBox.Show("Couldn't download the update", "Error", MessageBoxButtons.OK);
                return;
            }
            if (!checkUpdateresult.Mandatory)
            {
                if (MessageBox.Show("Do you want to restart now?", "Restart", MessageBoxButtons.YesNo) == DialogResult.No)
                    return;
            }
            string extractPath = autoUpdate.UnzipUpdate(downloadUpdateResult.DownloadPath);
            if (string.IsNullOrEmpty(extractPath))
                return;
            string scriptPath = autoUpdate.CreateUpdateScript(extractPath, Application.StartupPath, Application.ExecutablePath, 1);
            if (!File.Exists(scriptPath))
                return;
            Process.Start(scriptPath);
            Application.Exit();
        }
    }
}
