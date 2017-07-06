using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using NetOffice.OutlookApi.Enums;
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

        private void ClearOldVcfFiles(string folder)
        {
            DirectoryInfo di = new DirectoryInfo(folder);
            foreach (FileInfo file in di.GetFiles())
            {
                try
                {
                    file.Delete();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.ToString());
                }
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                try
                {
                    dir.Delete(true);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.ToString());
                }
            }
        }

        private void buttonCut_Click(object sender, EventArgs e)
        {
            string inputFile = textBoxInput.Text;
            string outputFolder = textBoxOutput.Text;
            if (!File.Exists(inputFile)) return;
            if (!Directory.Exists(outputFolder)) return;
            bool clearOldVcfFiles = checkBoxClearOldVcf.Checked;
            ResetProgressbar();
            var backgroundWorker = new BackgroundWorker()
            {
                WorkerReportsProgress = true
            };
            backgroundWorker.DoWork += (o, args) =>
            {
                if (clearOldVcfFiles)
                    ClearOldVcfFiles(outputFolder);
                int counter = CutVcf(backgroundWorker, inputFile, outputFolder);
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

        private int CutVcf(BackgroundWorker backgroundWorker, string inputFile, string outputFolder)
        {

            string textData = string.Empty;
            string name = string.Empty;
            bool flabegin = false;
            bool flagend = false;
            int counter = 0;

            string[] allLines = File.ReadAllLines(inputFile);

            for (int i = 0; i < allLines.Length; i++)
            {
                string text4 = allLines[i];
                if (text4 == "BEGIN:VCARD")
                {
                    flabegin = true;
                }
                if (text4 == "END:VCARD")
                {
                    flagend = true;
                }
                if (text4.StartsWith("N:")) // name
                {
                    text4 = "N;CHARSET=utf-8:" + text4.Substring(2);
                }
                if (text4.StartsWith("ORG:")) //company
                {
                    text4 = "ORG;CHARSET=utf-8:" + text4.Substring(4);
                }
                if (text4.StartsWith("FN:")) //full name
                {
                    name = text4.Substring(3);
                    name = CleanFileName(name);
                    text4 = "FN;CHARSET=utf-8:" + text4.Substring(3);
                }
                if (flabegin)
                {
                    textData = textData + Environment.NewLine + text4;
                }
                if (flagend)
                {
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        int noNameCount = Directory.GetFiles(outputFolder, "Noname_*.vcf", SearchOption.TopDirectoryOnly).Length;
                        name = "Noname_" + noNameCount.ToString();
                    }
                    else
                    {
                        int fileCount = Directory.GetFiles(outputFolder, name + ".vcf", SearchOption.TopDirectoryOnly).Length;
                        int filewithnumberCount = Directory.GetFiles(outputFolder, name + "_*.vcf", SearchOption.TopDirectoryOnly).Length;
                        int total = fileCount + filewithnumberCount;
                        if (total > 0)
                            name = name + "_" + total.ToString();
                    }
                    string filename = name + ".vcf";
                    string filePath = Path.Combine(outputFolder, filename);
                    File.WriteAllText(filePath, textData);
                    flabegin = false;
                    flagend = false;
                    textData = string.Empty;
                    name = string.Empty;
                    counter++;
                }
                backgroundWorker.ReportProgress(i * 100 / allLines.Length);
            }
            return counter;
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
                    ClearOldContact(backgroundWorker,outlookApplication);
                int counter = ImportContacts(backgroundWorker, outlookApplication, outputFolder);
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
        private void ResetProgressbar()
        {
            progressBar.Visible = true;
            progressBar.Value = 0;
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
        }

        private void ClearOldContact(BackgroundWorker backgroundWorker, Outlook.Application outlookApplication)
        {
            Outlook.MAPIFolder contactFolder = outlookApplication.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            int total = contactFolder.Items.Count;
            int remaining = total;
            int deleted = 0;
            while (remaining > 0)
            {
                var contact = (Outlook.ContactItem)contactFolder.Items[1];
                contact.Delete();
                Thread.Sleep(100);
                deleted++;
                remaining = contactFolder.Items.Count;
                backgroundWorker.ReportProgress(deleted * 100 / total);
            }
        }

        private int ImportContacts(BackgroundWorker backgroundWorker, Outlook.Application outlookApplication, string path)
        {
            int counter = 0;
            if (Directory.Exists(path))
            {
                string[] files = Directory.GetFiles(path, "*.vcf");
                for (int i = 0; i < files.Length; i++)
                {
                    var contact = outlookApplication.Session.OpenSharedItem(files[i]) as Outlook.ContactItem;
                    if (contact != null)
                    {
                        contact.Save();
                        counter++;
                    }
                    backgroundWorker.ReportProgress(i * 100 / files.Length);
                }
            }
            return counter;
        }

        private static string CleanFileName(string fileName)
        {
            string newFileName = fileName;
            var invalidChars = Path.GetInvalidFileNameChars();
            return invalidChars.Aggregate(newFileName, (current, c) => current.Replace(c.ToString(), ""));
        }

        private void linkLabelWebsite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("http://bbhcm.vn");
        }
    }
}
