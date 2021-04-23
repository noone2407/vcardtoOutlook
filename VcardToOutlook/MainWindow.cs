using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using NetOffice.OutlookApi.Enums;
using VcardToOutlook.Properties;
using Outlook = NetOffice.OutlookApi;

namespace VcardToOutlook
{
    public partial class MainWindow : Form
    {
        #region const
        string ContactFolder = "C:\\Contacts\\";
        readonly string[] sign = new string[] { "á", "à", "ả", "ã", "ạ", "â", "ấ", "ầ", "ẩ", "ẫ", "ậ", "ă", "ắ", "ằ", "ẳ", "ẵ", "ặ",
            "đ",
            "é","è","ẻ","ẽ","ẹ","ê","ế","ề","ể","ễ","ệ",
            "í","ì","ỉ","ĩ","ị",
            "ó","ò","ỏ","õ","ọ","ô","ố","ồ","ổ","ỗ","ộ","ơ","ớ","ờ","ở","ỡ","ợ",
            "ú","ù","ủ","ũ","ụ","ư","ứ","ừ","ử","ữ","ự",
            "ý","ỳ","ỷ","ỹ","ỵ"};
        readonly string[] nosign = new string[] { "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a",
            "d",
            "e","e","e","e","e","e","e","e","e","e","e",
            "i","i","i","i","i",
            "o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o",
            "u","u","u","u","u","u","u","u","u","u","u",
            "y","y","y","y","y"};
        #endregion

        public MainWindow()
        {
            InitializeComponent();

        }

        #region UI Events
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

            ResetProgressbar();
            var backgroundWorker = new BackgroundWorker()
            {
                WorkerReportsProgress = true
            };
            backgroundWorker.DoWork += (o, args) =>
            {
                string[] lines = CleanInputFile(inputFile);
                if (clearOldVcfFiles)
                    ClearOldVcfFiles(backgroundWorker, outputFolder);
                int counter = CutVcf(backgroundWorker, lines, outputFolder);
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
                    ClearOldContact(backgroundWorker, outlookApplication);
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

        private string[] CleanInputFile(string inputFile)
        {
            string text = File.ReadAllText(inputFile);
            text = text.Replace("=\r\n=", "="); // cut break line of quoted-printable
            text = text.Replace("\r\n", "\n"); // change crlf to lf
            return text.Split('\n');
        }
        private void linkLabelWebsite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://bbhcm.vn");
        }
        #endregion

        #region Private code

        private string NonUnicode(string text)
        {
            
            for (int i = 0; i < sign.Length; i++)
            {
                text = text.Replace(sign[i], nosign[i]);
                text = text.Replace(sign[i].ToUpper(), nosign[i].ToUpper());
            }
            return text;
        }
        private void ClearOldVcfFiles(BackgroundWorker backgroundWorker, string folder)
        {
            DirectoryInfo di = new DirectoryInfo(folder);
            var files = di.GetFiles();
            for (int i = 0; i < files.Length; i++)
            {
                try
                {
                    files[i].Delete();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.ToString());
                }
                backgroundWorker.ReportProgress(i * 100 / files.Length);
            }
        }

        private int CutVcf(BackgroundWorker backgroundWorker, string[] inputLines, string outputFolder)
        {
            if (inputLines.Length == 0)
            {
                return 0;
            }
            string textData = string.Empty;
            string name = string.Empty;
            bool flabegin = false;
            bool flagend = false;
            int counter = 0;
            string quotedPrintable = "ENCODING=QUOTED-PRINTABLE:";

            for (int i = 0; i < inputLines.Length; i++)
            {
                string line = inputLines[i];
                if (line == "BEGIN:VCARD")
                {
                    flabegin = true;
                }
                if (line == "END:VCARD")
                {
                    flagend = true;
                }
                if (line.StartsWith("N;")) // name
                {
                    if (line.Contains("ENCODING="))
                    {
                        int pos = line.IndexOf(quotedPrintable) + quotedPrintable.Length;
                        string text = line.Substring(pos).Replace(";", "");
                        string n = DecodeQuotedPrintables(text, "UTF-8");
                        n = NonUnicode(n);
                        line = "N;CHARSET=utf-8:" + n;
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            name = CleanFileName(n);
                        }
                    }
                }
                if (line.StartsWith("N:")) // name
                {
                    name = line.Substring(2).Replace(";", "");
                    line = "N;CHARSET=utf-8:" + line.Substring(2);
                    name = CleanFileName(name);
                }
                if (line.StartsWith("FN;")) // full name
                {
                    if (line.Contains("ENCODING="))
                    {
                        int pos = line.IndexOf(quotedPrintable) + quotedPrintable.Length;
                        string text = line.Substring(pos).Replace(";", "");
                        string fn = DecodeQuotedPrintables(text, "UTF-8");
                        fn = NonUnicode(fn);
                        line = "FN;CHARSET=utf-8:" + fn;
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            name = CleanFileName(fn);
                        }
                    }
                }
                if (line.StartsWith("FN:")) //full name
                {
                    name = line.Substring(3);
                    line = "FN;CHARSET=utf-8:" + line.Substring(3);
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        name = CleanFileName(name);
                    }
                }
                if (line.StartsWith("ORG;")) // company
                {
                    if (line.Contains("ENCODING="))
                    {
                        int pos = line.IndexOf(quotedPrintable) + quotedPrintable.Length;
                        string text = line.Substring(pos).Replace(";", "");
                        string org = DecodeQuotedPrintables(text, "UTF-8");
                        org = NonUnicode(org);
                        line = "ORG;CHARSET=utf-8:" + org;
                    }
                }
                if (line.StartsWith("ORG:")) //company
                {
                    line = "ORG;CHARSET=utf-8:" + line.Substring(4);
                }
                if (flabegin) // begin:vcard 
                {
                    if (string.IsNullOrEmpty(textData))
                    {
                        textData = line;
                    }
                    else
                    {
                        textData = textData + Environment.NewLine + line;
                    }

                }
                if (flagend)  // end:vcard 
                {
                    if (string.IsNullOrWhiteSpace(name)) // emtpy name
                    {
                        int noNameCount = Directory.GetFiles(outputFolder, "Noname_*.vcf", SearchOption.TopDirectoryOnly).Length;
                        name = "Noname_" + noNameCount.ToString();
                    }
                    else  // search for duplicated name
                    {
                        int fileCount = Directory.GetFiles(outputFolder, name + ".vcf", SearchOption.TopDirectoryOnly).Length;
                        int filewithnumberCount = Directory.GetFiles(outputFolder, name + "_*.vcf", SearchOption.TopDirectoryOnly).Length;
                        int total = fileCount + filewithnumberCount;
                        if (total > 0) // add number to duplicated name
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
                backgroundWorker.ReportProgress(i * 100 / inputLines.Length);
            }
            return counter;
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

        private string CleanFileName(string fileName)
        {
            string newFileName = fileName;
            var invalidChars = Path.GetInvalidFileNameChars();
            return invalidChars.Aggregate(newFileName, (current, c) => current.Replace(c.ToString(), ""));
        }

        private  string DecodeQuotedPrintables(string input, string charSet)
        {
            if (string.IsNullOrEmpty(charSet))
            {
                var charSetOccurences = new Regex(@"=\?.*\?Q\?", RegexOptions.IgnoreCase);
                var charSetMatches = charSetOccurences.Matches(input);
                foreach (Match match in charSetMatches)
                {
                    charSet = match.Groups[0].Value.Replace("=?", "").Replace("?Q?", "");
                    input = input.Replace(match.Groups[0].Value, "").Replace("?=", "");
                }
            }

            Encoding enc = new ASCIIEncoding();
            if (!string.IsNullOrEmpty(charSet))
            {
                try
                {
                    enc = Encoding.GetEncoding(charSet);
                }
                catch
                {
                    enc = new ASCIIEncoding();
                }
            }
            var arr = new List<byte>();
            foreach (string s in input.Split('='))
            {
                arr.AddRange(StringToByteArray(s));
            }
            string output = enc.GetString(arr.ToArray());
            return output;
        }
        private byte[] StringToByteArray(string hex)
        {
            return Enumerable.Range(0, hex.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
                             .ToArray();
        }
        #endregion
    }
}
