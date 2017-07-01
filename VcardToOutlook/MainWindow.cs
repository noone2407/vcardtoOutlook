using System;
using System.IO;
using System.Windows.Forms;
using OutlookApi;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
namespace VcardToOutlook
{
    public partial class MainWindow : Form
    {

        Outlook.Application App;
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
            var ofdSource = new OpenFileDialog { Filter = "Vcf file|*.vcf" };
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

        private void buttonClearFolder_Click(object sender, EventArgs e)
        {
            DirectoryInfo di = new DirectoryInfo(ContactFolder);
            foreach (FileInfo file in di.GetFiles())
            {
                try
                {
                    file.Delete();
                }
                catch (Exception) { }
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                try
                {
                    dir.Delete(true);
                }
                catch (Exception) { }
            }
        }

        private void buttonCut_Click(object sender, EventArgs e)
        {
            string inputFile = textBoxInput.Text;
            string OutputFolder = textBoxOutput.Text;

            if (!File.Exists(inputFile)) return;
            if (!Directory.Exists(OutputFolder)) return;


            string textData = string.Empty;
            string name = string.Empty;
            string filename = string.Empty;
            bool flabegin = false;
            bool flagend = false;
            int counter = 0;

            string[] allLines = File.ReadAllLines(inputFile);
            progressBar.Visible = true;
            progressBar.Value = 0;
            progressBar.Maximum = allLines.Length;
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
                        int noNameCount = Directory.GetFiles(OutputFolder, "Noname_*.vcf", SearchOption.TopDirectoryOnly).Length;
                        name = "Noname_" + noNameCount.ToString();
                    }
                    else
                    {
                        int fileCount = Directory.GetFiles(OutputFolder, name + ".vcf", SearchOption.TopDirectoryOnly).Length;
                        int filewithnumberCount = Directory.GetFiles(OutputFolder, name + "_*.vcf", SearchOption.TopDirectoryOnly).Length;
                        int total = fileCount + filewithnumberCount;
                        if (total > 0)
                            name = name + "_" + total.ToString();
                    }
                    filename = name + ".vcf";
                    string filePath = Path.Combine(OutputFolder, filename);
                    if (File.Exists(filePath))
                    {
                        MessageBox.Show(filePath);
                    }
                    File.WriteAllText(filePath, textData);
                    flabegin = false;
                    flagend = false;
                    textData = string.Empty;
                    name = string.Empty;
                    filename = string.Empty;
                    counter++;
                }
            }
            MessageBox.Show("Your VCard was split into " + counter.ToString() + " files.", "Success", MessageBoxButtons.OK);
            progressBar.Visible = false;
        }

        private void buttonImport_Click(object sender, EventArgs e)
        {
            string OutputFolder = textBoxOutput.Text;
            App = new Outlook.Application();
            string folderPath = App.Session.DefaultStore.GetRootFolder().FolderPath + @"\Contacts\Key Contacts";
            Outlook.Folder folder = GetFolder(folderPath);
            ImportContacts(OutputFolder, folder);
        }

        private void buttonAbout_Click(object sender, EventArgs e)
        {
            var about = new About();
            about.ShowDialog();
        }

        private void ImportContacts(string path, Outlook.Folder targetFolder)
        {
            Outlook.ContactItem contact;
            progressBar.Visible = true;
            progressBar.Value = 0;
            if (Directory.Exists(path))
            {
                string[] files = Directory.GetFiles(path, "*.vcf");
                progressBar.Maximum = files.Length;
                int counter = 0;
                foreach (string file in files)
                {
                    contact = App.Session.OpenSharedItem(file) as Outlook.ContactItem;
                    contact.Save();
                    counter++;
                }
                MessageBox.Show(string.Format("Imported {0}contact(s) to outlook.", counter.ToString()), "Success", MessageBoxButtons.OK);
                progressBar.Visible = false;
            }
        }
        private Outlook.Folder GetFolder(string folderPath)
        {
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                String[] folders = folderPath.Split(backslash.ToCharArray());
                folder = App.Session.Folders[folders[0]] as Outlook.Folder;
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        var subFolders = folder.Folders;
                        folder = subFolders[folders[i]] as Outlook.Folder;
                        if (folder == null)
                        {
                            return null;
                        }
                    }
                }
                return folder;
            }
            catch { return null; }
        }
        private static string CleanFileName(string fileName)
        {
            string newFileName = fileName;
            var invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                newFileName = newFileName.Replace(c.ToString(), "");
            }
            return newFileName;
        }



    }
}
