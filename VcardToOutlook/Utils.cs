using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.ComponentModel;
using System.Text.RegularExpressions;
using VcardToOutlook.Properties;
using Outlook = NetOffice.OutlookApi;
using System.Diagnostics;
using NetOffice.OutlookApi.Enums;

namespace VcardToOutlook
{
    static class Utils
    {
        #region const
        private static readonly string[] sign = new string[] { "á", "à", "ả", "ã", "ạ", "â", "ấ", "ầ", "ẩ", "ẫ", "ậ", "ă", "ắ", "ằ", "ẳ", "ẵ", "ặ",
            "đ",
            "é","è","ẻ","ẽ","ẹ","ê","ế","ề","ể","ễ","ệ",
            "í","ì","ỉ","ĩ","ị",
            "ó","ò","ỏ","õ","ọ","ô","ố","ồ","ổ","ỗ","ộ","ơ","ớ","ờ","ở","ỡ","ợ",
            "ú","ù","ủ","ũ","ụ","ư","ứ","ừ","ử","ữ","ự",
            "ý","ỳ","ỷ","ỹ","ỵ"};
        private static readonly string[] nosign = new string[] { "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a",
            "d",
            "e","e","e","e","e","e","e","e","e","e","e",
            "i","i","i","i","i",
            "o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o",
            "u","u","u","u","u","u","u","u","u","u","u",
            "y","y","y","y","y"};
        #endregion


        public static int CutVcf(BackgroundWorker backgroundWorker, string[] inputLines, string outputFolder, bool removeVietnameseSign)
        {
            if (inputLines.Length == 0)
            {
                return 0;
            }
            string textData = string.Empty;
            string name = string.Empty; // for file name
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
                        if (removeVietnameseSign)
                        {
                            n = RemoveVietnameseSign(n);
                            line = "N;" + n;
                        }
                        else
                        {
                            line = "N;CHARSET=utf-8:" + n;
                        }
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            name = CleanInvalidFileNameChars(n);
                        }
                    }
                }
                if (line.StartsWith("N:")) // name
                {
                    name = line.Substring(2).Replace(";", "");
                    name = CleanInvalidFileNameChars(name);
                    if (removeVietnameseSign)
                    {
                        name = RemoveVietnameseSign(name);
                        line = "N;" + name;
                    }
                    else
                    {
                        line = "N;CHARSET=utf-8:" + name;
                    }
                }
                if (line.StartsWith("FN;")) // full name
                {
                    if (line.Contains("ENCODING="))
                    {
                        int pos = line.IndexOf(quotedPrintable) + quotedPrintable.Length;
                        string text = line.Substring(pos).Replace(";", "");
                        string fn = DecodeQuotedPrintables(text, "UTF-8");
                        if (removeVietnameseSign)
                        {
                            fn = RemoveVietnameseSign(fn);
                            line = "FN;" + fn;
                        }
                        else
                        {
                            line = "FN;CHARSET=utf-8:" + fn;
                        }
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            name = CleanInvalidFileNameChars(fn);
                        }
                    }
                }
                if (line.StartsWith("FN:")) //full name
                {
                    name = line.Substring(3);
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        if (removeVietnameseSign)
                        {
                            name = RemoveVietnameseSign(name);
                            line = "FN;" + name;
                        }
                        else
                        {
                            line = "FN;CHARSET=utf-8:" + name;
                        }
                    }

                }
                if (line.StartsWith("ORG;")) // company
                {
                    if (line.Contains("ENCODING="))
                    {
                        int pos = line.IndexOf(quotedPrintable) + quotedPrintable.Length;
                        string text = line.Substring(pos).Replace(";", "");
                        string org = DecodeQuotedPrintables(text, "UTF-8");
                        if (removeVietnameseSign)
                        {
                            org = RemoveVietnameseSign(org);
                            line = "ORG;" + org;
                        }
                        else
                        {
                            line = "ORG;CHARSET=utf-8:" + org;
                        }
                    }
                }
                if (line.StartsWith("ORG:")) //company
                {
                    line = "ORG;CHARSET=utf-8:" + line.Substring(4);
                }
                if (flabegin) // has begin:vcard flag
                {
                    if (!string.IsNullOrEmpty(textData)) // add new line if there are data
                    {
                        textData = textData + Environment.NewLine + line;
                    }
                    else
                    {
                        textData = line;
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

        public static int ImportContacts(BackgroundWorker backgroundWorker, Outlook.Application outlookApplication, string path)
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

        public static void ClearOldContact(BackgroundWorker backgroundWorker, Outlook.Application outlookApplication)
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

        public static void ClearOldVcfFiles(BackgroundWorker backgroundWorker, string folder)
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

        public static string[] CleanInputFile(string inputFile)
        {
            string text = File.ReadAllText(inputFile);
            text = text.Replace("=\r\n=", "="); // cut break line of quoted-printable
            text = text.Replace("\r\n", "\n"); // change crlf to lf
            return text.Split('\n');
        }

        private static string CleanInvalidFileNameChars(string fileName)
        {
            string newFileName = fileName;
            var invalidChars = Path.GetInvalidFileNameChars();
            return invalidChars.Aggregate(newFileName, (current, c) => current.Replace(c.ToString(), ""));
        }

        private static string DecodeQuotedPrintables(string input, string charSet)
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
        private static byte[] StringToByteArray(string hex)
        {
            return Enumerable.Range(0, hex.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
                             .ToArray();
        }
        private static string RemoveVietnameseSign(string text)
        {
            for (int i = 0; i < sign.Length; i++)
            {
                text = text.Replace(sign[i], nosign[i]);
                text = text.Replace(sign[i].ToUpper(), nosign[i].ToUpper());
            }
            return text;
        }
    }
}
