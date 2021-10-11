using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace VcardToOutlook.AutoUpdate
{

    internal class AutoUpdateHelper
    {
        const string xmlUpdateUrl = "https://docs.google.com/document/d/14j2KqWDLu3ePJWRGV37-ApZvxfStET7gjGHUZJOq4hw/export?format=txt";
        GoogleDownloader fileDownloader;
        internal AutoUpdateHelper()
        {
            fileDownloader = new GoogleDownloader();
            // This callback is triggered for DownloadFileAsync only
            fileDownloader.DownloadProgressChanged += FileDownloader_DownloadProgressChanged;
            // This callback is triggered for both DownloadFile and DownloadFileAsync
            fileDownloader.DownloadFileCompleted += FileDownloader_DownloadFileCompleted;
        }

        internal CheckUpdateResult CheckUpdate()
        {
            var result = new CheckUpdateResult();
            string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Debug.WriteLine(path);
            fileDownloader.DownloadFile(xmlUpdateUrl, path);
            if (File.Exists(path))
            {
                string content = File.ReadAllText(path);
                if (!string.IsNullOrEmpty(content))
                {
                    ReadAutoUpdateXml(content, ref result);
                }
                File.Delete(path);
            }
            else
            {
                result.Success = false;
            }
            return result;
        }

        internal DownloadUpdateResult DownloadUpdate(string url)
        {
            var result = new DownloadUpdateResult();
            string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Debug.WriteLine(url);
            Debug.WriteLine(path);
            fileDownloader.DownloadFile(url, path);
            if (File.Exists(path))
            {
                result.Success = true;
                result.DownloadPath = path;
            }
            else
            {
                result.Success = false;
            }
            return result;
        }

        internal string UnzipUpdate(string zipPath)
        {
            string extractPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            if (Directory.Exists(extractPath))
                Directory.Delete(extractPath);
            ZipFile.ExtractToDirectory(zipPath, extractPath);
            File.Delete(zipPath);
            return extractPath;
        }

        internal string CreateUpdateScript(string extractedPath, string destinationPath, string exeName, int delayTime)
        {
            string scriptPath = Path.Combine(Path.GetTempPath(), $"autoupdate_{Path.GetRandomFileName()}.bat");
            string content = $"@echo off\r\n"; // echo off
            content += $"timeout {delayTime} > NUL\r\n"; // delay before action
            content += $"del /q {destinationPath}\\*\r\n"; // clear all files in app folder
            content += $"xcopy {extractedPath} {destinationPath} /c /q\r\n"; // copy all file in extracted folder to app folder
            content += $"del /q {extractedPath}\r\n"; // clean all contents of extracted folder
            content += $"rmdir {extractedPath}\r\n"; // delete extracted folder
            content += $"start {Path.Combine(destinationPath, exeName)} {scriptPath}\r\n"; // run new app
            content += $"(goto) 2>nul & del \"%~f0\""; // selft delete update script
            File.WriteAllText(scriptPath, content);
            return scriptPath;
        }

        private void FileDownloader_DownloadProgressChanged(object sender, GoogleDownloader.DownloadProgress progress)
        {
            Debug.WriteLine("Progress changed " + progress.BytesReceived + " " + progress.TotalBytesToReceive);
        }

        private void FileDownloader_DownloadFileCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            Console.WriteLine("Download completed");
        }

        private void ReadAutoUpdateXml(string content, ref CheckUpdateResult result)
        {
            try
            {
                var doc = new XmlDocument();
                doc.LoadXml(content);
                foreach (XmlElement elm in doc.SelectSingleNode("/item"))
                {
                    if (elm.Name.Equals("version"))
                        result.Version = Version.Parse(elm.InnerText);
                    if (elm.Name.Equals("url"))
                        result.Url = elm.InnerText;
                    if (elm.Name.Equals("mandatory"))
                        result.Mandatory = bool.Parse(elm.InnerText);
                }
                result.Success = (result.Version.Major > 0 && !string.IsNullOrEmpty(result.Url));
            }
            catch (Exception ex)
            {
                result.Success = false;
                Debug.WriteLine(ex.ToString());
            }

        }
    }
}
