using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace VcardToOutlook.AutoUpdate
{
    public class AutoUpdateScript
    {
        public static void CheckForUpdate()
        {
            var autoUpdate = new AutoUpdateHelper();
            var checkUpdateresult = autoUpdate.CheckUpdate();
            if (!checkUpdateresult.Success)
                return;
            if (!(checkUpdateresult.Version > new Version(Application.ProductVersion)))
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
