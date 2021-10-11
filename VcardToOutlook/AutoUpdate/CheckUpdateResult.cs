using System;

namespace VcardToOutlook.AutoUpdate
{
    internal class CheckUpdateResult
    {
        internal bool Success { get; set; }
        internal Version Version { get; set; }
        internal string Url { get; set; }
        internal bool Mandatory { get; set; }
    }
}
