using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;
using System.Linq;

using RaySo.Properties;
using RAD.ClipMon.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RaySo
{
    public partial class Form1 : Form
    {
        #region VARIABLES

        private NotifyIcon applicationIcon;

        private string lastCodeSnippet;
        private string lastCodeSnippetUrl;

        private string[] applications = { "Code", "devenv" };

        IntPtr _ClipboardViewerNext;

        #endregion

        #region FORM

        public Form1()
        {
            applicationIcon = new NotifyIcon();
            applicationIcon.Text = "Ray.so for Windows";
            applicationIcon.Icon = new Icon(Resources.AppIcon, 128, 128);

            applicationIcon.ContextMenuStrip = new ContextMenuStrip();
            applicationIcon.ContextMenuStrip.Items.Add("Open the last code", null, OpenCodeSnippetUrl);
            applicationIcon.ContextMenuStrip.Items.Add("-");
            applicationIcon.ContextMenuStrip.Items.Add("Exit", null, OnExit);

            applicationIcon.BalloonTipClicked += (s, e) => OpenCodeSnippetUrl(s, e);
            applicationIcon.Visible = true;
        }

        #endregion

        #region WORKING WITH CODE

        private string Base64Encode(string plainText)
        {
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            return Convert.ToBase64String(plainTextBytes);
        }

        private string RemoveExcessTabs(string code)
        {
            List<string> lines = new List<string>(code.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None));
            List<string> newLines = new List<string>();

            string firstLine = lines[0];
            Match match = Regex.Match(firstLine, @"^\s+");

            if (match.Success)
            {
                int excessSpaceLength = match.Value.Length;

                foreach (string line in lines)
                {
                    string newLine = Regex.Replace(line, @"^\s{" + excessSpaceLength + @"}", "");
                    newLines.Add(newLine);
                }

                return String.Join("\n", newLines.ToArray());
            }
            else
            {
                return code;
            }
        }

        private string GenerateCodeSnippetUrl(string theme, string code, string language)
        {
            return $"https://ray.so/?title=&theme={ theme }&spacing=32&background=true&darkMode=true&code={ code }&language={ language }";
        }

        private void OpenCodeSnippetUrl(object s, EventArgs e)
        {
            if (lastCodeSnippetUrl == null) return;

            Process.Start(lastCodeSnippetUrl);
        }

        #endregion

        #region WORKING WITH CLIPBOARD

        public string GetActiveWindowName()
        {
            IntPtr activatedHandle = GetForegroundWindow();

            Process[] processes = Process.GetProcesses();

            foreach (Process process in processes)
            {
                if (activatedHandle == process.MainWindowHandle)
                {
                    return process.ProcessName;
                }
            }

            return null;
        }

        private bool ClipboardSearch(IDataObject iData)
        {
            if (!iData.GetDataPresent(DataFormats.Text))
                return false;

            string textFromClipboard = (string)iData.GetData(DataFormats.Text);

            Regex codeParts = new Regex("[{}><=!'`\"():;]", RegexOptions.Multiline);
            Match codePartsMatch = codeParts.Match(textFromClipboard);

            if (codePartsMatch.Success)
            {
                if (lastCodeSnippet != textFromClipboard)
                {
                    lastCodeSnippet = textFromClipboard;

                    string code = RemoveExcessTabs(textFromClipboard);
                    string codeInBase64 = Base64Encode(code);
                    string codeUriSafe = HttpUtility.UrlEncode(codeInBase64);

                    lastCodeSnippetUrl = GenerateCodeSnippetUrl("breeze", codeUriSafe, "auto");

                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void GetClipboardData()
        {
            IDataObject iData = Clipboard.GetDataObject();

            string currentWindowName = GetActiveWindowName();

            if (ClipboardSearch(iData) && applications.Contains(currentWindowName))
                applicationIcon.ShowBalloonTip(5000, "Clipboard", "Click here to get an image of code", ToolTipIcon.Info);
        }

        #endregion

        #region OTHER

        private void OnExit(object sender, EventArgs e)
        {
            applicationIcon.Dispose();
            Application.Exit();
        }

        protected override void WndProc(ref Message m)
        {
            switch ((Msgs)m.Msg)
            {
                case Msgs.WM_DRAWCLIPBOARD:
                    GetClipboardData();
                    User32.SendMessage(_ClipboardViewerNext, m.Msg, m.WParam, m.LParam);
                    break;
                case Msgs.WM_CHANGECBCHAIN:
                    if (m.WParam == _ClipboardViewerNext)
                    {
                        _ClipboardViewerNext = m.LParam;
                    }
                    else
                    {
                        User32.SendMessage(_ClipboardViewerNext, m.Msg, m.WParam, m.LParam);
                    }
                    break;
                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false;
            ShowInTaskbar = false;
            _ClipboardViewerNext = User32.SetClipboardViewer(this.Handle);
            base.OnLoad(e);
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr GetForegroundWindow();

        #endregion
    }
}