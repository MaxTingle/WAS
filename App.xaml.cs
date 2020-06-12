using System;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;

namespace WorkingAttenuationSoftware
{
    public partial class App : System.Windows.Application
    {
        private SettingWindow settingWindow = null;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            #region Initialize settings tray icon
            NotifyIcon trayIcon = new NotifyIcon();
            trayIcon.Text = "WAS";
            trayIcon.Visible = true;
            trayIcon.Icon = WorkingAttenuationSoftware.Properties.Resources.TrayIcon;
            trayIcon.Click += this._TrayIconClicked;
            #endregion

            #region Initialize service
            //TODO: Create & start service/loop/whatever to listen for mic/app sound
            #endregion
        }

        private void _TrayIconClicked(Object sender, EventArgs e)
        {
            if(this.settingWindow == null)
            {
                this.settingWindow = new SettingWindow();
            }

            this.settingWindow.Show();
        }
    }
}