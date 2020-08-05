using System;
using System.Windows;
using System.Windows.Forms;
using WorkingAttenuationSoftware.Controls;

namespace WorkingAttenuationSoftware
{
    public partial class App : System.Windows.Application
    {
        private SettingWindow _SettingWindow = null;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            #region Initialize settings tray icon
            NotifyIcon trayIcon = new NotifyIcon
            {
                Text = "WAS",
                Visible = true,
                Icon = WorkingAttenuationSoftware.Properties.Resources.TrayIcon
            };
            trayIcon.Click += this._TrayIconClicked;
            #endregion

            #region Initialize service
            AudioReducer reducer = new AudioReducer(10.0f);
            reducer.StoreCurrentVolumeLevels();
            //TODO: Create & start service/loop/whatever to listen for mic/app sound
            #endregion
        }

        private void _TrayIconClicked(Object sender, EventArgs e)
        {
            if(this._SettingWindow == null)
            {
                this._SettingWindow = new SettingWindow();
                this._SettingWindow.Closed += (closeSender, closeE) =>
                {
                    this._SettingWindow = null;
                };
            }

            this._SettingWindow.Show();
        }
    }
}