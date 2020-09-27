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
            //TODO: Add ui:
            //TODO: Add mic selection
            //TODO: Add whitelist program selection
            //TODO: Add voip program selection
            //TODO: Add minimum noise length (seconds) input
            //TODO: Add minimum mic volume input
            //TODO: Add restore volume after input
            trayIcon.Click += this._TrayIconClicked;
            #endregion

            #region Initialize service
            AudioReducer reducer = new AudioReducer(0.4, 0.3, 8, 3);
            reducer.AddVOIPApp("discord.exe");
            reducer.ListenForIncomingVOIP();
            reducer.ListenForMic();
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