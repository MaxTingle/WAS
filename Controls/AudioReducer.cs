using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Media;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Timers;

namespace WorkingAttenuationSoftware.Controls
{
    public sealed class AudioReducer
    {
        public struct AppVolume
        {
            public int ProcessID;
            public string Path;
            public string Name;
            public float Volume;
        }

        /*
         * Settings
         */

        private double _MinimumNoiseLength;
        private int _MinimumMicVolume;
        private double _RestoreVolumeAfter;
        private int _MicDeviceNumber = 0;
        private double _Reduction;
        private List<string> _AppWhitelist = new List<string>();
        private List<string> _VoipApps = new List<string>();
        private bool _Debug = true;

        private bool _ReductionApplied;
        private List<AppVolume> _AppVolumes = new List<AppVolume>();
        private WaveInEvent _MicListener;

        private long _LastReductionAttemptMicroTime;
        private long _LastAudioMicroTime;
        private long _FirstAudioMicroTime;

        public AudioReducer(double reduction, double minimumNoiseLength, int minimumMicVolume, double restoreVolumeAfter)
        {
            this.SetReduction(reduction);
            this.SetMinimumNoiseLength(minimumNoiseLength);
            this.SetMinimumMicVolume(minimumMicVolume);
            this.SetRestoreVolumeAfter(restoreVolumeAfter);
        }

        /**
         * Sets the NAudio device number which represents the mic we want to listen to
         */
        public void SetMic(int deviceNumber)
        {
            this._MicDeviceNumber = deviceNumber;
        }

        /**
         * Sets the % of 100 that we should reduce application volumes by
         */
        public void SetReduction(double reduction)
        {
            this._Reduction = reduction;
        }

        /**
         * Sets the amount of seconds to wait for before restoring the original volumes
         */
        public void SetRestoreVolumeAfter(double restoreVolumeAfter)
        {
            this._RestoreVolumeAfter = restoreVolumeAfter;
        }

        /**
         * Sets the minimum volume (As a % of 100) the mic audio can be before we consider it as noise
         */
        public void SetMinimumMicVolume(int minimumMicVolume)
        {
            this._MinimumMicVolume = minimumMicVolume;
        }

        /**
         * Sets the minimum amount of time (In seconds) we need to hear speaker/mic sound for
         * before we actually reduce the volumes of applications
         */
        public void SetMinimumNoiseLength(double minimumNoiseLength)
        {
            this._MinimumNoiseLength = minimumNoiseLength;
        }

        /**
         * Resets the app whitelist so no apps will be safe from volume reudctions
         */
        public void ResetAppWhitelist()
        {
            this._AppWhitelist.Clear();
        }

        /**
         * Adds a voip app to listen for output sound of
         */
        public void AddVOIPApp(string path)
        {
            this._VoipApps.Add(path);
            this.WhitelistApp(path);
        }

        /**
         * Adds an app to the whitelist so its volume won't be reduced
         */
        public void WhitelistApp(string path)
        {
            this._WriteIfDebugging("Ignoring volume from " + path);
            this._AppWhitelist.Add(path);
        }

        /**
         * Listens for audio coming out of known voip applications
         */
        public void ListenForIncomingVOIP()
        {
            //TODO: List of known applications + customisation
            //TODO: This, use ListenIfNoisy
        }

        /**
         * Starts listening to the assigned device
         */
        public void ListenForMic()
        {
            this._MicListener = new WaveInEvent();
            this._MicListener.DeviceNumber = this._MicDeviceNumber;
            this._MicListener.DataAvailable += (object sender, WaveInEventArgs e) =>
            {
                //Calculate maximum volume
                float maximumVolumeInBytes = 0;
                for (int index = 0; index < e.BytesRecorded; index += 2)
                {
                    short sample = (short)((e.Buffer[index + 1] << 8) | e.Buffer[index + 0]);
                    
                    var sample32 = sample / 32768f;
                    if (sample32 < 0)
                    {
                        sample32 = -sample32;
                    }
                    
                    if (sample32 > maximumVolumeInBytes)
                    {
                        maximumVolumeInBytes = sample32;
                    }
                }

                float micVolume = maximumVolumeInBytes * 100;

                if (micVolume > this._MinimumMicVolume)
                {
                    this._WriteIfDebugging("Heard mic at qualifying volume: " + micVolume);
                    this.ApplyReductionIfNoisy();
                }
            };

            this._MicListener.BufferMilliseconds = 50;
            this._MicListener.StartRecording();
        }

        /**
         * Stops listening to the mic so the device can be changed
         */
        public void StopListeningToMic()
        {
            this._MicListener.StopRecording();
            this._MicListener = null;
        }

        /**
         * Applies volume reductions if sound has been heard for a long enough period
         * 
         * @See ApplyReduction
         */
        public void ApplyReductionIfNoisy()
        {
            long now = DateTimeOffset.Now.ToUnixTimeMilliseconds();
            double minimumNoiseLengthMicro = 1000 * this._MinimumNoiseLength;

            if (now - this._LastAudioMicroTime > minimumNoiseLengthMicro)
            {
                this._WriteIfDebugging("Not noisy for long enough, last noise was " + (now - this._LastAudioMicroTime) + " ago");
                this._FirstAudioMicroTime = now;
            }
            else
            {
                if (now - this._FirstAudioMicroTime >= minimumNoiseLengthMicro)
                {
                    this.ApplyReduction();
                }
                else
                {
                    this._WriteIfDebugging("Not noisy for long enough - first noise was " + (now - this._FirstAudioMicroTime) + " ago");
                }
            }

            this._LastAudioMicroTime = now;
        }

        /**
         * Stores the current volume level of all applications and reduces their volumes by the reduction percent
         */
        public void ApplyReduction()
        {
            this._LastReductionAttemptMicroTime = DateTimeOffset.Now.ToUnixTimeMilliseconds();

            if (this._ReductionApplied)
            {
                this._WriteIfDebugging("Already reduced volume levels");
                return;
            }

            this._ReductionApplied = true;
            this.StoreCurrentVolumeLevels();
            this._WriteIfDebugging("Reducing volume levels");

            foreach (AppVolume appVolume in this.GetStoredVolumeLevels())
            {
                if (!this._IsWhitelisted(appVolume))
                {
                    double newVolume = appVolume.Volume - (appVolume.Volume * this._Reduction);
                    this._WriteIfDebugging("Setting " + appVolume.Name + " from " + appVolume.Volume + " to " + newVolume);
                    VolumeMixer.SetApplicationVolume(appVolume.ProcessID, (float)newVolume);
                }
            }

            //Wait for X to reset the timers
            var resetTimer = new Timer();
            resetTimer.Interval = 100;
            resetTimer.AutoReset = true;
            resetTimer.Elapsed += (object sender, ElapsedEventArgs e) =>
            {
                if(DateTimeOffset.Now.ToUnixTimeMilliseconds() - this._LastReductionAttemptMicroTime >= 1000 * this._RestoreVolumeAfter)
                {
                    this.RemoveReduction();
                    resetTimer.Stop();
                    resetTimer.Dispose();
                }
            };
            resetTimer.Start();
        }

        /**
         * Resets all reduced applications to the stored volume level
         */
        public void RemoveReduction()
        {
            if (!this._ReductionApplied)
            {
                return;
            }

            this._WriteIfDebugging("Restoring volume levels");
            foreach (AppVolume appVolume in this.GetStoredVolumeLevels())
            {
                if (!this._IsWhitelisted(appVolume))
                {
                    VolumeMixer.SetApplicationVolume(appVolume.ProcessID, appVolume.Volume);
                }
            }

            this._ReductionApplied = false;
        }

        /**
         * Stores the volume levels for all applications
         */
        public void StoreCurrentVolumeLevels()
        {
            this._WriteIfDebugging("Storing all volume levels");
            this._AppVolumes = AudioReducer.GetAppVolumes();
        }

        /**
         * Stores volume levels for any applications we haven't already stored them for
         */
        public void StoreNewVolumeLevels()
        {
            this._WriteIfDebugging("Updating volume level list");

            foreach (AppVolume appVolume in AudioReducer.GetAppVolumes())
            {
                if(!this._AppVolumes.Any(existingAppVolume => existingAppVolume.ProcessID == appVolume.ProcessID && existingAppVolume.Path == appVolume.Path)) {
                    this._AppVolumes.Add(appVolume);
                }
            }
        }

        /**
         * Gets the last stored volume levels
         */
        public List<AppVolume> GetStoredVolumeLevels()
        {
            return this._AppVolumes;
        }

        /**
         * Gets the volume levels for all current processes
         */
        public static List<AppVolume> GetAppVolumes()
        {
            List<AppVolume> appVolumes = new List<AppVolume>();

            foreach (Process process in Process.GetProcesses())
            {
                float? processVolume = VolumeMixer.GetApplicationVolume(process.Id);

                if (processVolume != null)
                {
                    String processPath = process.StartInfo.FileName;

                    try
                    {
                        processPath = process.MainModule.FileName;
                    }
                    //Some processes can't be accessed, like the system sounds process (Idle)
                    catch(Exception e)
                    {
                        Console.WriteLine("Failed accessing filename for process \"" + process.ProcessName + "\": " + e.Message);
                    }

                    appVolumes.Add(new AppVolume
                    {
                        ProcessID = process.Id,
                        Path = processPath,
                        Name = process.ProcessName,
                        Volume = (float)processVolume
                    });
                }
            }

            return appVolumes;
        }

        private void _WriteIfDebugging(string message)
        {
            if(this._Debug)
            {
                Console.WriteLine("[AudioReducer] " + message);
            }
        }

        private bool _IsWhitelisted(AppVolume appVolume)
        {
            foreach(string path in this._AppWhitelist)
            {
                if(this._DoesPathWildcardMatch(appVolume.Path, path))
                {
                    return true;
                }
            }

            return false;
        }

        private bool _DoesPathWildcardMatch(string path, string wildcardPath)
        {
            return path == wildcardPath || (!wildcardPath.Contains("/") && path.Replace('\\', '/').Split('/').Last().ToLower() == wildcardPath);
        }
    }
}