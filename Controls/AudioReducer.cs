using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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

        /**
         * Pinvoking helper class to avoid accessing restricted Process MainModule,
         * which spits out an Access is Denied exception every all
         * 
         * @see https://stackoverflow.com/a/46004513
         */
        private static class PathFinder
        {
            [Flags]
            private enum ProcessAccessFlags : uint
            {
                QueryLimitedInformation = 0x00001000
            }

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern bool QueryFullProcessImageName(
                [In] IntPtr hProcess,
                [In] int dwFlags,
                [Out] StringBuilder lpExeName,
                ref int lpdwSize
            );

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern IntPtr OpenProcess(
                ProcessAccessFlags processAccess,
                bool bInheritHandle,
                int processId
            );

            public static String GetProcessFilename(Process p)
            {
                int capacity = 2000;
                StringBuilder builder = new StringBuilder(capacity);
                IntPtr ptr = PathFinder.OpenProcess(ProcessAccessFlags.QueryLimitedInformation, false, p.Id);
                if (!PathFinder.QueryFullProcessImageName(ptr, 0, builder, ref capacity))
                {
                    return String.Empty;
                }

                return builder.ToString();
            }
        }

        /*
         * Settings
         */

        private double _MinimumNoiseLength = 0;
        private double _MinimumMicVolume = 0;
        private double _MinimumAppVolume = 0;
        private double _RestoreVolumeAfter = 0;
        private double _Reduction;

        private int _MicDeviceNumber = 0;
        private List<string> _AppWhitelist = new List<string>();
        private List<string> _VoipApps = new List<string>();

        private bool _Debug = true;

        private bool _ReductionApplied;
        private List<AppVolume> _AppVolumes = new List<AppVolume>();
        private WaveInEvent _MicListener;

        private long _LastReductionAttemptMicroTime;
        private long _LastAudioMicroTime;
        private long _FirstAudioMicroTime;

        private const int _PIDRecheckInterval = 5000;
        private const int _VolumeRecheckInterval = 50;

        public AudioReducer(double reduction)
        {
            this.SetReduction(reduction);
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
         * Sets the amount of ms to wait for before restoring the original volumes
         */
        public void SetRestoreVolumeAfter(double restoreVolumeAfter)
        {
            this._RestoreVolumeAfter = restoreVolumeAfter;
        }

        /**
         * Sets the minimum volume (As a % of 100) the mic audio can be before we consider it as noise
         */
        public void SetMinimumMicVolume(double minimumMicVolume)
        {
            this._MinimumMicVolume = minimumMicVolume;
        }

        /**
         * Sets the minimum volume (As a % of 100) the voip apps must be before we consider them noisy
         */
        public void SetMinimumAppVolume(double minimumAppVolume)
        {
            this._MinimumAppVolume = minimumAppVolume;
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
            int msSincePidCheck = AudioReducer._PIDRecheckInterval;
            int msSinceCheck = 0;

            List<float> valuesSinceCheck = new List<float>();
            Dictionary<int, string> voipPids = new Dictionary<int, string>();

            Timer voipChecker = new Timer();
            voipChecker.AutoReset = true;
            voipChecker.Interval = AudioReducer._VolumeRecheckInterval;
            voipChecker.Elapsed += (sender, e) =>
            {
                msSinceCheck += (int)voipChecker.Interval;
                msSincePidCheck += (int)voipChecker.Interval;

                //Re-check voip apps for new/removed processes
                if (msSincePidCheck > AudioReducer._PIDRecheckInterval)
                {
                    msSincePidCheck = 0;
                    Process[] currentProcesses = Process.GetProcesses();

                    //Remove pids which are no longer valid
                    foreach (KeyValuePair<int, string> voipPid in voipPids.ToList())
                    {
                        if (!currentProcesses.Any(process => process.Id == voipPid.Key))
                        {
                            voipPids.Remove(voipPid.Key);
                            this._WriteIfDebugging("Voip app " + voipPid.Value + " (" + voipPid.Key + ") process lost");
                        }
                    }

                    //Load the pid by scanning the process list
                    foreach (String voipAppPath in this._VoipApps)
                    {
                        foreach (Process process in currentProcesses)
                        {
                            if (!voipPids.ContainsKey(process.Id))
                            {
                                String processPath = PathFinder.GetProcessFilename(process);

                                if (this._DoesPathWildcardMatch(processPath, voipAppPath))
                                {
                                    voipPids[process.Id] = voipAppPath;
                                    this._WriteIfDebugging("Voip app " + voipAppPath + " found as process " + process.Id);
                                }
                            }
                        }
                    }
                }

                //Update average value lists
                foreach(KeyValuePair<int, string> voipPid in voipPids)
                {
                    float? currentVolume = VolumeMixer.GetApplicationPeak(voipPid.Key);
                    if (currentVolume.HasValue)
                    {
                        valuesSinceCheck.Add(currentVolume.Value);
                    }
                }

                //If we've reached our check threshold check what the values have been
                if (msSinceCheck >= this._MinimumNoiseLength * 1000)
                {
                    double average = 0;
                    if (valuesSinceCheck.Count() > 0)
                    {
                        average = valuesSinceCheck.Average();
                    }

                    if(average > this._MinimumAppVolume)
                    {
                        this._WriteIfDebugging("Heard apps at qualifying average volume: " + average);
                        this.ApplyReductionIfNoisy();
                    }

                    valuesSinceCheck.Clear();
                    msSinceCheck = 0;
                }
            };
            voipChecker.Start();
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

                if (maximumVolumeInBytes > this._MinimumMicVolume)
                {
                    this._WriteIfDebugging("Heard mic at qualifying volume: " + maximumVolumeInBytes);
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

                    if (!this._Debug)
                    {
                        VolumeMixer.SetApplicationVolume(appVolume.ProcessID, (float)newVolume);
                    }
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
                    appVolumes.Add(new AppVolume
                    {
                        ProcessID = process.Id,
                        Path = PathFinder.GetProcessFilename(process),
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