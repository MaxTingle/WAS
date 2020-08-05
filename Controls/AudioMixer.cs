using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Documents;
using System.Collections.Generic;
using System;

namespace WorkingAttenuationSoftware.Controls
{
    public struct AppVolume
    {
        public int ProcessID;
        public string Path;
        public float Volume;
    }

    public static class AudioMixer
    {
        //Lib to use https://docs.microsoft.com/en-gb/windows/win32/coreaudio/wasapi

        //Find the class you wanna use https://docs.microsoft.com/en-us/windows/win32/api/audioclient/nn-audioclient-isimpleaudiovolume
        //Get the guid from https://www.magnumdb.com/
        //Reference the guid, put the methods in

        //Do it for you libs
        //https://stackoverflow.com/questions/20938934/controlling-applications-volume-by-process-id
        //https://stackoverflow.com/questions/14306048/controlling-volume-mixer
        //https://archive.codeplex.com/?p=netcoreaudio

        public static List<AppVolume> GetAppVolumes()
        {
            List<AppVolume> appVolumes = new List<AppVolume>();

            //TODO: Replace loop through process list with api call to get all audio sessions
            foreach(Process process in Process.GetProcesses())
            {
                float? processVolume = AudioMixer.GetVolumeLevel(process.Id);

                if (processVolume != null)
                {
                    appVolumes.Add(new AppVolume {
                        ProcessID = process.Id,
                        Path = process.ProcessName,
                        Volume = (float)processVolume
                    });
                }
            }

            return appVolumes;
        }

        public static float? GetVolumeLevel(int processId)
        {

        }

        [Guid("87ce5498-68d6-44e5-9215-6da47ef883d8"), ComImport(ComInterfaceType.InterfaceIsIUnknown)]
        internal class ISimpleAudioVolume
        {

        }
    }
}