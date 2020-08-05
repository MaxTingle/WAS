using System.Collections.Generic;
using System.Linq;

namespace WorkingAttenuationSoftware.Controls
{
    public sealed class AudioReducer
    {
        private float _Reduction;
        private List<string> _AppWhitelist = new List<string>();
        private List<AppVolume> _AppVolumes = new List<AppVolume>();

        public AudioReducer(float reduction)
        {
            this._Reduction = reduction;
        }

        public void ReduceDevices()
        {
            this.StoreCurrentVolumeLevels();
        }

        public void SetReduction(float reduction)
        {
            this._Reduction = reduction;
        }

        public void StoreCurrentVolumeLevels()
        {
            this._AppVolumes = AudioMixer.GetAppVolumes();
        }

        public void StoreNewVolumeLevels()
        {
            foreach(AppVolume appVolume in AudioMixer.GetAppVolumes())
            {
                if(!this._AppVolumes.Any(existingAppVolume => existingAppVolume.ProcessID == appVolume.ProcessID && existingAppVolume.Path == appVolume.Path)) {
                    this._AppVolumes.Add(appVolume);
                }
            }
        }

        public void ResetAppWhitelist()
        {
            this._AppWhitelist.Clear();
        }

        public void WhitelistApp(string path)
        {
            this._AppWhitelist.Add(path);
        }
    }
}