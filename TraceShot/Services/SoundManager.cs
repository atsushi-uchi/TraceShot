using System.IO;
using System.Media;

namespace TraceShot.Services
{
    public class SoundManager
    {
        public static SoundManager Instance { get; } = new SoundManager();

        private SoundPlayer? _shutterPlayer;
        private SoundPlayer? _voiceStartPlayer;

        private SoundManager()
        {
            try
            {
                string soundPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Resources\shutter.wav");
                if (File.Exists(soundPath))
                {
                    _shutterPlayer = new SoundPlayer(soundPath);
                    _shutterPlayer.Load(); // 💡 先に読み込んでおくことで再生時の遅延を無くす
                }
            }
            catch { /* ロギングなど */ }

            try
            {
                string soundPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Resources\voice_start.wav");
                if (File.Exists(soundPath))
                {
                    _voiceStartPlayer = new SoundPlayer(soundPath);
                    _voiceStartPlayer.Load(); // 💡 先に読み込んでおくことで再生時の遅延を無くす
                }
            }
            catch { /* ロギングなど */ }
        }

        public void PlayShutter() 
        {
            try
            {
                if (_shutterPlayer != null)
                {
                    _shutterPlayer.Play();
                }
                else
                {
                    SystemSounds.Asterisk.Play();
                }
            }
            catch
            {
                SystemSounds.Beep.Play();
            }
        }

        public void VoiceStartShutter() 
        {
            try
            {
                if (_voiceStartPlayer != null)
                {
                    _voiceStartPlayer.Play();
                }
                else
                {
                    SystemSounds.Asterisk.Play();
                }
            }
            catch
            {
                SystemSounds.Beep.Play();
            }
        }
    }
}