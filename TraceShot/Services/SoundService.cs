using System.IO;
using System.Media;

namespace TraceShot.Services
{
    public class SoundService
    {
        public static SoundService Instance { get; } = new SoundService();

        private SoundPlayer? _shutterPlayer;
        private SoundPlayer? _voiceStartPlayer;
        private string folder = @"Resources\Sounds\";

        private SoundService()
        {
            try
            {
                string filename = "Camera-Phone03-1.wav";
                string soundPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{folder}{filename}");
                if (File.Exists(soundPath))
                {
                    _shutterPlayer = new SoundPlayer(soundPath);
                    _shutterPlayer.Load(); // 💡 先に読み込んでおくことで再生時の遅延を無くす
                }
            }
            catch { /* ロギングなど */ }

            try
            {
                string filename = "Camera-Phone03-2.wav";
                string soundPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{folder}{filename}");
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
                    SystemSounds.Exclamation.Play();
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
                    SystemSounds.Hand.Play();
                }
            }
            catch
            {
                SystemSounds.Beep.Play();
            }
        }
    }
}