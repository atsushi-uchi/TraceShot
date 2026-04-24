using System.IO;
using System.Windows.Media;

namespace TraceShot.Services
{
    public class SoundService
    {
        public static SoundService Instance { get; } = new SoundService();

        private MediaPlayer _shutterPlayer = new();
        private MediaPlayer _voiceStartPlayer = new();

        // 0.0 ～ 1.0 の範囲。設定値（0～100など）から変換してセット
        public double Volume
        {
            get => _shutterPlayer.Volume;
            set
            {
                _shutterPlayer.Volume = value;
                _voiceStartPlayer.Volume = value;
            }
        }

        private SoundService()
        {
            InitializePlayer(_shutterPlayer, "Camera-Phone03-1.wav");
            InitializePlayer(_voiceStartPlayer, "Camera-Phone03-2.wav");
            // デフォルト音量を 0.5 (50%) くらいにしておく例
            Volume = 0.5;
        }

        private void InitializePlayer(MediaPlayer player, string fileName)
        {
            try
            {
                string folder = @"Resources\Sounds\";
                string soundPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{folder}{fileName}");
                if (File.Exists(soundPath))
                {
                    player.Open(new Uri(soundPath));
                }
            }
            catch { /* ロギング */ }
        }

        public void PlayShutter() => Play(_shutterPlayer);
        public void VoiceStartShutter() => Play(_voiceStartPlayer);

        private void Play(MediaPlayer player)
        {
            try
            {
                // MediaPlayerは再生が終わっても位置が最後で止まるため、
                // 再生するたびに位置を最初に戻す必要があります
                player.Stop();
                player.Play();
            }
            catch
            {
                System.Media.SystemSounds.Beep.Play();
            }
        }
    }
}