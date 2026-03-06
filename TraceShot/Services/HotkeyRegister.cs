using NHotkey;
using NHotkey.Wpf;
using System.Windows;
using System.Windows.Input;

namespace TraceShot.Services
{
    public static class HotkeyRegister
    {
        private const string Name = "Bookmark";


        public static string RegisterBookmark(Key key, ModifierKeys mod, EventHandler<HotkeyEventArgs> handler)
        {
            if (TryRegister(key, mod, handler))
            {
                return Format(key, mod);
            }

            System.Windows.MessageBox.Show("ホットキーを登録できませんでした（すべて使用中の可能性があります）。",
                "Hotkey Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return "";
        }

        private static bool TryRegister(Key key, ModifierKeys mod, EventHandler<HotkeyEventArgs> handler)
        {
            try
            {
                HotkeyManager.Current.AddOrReplace(
                    Name,
                    new KeyGesture(key, mod),
                    handler);

                return true;
            }
            catch (HotkeyAlreadyRegisteredException)
            {
                return false;
            }
        }

        public static string Format(Key key, ModifierKeys mod)
        {
            var p = new List<string>();
            if (mod.HasFlag(ModifierKeys.Control)) p.Add("Ctrl");
            if (mod.HasFlag(ModifierKeys.Shift)) p.Add("Shift");
            if (mod.HasFlag(ModifierKeys.Alt)) p.Add("Alt");
            if (mod.HasFlag(ModifierKeys.Windows)) p.Add("Win");
            p.Add(key == Key.Snapshot ? "PrintScreen" : key.ToString());
            return string.Join("+", p);
        }
    }
}