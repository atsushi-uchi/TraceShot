using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using NHotkey;
using NHotkey.Wpf;

namespace TraceShot.Services
{
    public static class HotkeyRegister
    {
        private const string Name = "Bookmark";

        public static string RegisterBookmark(EventHandler<HotkeyEventArgs> handler)
        {
            foreach (var (key, mod) in Candidates())
            {
                if (TryRegister(key, mod, handler))
                {
                    return Format(key, mod);
                }
            }

            System.Windows.MessageBox.Show("ホットキーを登録できませんでした（すべて使用中の可能性があります）。",
                "Hotkey Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return "";
        }

        private static IEnumerable<(Key key, ModifierKeys mod)> Candidates()
        {
            //yield return (Key.Snapshot, ModifierKeys.None);  // PrintScreen
            yield return (Key.F12, ModifierKeys.None);         // F12
            yield return (Key.F12, ModifierKeys.Control);      // Ctrl+F12
            yield return (Key.Snapshot, ModifierKeys.Control); // Ctrl+PrintScreen

            for (var k = Key.F11; k >= Key.F1; k--)            // F11→F1
                yield return (k, ModifierKeys.None);
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

        private static string Format(Key key, ModifierKeys mod)
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