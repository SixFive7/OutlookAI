using System;
using System.Drawing;
using Microsoft.Win32;

namespace OutlookAI.Services
{
    public static class ThemeService
    {
        public static bool IsDarkMode { get; private set; }

        public static Color Background { get; private set; }
        public static Color ControlBackground { get; private set; }
        public static Color TextBoxBackground { get; private set; }
        public static Color Text { get; private set; }
        public static Color SecondaryText { get; private set; }
        public static Color Accent { get; private set; }
        public static Color Border { get; private set; }
        public static Color StatusError { get; private set; }
        public static Color StatusSuccess { get; private set; }
        public static Color LinkError { get; private set; }
        public static Color ButtonFace { get; private set; }
        public static Color ButtonText { get; private set; }
        public static Color ResultBackground { get; private set; }

        static ThemeService()
        {
            SetLightDefaults();
            try { Detect(); } catch { }

            try
            {
                SystemEvents.UserPreferenceChanged += (s, e) =>
                {
                    if (e.Category == UserPreferenceCategory.General)
                    {
                        try { Detect(); } catch { }
                    }
                };
            }
            catch { }
        }

        private static void SetLightDefaults()
        {
            IsDarkMode = false;
            Background = Color.FromArgb(250, 249, 248);
            ControlBackground = SystemColors.Control;
            TextBoxBackground = SystemColors.Window;
            Text = SystemColors.ControlText;
            SecondaryText = Color.Gray;
            Accent = Color.FromArgb(0, 120, 212);
            Border = SystemColors.ControlDark;
            StatusError = Color.DarkRed;
            StatusSuccess = Color.DarkGreen;
            LinkError = Color.IndianRed;
            ButtonFace = SystemColors.ButtonFace;
            ButtonText = SystemColors.ControlText;
            ResultBackground = SystemColors.Window;
        }

        public static void Detect()
        {
            IsDarkMode = DetectDarkMode();

            if (IsDarkMode)
            {
                Background = Color.FromArgb(32, 32, 32);
                ControlBackground = Color.FromArgb(45, 45, 48);
                TextBoxBackground = Color.FromArgb(51, 51, 55);
                Text = Color.FromArgb(230, 230, 230);
                SecondaryText = Color.FromArgb(160, 160, 160);
                Accent = Color.FromArgb(75, 156, 245);
                Border = Color.FromArgb(70, 70, 74);
                StatusError = Color.FromArgb(255, 120, 120);
                StatusSuccess = Color.FromArgb(120, 220, 120);
                LinkError = Color.FromArgb(255, 140, 140);
                ButtonFace = Color.FromArgb(55, 55, 60);
                ButtonText = Color.FromArgb(230, 230, 230);
                ResultBackground = Color.FromArgb(40, 40, 44);
            }
            else
            {
                SetLightDefaults();
            }
        }

        private static bool DetectDarkMode()
        {
            try
            {
                foreach (var ver in new[] { "16.0", "17.0", "15.0" })
                {
                    using (var key = Registry.CurrentUser.OpenSubKey($@"SOFTWARE\Microsoft\Office\{ver}\Common"))
                    {
                        var theme = key?.GetValue("UI Theme");
                        if (theme is int themeValue)
                        {
                            if (themeValue == 4 || themeValue == 5)
                                return true;
                            if (themeValue == 0 || themeValue == 3)
                                return false;
                        }
                    }
                }
            }
            catch { }

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    var value = key?.GetValue("AppsUseLightTheme");
                    if (value is int lightTheme)
                        return lightTheme == 0;
                }
            }
            catch { }

            return false;
        }
    }
}
