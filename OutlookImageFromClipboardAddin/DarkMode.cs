using Microsoft.Win32;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace OutlookImageFromClipboardAddin
{
    public static class WindowsApiHelper
    {
        public enum DisplayAffinity : uint
        {
            None = 0,
            Monitor = 1
        }

        [DllImport("user32.dll")]
        public static extern bool SetWindowDisplayAffinity(IntPtr hwnd, DisplayAffinity affinity);


        public static bool IsLightMode()
        {
            const String DWM_KEY = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            using (RegistryKey dwmKey = Registry.CurrentUser.OpenSubKey(DWM_KEY, RegistryKeyPermissionCheck.ReadSubTree))
            {
                if (dwmKey is null) return true;

                Object accentColorObj = dwmKey.GetValue("AppsUseLightTheme");
                if (accentColorObj is 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public static bool IsDarkMode()
        {
            return !IsLightMode();
        }

        [DllImport("uxtheme.dll", SetLastError = true, ExactSpelling = true, CharSet = CharSet.Unicode)]

        public static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);

        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        public const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
        public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

        public static bool UseImmersiveDarkMode(IntPtr handle, bool enabled)
        {
            if (IsWindows10OrGreater(17763))
            {
                var attribute = DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1;
                if (IsWindows10OrGreater(18985))
                {
                    attribute = DWMWA_USE_IMMERSIVE_DARK_MODE;
                }

                int useImmersiveDarkMode = enabled ? 1 : 0;
                return DwmSetWindowAttribute(handle, (int)attribute, ref useImmersiveDarkMode, sizeof(int)) == 0;
            }

            return false;
        }

        private static bool IsWindows10OrGreater(int build = -1)
        {
            return Environment.OSVersion.Version.Major >= 10 && Environment.OSVersion.Version.Build >= build;
        }

        public static Color GetAccentColor()
        {
            const String DWM_KEY = @"Software\Microsoft\Windows\DWM";
            using (RegistryKey dwmKey = Registry.CurrentUser.OpenSubKey(DWM_KEY, RegistryKeyPermissionCheck.ReadSubTree))
            {
                const String KEY_EX_MSG = "The \"HKCU\\" + DWM_KEY + "\" registry key does not exist.";
                if (dwmKey is null) throw new InvalidOperationException(KEY_EX_MSG);

                Object accentColorObj = dwmKey.GetValue("AccentColor");
                if (accentColorObj is Int32 accentColorDword)
                {
                    return ParseDWordColor(accentColorDword);
                }
                else
                {
                    const String VALUE_EX_MSG = "The \"HKCU\\" + DWM_KEY + "\\AccentColor\" registry key value could not be parsed as an ABGR color.";
                    throw new InvalidOperationException(VALUE_EX_MSG);
                }
            }

        }

        public static Color GetReadableForeColor(Color c)
        {
            return GetReadableForeColor(c, Color.Black, Color.White);
        }

        public static Color GetReadableForeColor(Color c, Color MinColor, Color MaxColor)
        {
            return (((c.R + c.B + c.G) / 3) > 128) ? MinColor : MaxColor;
        }

        private static Color ParseDWordColor(Int32 color)
        {
            Byte
                a = (byte)((color >> 24) & 0xFF),
                b = (byte)((color >> 16) & 0xFF),
                g = (byte)((color >> 8) & 0xFF),
                r = (byte)((color >> 0) & 0xFF);

            return Color.FromArgb(a, r, g, b);
        }
    }
    public class CustomColorTable : ProfessionalColorTable
    {
        public static Color dark1 = Color.FromArgb(20, 20, 20);
        public static Color dark2 = Color.FromArgb(45, 45, 45);
        public static Color dark3 = Color.FromArgb(60, 60, 60);
        public static Color dark4 = Color.FromArgb(80, 80, 80);
        public static Color dark5 = Color.FromArgb(110, 110, 110);

        public static Color light1 = Color.White;

        public static Color accent = WindowsApiHelper.GetAccentColor();
        //public static Color accent = Color.Red;

        public override Color ButtonPressedHighlight
        {
            get { return accent; }
        }

        public override Color ButtonPressedGradientBegin
        {
            get { return accent; }
        }
        public override Color ButtonPressedGradientMiddle
        {
            get { return accent; }
        }
        public override Color ButtonPressedGradientEnd
        {
            get { return accent; }
        }

        public override Color ButtonSelectedBorder
        {
            get { return accent; }
        }

        public override Color ButtonSelectedHighlight
        {
            get { return accent; }
        }
        public override Color ButtonSelectedHighlightBorder
        {
            get { return accent; }
        }

        public override Color ButtonSelectedGradientBegin
        {
            get { return accent; }
        }
        public override Color ButtonSelectedGradientMiddle
        {
            get { return accent; }
        }
        public override Color ButtonSelectedGradientEnd
        {
            get { return accent; }
        }


        public override Color ButtonCheckedGradientBegin
        {
            get { return dark2; }
        }
        public override Color ButtonCheckedGradientMiddle
        {
            get { return dark2; }
        }
        public override Color ButtonCheckedGradientEnd
        {
            get { return dark2; }
        }


        public override Color ButtonCheckedHighlight
        {
            get { return Color.Yellow; }
        }

        public override Color ToolStripGradientBegin
        {
            get { return dark2; }
        }
        public override Color ToolStripGradientMiddle
        {
            get { return dark2; }
        }
        public override Color ToolStripGradientEnd
        {
            get { return dark2; }
        }

        public override Color ToolStripContentPanelGradientBegin
        {
            get { return dark2; }
        }
        public override Color ToolStripContentPanelGradientEnd
        {
            get { return dark2; }
        }
        public override Color GripDark
        {
            get { return dark1; }
        }

        public override Color GripLight
        {
            get { return dark3; }
        }
        public override Color SeparatorDark
        {
            get { return dark1; }
        }
        public override Color SeparatorLight
        {
            get { return dark3; }
        }
        public override Color MenuBorder
        {
            get { return dark3; }
        }
        public override Color MenuItemBorder
        {
            get { return accent; }
        }
        public override Color MenuItemSelected
        {
            get { return accent; }
        }
        public override Color StatusStripGradientBegin
        {
            get { return dark1; }
        }
        public override Color StatusStripGradientEnd
        {
            get { return dark1; }
        }

        public override Color CheckBackground
        {
            get { return accent; }
        }

        public override Color CheckSelectedBackground
        {
            get { return accent; }
        }

        public override Color CheckPressedBackground
        {
            get { return Color.Lime; }
        }


        public override Color MenuItemPressedGradientBegin
        {
            get { return dark3; }
        }

        public override Color MenuItemPressedGradientEnd
        {
            get { return dark3; }
        }

        public override Color MenuItemPressedGradientMiddle
        {
            get { return dark3; }
        }


        public override Color ToolStripDropDownBackground
        {
            get { return dark2; }
        }
        public override Color ImageMarginGradientBegin
        {
            get { return dark2; }
        }
        public override Color ImageMarginGradientMiddle
        {
            get { return dark2; }
        }
        public override Color ImageMarginGradientEnd
        {
            get { return dark2; }
        }



        public override Color ToolStripBorder
        {
            get { return dark3; }
        }
    }
}
