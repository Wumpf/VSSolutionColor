using Microsoft.VisualStudio.Shell;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace SolutionColor
{
    internal static class VSUtils
    {
        public static EnvDTE80.DTE2 GetDTE()
        {
            var dte = Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;
            if (dte == null)
            {
                throw new System.Exception("Failed to retrieve DTE2!");
            }
            return dte;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        public static void SetTitleBarColor(string hexColor)
        {
            try
            {
                IntPtr active = GetActiveWindow();
                var activeWindow = Application.Current.Windows.OfType<Window>()
                    .SingleOrDefault(window => new WindowInteropHelper(window).Handle == active);

                // Apply knowledge of basic Visual Studio 2015 window structure.
                var windowContentPresenter = VisualTreeHelper.GetChild(activeWindow, 0);

                var rootGrid = VisualTreeHelper.GetChild(windowContentPresenter, 0);
                var mainWindowTitleBar = VisualTreeHelper.GetChild(rootGrid, 0);

                var brushConverter = new BrushConverter();
                var newBackground = (Brush)brushConverter.ConvertFrom(hexColor);

                System.Reflection.PropertyInfo propertyInfo = mainWindowTitleBar.GetType().GetProperty("Background");
                propertyInfo.SetValue(mainWindowTitleBar, newBackground, null);
            }
            catch(Exception e)
            {
                MessageBox.Show("Failed to set titlebar color", "Failed to set the color of the title bar:\n" + e.ToString());
            }
        }
    }
}
