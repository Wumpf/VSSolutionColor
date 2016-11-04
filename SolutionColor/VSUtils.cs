using Microsoft.VisualStudio.Shell;
using System;
using System.Drawing;
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

        public static string GetCurrentSolutionPath()
        {
            return GetDTE().Solution.FileName;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        private static DependencyObject cachedTitleBar = null;
        private static object defaultBackgroundValue = null;
        private const string ColorPropertyName = "Background";

        private static DependencyObject GetTitleBar()
        {
            if (cachedTitleBar == null)
            {
                IntPtr active = GetActiveWindow();
                var activeWindow = Application.Current.Windows.OfType<Window>()
                    .SingleOrDefault(window => new WindowInteropHelper(window).Handle == active);

                // Apply knowledge of basic Visual Studio 2015 window structure.
                var windowContentPresenter = VisualTreeHelper.GetChild(activeWindow, 0);
                var rootGrid = VisualTreeHelper.GetChild(windowContentPresenter, 0);

                cachedTitleBar = VisualTreeHelper.GetChild(rootGrid, 0);
                if (cachedTitleBar != null)
                {
                    System.Reflection.PropertyInfo propertyInfo = cachedTitleBar.GetType().GetProperty(ColorPropertyName);
                    defaultBackgroundValue = propertyInfo.GetValue(cachedTitleBar);
                }
            }

            return cachedTitleBar;
        }

        public static void SetTitleBarColor(System.Drawing.Color color)
        {
            try
            {
                var titleBar = GetTitleBar();

                var newBackground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B));

                System.Reflection.PropertyInfo propertyInfo = titleBar.GetType().GetProperty(ColorPropertyName);
                propertyInfo.SetValue(titleBar, newBackground, null);
            }
            catch(Exception e)
            {
                MessageBox.Show("Failed to set the color of the title bar:\n" + e.ToString(), "Failed to set Titlebar Color");
            }
        }

        public static void ResetTitleBarColor()
        {
            try
            {
                var titleBar = GetTitleBar();

                System.Reflection.PropertyInfo propertyInfo = titleBar.GetType().GetProperty(ColorPropertyName);
                propertyInfo.SetValue(titleBar, defaultBackgroundValue, null);
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to reset the color of the title bar:\n" + e.ToString(), "Failed to reset Titlebar Color");
            }
        }

        public static System.Drawing.Color TryGetTitleBarColor()
        {
            try
            {
                var titleBar = GetTitleBar();

                System.Reflection.PropertyInfo propertyInfo = titleBar.GetType().GetProperty(ColorPropertyName);
                var colorBrush = propertyInfo.GetValue(titleBar) as SolidColorBrush;
                if (colorBrush != null)
                {
                    return System.Drawing.Color.FromArgb(colorBrush.Color.A, colorBrush.Color.R, colorBrush.Color.G, colorBrush.Color.B);
                }
                else
                    return System.Drawing.Color.Black;
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to get the color of the title bar:\n" + e.ToString(), "Failed to get Titlebar Color");
            }

            return System.Drawing.Color.Black;
        }
    }
}
