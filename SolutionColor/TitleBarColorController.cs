using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;

namespace SolutionColor
{
    /// <summary>
    /// All the important magic to manipulate the main window's title bar happens here.
    /// </summary>
    public class TitleBarColorController
    {
        private DependencyObject titleBar = null;
        private TextBlock titleBarTextBox = null;

        private object defaultBackgroundValue = null;
        private const string ColorPropertyName = "Background";
        private System.Windows.Media.Brush defaultTextForeground = null;

        public TitleBarColorController()
        {
            // Too early to gather widget pointers since window might not be up yet.
            //UpdateWidgetPointer();
        }

        /// <summary>
        /// Uses knowledge of the VS window structure to retrieve pointers to the titlebar and its text element.
        /// Don't call before the application window isn't up yet.
        /// </summary>
        private void UpdateWidgetPointer()
        {
            if (titleBar != null && titleBarTextBox != null)
                return;

            try
            {
                var mainWindow = Application.Current.MainWindow;

                // Apply knowledge of basic Visual Studio 2015 window structure.
                var windowContentPresenter = VisualTreeHelper.GetChild(mainWindow, 0);
                var rootGrid = VisualTreeHelper.GetChild(windowContentPresenter, 0);

                titleBar = VisualTreeHelper.GetChild(rootGrid, 0);
                System.Reflection.PropertyInfo propertyInfo = titleBar.GetType().GetProperty(ColorPropertyName);
                defaultBackgroundValue = propertyInfo.GetValue(titleBar);

                var dockPanel = VisualTreeHelper.GetChild(titleBar, 0);

                titleBarTextBox = VisualTreeHelper.GetChild(dockPanel, 3) as TextBlock;
                defaultTextForeground = titleBarTextBox.Foreground;
            }
            catch  { }
        }

        /// <summary>
        /// Tries to set a given color to the titlebar. Will color the text either black or white dependingon the color's brightness.
        /// Opens message box if something goes wrong.
        /// </summary>
        public void SetTitleBarColor(System.Drawing.Color color)
        {
            UpdateWidgetPointer();

            try
            {
                if (titleBar != null)
                {
                    System.Reflection.PropertyInfo propertyInfo = titleBar.GetType().GetProperty(ColorPropertyName);
                    propertyInfo.SetValue(titleBar, new SolidColorBrush(System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B)), null);
                }

                if(titleBarTextBox != null)
                {
                    float luminance = 0.299f * color.R + 0.587f * color.G + 0.114f * color.B;
                    if(luminance > 128.0f)
                        titleBarTextBox.Foreground = new SolidColorBrush(Color.FromRgb(0,0,0));
                    else
                        titleBarTextBox.Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                }
            }
            catch(Exception e)
            {
                MessageBox.Show("Failed to set the color of the title bar:\n" + e.ToString(), "Failed to set Titlebar Color");
            }
        }

        /// <summary>
        /// Resets titlebar (and text) color to the default that was saved in the first successful UpdateWidgetPointer() call.
        /// </summary>
        public void ResetTitleBarColor()
        {
            // If we don't have the widget pointers yet, we also don't have a color to reset to.
            //UpdateWidgetPointer();

            try
            {
                if (titleBar != null)
                {
                    System.Reflection.PropertyInfo propertyInfo = titleBar.GetType().GetProperty(ColorPropertyName);
                    propertyInfo.SetValue(titleBar, defaultBackgroundValue, null);
                }

                if (titleBarTextBox != null)
                {
                    titleBarTextBox.Foreground = defaultTextForeground;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to reset the color of the title bar:\n" + e.ToString(), "Failed to reset Titlebar Color");
            }
        }

        /// <summary>
        /// Tries to retrieve the color of the title bar. Falls back to black if it fails.
        /// </summary>
        public System.Drawing.Color TryGetTitleBarColor()
        {
            UpdateWidgetPointer();

            try
            {
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
