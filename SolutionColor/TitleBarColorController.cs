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
        private DependencyObject titleBarContainer = null;
        private TextBlock titleBarTextBox = null;

        private object defaultBackgroundValue = null;
        private const string ColorPropertyName = "Background";
        private System.Windows.Media.Brush defaultTextForeground = null;

        private TitleBarColorController()
        {
        }

        /// <summary>
        /// Uses knowledge of the VS window structure to retrieve pointers to the titlebar and its text element.
        /// </summary>
        static public TitleBarColorController CreateFromWindow(Window window)
        {
            TitleBarColorController newController = new TitleBarColorController();
            try
            {
                // Apply knowledge of basic Visual Studio 2015/2017/2019 window structure.

                if (window == Application.Current.MainWindow)
                {
                    var windowContentPresenter = VisualTreeHelper.GetChild(window, 0);
                    var rootGrid = VisualTreeHelper.GetChild(windowContentPresenter, 0);

                    newController.titleBarContainer = VisualTreeHelper.GetChild(rootGrid, 0);

                    // Note that this part doesn't work for the VS2019 main windows as there is simply no title text like this.
                    // However docked-out code windows are just like in previous versions.
                    var dockPanel = VisualTreeHelper.GetChild(newController.titleBarContainer, 0);
                    newController.titleBarTextBox = VisualTreeHelper.GetChild(dockPanel, 3) as TextBlock;
                }
                else
                {
                    var windowContentPresenter = VisualTreeHelper.GetChild(window, 0);
                    var rootGrid = VisualTreeHelper.GetChild(windowContentPresenter, 0);
                    var rootDockPanel = VisualTreeHelper.GetChild(rootGrid, 0);
                    var titleBarContainer = VisualTreeHelper.GetChild(rootDockPanel, 0);
                    var titleBar = VisualTreeHelper.GetChild(titleBarContainer, 0);
                    var border = VisualTreeHelper.GetChild(titleBar, 0);
                    var contentPresenter = VisualTreeHelper.GetChild(border, 0);
                    var grid = VisualTreeHelper.GetChild(contentPresenter, 0);

                    newController.titleBarContainer = grid;

                    newController.titleBarTextBox = VisualTreeHelper.GetChild(grid, 1) as TextBlock;
                }

                if (newController.titleBarContainer != null)
                {
                    System.Reflection.PropertyInfo propertyInfo = newController.titleBarContainer.GetType().GetProperty(ColorPropertyName);
                    newController.defaultBackgroundValue = propertyInfo.GetValue(newController.titleBarContainer);
                }

                if (newController.titleBarTextBox != null)
                    newController.defaultTextForeground = newController.titleBarTextBox.Foreground;
            }
            catch
            {
                return null;
            }

            if (newController.titleBarContainer == null)
                return null;

            return newController;
        }

        /// <summary>
        /// Tries to set a given color to the titlebar. Will color the text either black or white dependingon the color's brightness.
        /// Opens message box if something goes wrong.
        /// </summary>
        public void SetTitleBarColor(System.Drawing.Color color)
        {
            try
            {
                if (titleBarContainer != null)
                {
                    System.Reflection.PropertyInfo propertyInfo = titleBarContainer.GetType().GetProperty(ColorPropertyName);
                    propertyInfo.SetValue(titleBarContainer, new SolidColorBrush(System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B)), null);
                }

                if (titleBarTextBox != null)
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
            try
            {
                if (titleBarContainer != null)
                {
                    System.Reflection.PropertyInfo propertyInfo = titleBarContainer.GetType().GetProperty(ColorPropertyName);
                    propertyInfo.SetValue(titleBarContainer, defaultBackgroundValue, null);
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
            try
            {
                System.Reflection.PropertyInfo propertyInfo = titleBarContainer.GetType().GetProperty(ColorPropertyName);
                var colorBrush = propertyInfo.GetValue(titleBarContainer) as SolidColorBrush;
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
