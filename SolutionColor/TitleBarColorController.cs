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
        private DependencyObject mainMenuControl = null;
        private DependencyObject mainMenuItemsWrapperControl = null;
        private TextBlock titleBarTextBox = null;

        private object defaultBackgroundValue = null;
        private Brush defaultTextForeground = null;
        private object defaultMenuBackgroundValue = null;
        private Style defaultMenuItemStyle = null;
        private const string ColorPropertyName = "Background";
        private const string ForegroundPropertyName = "Foreground";

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
                    try
                    {
                        var dockPanel = VisualTreeHelper.GetChild(newController.titleBarContainer, 0);
                        newController.titleBarTextBox = VisualTreeHelper.GetChild(dockPanel, 3) as TextBlock;
                    }
                    catch
                    {
                        // We can do without the text box!
                    }

                    // In VS2019+ the main menu has been integrated with the title bar.
                    // We can set the opacity to 0 and color the text as we did/do with the title text.
                    newController.mainMenuControl = GetDecendantFirstInLine(newController.titleBarContainer, 6);
                    if (newController.mainMenuControl != null)
                    {
                        newController.mainMenuItemsWrapperControl = GetDecendantFirstInLine(newController.mainMenuControl, 3);
                        if (newController.mainMenuItemsWrapperControl is MenuItem) // Nestedness of the layout changed a bit over different versions;
                            newController.mainMenuItemsWrapperControl = GetDecendantFirstInLine(newController.mainMenuControl, 2);

                        System.Reflection.PropertyInfo propertyInfo = newController.mainMenuControl.GetType().GetProperty(ColorPropertyName);
                        newController.defaultMenuBackgroundValue = propertyInfo.GetValue(newController.mainMenuControl);
                        if (newController.mainMenuItemsWrapperControl != null && VisualTreeHelper.GetChildrenCount(newController.mainMenuItemsWrapperControl) > 0)
                        {
                            var child = VisualTreeHelper.GetChild(newController.mainMenuItemsWrapperControl, 0);
                            newController.defaultMenuItemStyle = child.GetType().GetProperty("Style")?.GetValue(child) as Style;
                        }
                    }
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
                float luminance = 0.299f * color.R + 0.587f * color.G + 0.114f * color.B;
                var textColor = (luminance > 128.0f) ? Color.FromRgb(0, 0, 0) : Color.FromRgb(255, 255, 255);
                var textBrush = new SolidColorBrush(textColor);

                if (titleBarContainer != null)
                {
                    System.Reflection.PropertyInfo propertyInfo = titleBarContainer.GetType().GetProperty(ColorPropertyName);
                    propertyInfo.SetValue(titleBarContainer, new SolidColorBrush(Color.FromArgb(color.A, color.R, color.G, color.B)), null);
                }

                if (titleBarTextBox != null)
                {
                    titleBarTextBox.Foreground = textBrush;
                }

                if (mainMenuControl != null)
                {
                    System.Reflection.PropertyInfo propertyInfo = mainMenuControl.GetType().GetProperty(ColorPropertyName);
                    propertyInfo.SetValue(mainMenuControl, new SolidColorBrush(Colors.Transparent));
                }

                if (mainMenuItemsWrapperControl != null)
                {
                    var newMenuItemStyle = CreateNewMenuItemStyle(mainMenuItemsWrapperControl, textBrush);
                    if (newMenuItemStyle != null)
                    {
                        ApplyStyleOnAllChildren(mainMenuItemsWrapperControl, newMenuItemStyle);
                    }
                }
            }
            catch (Exception e)
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

                if (mainMenuControl != null)
                {
                    System.Reflection.PropertyInfo propertyInfo = mainMenuControl.GetType().GetProperty(ColorPropertyName);
                    propertyInfo.SetValue(mainMenuControl, defaultMenuBackgroundValue);
                }

                if (mainMenuItemsWrapperControl != null)
                {
                    ApplyStyleOnAllChildren(mainMenuItemsWrapperControl, defaultMenuItemStyle);
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

        /// <summary>
        /// Creates a new Style with the supplied <paramref name="newTextBrush"/> based on the Style of the first Child object of <paramref name="menuItemWrapper"/>
        /// </summary>
        /// <param name="menuItemWrapper"></param>
        /// <param name="newTextBrush"></param>
        /// <returns></returns>
        private Style CreateNewMenuItemStyle(DependencyObject menuItemWrapper, SolidColorBrush newTextBrush)
        {
            if (defaultMenuItemStyle == null)
                return null;

            var newStyle = new Style(defaultMenuItemStyle.TargetType, defaultMenuItemStyle);
            foreach (var setter in defaultMenuItemStyle.Setters)
            {
                if ((setter is Setter) && ((setter as Setter).Property.ToString() == ForegroundPropertyName))
                {
                    newStyle.Setters.Remove(setter);
                    newStyle.Setters.Add(new Setter((setter as Setter).Property, newTextBrush));
                }
            }
            return newStyle;
        }

        /// <summary>
        /// Applies the <paramref name="styleToApply"/> to all children of <paramref name="menuItemWrapper"/>
        /// </summary>
        /// <param name="menuItemWrapper">The object that has all the relevant MenuItems as children</param>
        /// <param name="styleToApply">The new style to apply to all children</param>
        private void ApplyStyleOnAllChildren(DependencyObject menuItemWrapper, Style styleToApply)
        {
            if (menuItemWrapper == null)
                throw new ArgumentNullException(nameof(menuItemWrapper));

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(menuItemWrapper); i++)
            {
                var menuItem = VisualTreeHelper.GetChild(menuItemWrapper, i) as MenuItem;
                if (menuItem != null)
                {
                    menuItem.Style = styleToApply;
                }
            }
        }

        /// <summary>
        /// Call VisualTreeHelper.GetChild(ref, 0) multiple times, to get a first-in-line decendant x levels deep.
        /// </summary>
        /// <param name="reference">The parent visual to get the decendant of</param>
        /// <param name="levelsDeep">The amount of levels deep (if 1 is passed, this method behaves as GetChild).</param>
        static private DependencyObject GetDecendantFirstInLine(DependencyObject reference, int levelsDeep)
        {
            while ((reference != null) && (levelsDeep > 0))
            {
                if (VisualTreeHelper.GetChildrenCount(reference) < 1)
                    return null;
                reference = VisualTreeHelper.GetChild(reference, 0);
                levelsDeep--;
            }
            return reference;
        }

    }
}
