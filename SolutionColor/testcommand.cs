//------------------------------------------------------------------------------
// <copyright file="testcommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Runtime.InteropServices;
using System.Windows;
using System.Linq;
using System.Windows.Interop;
using System.Windows.Media;

namespace SolutionColor
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class testcommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("00d80876-3407-4666-bf62-7262028ea83b");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="testcommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private testcommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MessUpTitleBarColor, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static testcommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new testcommand(package);
        }

        [DllImport("user32.dll")]
        static extern IntPtr GetActiveWindow();

        private void MessUpTitleBarColor(object sender, EventArgs e)
        {
            IntPtr active = GetActiveWindow();
            var activeWindow = Application.Current.Windows.OfType<Window>()
                .SingleOrDefault(window => new WindowInteropHelper(window).Handle == active);


            // Apply knowledge of basic Visual Studio 2015 window structure.
            var windowContentPresenter = VisualTreeHelper.GetChild(activeWindow, 0);
            var rootGrid = VisualTreeHelper.GetChild(windowContentPresenter, 0);
            var mainWindowTitleBar = VisualTreeHelper.GetChild(rootGrid, 0);

            var brushConverter = new BrushConverter();
            var newBackground = (Brush)brushConverter.ConvertFrom("#ff0000");

            System.Reflection.PropertyInfo propertyInfo = mainWindowTitleBar.GetType().GetProperty("Background");
            propertyInfo.SetValue(mainWindowTitleBar, newBackground, null);
        }
    }
}
