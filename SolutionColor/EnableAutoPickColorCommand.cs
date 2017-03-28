using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;

namespace SolutionColor
{
    /// <summary>
    /// Command enable/disable automatic color picking.
    /// </summary>
    internal sealed class EnableAutoPickColorCommand
    {
        public const int CommandId = 0x0102;

        private SolutionColorPackage package;
        private MenuCommand menuItem;

        /// <summary>
        /// Initializes a new instance of the <see cref="EnableAutoPickColorCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private EnableAutoPickColorCommand(SolutionColorPackage package)
        {
            this.package = package;
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            OleMenuCommandService commandService = ((IServiceProvider)package).GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(SolutionColorPackage.ToolbarCommandSetGuid, CommandId);
                menuItem = new MenuCommand(this.Execute, menuCommandID);
                commandService.AddCommand(menuItem);

                menuItem.Checked = package.Settings.IsAutomaticColorPickEnabled();
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static EnableAutoPickColorCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(SolutionColorPackage package)
        {
            Instance = new EnableAutoPickColorCommand(package);
        }

        private void Execute(object sender, EventArgs e)
        {
            menuItem.Checked = !menuItem.Checked;
            package.Settings.SetAutomaticColorPickEnabled(menuItem.Checked);
        }
    }
}
