using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;

namespace SolutionColor
{
    /// <summary>
    /// Command to reset the title bar color.
    /// </summary>
    internal sealed class ResetColorCommand
    {
        public const int CommandId = 0x0101;

        private SolutionColorPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResetColorCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ResetColorCommand(SolutionColorPackage package)
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
                var menuItem = new MenuCommand(this.Execute, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ResetColorCommand Instance
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
            Instance = new ResetColorCommand(package);
        }

        private void Execute(object sender, EventArgs e)
        {
            package.TitleBarColorControl.ResetTitleBarColor();
            package.Settings.RemoveSolutionColorSetting(VSUtils.GetCurrentSolutionPath());
        }
    }
}
