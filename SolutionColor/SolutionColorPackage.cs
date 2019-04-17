using System;
using System.Windows;
using System.Windows.Automation;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using System.Collections.Generic;
using System.Windows.Interop;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace SolutionColor
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(SolutionColorPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(Microsoft.VisualStudio.VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class SolutionColorPackage : AsyncPackage
    {
        public const string PackageGuidString = "8fa74b3d-8744-465c-b06e-a719e1d63ddf";

        public static readonly Guid ToolbarCommandSetGuid = new Guid("00d80876-3407-4666-bf62-7262028ea83b");

        public SolutionColorSettingStore Settings { get; private set; } = new SolutionColorSettingStore();


        /// <summary>
        /// Store process id, since we use this on a very regular basis (whenever any windows opens anywhere...) and we don't want to do GetCurrentProcess every time.
        /// </summary>
        private readonly int currentProcessId = System.Diagnostics.Process.GetCurrentProcess().Id;

        private Dictionary<Window, TitleBarColorController> windowTitleBarController = new Dictionary<Window, TitleBarColorController>();

        /// <summary>
        /// Currently scheduled call to UpdateTitleBarControllerList
        /// </summary>
        private Task scheduledUpdateControllerOperation = null;


        /// <summary>
        /// Listener to opened solutions. Sets title bar color settings in effect if any.
        /// </summary>
        private class SolutionOpenListener : SolutionListener
        {
            private SolutionColorPackage package;

            public SolutionOpenListener(SolutionColorPackage package) : base(package)
            {
                this.package = package;
            }

            /// <summary>
            /// IVsSolutionEvents3.OnAfterOpenSolution is called BEFORE a solution is fully loaded.
            /// This is different from EnvDTE.Events.SolutionEvents.Opened which is loaded after a resolution was loaded due to historical reasons:
            /// In earlier Visual Studio versions there was no asynchronous loading, so a solution was either fully loaded or not at all.
            /// </summary>
            public override int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                // Check if we already saved something for this solution.
                string solutionPath = VSUtils.GetCurrentSolutionPath();
                System.Drawing.Color color;
                if (package.Settings.GetSolutionColorSetting(solutionPath, out color))
                    package.SetTitleBarColor(color);
                else if (package.Settings.IsAutomaticColorPickEnabled())
                {
                    color = RandomColorGenerator.RandomColor.GetColor(RandomColorGenerator.ColorScheme.Random, VSUtils.IsUsingDarkTheme() ? RandomColorGenerator.Luminosity.Dark : RandomColorGenerator.Luminosity.Light);
                    package.SetTitleBarColor(color);
                    package.Settings.SaveOrOverwriteSolutionColor(solutionPath, color);
                }

                return 0;
            }

            public override int OnAfterCloseSolution(object pUnkReserved)
            {
                package.ResetTitleBarColor();
                return 0;
            }
        }

        private SolutionListener listener;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            await JoinableTaskFactory.SwitchToMainThreadAsync();    // Command constructors need to be on main thread.
            PickColorCommand.Initialize(this);
            ResetColorCommand.Initialize(this);
            EnableAutoPickColorCommand.Initialize(this);

            listener = new SolutionOpenListener(this);

            await UpdateTitleBarControllerListAsync();

            // Window events won't give us events for undocking so we can't use that.
            //var dte = VSUtils.GetDTE();
            //windowEvents = dte.Events.WindowEvents;
            //windowEvents.WindowCreated += (Window) => CheckForNewController();
            //windowEvents.WindowClosing += (Window) => CheckForNewController();

            // Instead we're using a bigger gun: The Automation framework!
            // Sadly, it is not enough to listen to child windows of VS since code window popups are direct children of the desktop in terms of UI.
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, AutomationElement.RootElement, TreeScope.Children, OnWindowOpenedClosed);
            // Weirdly, this doesn't apply to ALL child windows, and some are children of the main window after all. (Repro: Create a window out of two non-code views)
            var windowHandle = new WindowInteropHelper(Application.Current.MainWindow).Handle;
            var windowAutomationElement = AutomationElement.FromHandle(windowHandle);
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, windowAutomationElement, TreeScope.Subtree, OnWindowOpenedClosed);

            // Cant use TreeScope.Children on WindowClosedEvent.
            Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent, AutomationElement.RootElement, TreeScope.Subtree, OnWindowOpenedClosed);
        }

        protected override void Dispose(bool disposing)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (disposing)
                listener.Dispose();
            base.Dispose(disposing);
        }

        private void OnWindowOpenedClosed(object sender, AutomationEventArgs args)
        {
            // Ignore if different process.
            var element = sender as AutomationElement;
            if (element == null || element.Current.ProcessId != currentProcessId)
                return;

            // Since we're scheduling to the main thread which should be the only window creating thread, this should be thread safe.
            // (if not, it wouldn't be too tragic - as long as we don't spam the scheduler with more updates than necessary we're fine)
            if (scheduledUpdateControllerOperation == null || scheduledUpdateControllerOperation.IsCompleted)
            {
                scheduledUpdateControllerOperation = UpdateTitleBarControllerListAsync();
            }
        }

        private async Task UpdateTitleBarControllerListAsync()
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();

            Window[] windows = new Window[Application.Current.Windows.Count];
            Application.Current.Windows.CopyTo(windows, 0);

            // Destroy old ones.
            windowTitleBarController = windowTitleBarController.Where(x => windows.Contains(x.Key))
                                                               .ToDictionary(x=>x.Key, x=>x.Value);

            // Look for new ones to add.
            foreach (Window window in windows)
            {
                if (windowTitleBarController.ContainsKey(window))
                    continue;

                var newController = TitleBarColorController.CreateFromWindow(window);
                if (newController != null)
                {
                    windowTitleBarController.Add(window, newController);

                    // Check if we already saved something for this solution.
                    // Do this in here since we call UpdateTitleBarControllerList fairly regularly and in the most cases won't have any new controllers.
                    System.Drawing.Color color;
                    if (Settings.GetSolutionColorSetting(VSUtils.GetCurrentSolutionPath(), out color))
                        newController.SetTitleBarColor(color);
                }
            }
        }

        #region Color Manipulation

        public void SetTitleBarColor(System.Drawing.Color color)
        {
            foreach (var bar in windowTitleBarController)
                bar.Value.SetTitleBarColor(color);
        }

        public void ResetTitleBarColor()
        {
            foreach(var bar in windowTitleBarController)
                bar.Value.ResetTitleBarColor(); 
        }

        public System.Drawing.Color GetMainTitleBarColor()
        {
            TitleBarColorController titleBar;
            if (windowTitleBarController.TryGetValue(Application.Current.MainWindow, out titleBar))
                return titleBar.TryGetTitleBarColor();
            else
                return System.Drawing.Color.Black;
        }

        #endregion
    }
}
