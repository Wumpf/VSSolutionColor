using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

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
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(SolutionColorPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(Microsoft.VisualStudio.VSConstants.UICONTEXT.NoSolution_string)]
    public sealed class SolutionColorPackage : Package
    {
        public const string PackageGuidString = "8fa74b3d-8744-465c-b06e-a719e1d63ddf";

        public static readonly Guid ToolbarCommandSetGuid = new Guid("00d80876-3407-4666-bf62-7262028ea83b");

        public SolutionColorSettingStore Settings { get; private set; } = new SolutionColorSettingStore();

        public TitleBarColorController TitleBarColorControl { get; private set; }


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
                // Check if we already saved something for this solution.
                System.Drawing.Color color;
                if (package.Settings.GetSolutionColorSetting(VSUtils.GetCurrentSolutionPath(), out color))
                    package.TitleBarColorControl.SetTitleBarColor(color);
                return 0;
            }
        }

        private SolutionListener listener;
        

        /// <summary>
        /// Initializes a new instance of the <see cref="PickColorCommand"/> class.
        /// </summary>
        public SolutionColorPackage()
        {
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            TitleBarColorControl = new TitleBarColorController();
            PickColorCommand.Initialize(this);
            ResetColorCommand.Initialize(this);

            listener = new SolutionOpenListener(this);

            base.Initialize();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                listener.Dispose();
            base.Dispose(disposing);
        }

        #endregion
    }
}
