using Microsoft.VisualStudio.Shell;

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
    }
}
