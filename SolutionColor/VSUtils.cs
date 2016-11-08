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
    }
}
