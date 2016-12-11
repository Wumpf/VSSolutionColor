using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace SolutionColor
{
    /// <summary>
    /// Abstract listener for solution events.
    /// This is much more powerful than EnvDTE.Events.SolutionEvents!
    /// 
    /// More events, intenionally not implemented: IVsSolutionEvents4, IVsSolutionEvents, IVsSolutionEvents2
    /// </summary>
    abstract public class SolutionListener : IVsSolutionEvents3, IDisposable
    {
        private IVsSolution solutionService;
        private uint eventsCookie = (uint)Constants.VSCOOKIE_NIL;
        private bool isDisposed;

        protected SolutionListener(IServiceProvider serviceProvider)
        {
            solutionService = serviceProvider.GetService(typeof(SVsSolution)) as IVsSolution;
            if (solutionService == null)
            {
                throw new InvalidOperationException();
            }

            ErrorHandler.ThrowOnFailure(solutionService.AdviseSolutionEvents(this, out eventsCookie));
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                if (disposing && solutionService != null && eventsCookie != (uint)Constants.VSCOOKIE_NIL)
                {
                    ErrorHandler.ThrowOnFailure(solutionService.UnadviseSolutionEvents(eventsCookie));
                    eventsCookie = (uint)Constants.VSCOOKIE_NIL;
                }
                isDisposed = true;
            }
        }

        #region Event Impls

        public virtual int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            return 0;
        }

        public virtual int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return 0;
        }

        public virtual int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            return 0;
        }

        public virtual int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return 0;
        }

        public virtual int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            return 0;
        }

        public virtual int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return 0;
        }

        public virtual int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            return 0;
        }

        public virtual int OnBeforeCloseSolution(object pUnkReserved)
        {
            return 0;
        }

        public virtual int OnAfterCloseSolution(object pUnkReserved)
        {
            return 0;
        }

        public virtual int OnAfterMergeSolution(object pUnkReserved)
        {
            return 0;
        }

        public virtual int OnBeforeOpeningChildren(IVsHierarchy pHierarchy)
        {
            return 0;
        }

        public virtual int OnAfterOpeningChildren(IVsHierarchy pHierarchy)
        {
            return 0;
        }

        public virtual int OnBeforeClosingChildren(IVsHierarchy pHierarchy)
        {
            return 0;
        }

        public virtual int OnAfterClosingChildren(IVsHierarchy pHierarchy)
        {
            return 0;
        }

        public virtual int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            return 0;
        }

        #endregion
    }
}
