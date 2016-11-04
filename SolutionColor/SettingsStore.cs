using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
using System.IO;

namespace SolutionColor
{
    public class SolutionColorSettingStore
    {
        private const string CollectionName = "SolutionColorSettings";

        public SolutionColorSettingStore()
        {
        }

        public void SaveOrOverwriteSolutionColor(string solutionPath, System.Drawing.Color color)
        {
            if (string.IsNullOrEmpty(solutionPath)) return;
            solutionPath = Path.GetFullPath(solutionPath);

            byte[] colorBytes = { color.A, color.R, color.G, color.B };

            var settingsStore = GetSettingsStore();
            settingsStore.SetMemoryStream(CollectionName, solutionPath, new MemoryStream(colorBytes));
        }

        public void RemoveSolutionColorSetting(string solutionPath)
        {
            if (string.IsNullOrEmpty(solutionPath)) return;
            solutionPath = Path.GetFullPath(solutionPath);

            var settingsStore = GetSettingsStore();
            if (settingsStore.PropertyExists(CollectionName, solutionPath))
                settingsStore.DeleteProperty(CollectionName, solutionPath);
        }

        public bool GetSolutionColorSetting(string solutionPath, out System.Drawing.Color color)
        {
            color = System.Drawing.Color.Black;

            if (string.IsNullOrEmpty(solutionPath)) return false;
            solutionPath = Path.GetFullPath(solutionPath);

            var settingsStore = GetSettingsStore();
            if (settingsStore.PropertyExists(CollectionName, solutionPath))
            {
                MemoryStream colorBytes = settingsStore.GetMemoryStream(CollectionName, solutionPath);
                color = System.Drawing.Color.FromArgb(colorBytes.ReadByte(), colorBytes.ReadByte(), colorBytes.ReadByte(), colorBytes.ReadByte());
                return true;
            }
            else
                return false;
        }

        private WritableSettingsStore GetSettingsStore()
        {
            var settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
            WritableSettingsStore settingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

            // Ensure our settings collection exists.
            if (!settingsStore.CollectionExists(CollectionName))
                settingsStore.CreateCollection(CollectionName);

            return settingsStore;
        }
    }
}
