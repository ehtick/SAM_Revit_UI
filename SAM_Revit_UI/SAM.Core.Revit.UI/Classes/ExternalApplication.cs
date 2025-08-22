using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace SAM.Core.Revit.UI
{
    public class ExternalApplication : IExternalApplication
    {
        public static string TabName { get; } = "SAM";

        public static WindowHandle WindowHandle { get; } = new WindowHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);

        public static string GetAssemblyPath()
        {
            return Assembly.GetExecutingAssembly()?.Location;
        }

        public string GetAssemblyDirectory()
        {
            return System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly()?.Location);
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            application.CreateRibbonTab(TabName);

            List<Assembly> assemblies = new List<Assembly>();

            string directory = GetAssemblyDirectory();
            if (!string.IsNullOrWhiteSpace(directory) && System.IO.Directory.Exists(directory))
            {
                string[] paths = System.IO.Directory.GetFiles(directory, "*.Revit.UI.dll");
                if (paths != null)
                {
                    foreach (string path in paths)
                    {
                        Assembly assembly = Assembly.LoadFrom(path);
                        if (assembly == null)
                        {
                            continue;
                        }

                        assemblies.Add(assembly);
                    }
                }
            }

            List<ISAMRibbonItemData> sAMRibbonItemDatas = new List<ISAMRibbonItemData>();
            foreach (Assembly assembly in assemblies)
            {
                Type[] types = assembly?.GetExportedTypes();
                if (types == null)
                {
                    continue;
                }

                foreach (Type type in types)
                {
                    if (!typeof(ISAMRibbonItemData).IsAssignableFrom(type) || type.IsAbstract)
                    {
                        continue;
                    }

                    ISAMRibbonItemData sAMRibbonItemData = Activator.CreateInstance(type) as ISAMRibbonItemData;
                    if (sAMRibbonItemData == null)
                    {
                        continue;
                    }

                    sAMRibbonItemDatas.Add(sAMRibbonItemData);
                }
            }

            sAMRibbonItemDatas.Sort((x, y) => x.Index.CompareTo(y.Index));

            foreach (ISAMRibbonItemData sAMRibbonItemData in sAMRibbonItemDatas)
            {
                string ribbonPanelName = sAMRibbonItemData.RibbonPanelName;
                if (string.IsNullOrWhiteSpace(ribbonPanelName))
                {
                    ribbonPanelName = "General";
                }

                RibbonPanel ribbonPanel = application.GetRibbonPanels(TabName)?.Find(x => x.Name == ribbonPanelName);
                if (ribbonPanel == null)
                {
                    ribbonPanel = application.CreateRibbonPanel(TabName, ribbonPanelName);
                }

                sAMRibbonItemData.Create(ribbonPanel);
            }

            return Result.Succeeded;
        }
    }
}
