﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using SAM.Core.Revit.UI.Properties;
using System.Windows.Media.Imaging;

namespace SAM.Core.Revit.UI
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Wiki : PushButtonExternalCommand
    {
        public override string RibbonPanelName => "General";

        public override int Index => 0;

        public override BitmapSource BitmapSource => Windows.Convert.ToBitmapSource(Resources.SAM_Small);

        public override string Text => "Info";

        public override string ToolTip => "Info";

        public override string AvailabilityClassName => typeof(AlwaysAvailableExternalCommandAvailability).FullName;

        public override void Execute()
        {
            Query.StartProcess("https://github.com/HoareLea/SAM/wiki/00-Home");
        }
    }
}
