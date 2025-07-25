﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SAM.Analytical.Revit.UI.Properties;
using SAM.Core.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;

namespace SAM.Analytical.Revit.UI
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DeleteViews : PushButtonExternalCommand
    {
        public override string RibbonPanelName => "Project Setup";

        public override int Index => 14;

        public override BitmapSource BitmapSource => Core.Windows.Convert.ToBitmapSource(Resources.SAM_Small);

        public override string Text => "Delete\nViews";

        public override string ToolTip => "Delete Views";

        public override string AvailabilityClassName => null;

        public override void Execute()
        {
            Document document = Document;
            if (document == null)
            {
                return;
            }

            List<View> views = new FilteredElementCollector(document).OfClass(typeof(View)).Cast<View>().ToList();
            if (views == null || views.Count == 0)
            {
                return;
            }

            for (int i = views.Count - 1; i >= 0; i--)
            {
                View view = views[i];
                if (view == null)
                {
                    views.RemoveAt(i);
                    continue;
                }

                if (view.ViewType != ViewType.FloorPlan)
                {
                    views.RemoveAt(i);
                    continue;
                }

                if (!view.IsTemplate)
                {
                    views.RemoveAt(i);
                    continue;
                }

            }

            List<string> templateNames = new List<string>() { "00_Reference" };

            using (Core.Windows.Forms.TreeViewForm<View> treeViewForm = new Core.Windows.Forms.TreeViewForm<View>("Select Templates", views, (View view) => view.Name, null, (View view) => !templateNames.Contains(view.Name)))
            {
                if (treeViewForm.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    return;
                }

                templateNames = treeViewForm.SelectedItems?.ConvertAll(x => x.Name);
            }

            if (templateNames == null || templateNames.Count == 0)
            {
                return;
            }

            using (Transaction transaction = new Transaction(document, "Delete Views"))
            {
                transaction.Start();

                foreach (string templateName in templateNames)
                {
                    List<ElementId> elementIds = Core.Revit.Modify.DeleteViews(document, templateName, false, new ViewType[] { ViewType.FloorPlan });
                }

                transaction.Commit();
            }
        }
    }
}
