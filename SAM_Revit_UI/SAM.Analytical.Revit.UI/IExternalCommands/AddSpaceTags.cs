﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SAM.Analytical.Revit.UI.Properties;
using SAM.Core.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

namespace SAM.Analytical.Revit.UI
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AddSpaceTags : PushButtonExternalCommand
    {
        public override string RibbonPanelName => "Project Setup";

        public override int Index => 12;

        public override BitmapSource BitmapSource => Core.Windows.Convert.ToBitmapSource(Resources.SAM_Small);

        public override string Text => "Add\nSpace Tags";

        public override string ToolTip => "Add Space Tags";

        public override string AvailabilityClassName => null;

        public override Result Execute(ExternalCommandData externalCommandData, ref string message, ElementSet elementSet)
        {
            Document document = externalCommandData?.Application?.ActiveUIDocument?.Document;
            if (document == null)
            {
                return Result.Failed;
            }

            List<Autodesk.Revit.DB.Mechanical.SpaceTagType> spaceTagTypes = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_MEPSpaceTags).WhereElementIsElementType().Cast<Autodesk.Revit.DB.Mechanical.SpaceTagType>().ToList();
            if(spaceTagTypes == null || spaceTagTypes.Count == 0)
            {
                return Result.Failed;
            }

            List<Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>>> tuples = new List<Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>>>();
            tuples.Add(new Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>>(spaceTagTypes.Find(x => x.Id.IntegerValue == 1107209), new List<string>() { "RiserCLG", "RiserHTG", "RiserVNT" }));
            tuples.Add(new Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>>(spaceTagTypes.Find(x => x.Id.IntegerValue == 1107211), new List<string>() { "ICType"}));
            tuples.Add(new Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>>(spaceTagTypes.Find(x => x.Id.IntegerValue == 1107215), new List<string>() { "RefExhaust", "RefSupply" }));
            tuples.Add(new Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>>(spaceTagTypes.Find(x => x.Id.IntegerValue == 1107203), new List<string>() { "SAM Model" }));
            tuples.Add(new Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>>(spaceTagTypes.Find(x => x.Id.IntegerValue == 1107213), new List<string>() { "NoPeople" }));
            tuples.Add(new Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>>(spaceTagTypes.Find(x => x.Id.IntegerValue == 1107219), new List<string>() { "Heating Load" }));
            tuples.Add(new Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>>(spaceTagTypes.Find(x => x.Id.IntegerValue == 1107217), new List<string>() { "Cooling Load" }));

            using (Transaction transaction = new Transaction(document, "Add Space Tags"))
            {
                transaction.Start();

                foreach(Tuple<Autodesk.Revit.DB.Mechanical.SpaceTagType, List<string>> tuple in tuples)
                {
                    if(tuple.Item1 == null || tuple.Item2 == null || tuple.Item2.Count == 0)
                    {
                        continue;
                    }

                    List<string> templateNames = tuple.Item2;
                    templateNames.RemoveAll(x => string.IsNullOrWhiteSpace(x));
                    if(templateNames == null || templateNames.Count == 0)
                    {
                        continue;
                    }

                    List<Autodesk.Revit.DB.Mechanical.SpaceTag> spaceTags = Core.Revit.Modify.TagSpaces(document, templateNames, tuple.Item1.Id, new ViewType[] { ViewType.FloorPlan });

                }

                transaction.Commit();
            }

                return Result.Succeeded;
        }
    }
}
