using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SAM.Core.Revit.UI
{
    public abstract class ExternalCommand : IExternalCommand
    {
        public abstract Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements);
    }
}
