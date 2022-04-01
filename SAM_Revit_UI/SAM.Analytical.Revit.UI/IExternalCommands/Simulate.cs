﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Analysis;
using Autodesk.Revit.UI;
using SAM.Analytical.Revit.UI.Properties;
using SAM.Core.Revit;
using SAM.Core.Revit.UI;
using SAM.Core.Tas;
using SAM.Weather;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

namespace SAM.Analytical.Revit.UI
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Simulate : PushButtonExternalCommand
    {
        public override string RibbonPanelName => "Tas";

        public override int Index => 17;

        public override BitmapSource BitmapSource => Core.Windows.Convert.ToBitmapSource(Resources.SAM_Simulate, 32, 32);

        public override string Text => "Simulate";

        public override string ToolTip => "Simulate";

        public override string AvailabilityClassName => null;

        public override Result Execute(ExternalCommandData externalCommandData, ref string message, ElementSet elementSet)
        {
            Document document = externalCommandData?.Application?.ActiveUIDocument?.Document;
            if (document == null)
            {
                return Result.Failed;
            }

            string path = document.PathName;
            if (string.IsNullOrWhiteSpace(path))
            {
                string name = document.Title;
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = "000000_SAM_AnalyticalModel";
                }

                using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                {
                    folderBrowserDialog.Description = "Select Directory";
                    folderBrowserDialog.ShowNewFolderButton = true;
                    if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
                    {
                        return Result.Cancelled;
                    }

                    path = System.IO.Path.Combine(folderBrowserDialog.SelectedPath, name + ".rvt");
                }

                if (string.IsNullOrWhiteSpace(path))
                {
                    return Result.Failed;
                }

                document.SaveAs(path);
            }

            string projectName = null;
            string outputDirectory = null;
            bool unmetHours = false;
            WeatherData weatherData = null;
            SolarCalculationMethod solarCalculationMethod = SolarCalculationMethod.None;

            using (Forms.SimulateForm simulateForm = new Forms.SimulateForm(System.IO.Path.GetFileNameWithoutExtension(path), System.IO.Path.GetDirectoryName(path)))
            {
                if(simulateForm.ShowDialog() != DialogResult.OK)
                {
                    return Result.Cancelled;
                }

                projectName = simulateForm.ProjectName;
                outputDirectory = simulateForm.OutputDirectory;
                unmetHours = simulateForm.UnmetHours;
                weatherData = simulateForm.WeatherData;
                solarCalculationMethod = simulateForm.SolarCalculationMethod;
            }

            if (weatherData == null)
            {
                return Result.Failed;
            }

            AnalyticalModel analyticalModel = null;

            string path_TBD = System.IO.Path.Combine(outputDirectory, projectName + ".tbd");

            bool simulate = false;

            Dictionary<Guid, ElementId> dictionary = new Dictionary<Guid, ElementId>();
            using (Core.Windows.SimpleProgressForm simpleProgressForm = new Core.Windows.SimpleProgressForm("Preparing Model", string.Empty, 6))
            {
                simpleProgressForm.Increment("Converting Model");
                using (Transaction transaction = new Transaction(document, "Convert Model"))
                {
                    transaction.Start();
                    analyticalModel = Convert.ToSAM_AnalyticalModel(document, new ConvertSettings(true, true, false));
                    List<Panel> panels = analyticalModel?.GetPanels();
                    if (panels != null)
                    {
                        foreach(Panel panel in panels)
                        {
                            EnergyAnalysisSurface energyAnalysisSurface = Core.Revit.Query.Element<EnergyAnalysisSurface>(document, panel);
                            HostObject hostObject = Core.Revit.Query.Element(document, energyAnalysisSurface?.CADObjectUniqueId, energyAnalysisSurface?.CADLinkUniqueId) as HostObject;
                            if(hostObject != null)
                            {
                                dictionary[panel.Guid] = hostObject.Id;
                            }

                            List<Aperture> apertures = panel.Apertures;
                            if(apertures != null)
                            {
                                foreach(Aperture aperture in apertures)
                                {
                                    EnergyAnalysisOpening energyAnalysisOpening = Core.Revit.Query.Element<EnergyAnalysisOpening>(document, aperture);
                                    FamilyInstance familyInstance = Core.Revit.Query.Element(energyAnalysisOpening) as FamilyInstance;
                                    if(familyInstance != null)
                                    {
                                        dictionary[aperture.Guid] = familyInstance.Id;
                                    }
                                }
                            }
                        }
                    }
                    
                    transaction.RollBack();
                }

                if (analyticalModel == null)
                {
                    MessageBox.Show("Could not convert to AnalyticalModel");
                    return Result.Failed;
                }

                if (System.IO.File.Exists(path_TBD))
                {
                    System.IO.File.Delete(path_TBD);
                }

                List<int> hoursOfYear = Analytical.Query.DefaultHoursOfYear();

                //Run Solar Calculation for cooling load

                simpleProgressForm.Increment("Solar Calculations");
                if(solarCalculationMethod != SolarCalculationMethod.None)
                {
                    SolarCalculator.Modify.Simulate(analyticalModel, hoursOfYear.ConvertAll(x => new DateTime(2018, 1, 1).AddHours(x)), Core.Tolerance.MacroDistance, Core.Tolerance.MacroDistance, 0.012, Core.Tolerance.Distance);
                }

                using (SAMTBDDocument sAMTBDDocument = new SAMTBDDocument(path_TBD))
                {
                    TBD.TBDDocument tBDDocument = sAMTBDDocument.TBDDocument;

                    simpleProgressForm.Increment("Updating WeatherData");
                    Weather.Tas.Modify.UpdateWeatherData(tBDDocument, weatherData);

                    TBD.Calendar calendar = tBDDocument.Building.GetCalendar();

                    List<TBD.dayType> dayTypes = Query.DayTypes(calendar);
                    if (dayTypes.Find(x => x.name == "HDD") == null)
                    {
                        TBD.dayType dayType = calendar.AddDayType();
                        dayType.name = "HDD";
                    }

                    if (dayTypes.Find(x => x.name == "CDD") == null)
                    {
                        TBD.dayType dayType = calendar.AddDayType();
                        dayType.name = "CDD";
                    }

                    simpleProgressForm.Increment("Converting to TBD");
                    Tas.Convert.ToTBD(analyticalModel, tBDDocument);

                    simpleProgressForm.Increment("Updating Zones");
                    Tas.Modify.UpdateZones(tBDDocument.Building, analyticalModel, true);

                    simpleProgressForm.Increment("Updating Shading");
                    simulate = Tas.Modify.UpdateShading(tBDDocument, analyticalModel);

                    sAMTBDDocument.Save();
                }
            }

            List<DesignDay> heatingDesignDays = new List<DesignDay>() { Analytical.Query.HeatingDesignDay(weatherData) };
            List<DesignDay> coolingDesignDays = new List<DesignDay>() { Analytical.Query.CoolingDesignDay(weatherData) };

            SurfaceOutputSpec surfaceOutputSpec = new SurfaceOutputSpec("Tas.Simulate")
            {
                SolarGain = true,
                Conduction = true,
                ApertureData = false,
                Condensation = false,
                Convection = false,
                LongWave = false,
                Temperature = false
            };

            List<SurfaceOutputSpec> surfaceOutputSpecs = new List<SurfaceOutputSpec>() { surfaceOutputSpec };

            analyticalModel = Tas.Modify.RunWorkflow(analyticalModel, path_TBD, null, null, heatingDesignDays, coolingDesignDays, surfaceOutputSpecs, unmetHours, simulate, false);

            List<Core.ISAMObject> results = null;

            AdjacencyCluster adjacencyCluster = null;
            if (analyticalModel != null)
            {
                adjacencyCluster = analyticalModel?.AdjacencyCluster;
                if (adjacencyCluster != null)
                {
                    results = new List<Core.ISAMObject>();
                    adjacencyCluster.GetObjects<SpaceSimulationResult>()?.ForEach(x => results.Add(x));
                    adjacencyCluster.GetObjects<ZoneSimulationResult>()?.ForEach(x => results.Add(x));
                    adjacencyCluster.GetObjects<AdjacencyClusterSimulationResult>()?.ForEach(x => results.Add(x));
                    adjacencyCluster.GetPanels()?.ForEach(x => results.Add(x));
                }
            }

            if (adjacencyCluster != null && results != null && results.Count != 0)
            {
                using (Core.Windows.SimpleProgressForm simpleProgressForm = new Core.Windows.SimpleProgressForm("Inserting Results", string.Empty, results.Count + 3))
                {
                    simpleProgressForm.Increment("Initialization");

                    ConvertSettings convertSettings = new ConvertSettings(false, true, false);

                    using (Transaction transaction = new Transaction(document, "Simulate"))
                    {
                        transaction.Start();

                        foreach (Core.ISAMObject sAMObject in results)
                        {
                            simpleProgressForm.Increment(sAMObject?.Name == null ? "???" : sAMObject.Name);

                            if (sAMObject is SpaceSimulationResult)
                            {
                                Convert.ToRevit(adjacencyCluster, (SpaceSimulationResult)sAMObject, document, convertSettings)?.Cast<Element>().ToList();
                            }
                            else if (sAMObject is ZoneSimulationResult)
                            {
                                Convert.ToRevit(adjacencyCluster, (ZoneSimulationResult)sAMObject, document, convertSettings)?.Cast<Element>().ToList();
                            }
                            else if (sAMObject is AdjacencyClusterSimulationResult)
                            {
                                Convert.ToRevit((AdjacencyClusterSimulationResult)sAMObject, document, convertSettings);
                            }
                            else if(sAMObject is Panel)
                            {
                                Panel panel = (Panel)sAMObject;

                                ElementId elementId = null;

                                if (dictionary.TryGetValue(panel.Guid, out elementId))
                                {
                                    Core.Revit.Modify.SetValues(document.GetElement(elementId), panel, ActiveSetting.Setting);
                                }

                                List<Aperture> apertures = panel.Apertures;
                                if(apertures != null)
                                {
                                    foreach(Aperture aperture in apertures)
                                    {
                                        if (dictionary.TryGetValue(aperture.Guid, out elementId))
                                        {
                                            Core.Revit.Modify.SetValues(document.GetElement(elementId), aperture, ActiveSetting.Setting);
                                        }
                                    }
                                }
                            }
                        }

                        simpleProgressForm.Increment("Coping Parameters");

                        Modify.CopySpatialElementParameters(document, Tool.TAS);

                        simpleProgressForm.Increment("Finishing");

                        transaction.Commit();
                    }
                }
            }

            return Result.Succeeded;
        }

        //public override Result Execute(ExternalCommandData externalCommandData, ref string message, ElementSet elementSet)
        //{
        //    Document document = externalCommandData?.Application?.ActiveUIDocument?.Document;
        //    if (document == null)
        //    {
        //        return Result.Failed;
        //    }

        //    WeatherData weatherData = null;
        //    using (OpenFileDialog openFileDialog = new OpenFileDialog())
        //    {
        //        openFileDialog.Filter = "epw files (*.epw)|*.epw|TAS TBD files (*.tbd)|*.tbd|TAS TSD files (*.tsd)|*.tsd|TAS TWD files (*.twd)|*.twd|All files (*.*)|*.*";
        //        openFileDialog.FilterIndex = 1;
        //        openFileDialog.RestoreDirectory = true;

        //        if (openFileDialog.ShowDialog() != DialogResult.OK)
        //        {
        //            return Result.Cancelled;
        //        }

        //        string path_WeatherData = openFileDialog.FileName;
        //        string extension = System.IO.Path.GetExtension(path_WeatherData).ToLower().Trim();
        //        if (string.IsNullOrWhiteSpace(extension))
        //        {
        //            return Result.Failed;
        //        }

        //        try
        //        {
        //            if (extension.EndsWith("epw"))
        //            {
        //                weatherData = Weather.Convert.ToSAM(path_WeatherData);
        //            }
        //            else
        //            {
        //                List<WeatherData> weatherDatas = Weather.Tas.Convert.ToSAM_WeatherDatas(path_WeatherData);
        //                if (weatherDatas == null || weatherDatas.Count == 0)
        //                {
        //                    return Result.Failed;
        //                }

        //                if (weatherDatas.Count == 1)
        //                {
        //                    weatherData = weatherDatas[0];
        //                }
        //                else
        //                {
        //                    weatherDatas.Sort((x, y) => x.Name.CompareTo(y.Name));

        //                    using (Core.Windows.Forms.ComboBoxForm<WeatherData> comboBoxForm = new Core.Windows.Forms.ComboBoxForm<WeatherData>("Select Weather Data", weatherDatas, (WeatherData x) => x.Name))
        //                    {
        //                        if (comboBoxForm.ShowDialog() != DialogResult.OK)
        //                        {
        //                            return Result.Cancelled;
        //                        }

        //                        weatherData = comboBoxForm.SelectedItem;
        //                    }

        //                }

        //            }
        //        }
        //        catch (Exception exception)
        //        {
        //            weatherData = null;
        //        }
        //    }

        //    if (weatherData == null)
        //    {
        //        return Result.Failed;
        //    }

        //    string path = document.PathName;
        //    if (string.IsNullOrWhiteSpace(path))
        //    {
        //        string name = document.Title;
        //        if (string.IsNullOrWhiteSpace(name))
        //        {
        //            name = "000000_SAM_AnalyticalModel";
        //        }

        //        using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
        //        {
        //            folderBrowserDialog.Description = "Select Directory";
        //            folderBrowserDialog.ShowNewFolderButton = true;
        //            if (folderBrowserDialog.ShowDialog() != DialogResult.OK)
        //            {
        //                return Result.Cancelled;
        //            }

        //            path = System.IO.Path.Combine(folderBrowserDialog.SelectedPath, name + ".rvt");
        //        }

        //        if (string.IsNullOrWhiteSpace(path))
        //        {
        //            return Result.Failed;
        //        }

        //        document.SaveAs(path);
        //    }

        //    AnalyticalModel analyticalModel = null;

        //    using (Transaction transaction = new Transaction(document, "Convert Model"))
        //    {
        //        transaction.Start();
        //        analyticalModel = Convert.ToSAM_AnalyticalModel(document, new ConvertSettings(true, true, false));
        //        transaction.RollBack();
        //    }

        //    if (analyticalModel == null)
        //    {
        //        return Result.Failed;
        //    }

        //    string path_gbXML = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), System.IO.Path.GetFileNameWithoutExtension(path));
        //    if (System.IO.File.Exists(path_gbXML))
        //    {
        //        System.IO.File.Delete(path_gbXML);
        //    }

        //    bool exported = false;


        //    exported = document.TogbXML(path_gbXML);
        //    if (!exported)
        //    {
        //        return Result.Failed;
        //    }

        //    path_gbXML = path_gbXML + ".xml";

        //    string path_T3D = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), System.IO.Path.GetFileNameWithoutExtension(path) + ".t3d");

        //    exported = Core.Tas.Convert.ToT3D(path_T3D, path_gbXML, true, true, true, false);
        //    if (!exported)
        //    {
        //        return Result.Failed;
        //    }

        //    analyticalModel = Tas.Query.UpdateT3D(analyticalModel, path_T3D);

        //    string path_TBD = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), System.IO.Path.GetFileNameWithoutExtension(path) + ".tbd");

        //    using (SAMTBDDocument sAMTBDDocument = new SAMTBDDocument(path_TBD))
        //    {
        //        TBD.TBDDocument tBDDocument = sAMTBDDocument.TBDDocument;

        //        Weather.Tas.Modify.UpdateWeatherData(tBDDocument, weatherData);

        //        double latitude_TBD = Core.Query.Round(analyticalModel.Location.Latitude, 0.01);
        //        double longitude_TBD = Core.Query.Round(analyticalModel.Location.Longitude, 0.01);

        //        double latitude_WeatherData = Core.Query.Round(weatherData.Latitude, 0.01);
        //        double longitude_WeatherDate = Core.Query.Round(weatherData.Longitude, 0.01);

        //        //if (Math.Abs(latitude_TBD - latitude_WeatherData) > 0.01 || Math.Abs(longitude_TBD - longitude_WeatherDate) > 0.01)
        //        //{
        //        //    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "WeatherData Longitude or Latitude mismatch");
        //        //}

        //        TBD.Calendar calendar = tBDDocument.Building.GetCalendar();

        //        TBD.dayType dayType = null;

        //        dayType = calendar.AddDayType();
        //        dayType.name = "HDD";

        //        dayType = calendar.AddDayType();
        //        dayType.name = "CDD";

        //        sAMTBDDocument.Save();
        //    }

        //    exported = Tas.Convert.ToTBD(path_T3D, path_TBD, 1, 365, 15, true);
        //    if (!exported)
        //    {
        //        return Result.Failed;
        //    }

        //    AdjacencyCluster adjacencyCluster = null;

        //    using (SAMTBDDocument sAMTBDDocument = new SAMTBDDocument(path_TBD))
        //    {
        //        Tas.Query.UpdateFacingExternal(analyticalModel, sAMTBDDocument);

        //        Tas.Modify.AssignAdiabaticConstruction(sAMTBDDocument.TBDDocument.Building, "Adiabatic", new string[] { "-unzoned", "-internal", "-exposed" }, false, true);

        //        Tas.Modify.UpdateBuildingElements(sAMTBDDocument, analyticalModel);

        //        adjacencyCluster = analyticalModel.AdjacencyCluster;
        //        Tas.Modify.UpdateThermalParameters(adjacencyCluster, sAMTBDDocument.TBDDocument?.Building);
        //        analyticalModel = new AnalyticalModel(analyticalModel, adjacencyCluster);

        //        Tas.Modify.UpdateZones(analyticalModel, sAMTBDDocument, true);

        //        //Update.DesignDays Missing!!

        //        sAMTBDDocument.Save();
        //    }

        //    Tas.Query.Sizing(path_TBD, analyticalModel, false, true);

        //    analyticalModel = Tas.Modify.UpdateDesignLoads(path_TBD, analyticalModel);

        //    bool hasWeatherData = false;
        //    using (SAMTBDDocument sAMTBDDocument = new SAMTBDDocument(path_TBD))
        //    {
        //        TBD.TBDDocument tBDDocument = sAMTBDDocument.TBDDocument;

        //        hasWeatherData = tBDDocument?.Building.GetWeatherYear() != null;
        //    }

        //    if (!hasWeatherData)
        //    {
        //        MessageBox.Show("Could not complete simulation. TBD file has no Weather Data");
        //        return Result.Failed;
        //    }

        //    string path_TSD = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), System.IO.Path.GetFileNameWithoutExtension(path) + ".tsd");

        //    bool result = Tas.Modify.Simulate(path_TBD, path_TSD, 1, 365);



        //    adjacencyCluster = analyticalModel.AdjacencyCluster;
        //    Tas.Modify.AddResults(path_TSD, adjacencyCluster);
        //    analyticalModel = new AnalyticalModel(analyticalModel, adjacencyCluster);

        //    analyticalModel.ToRevit(document, new ConvertSettings(true, true, false));

        //    return Result.Succeeded;
        //}
    }
}