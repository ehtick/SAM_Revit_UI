﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SAM.Core.Revit.UI;
using SAM.Analytical.Revit.UI.Properties;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using System;
using NetOffice.ExcelApi;
using System.Collections.Generic;
using System.Linq;
using SAM.Core.Revit;

namespace SAM.Analytical.Revit.UI
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class GenerateSharedParametersFile : PushButtonExternalCommand
    {
        public override string RibbonPanelName => "Shared Parameters";

        public override int Index => 4;

        public override BitmapSource BitmapSource => Core.Windows.Convert.ToBitmapSource(Resources.SAM_GenerateSharedParametersFile, 32, 32);

        public override string Text => "Generate";

        public override string ToolTip => "Generate Shared Parameters File";

        public override string AvailabilityClassName => typeof(AlwaysAvailableExternalCommandAvailability).FullName;

        public override void Execute()
        {
            Autodesk.Revit.ApplicationServices.Application application = ExternalCommandData?.Application?.Application;
            if (application == null)
            {
                return;
            }

            string path_Excel = null;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Workbook|*.xlsm;*.xlsx";
                openFileDialog.Title = "Select Excel file";
                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                path_Excel = openFileDialog.FileName;
            }

            if (string.IsNullOrEmpty(path_Excel))
            {
                return;
            }

            string path_SharedParametersFile = null;

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Text file|*.txt;*.txt";
                saveFileDialog.Title = "Select Shared Parameter file";
                if (saveFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                path_SharedParametersFile = saveFileDialog.FileName;
            }

            if (string.IsNullOrEmpty(path_SharedParametersFile))
            {
                return;
            }

            System.IO.File.WriteAllText(path_SharedParametersFile, string.Empty);

            Func<Worksheet, bool> func = new Func<Worksheet, bool>((Worksheet worksheet) =>
            {
                if (worksheet == null)
                {
                    return false;
                }

                int index_Group = 2;
                int index_Guid = 1;
                int index_ParameterType = 8;
                int index_Name = 7;

                List<int> indexes = new List<int>() { index_Group, index_Guid, index_ParameterType, index_Name };

                object[,] objects = worksheet.Range(worksheet.Cells[5, 1], worksheet.Cells[worksheet.UsedRange.Rows.Count, indexes.Max()]).Value as object[,];

                object[,] guids = worksheet.Range(worksheet.Cells[5, 1], worksheet.Cells[worksheet.UsedRange.Rows.Count, 1]).Value as object[,];

                if (objects == null || objects.GetLength(0) <= 1 || objects.GetLength(1) < indexes.Max())
                {
                    return false;
                }

                using (SharedParameterFileWrapper sharedParameterFileWrapper = new SharedParameterFileWrapper(application))
                {
                    sharedParameterFileWrapper.Open(path_SharedParametersFile);

                    List<string> names = new List<string>();
                    using (Core.Windows.Forms.ProgressForm progressForm = new Core.Windows.Forms.ProgressForm("Creating Shared Parameters", objects.GetLength(0)))
                    {
                        for (int i = 1; i <= objects.GetLength(0); i++)
                        {
                            if (string.IsNullOrEmpty(objects[i, index_Name] as string) || string.IsNullOrEmpty(objects[i, index_ParameterType] as string))
                            {
                                progressForm.Update("???");
                                continue;
                            }

                            string name = objects[i, index_Name] as string;
                            if (string.IsNullOrEmpty(name))
                                progressForm.Update("???");
                            else
                                progressForm.Update(name);

                            string parameterTypeString = objects[i, index_ParameterType] as string;
                            parameterTypeString = parameterTypeString.Replace(" ", string.Empty);

#if Revit2017 || Revit2018 || Revit2019 || Revit2020 || Revit2021 || Revit2022
                                ParameterType parameterType = ParameterType.Invalid;
                                if (Enum.TryParse(parameterTypeString, out parameterType))
#else
                            ForgeTypeId forgeTypeId = Core.Revit.Query.ForgeTypeId(parameterTypeString);
                            if (forgeTypeId != null)
#endif
                            {
                                name = name.Trim();

                                if (names.IndexOf(name) < 0)
                                {
                                    names.Add(name);

#if Revit2017 || Revit2018 || Revit2019 || Revit2020 || Revit2021 || Revit2022
                                    ExternalDefinitionCreationOptions externalDefinitionCreationOptions = new ExternalDefinitionCreationOptions(name, parameterType);
#else
                                    ExternalDefinitionCreationOptions externalDefinitionCreationOptions = new ExternalDefinitionCreationOptions(name, forgeTypeId);
#endif
                                    string guid_String = objects[i, index_Guid] as string;
                                    if (!string.IsNullOrEmpty(guid_String))
                                    {
                                        Guid aGuid;
                                        if (Guid.TryParse(guid_String, out aGuid))
                                            externalDefinitionCreationOptions.GUID = aGuid;
                                    }
                                    else
                                    {
                                        externalDefinitionCreationOptions.GUID = Guid.NewGuid();
                                    }

                                    string group = objects[i, index_Group] as string;
                                    if (string.IsNullOrEmpty(group))
                                        group = "General";

                                    ExternalDefinition externalDefinition = sharedParameterFileWrapper.Create(group, externalDefinitionCreationOptions) as ExternalDefinition;
                                    if (externalDefinition != null)
                                    {
                                        guids[i, 1] = externalDefinition.GUID.ToString();
                                    }

                                }
                            }
                        }
                    }

                    sharedParameterFileWrapper.Close();
                }

                return true;

            });

            Core.Excel.Modify.Edit(path_Excel, "Live", func);

        }
    }
}
