﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using SAM.Core.Revit.UI;
using SAM.Analytical.Revit.UI.Properties;
using System.Windows.Media.Imaging;
using System.Windows.Forms;
using System.Collections.Generic;
using System;
using System.Linq;

namespace SAM.Analytical.Revit.UI
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AddParameters : PushButtonExternalCommand
    {
        public override string RibbonPanelName => "Shared Parameters";

        public override int Index => 5;

        public override BitmapSource BitmapSource => Core.Windows.Convert.ToBitmapSource(Resources.SAM_AddParameters, 32, 32);

        public override string Text => "Add\nParameters";

        public override string ToolTip => "Add Parameters";

        public override string AvailabilityClassName => null;

        public override void Execute()
        {
            Document document = Document;
            if (document == null)
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

            if (string.IsNullOrWhiteSpace(path_Excel))
            {
                return;
            }

            object[,] objects = Core.Excel.Query.Values(path_Excel, "Live");
            if (objects == null || objects.GetLength(0) <= 1 || objects.GetLength(1) < 11)
            {
                return;
            }

            int index_Group = 2;
            int index_Guid = 1;
            int index_ParameterType = 8;
            int index_Name = 7;

            string[] unselected = new string[] { "DetailItem_AHU", "Space_Security", "Construction_CFD", "Space_LightingElec", "Space_DHW", "Space_Electrical", "Plant_Electrical", "DetailItem_Emitter", "Space_FireAlarm", "Construction_Detail", "Space_Data", "DetailItem_Benchmark", "DetailItem_ICData", "DetailItem_MEPInput", "DetailItem_Profiles", "DetailItem_Material", "Architect_Required" };
            List<string> names_Selected = Query.ParameterNames(objects, index_Group, index_Name, unselected);
            if (names_Selected == null || names_Selected.Count == 0)
            {
                return;
            }

            string path_SharedParametersFile = ExternalCommandData.Application.Application.SharedParametersFilename;

            string path_SharedParametersFile_Temp = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllText(path_SharedParametersFile_Temp, string.Empty);
            ExternalCommandData.Application.Application.SharedParametersFilename = path_SharedParametersFile_Temp;

            using (Transaction transaction = new Transaction(document, "Add Parameters"))
            {
                transaction.Start();

                using (Core.Revit.SharedParameterFileWrapper sharedParameterFileWrapper = new Core.Revit.SharedParameterFileWrapper(ExternalCommandData.Application.Application))
                {
                    sharedParameterFileWrapper.Open();

                    BindingMap bindingMap = document.ParameterBindings;
                    List<string> names = new List<string>();

                    using (Core.Windows.Forms.ProgressForm progressForm = new Core.Windows.Forms.ProgressForm("Creating Shared Parameters", objects.GetLength(0)))
                    {
                        for (int i = 1; i <= objects.GetLength(0); i++)
                        {
                            if (!string.IsNullOrEmpty(objects[i, index_Name] as string) && !string.IsNullOrEmpty(objects[i, index_ParameterType] as string))
                            {
                                string name = objects[i, index_Name] as string;
                                if (string.IsNullOrEmpty(name))
                                    progressForm.Update("???");
                                else
                                    progressForm.Update(name);

                                if (!names_Selected.Contains(name))
                                {
                                    continue;
                                }

                                string parameterTypeString = objects[i, index_ParameterType] as string;
                                parameterTypeString = parameterTypeString.Replace(" ", string.Empty);

#if Revit2020 || Revit2021 || Revit2022
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

                                        Definition definition = sharedParameterFileWrapper.Find(name);
                                        if (definition == null)
                                        {

#if Revit2017 || Revit2018 || Revit2019 || Revit2020 || Revit2021 || Revit2022
                                            ExternalDefinitionCreationOptions externalDefinitionCreationOptions = new ExternalDefinitionCreationOptions(name, parameterType);
#else
                                            ExternalDefinitionCreationOptions externalDefinitionCreationOptions = new ExternalDefinitionCreationOptions(name, forgeTypeId);
#endif

                                            string guid_String = objects[i, index_Guid] as string;
                                            if (!string.IsNullOrEmpty(guid_String))
                                            {
                                                Guid guid;
                                                if (Guid.TryParse(guid_String, out guid))
                                                {
                                                    externalDefinitionCreationOptions.GUID = guid;
                                                }
                                            }

                                            string parameterGroup = objects[i, index_Group] as string;
                                            if (!string.IsNullOrWhiteSpace(parameterGroup))
                                            {
                                                definition = sharedParameterFileWrapper.Create(parameterGroup, externalDefinitionCreationOptions);
                                            }
                                        }

                                        if (definition != null)
                                        {
                                            if (objects[i, 13] is string && objects[i, 12] is string)
                                            {
                                                string group = objects[i, 12] as string;

#if Revit2017 || Revit2018 || Revit2019 || Revit2020 || Revit2021 || Revit2022 || Revit2023 || Revit2024
                                                BuiltInParameterGroup builtInParameterGroup;
                                                if (Enum.TryParse("PG_" + group, out builtInParameterGroup))
                                                {
                                                    CategorySet categorySet = new CategorySet();

                                                    string[] categoryNames = (objects[i, 13] as string).Split(',');
                                                    foreach (string categoryName in categoryNames)
                                                    {
                                                        if (string.IsNullOrEmpty(categoryName))
                                                            continue;

                                                        BuiltInCategory builtInCategory;
                                                        if (Enum.TryParse("OST_" + categoryName.Trim().Replace(" ", string.Empty), out builtInCategory))
                                                        {
                                                            Category category = document.Settings.Categories.Cast<Category>().ToList().Find(x => x.Id.IntegerValue == (int)builtInCategory);
                                                            if (category != null)
                                                                categorySet.Insert(category);
                                                        }
                                                    }

                                                    if (categorySet.Size > 0)
                                                    {
                                                        string instance = objects[i, 14] as string;

                                                        if (string.IsNullOrEmpty(instance))
                                                            continue;

                                                        Autodesk.Revit.DB.Binding binding = null;
                                                        if (instance != null && instance.Trim().ToUpper() == "INSTANCE")
                                                            binding = ExternalCommandData.Application.Application.Create.NewInstanceBinding(categorySet);
                                                        else
                                                            binding = ExternalCommandData.Application.Application.Create.NewTypeBinding(categorySet);

                                                        bindingMap.Insert(definition, binding, builtInParameterGroup);
                                                    }
                                                }
#else
                                                ForgeTypeId groupTypeId = Revit.Query.GroupTypeId(group);
                                                if (groupTypeId != null)
                                                {
                                                    CategorySet categorySet = new CategorySet();

                                                    string[] categoryNames = (objects[i, 13] as string).Split(',');
                                                    foreach (string categoryName in categoryNames)
                                                    {
                                                        if (string.IsNullOrEmpty(categoryName))
                                                            continue;

                                                        BuiltInCategory builtInCategory;
                                                        if (Enum.TryParse("OST_" + categoryName.Trim().Replace(" ", string.Empty), out builtInCategory))
                                                        {
#if Revit2017 || Revit2018 || Revit2019 || Revit2020 || Revit2021 || Revit2022 || Revit2023 || Revit2024
                                                            Category category = document.Settings.Categories.Cast<Category>().ToList().Find(x => x.Id.IntegerValue == (int)builtInCategory);
#else
                                                            Category category = document.Settings.Categories.Cast<Category>().ToList().Find(x => x.Id.Value == (long)builtInCategory);
#endif

                                                            if (category != null)
                                                                categorySet.Insert(category);
                                                        }
                                                    }

                                                    if (categorySet.Size > 0)
                                                    {
                                                        string instance = objects[i, 14] as string;

                                                        if (string.IsNullOrEmpty(instance))
                                                            continue;

                                                        Autodesk.Revit.DB.Binding binding = null;
                                                        if (instance != null && instance.Trim().ToUpper() == "INSTANCE")
                                                            binding = ExternalCommandData.Application.Application.Create.NewInstanceBinding(categorySet);
                                                        else
                                                            binding = ExternalCommandData.Application.Application.Create.NewTypeBinding(categorySet);

                                                        bindingMap.Insert(definition, binding, groupTypeId);
                                                    }
                                                }
#endif
                                            }
                                        }
                                    }
                                }

                            }
                            else
                            {
                                progressForm.Update("???");
                            }
                        }
                    }
                }

                transaction.Commit();
            }

            ExternalCommandData.Application.Application.SharedParametersFilename = path_SharedParametersFile;
            ExternalCommandData.Application.Application.OpenSharedParameterFile();

            System.IO.File.Delete(path_SharedParametersFile_Temp);
        }
    }
}
