#region Header
// Revit API .NET Labs
//
// Copyright (C) 2007-2019 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software
// for any purpose and without fee is hereby granted, provided
// that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
#endregion // Header

#region Namespaces
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

using Newtonsoft.Json;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;

using DesignAutomationFramework;
using Autodesk.Revit.DB.Events;

using libxl;

#endregion // Namespaces

namespace ExportImportExcel
{


    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class HandleParamsExcel : IExternalDBApplication
    {
        public const string FireRating = "Fire Rating";
        public const string Comments = "Comments";

        public ExternalDBApplicationResult OnStartup(ControlledApplication app)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += HandleDesignAutomationReadyEvent;
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnShutdown(ControlledApplication app)
        {
            return ExternalDBApplicationResult.Succeeded;
        }

        public void HandleDesignAutomationReadyEvent(object sender, DesignAutomationReadyEventArgs e)
        {
            e.Succeeded = ProcessParameters(e.DesignAutomationData);
        }

        public bool ProcessParameters(DesignAutomationData data)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            Application rvtApp = data.RevitApp;
            if (rvtApp == null)
                throw new InvalidDataException(nameof(rvtApp));

            InputParams inputParams = InputParams.Parse("params.json");
            if (inputParams == null)
                throw new InvalidDataException("Cannot parse out input parameters correctly.");

            Console.WriteLine("Got the input json file sucessfully.");

            var cloudModelPath = ModelPathUtils.ConvertCloudGUIDsToCloudPath(ModelPathUtils.CloudRegionUS, inputParams.ProjectGuid, inputParams.ModelGuid);
            rvtApp.FailuresProcessing += OnFailuresProcessing;

            Console.WriteLine("Revit starts openning Revit Cloud Model");
            Document doc = rvtApp.OpenDocumentFile(cloudModelPath, new OpenOptions());
            if (doc == null)
                throw new InvalidOperationException("Could not open Revit Cloud Model.");

            Console.WriteLine("Revit Cloud Model is opened");
            return ExportToExcel(rvtApp, doc, inputParams.IncludeFireRating, inputParams.IncludeComments);
        }


        public static bool ExportToExcel( Application rvtApp, Document doc, bool includeFireRating, bool includeComments )
        {
            var outputDir = (!doc.IsModelInCloud && doc.PathName != "")? Path.GetDirectoryName(doc.PathName): Directory.GetCurrentDirectory();
            Console.WriteLine("outputDir is: " + outputDir);

            try
            {
                List<View3D> views = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().Where(vw => (
                    !vw.IsTemplate && vw.CanBePrinted)
                ).ToList();
                Console.WriteLine("the number of views: " + views.Count);

                STLExportOptions options = new STLExportOptions();
                options.ExportBinary = true;
                foreach (var item in views)
                {
                    Console.WriteLine("Start Exporting view name: " + item.Name);
                    options.ViewId = item.Id;
                    if (doc.Export(outputDir, item.Name+".stl", options))
                    {
                        var file = Path.Combine(outputDir, item.Name + ".stl");
                        if (File.Exists(file))
                            Console.WriteLine("ExportToSTL successful at: " + file);
                        else
                            Console.WriteLine("ExportToSTL missing output at: " + file);
                    }else
                        Console.WriteLine("ExportToSTL: failed");
                }

                return true;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Exception ExportToSTL: " + ex.Message);
                Console.WriteLine("StackTrace: " + ex.StackTrace);
                return false;
            }
        }

    
        // Simple upgrade failure processor that ignores all warnings and resolve all resolvable errors
        private void OnFailuresProcessing(object sender, FailuresProcessingEventArgs e)
        {
            var fa = e?.GetFailuresAccessor();

            fa.DeleteAllWarnings(); // Ignore all upgrade warnings
            var failures = fa.GetFailureMessages();
            if (!failures.Any())
            {
                return;
            }

            failures = failures.Where(fail => fail.HasResolutions()).ToList();
            fa.ResolveFailures(failures);
        }
    }


    /// <summary>
    /// InputParams is used to parse the input Json parameters
    /// </summary>
    internal class InputParams
    {
        public bool Export { get; set; } = false;
        public bool IncludeFireRating { get; set; } = true;
        public bool IncludeComments { get; set; } = true;

        public string Region { get; set; } = ModelPathUtils.CloudRegionUS;
        [JsonProperty(PropertyName = "projectGuid", Required = Required.Default)]
        public Guid ProjectGuid { get; set; } 
        [JsonProperty(PropertyName = "modelGuid", Required = Required.Default)]
        public Guid ModelGuid { get; set; }

        static public InputParams Parse(string jsonPath)
        {
            try
            {
                if (!File.Exists(jsonPath))
                    return new InputParams { Export = false, IncludeFireRating = true, IncludeComments = true, Region = ModelPathUtils.CloudRegionUS, ProjectGuid = new Guid(""), ModelGuid= new Guid("") };

                string jsonContents = File.ReadAllText(jsonPath);
                return JsonConvert.DeserializeObject<InputParams>(jsonContents);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Exception when parsing json file: " + ex);
                return null;
            }
        }
    }

}
