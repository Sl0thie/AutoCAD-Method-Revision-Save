using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.IO;
using Serilog;
using Serilog.Sinks.File;

namespace AutoCAD_Method___Revision_Save
{
    public class RevisionSave : IExtensionApplication
    {
        /// <summary>
        /// 
        /// </summary>
        public void Initialize()
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File("logs/myapp.txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();

            Log.Information("Initialize");

            try
            {
                AcadApplication cadApp = (AcadApplication)Application.AcadApplication;
                AcadToolbar tb = cadApp.MenuGroups.Item(0).Toolbars.Add("External Methods");
                tb.Dock(Autodesk.AutoCAD.Interop.Common.AcToolbarDockStatus.acToolbarDockLeft);
                string basePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                AcadToolbarItem tbButton0 = tb.AddToolbarButton(0, "Revision Save", "Saves the file while updating the revision numbers in the file name.", "REVISIONSAVE\n", null);
                tbButton0.SetBitmaps(basePath + "/Save16.bmp", basePath + "/Save32.bmp");
            }
            catch (System.Exception ex)
            {
                Log.Error(ex, "Something went wrong");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public void Terminate()
        {
            //throw new NotImplementedException();

            Log.Information("Terminate");
        }

        /// <summary>
        /// 
        /// </summary>
        [CommandMethod("REVISIONSAVE", CommandFlags.Session)]
        public void Cmd6()
        {
            // Get AutoCAD objects.
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;

            if (doc.IsNamedDrawing)
            {
                string originalFileName = doc.Name;
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Drawing name is {originalFileName}.\n");

                if (originalFileName.Contains(" - revision "))
                {
                    string newFileName;

                    try
                    {
                        string versionNumbers = originalFileName.Substring(originalFileName.IndexOf(" - revision ") + 12);
                        versionNumbers = versionNumbers.Substring(0, versionNumbers.Length - 4);
                        int major = Convert.ToInt32(versionNumbers.Substring(0, versionNumbers.IndexOf(".")));
                        versionNumbers = versionNumbers.Substring(versionNumbers.IndexOf(".") + 1);
                        int minor = Convert.ToInt32(versionNumbers.Substring(0, versionNumbers.IndexOf(".")));
                        versionNumbers = versionNumbers.Substring(versionNumbers.IndexOf(".") + 1);
                        int revision = Convert.ToInt32(versionNumbers);
                        string path = originalFileName.Substring(0, originalFileName.IndexOf(" - revision "));
                        newFileName = path + " - revision " + major.ToString() + "." + minor.ToString() + "." + (revision + 1).ToString() + ".dwg";
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"New file name is {newFileName}.\n");
                    }
                    catch (System.Exception ex)
                    {
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Revision Save method failed.\n {ex.Message}\n");
                        return;
                    }

                    using (Database db = doc.Database)
                    {
                        db.SaveAs(newFileName, DwgVersion.Current);
                        _ = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Open(newFileName, false);
                        db.CloseInput(true);                   
                    }

                    doc.CloseAndSave(originalFileName);





                    //doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                    //using (Database db = doc.Database)
                    //{
                    //    db.SaveAs(newFileName, DwgVersion.Current);
                    //}
                }
                else
                {
                    string newFileName;
                    string reFileName;

                    try
                    {
                        newFileName = originalFileName.Substring(0, originalFileName.Length - 4) + " - revision 1.0.1.dwg";
                        reFileName = originalFileName.Substring(0, originalFileName.Length - 4) + " - revision 1.0.0.dwg";

                        Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"New file name is {newFileName}.\n");
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Renaming original file to {reFileName}.\n");
                    }
                    catch (System.Exception ex)
                    {
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Revision Save method failed.\n {ex.Message}\n");
                        return;
                    }

                    using (Database db = doc.Database)
                    {
                        db.SaveAs(newFileName, DwgVersion.Current);

                        _ = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Open(newFileName, false);



                        db.CloseInput(true);
                    }
                   
                    
                    doc.CloseAndSave(originalFileName);
                    File.Move(originalFileName, reFileName);
                }
            }
            else
            {
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Drawing {doc.Name} is not a named drawing.\n");
            }
        }
    }
}
