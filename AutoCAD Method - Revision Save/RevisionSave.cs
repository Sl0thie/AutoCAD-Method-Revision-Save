namespace AutoCAD_Method___Revision_Save
{
    using System;
    using System.IO;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Interop;
    using Autodesk.AutoCAD.Runtime;
    using Serilog;
    using Serilog.Sinks.File;

    /// <summary>
    /// RevisionSave is an Add-In for AutoCAD.
    /// </summary>
    public class RevisionSave : IExtensionApplication
    {
        /// <summary>
        /// Initialize method is called by AutoCAD during startup.
        /// </summary>
        public void Initialize()
        {
            // Create and start logger.
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File("logs/AutoCAD Revision Save - .txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();
            Log.Information("Initialize");

            // Create a toolbar button and include it in AutoCAD.
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
        /// Terminate method is called during shutdown.
        /// </summary>
        public void Terminate()
        {
            Log.Information("Terminate");
        }

        /// <summary>
        /// RevisionSaveCommand method is fired when the toolbar button is pressed.
        /// It checks to see if a revision number is present and if so it increments the number and saves the file.
        /// Or it add a revision number and saves the file.
        /// </summary>
        [CommandMethod("REVISIONSAVE", CommandFlags.Session)]
        public void RevisionSaveCommand()
        {
            // Get AutoCAD objects.
            Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;

            if (doc.IsNamedDrawing)
            {
                string originalFileName = doc.Name;
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Drawing name is {originalFileName}.\n");

                // Check if the file name already contains a revision number.
                if (originalFileName.Contains(" - revision "))
                {
                    string newFileName;

                    try
                    {
                        // Update revision number.
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
                        // Inform user of the failure via the AutoCAD console.
                        Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Revision Save method failed.\n {ex.Message}\n");
                        return;
                    }

                    using (Database db = doc.Database)
                    {
                        // Save current drawing to new file name.
                        db.SaveAs(newFileName, DwgVersion.Current);

                        // Close the original file.
                        db.CloseInput(true);
                        doc.CloseAndSave(originalFileName);

                        // Open the new file.
                        _ = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Open(newFileName, false);
                    }
                }
                else
                {
                    // Create the first revision number if one is not found.
                    string newFileName;
                    string reFileName;

                    try
                    {
                        // Add revision number.
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
                        // Save the drawing to the new file name.
                        db.SaveAs(newFileName, DwgVersion.Current);

                        // Close the original file.
                        db.CloseInput(true);
                        doc.CloseAndSave(originalFileName);

                        // Open the drawing under the new file name.
                        _ = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.Open(newFileName, false);

                        // Rename the original file name.
                        File.Move(originalFileName, reFileName);
                    }
                }
            }
            else
            {
                Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Drawing {doc.Name} is not a named drawing.\n");
            }
        }
    }
}
