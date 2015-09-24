using Microsoft.VisualStudio.Tools.Applications.Deployment;
using Microsoft.VisualStudio.Tools.Applications;
using System.IO;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Diagnostics;
using System.Collections.Generic;
using IWshRuntimeLibrary;

namespace Iren.ToolsExcel.PostDeployment
{
    public class FileCopyPDA : IAddInPostDeploymentAction
    {
        public void Execute(AddInPostDeploymentActionArgs args)
        {   
            XElement parameters = XElement.Parse(args.PostActionManifestXml);

            //configurabili
            string dataDirectory = @"Data\";
            string file = parameters.Attribute("filename").Value;
            string destPath = Environment.ExpandEnvironmentVariables(parameters.Attribute("destinationpath").Value);

            //statici
            string sourcePath = args.AddInPath;
            Uri deploymentManifestUri = args.ManifestLocation;
            string sourceFile = Path.Combine(sourcePath, dataDirectory, file);
            string destFile = Path.Combine(destPath, file);

            //creo cartella sul desktop e inserisco il link al file
            string desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "PSO");
            string desktopLink = Path.Combine(desktopPath, file + ".lnk");

            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                    if (!Directory.Exists(destPath))
                        Directory.CreateDirectory(destPath);

                    System.IO.File.Copy(sourceFile, destFile, true);

                    if(ServerDocument.IsCustomized(destFile))
                        ServerDocument.RemoveCustomization(destFile);

                    ServerDocument.AddCustomization(destFile, deploymentManifestUri);

                    if (!Directory.Exists(desktopPath))
                        Directory.CreateDirectory(desktopPath);

                    var shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(desktopLink);
                    shortcut.Description = "Collegamento all'applicazione " + file;
                    shortcut.TargetPath = destFile;
                    shortcut.Save();

                    break;
                case AddInInstallationStatus.Update:
                    string dirUPDATE = Path.Combine(destPath, "UPDATE");
                    string fileUPDATE = Path.Combine(dirUPDATE, file);
                    if (!Directory.Exists(dirUPDATE))
                        Directory.CreateDirectory(dirUPDATE);

                    System.IO.File.Copy(sourceFile, fileUPDATE, true);

                    if (ServerDocument.IsCustomized(fileUPDATE))
                        ServerDocument.RemoveCustomization(fileUPDATE);

                    ServerDocument.AddCustomization(fileUPDATE, deploymentManifestUri);

                    break;
                case AddInInstallationStatus.Uninstall:
                    if (System.IO.File.Exists(destFile))
                    {
                        //rimuovo tutti i file dell'installazione
                        System.IO.File.Delete(destFile);

                        if (System.IO.File.Exists(desktopLink))
                            System.IO.File.Delete(desktopLink);

                        string update = Path.Combine(desktopPath, "UPDATE");
                        if (Directory.Exists(update) && !Directory.EnumerateFileSystemEntries(update).Any())
                            Directory.Delete(update);

                        if (!Directory.EnumerateFileSystemEntries(destPath).Any())
                            Directory.Delete(destPath);
                    }
                    break;
            }
        }
    }
}