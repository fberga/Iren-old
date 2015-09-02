using Microsoft.VisualStudio.Tools.Applications;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Iren.ToolsExcel.Base;
using IWshRuntimeLibrary;
using System.Reflection;

namespace Iren.ToolsExcel.FileCopy
{
    public class ToolsExcelFileCopy : IAddInPostDeploymentAction
    {
        public Version Version { get { return Assembly.GetExecutingAssembly().GetName().Version; } }

        public void Execute(AddInPostDeploymentActionArgs args)
        {
            string dataDirectory = Path.Combine("Data", Simboli.nomeFile);
            //string file = @"ExcelWorkbook.xlsx";
            string sourcePath = args.AddInPath;
            Uri deploymentManifestUri = args.ManifestLocation;
            string destPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Simboli.nomeApplicazione);
            string linkPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string sourceFile = Path.Combine(sourcePath, dataDirectory);
            string destFile = Path.Combine(destPath, Simboli.nomeFile);
            string lnk = Path.Combine(linkPath, Simboli.nomeFile + ".lnk");

            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                case AddInInstallationStatus.Update:
                    
                    if (!Directory.Exists(destPath))
                        Directory.CreateDirectory(destPath);

                    System.IO.File.Copy(sourceFile, destFile);
                    ServerDocument.RemoveCustomization(destFile);
                    ServerDocument.AddCustomization(destFile, deploymentManifestUri);

                    //shortcut
                    if(System.IO.File.Exists(lnk))
                        System.IO.File.Delete(lnk);

                    var shell = new WshShell();
                    IWshShortcut shortcut = shell.CreateShortcut(lnk);
                    shortcut.TargetPath = destFile;
                    shortcut.IconLocation = destFile;
                    shortcut.Save();
                    break;
                case AddInInstallationStatus.Uninstall:
                    if (Directory.Exists(destPath))
                        Directory.Delete(destPath, true);
                    if (System.IO.File.Exists(lnk))
                        System.IO.File.Delete(lnk);
                    break;
            }
        }
    }
}
