using Microsoft.VisualStudio.Tools.Applications;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.FileCopy
{
    public class ToolsExcelFileCopy : IAddInPostDeploymentAction
    {
        public void Execute(AddInPostDeploymentActionArgs args)
        {
            string dataDirectory = @"Data\ExcelWorkbook.xlsx";
            string file = @"ExcelWorkbook.xlsx";
            string sourcePath = args.AddInPath;
            Uri deploymentManifestUri = args.ManifestLocation;
            string destPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string sourceFile = System.IO.Path.Combine(sourcePath, dataDirectory);
            string destFile = System.IO.Path.Combine(destPath, file);

            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                case AddInInstallationStatus.Update:
                    File.Copy(sourceFile, destFile);
                    ServerDocument.RemoveCustomization(destFile);
                    ServerDocument.AddCustomization(destFile, deploymentManifestUri);
                    break;
                case AddInInstallationStatus.Uninstall:
                    if (File.Exists(destFile))
                    {
                        File.Delete(destFile);
                    }
                    break;
            }
        }
    }
}
