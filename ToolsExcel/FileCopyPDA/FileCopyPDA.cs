using Microsoft.VisualStudio.Tools.Applications.Deployment;
using Microsoft.VisualStudio.Tools.Applications;
using System.IO;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Diagnostics;
using System.Collections.Generic;

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
            string destPath = parameters.Attribute("destinationpath").Value;

            //statici
            string sourcePath = args.AddInPath;
            Uri deploymentManifestUri = args.ManifestLocation;
            string sourceFile = System.IO.Path.Combine(sourcePath, dataDirectory, file);
            string destFile = System.IO.Path.Combine(destPath, file);

            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                case AddInInstallationStatus.Update:
                    File.Copy(sourceFile, destFile, true);
                    //ServerDocument.RemoveCustomization(destFile);
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