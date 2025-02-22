using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Amethyst.Plugins.Contract;
using Windows.Storage;
using RestSharp;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Markup;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Shapes;
using plugin_Kinect360.Fixes;
using Path = System.IO.Path;

namespace plugin_Kinect360;

internal class SetupData : ICoreSetupData
{
    public object PluginIcon => new PathIcon
    {
        Data = (Geometry)XamlBindingHelper.ConvertValue(typeof(Geometry),
            "M55.75,3.68a2.27,2.27,0,0,0-2.42-2.4c-8.5.05-17,0-25.5,0H2.38c-.22,0-.44,0-.66,0A1.84,1.84,0,0,0,0,3.12c0,2.4,0,4.81,0,7.21a2.06,2.06,0,0,0,.44,1.12,1.9,1.9,0,0,0,1.66.67H23.76c1.48,0,1.46,0,1.47,1.46,0,.38-.09.54-.52.61-2.79.46-5.57,1-8.35,1.43-.35.06-.47.21-.45.56,0,.64,0,1.28,0,1.92,0,.88.35,1.22,1.25,1.23H38.45c.9,0,1.24-.34,1.25-1.23,0-.64,0-1.28,0-1.92,0-.35-.1-.5-.45-.56L31.68,14.3c-1.28-.22-1.26-.22-1.3-1.51,0-.53.12-.7.68-.69,7.46,0,14.92,0,22.37,0a2.07,2.07,0,0,0,2.3-2.3C55.73,7.77,55.67,5.72,55.75,3.68ZM18.77,9a2.3,2.3,0,1,1,2.3-2.27A2.31,2.31,0,0,1,18.77,9ZM29.66,9a2.3,2.3,0,1,1,2.28-2.3A2.32,2.32,0,0,1,29.66,9ZM37.4,9a2.26,2.26,0,0,1-2.32-2.3,2.25,2.25,0,0,1,2.27-2.29,2.32,2.32,0,0,1,2.34,2.31A2.29,2.29,0,0,1,37.4,9Z")
    };

    public string GroupName => "kinect";
    public Type PluginType => typeof(ITrackingDevice);
}

internal class SdkInstaller : IDependencyInstaller
{
    public IDependencyInstaller.ILocalizationHost Host { get; set; }

    public List<IDependency> ListDependencies()
    {
        return
        [
            new KinectSdk
            {
                Host = Host,
                Name = Host?.RequestLocalizedString("/Plugins/Kinect360/Dependencies/Runtime/Name") ??
                       "Kinect for Xbox 360 SDK"
            }
        ];
    }

    public List<IFix> ListFixes()
    {
        return
        [
            new NotPoweredFix
            {
                Host = Host,
                Name = Host?.RequestLocalizedString( // Without the "fix" part
                    "/Plugins/Kinect360/Fixes/NotPowered/Name") ?? "Not Powered"
            },
            new NotReadyFix
            {
                Host = Host,
                Name = Host?.RequestLocalizedString( // Without the "fix" part
                    "/Plugins/Kinect360/Fixes/NotReady/Name") ?? "Not Ready"
            }
        ];
    }
}

internal class KinectSdk : IDependency
{
    private List<string> SdkFilesToInstall { get; } =
    [
        "KinectDrivers-v1.8-x64.WHQL.msi",
        "KinectRuntime-v1.8-x64.msi",
        "KinectSDK-v1.8-x64.msi"
    ];

    public IDependencyInstaller.ILocalizationHost Host { get; set; }

    public string Name { get; set; }
    public bool IsMandatory => true;

    public bool IsInstalled
    {
        get
        {
            try
            {
                // Well, this is pretty much all we need for the plugin to be loaded
                return File.Exists(@"C:\Windows\System32\Kinect10.dll");
            }
            catch (Exception)
            {
                // Access denied?
                return false;
            }
        }
    }

    public string InstallerEula
    {
        get
        {
            try
            {
                return File.ReadAllText(Path.Join(
                    Directory.GetParent(Assembly.GetExecutingAssembly().Location)!.FullName,
                    "Assets", "Resources", "eula.md"));
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }
    }

    public async Task<bool> Install(IProgress<InstallationProgress> progress, CancellationToken cancellationToken)
    {
        // Amethyst will handle this exception for us anyway
        cancellationToken.ThrowIfCancellationRequested();
        var dependenciesFolder = Path.Join(Directory.GetParent(
                Assembly.GetExecutingAssembly().Location)!.FullName,
            "Assets", "Resources", "Dependencies");

        await PathsHandler.Setup();

        // Copy to temp if amethyst is packaged
        // ReSharper disable once InvertIf
        // Create a shared folder with the dependencies
        var dependenciesFolderInternal = await PathsHandler.TemporaryFolder.CreateFolderAsync(
            Guid.NewGuid().ToString().ToUpper(), CreationCollisionOption.OpenIfExists);

        // Copy all driver files to Amethyst's local data folder
        new DirectoryInfo(Path.Join(Directory.GetParent(Assembly.GetExecutingAssembly().Location)!.FullName,
                "Assets", "Resources", "Dependencies"))
            .CopyToFolder(dependenciesFolderInternal.Path);

        // Update the installation paths
        dependenciesFolder = dependenciesFolderInternal.Path;

        // Finally install the packages
        return InstallFiles(SdkFilesToInstall.Select(x => Path.Join(
            dependenciesFolder, x)), progress, cancellationToken);

        // Apply other related fixes, non-critical
        // ReSharper disable once ConditionIsAlwaysTrueOrFalse
        // (await SetupPost(progress) || true); // TODO
    }

    private bool InstallFiles(IEnumerable<string> files,
        IProgress<InstallationProgress> progress, CancellationToken cancellationToken)
    {
        // Amethyst will handle this exception for us anyway
        cancellationToken.ThrowIfCancellationRequested();

        // Execute each install
        foreach (var installFile in files)
            try
            {
                // msi /qn /norestart
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = true,
                    StageTitle =
                        (Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Installing") ??
                         "Installing {0}...").Replace("{0}", Path.GetFileName(installFile))
                });

                var msiExecutableStart = new ProcessStartInfo
                {
                    FileName = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                        @"System32\msiexec.exe"),
                    WorkingDirectory = Directory.GetParent(installFile)!.FullName,
                    Arguments = $"/i {installFile} /quiet /qn /norestart ALLUSERS=1",
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true,
                    Verb = "runas"
                };

                var msiExecutable = Process.Start(msiExecutableStart);
                msiExecutable!.WaitForExit(60000);
            }
            catch (Exception e)
            {
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = true,
                    StageTitle =
                        (Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Exceptions/Other") ??
                         "Exception: {0}").Replace("{0}", e.Message)
                });

                return false;
            }

        return true;
    }
}