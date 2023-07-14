using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Amethyst.Plugins.Contract;
using Windows.Storage;
using RestSharp;
using Microsoft.UI.Xaml.Controls;

namespace plugin_Kinect360;

internal class SetupData : ICoreSetupData
{
    public object PluginIcon => new BitmapIcon
    {
        UriSource = new Uri(Path.Join(Directory.GetParent(
                Assembly.GetExecutingAssembly().Location)!.FullName,
            "Assets", "Resources", "icon.png"))
    };

    public string GroupName => "kinect";
    public Type PluginType => typeof(ITrackingDevice);
}

internal class SdkInstaller : IDependencyInstaller
{
    private const string WixDownloadUrl = // 34670186
        "https://github.com/wixtoolset/wix3/releases/download/wix3112rtm/wix311-binaries.zip";

    private const string RuntimeDownloadUrl = // 233219096
        "https://download.microsoft.com/download/E/1/D/E1DEC243-0389-4A23-87BF-F47DE869FC1A/KinectSDK-v1.8-Setup.exe";

    private const string ToolkitDownloadUrl = // 403021808
        "https://download.microsoft.com/download/D/0/6/D061A21C-3AF3-4571-8560-4010E96F0BC8/KinectDeveloperToolkit-v1.8.0-Setup.exe";

    private List<string> SdkFilesToInstall { get; } = new()
    {
        "KinectDrivers-v1.8-x64.WHQL.msi",
        "KinectRuntime-v1.8-x64.msi",
        "KinectSDK-v1.8-x64.msi"
    };

    private string TemporaryFolderName { get; } = Guid.NewGuid().ToString().ToUpper();

    public bool IsInstalled
    {
        get
        {
            try
            {
                // Well, this is pretty much all we need for the plugin to be loaded
                return File.Exists(
                    @"C:\Program Files\Microsoft SDKs\Kinect\v1.8\Assemblies\Microsoft.Kinect.dll");
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

    public bool ProvidesTools => true;

    public bool ToolsInstalled
    {
        get
        {
            try
            {
                // Well, this is pretty much all we need for the plugin to be loaded
                return File.Exists(
                    @"C:\Program Files\Microsoft SDKs\Kinect\Developer Toolkit v1.8.0\Tools\ToolkitBrowser\ToolkitBrowser.exe");
            }
            catch (Exception)
            {
                // Access denied?
                return false;
            }
        }
    }

    public IDependencyInstaller.ILocalizationHost Host { get; set; }

    public async Task<bool> Install(IProgress<InstallationProgress> progress)
    {
        return
            // Download and unpack WiX
            await SetupWix("WiXToolset", progress) &&

            // Download, unpack, and install the runtime
            await SetupSdk("WiXToolset", progress);

        // Apply other related fixes, non-critical
        // ReSharper disable once ConditionIsAlwaysTrueOrFalse
        // (await SetupPost(progress) || true); // TODO
    }

    public async Task<bool> InstallTools(IProgress<InstallationProgress> progress)
    {
        return
            // Download and unpack WiX
            await SetupWix("WiXToolset", progress) &&

            // Download, unpack, and install the runtime
            await SetupToolkit("WiXToolset", progress);
    }

    private async Task<StorageFolder> GetTempDirectory()
    {
        return await ApplicationData.Current.TemporaryFolder.CreateFolderAsync(
            TemporaryFolderName, CreationCollisionOption.OpenIfExists);
    }

    private async Task<bool> SetupWix(string outputFolder, IProgress<InstallationProgress> progress)
    {
        try
        {
            using var client = new RestClient();
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Downloading/WiX") ??
                             "Downloading WiX Toolset"
            });

            // Create a stream reader using the received Installer Uri
            await using var stream =
                await client.ExecuteDownloadStreamAsync(WixDownloadUrl, new RestRequest());

            // Replace or create our installer file
            var installerFile = await (await GetTempDirectory()).CreateFileAsync(
                "wix-binaries.zip", CreationCollisionOption.ReplaceExisting);

            // Create an output stream and push all the available data to it
            await using var fsInstallerFile = await installerFile.OpenStreamForWriteAsync();
            await stream.CopyToWithProgressAsync(fsInstallerFile, innerProgress =>
            {
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = false,
                    OverallProgress = innerProgress / 34670186.0,
                    StageTitle = Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Downloading/WiX") ??
                                 "Downloading WiX Toolset"
                });
            }); // The runtime will do the rest for us

            // Close the file to unlock it
            fsInstallerFile.Close();

            var sourceZip = Path.GetFullPath(Path.Combine((await GetTempDirectory()).Path, "wix-binaries.zip"));
            var tempDirectory = Path.GetFullPath(Path.Combine((await GetTempDirectory()).Path, outputFolder));

            if (File.Exists(sourceZip))
            {
                if (!Directory.Exists(tempDirectory))
                    Directory.CreateDirectory(tempDirectory);

                try
                {
                    // Extract the toolset
                    ZipFile.ExtractToDirectory(sourceZip, tempDirectory, true);
                }
                catch (Exception e)
                {
                    progress.Report(new InstallationProgress
                    {
                        IsIndeterminate = true,
                        StageTitle =
                            (Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Exceptions/WiX/Extraction") ??
                             "Toolset extraction failed! Exception: {0}").Replace("{0}", e.Message)
                    });

                    return false;
                }

                return true;
            }
        }
        catch (Exception e)
        {
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = (Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Exceptions/WiX/Installation") ??
                              "Toolset installation failed! Exception: {0}").Replace("{0}", e.Message)
            });
            return false;
        }

        return false;
    }

    private async Task<bool> SetupSdk(string wixFolder, IProgress<InstallationProgress> progress)
    {
        try
        {
            using var client = new RestClient();
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Downloading/Runtime") ??
                             "Downloading Kinect for Xbox 360 SDK..."
            });

            // Create a stream reader using the received Installer Uri
            await using var stream =
                await client.ExecuteDownloadStreamAsync(RuntimeDownloadUrl, new RestRequest());

            // Replace or create our installer file
            var installerFile = await (await GetTempDirectory()).CreateFileAsync(
                "kinect-setup.exe", CreationCollisionOption.ReplaceExisting);

            // Create an output stream and push all the available data to it
            await using var fsInstallerFile = await installerFile.OpenStreamForWriteAsync();
            await stream.CopyToWithProgressAsync(fsInstallerFile, innerProgress =>
            {
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = false,
                    OverallProgress = innerProgress / 233219096.0,
                    StageTitle = Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Downloading/Runtime") ??
                                 "Downloading Kinect for Xbox 360 SDK..."
                });
            }); // The runtime will do the rest for us

            // Close the file to unlock it
            fsInstallerFile.Close();

            return
                // Extract all runtime files for the installation
                await ExtractFiles(Path.GetFullPath(Path.Combine((await GetTempDirectory()).Path, wixFolder)),
                    installerFile.Path, Path.Join((await GetTempDirectory()).Path, "KinectRuntime"), progress) &&

                // Install the files using msi installers
                InstallFiles(Directory.GetFiles(Path.Join((await GetTempDirectory()).Path,
                        "KinectRuntime", "AttachedContainer"), "*.msi")
                    .Where(x => SdkFilesToInstall.Contains(Path.GetFileName(x))), progress);
        }
        catch (Exception e)
        {
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle =
                    (Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Exceptions/Runtime/Installation") ??
                     "SDK installation failed! Exception: {0}").Replace("{0}", e.Message)
            });
            return false;
        }
    }

    private async Task<bool> SetupToolkit(string wixFolder, IProgress<InstallationProgress> progress)
    {
        try
        {
            using var client = new RestClient();
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Downloading/Toolkit") ??
                             "Downloading Kinect for Xbox 360 Toolkit..."
            });

            // Create a stream reader using the received Installer Uri
            await using var stream =
                await client.ExecuteDownloadStreamAsync(ToolkitDownloadUrl, new RestRequest());

            // Replace or create our installer file
            var installerFile = await (await GetTempDirectory()).CreateFileAsync(
                "kinect-setup.exe", CreationCollisionOption.ReplaceExisting);

            // Create an output stream and push all the available data to it
            await using var fsInstallerFile = await installerFile.OpenStreamForWriteAsync();
            await stream.CopyToWithProgressAsync(fsInstallerFile, innerProgress =>
            {
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = false,
                    OverallProgress = innerProgress / 403021808.0,
                    StageTitle = Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Downloading/Toolkit") ??
                                 "Downloading Kinect for Xbox 360 Toolkit..."
                });
            }); // The runtime will do the rest for us

            // Close the file to unlock it
            fsInstallerFile.Close();

            return
                // Extract all runtime files for the installation
                await ExtractFiles(Path.GetFullPath(Path.Combine((await GetTempDirectory()).Path, wixFolder)),
                    installerFile.Path, Path.Join((await GetTempDirectory()).Path, "KinectToolkit"), progress) &&

                // Install the files using msi installers
                InstallFiles(Directory.GetFiles(Path.Join((await GetTempDirectory()).Path,
                    "KinectToolkit", "AttachedContainer"), "KinectToolkit-v1.8.0-x64.msi"), progress);
        }
        catch (Exception e)
        {
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle =
                    (Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Exceptions/Toolkit/Installation") ??
                     "Toolkit installation failed! Exception: {0}").Replace("{0}", e.Message)
            });
            return false;
        }
    }

    private Task<bool> SetupPost(IProgress<InstallationProgress> progress)
    {
        // TODO Merge fixes functionality from the installer
        return Task.FromResult(true);
    }

    private async Task<bool> ExtractFiles(string wixPath, string sourceFile, string outputFolder,
        IProgress<InstallationProgress> progress)
    {
        try
        {
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = (Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Unpacking") ??
                              "Unpacking {0}...").Replace("{0}", Path.GetFileName(sourceFile))
            });

            // dark.exe {sourceFile} -x {outDir}
            var procStart = new ProcessStartInfo
            {
                FileName = Path.Combine(wixPath, "dark.exe"),
                WorkingDirectory = (await GetTempDirectory()).Path,
                Arguments = $"\"{sourceFile}\" -x \"{outputFolder}\"",
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,

                // Verbose error handling
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = false,
                UseShellExecute = false,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            var proc = Process.Start(procStart);
            // Redirecting process output so that we can log what happened
            var stdout = new StringBuilder();
            var stderr = new StringBuilder();

            proc!.OutputDataReceived += (_, args) =>
            {
                if (args.Data != null)
                    stdout.AppendLine(args.Data);
            };
            proc.ErrorDataReceived += (_, args) =>
            {
                if (args.Data != null)
                    stderr.AppendLine(args.Data);
            };

            proc.BeginErrorReadLine();
            proc.BeginOutputReadLine();
            var hasExited = proc.WaitForExit(60000);

            // https://github.com/wixtoolset/wix3/blob/6b461364c40e6d1c487043cd0eae7c1a3d15968c/src/tools/dark/dark.cs#L54
            // Exit codes for DARK:
            // 
            // 0 - Success
            // 1 - Error
            // Just in case
            if (!hasExited)
            {
                // WTF
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = true,
                    StageTitle =
                        Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Dark/Error/Timeout") ??
                        "Failed to execute dark.exe in the allocated time!"
                });

                proc.Kill();
            }

            if (proc.ExitCode == 1)
            {
                // Assume WiX failed
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = true,
                    StageTitle =
                        (Host?.RequestLocalizedString("/Plugins/Kinect360/Stages/Dark/Error/Result") ??
                         "Dark.exe exited with error code: {0}").Replace("{0}", proc.ExitCode.ToString())
                });

                return false;
            }
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

    private bool InstallFiles(IEnumerable<string> files, IProgress<InstallationProgress> progress)
    {
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

public static class RestExtensions
{
    public static Task<byte[]> ExecuteDownloadDataAsync(this RestClient client, string baseUrl, RestRequest request)
    {
        client.Options.BaseUrl = new Uri(baseUrl);
        return client.DownloadDataAsync(request);
    }

    public static Task<Stream> ExecuteDownloadStreamAsync(this RestClient client, string baseUrl, RestRequest request)
    {
        client.Options.BaseUrl = new Uri(baseUrl);
        return client.DownloadStreamAsync(request);
    }
}

public static class StreamExtensions
{
    public static async Task CopyToWithProgressAsync(this Stream source,
        Stream destination, Action<long> progress = null, int bufferSize = 10240)
    {
        var buffer = new byte[bufferSize];
        var total = 0L;
        int amtRead;

        do
        {
            amtRead = 0;
            while (amtRead < bufferSize)
            {
                var numBytes = await source.ReadAsync(
                    buffer, amtRead, bufferSize - amtRead);
                if (numBytes == 0) break;
                amtRead += numBytes;
            }

            total += amtRead;
            await destination.WriteAsync(buffer, 0, amtRead);
            progress?.Invoke(total);
        } while (amtRead == bufferSize);
    }
}