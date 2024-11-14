using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.System;
using Amethyst.Plugins.Contract;
using NAudio.CoreAudioApi;
using plugin_Kinect360.PInvoke;

namespace plugin_Kinect360.Fixes;

internal class NotReadyFix : IFix
{
    public IDependencyInstaller.ILocalizationHost Host { get; set; }
    public string Name { get; set; } = string.Empty; // Set in ListFixes()

    public bool IsMandatory => IsNecessary; // Runtime check (set both to 1 to auto-apply during setup)
    public bool IsNecessary => MustFixNotReady(); // Runtime check (set both to 1 to auto-apply during setup)
    public string InstallerEula => string.Empty; // Don't show, check the KinectSdk one for reference

    public async Task<bool> Apply(IProgress<InstallationProgress> progress,
        CancellationToken cancellationToken, object arg = null)
    {
        Host?.Log($"Received fix application arguments: {arg}");

        try
        {
            var result = await CheckCoreIntegrity(progress);
            result &= CheckMicrophone(progress);
            result &= PreFixUnknownDevices();
            result &= await InstallDrivers(progress);
            result &= new DeviceTree(Host).RescanDevices();

            return result; // The fix
        }
        catch (Exception ex)
        {
            Host?.Log(ex);
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = ex.Message
            });

            return false;
        }
    }

    private bool MustFixNotReady()
    {
        try
        {
            var deviceTree = new DeviceTree(Host);
            var kinectDevices = deviceTree.DeviceNodes.Where(d => d.ClassGuid == DeviceClasses.KinectForWindows).Count(device =>
                device.GetProperty(DevRegProperty.HardwareId) == "USB\\VID_045E&PID_02B0&REV_0107" || // Kinect for Windows Device
                device.GetProperty(DevRegProperty.HardwareId) == "USB\\VID_045E&PID_02BB&REV_0100&MI_00" || // Kinect for Windows Audio Array
                device.GetProperty(DevRegProperty.HardwareId) == "USB\\VID_045E&PID_02BB&REV_0100&MI_01" || // Kinect for Windows Security Device
                device.GetProperty(DevRegProperty.HardwareId) == "USB\\VID_045E&PID_02AE&REV_010;"); // Kinect for Windows Camera

            // Check if less than 3 drivers have installed successfully
            if (kinectDevices < 4) return true;

            // Load Kinect10.dll (if installed and check for E_NUI_NOTREADY)
            var kinectAssembly = Assembly.LoadFrom("KinectHandler.dll");
            var kinectHandler = kinectAssembly.GetType("KinectHandler.KinectHandler");

            if (kinectHandler is null)
            {
                Host?.Log("Could not find type KinectHandler.KinectHandler in KinectHandler.dll");
                return false;
            }

            dynamic handler = Activator.CreateInstance(kinectHandler);
            if (handler is not null) return handler.DeviceStatus is 7;

            Host?.Log("Activator.CreateInstance returned null for KinectHandler.KinectHandler");
            return false;
        }
        catch (Exception ex)
        {
            Host?.Log(ex);
            return false;
        }
    }

    private async Task<bool> CheckCoreIntegrity(IProgress<InstallationProgress> progress)
    {
        progress.Report(new InstallationProgress
            { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/CheckingMemoryIntegrity") });

        // ReSharper disable once InvertIf
        if (NtDll.IsCodeIntegrityEnabled() || new DeviceTree(Host).DeviceNodes
                .Where(device => device.GetProperty(DevRegProperty.HardwareId).StartsWith("USB\\VID_045E&PID_02AE"))
                .Any(device => (device.GetStatus(out var code) & DeviceNodeStatus.HasProblem) != 0 || code == DeviceNodeProblemCode.DriverFailedLoad))
        {
            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/MemoryIntegrityEnabled") });
            await Launcher.LaunchUriAsync(new Uri(
                $"amethyst-app:crash-message#{Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Prompt/MustDisableMemoryIntegrity")}"));
            await Launcher.LaunchUriAsync(new Uri("windowsdefender://coreisolation"));
            return false; // Open Windows Security on the Core Isolation page and give up
        }

        return true;
    }

    private bool CheckMicrophone(IProgress<InstallationProgress> progress)
    {
        progress.Report(new InstallationProgress
            { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/CheckingKinectMicrophone") });

        // ReSharper disable once InvertIf
        if (new MMDeviceEnumerator().EnumerateAudioEndPoints(DataFlow.Capture,
                DeviceState.Disabled | DeviceState.Unplugged | DeviceState.Active)
            .Any(wasapi => wasapi.DeviceFriendlyName == "Kinect USB Audio"))
        {
            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/KinectV1MicrophoneFound") });

            // ReSharper disable once InvertIf
            if (new MMDeviceEnumerator().EnumerateAudioEndPoints(DataFlow.Capture,
                    DeviceState.Disabled | DeviceState.Unplugged)
                .Where(wasapi => wasapi.DeviceFriendlyName == "Kinect USB Audio")
                .Any(wasapi => wasapi.State != DeviceState.Active))
            {
                progress.Report(new InstallationProgress
                    { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/KinectMicrophoneDisabled") });
                return new MMDeviceEnumerator().EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Disabled)
                    .Any(wasapi => wasapi.DeviceFriendlyName == "Kinect USB Audio" && DevicePolicy.SetAudioEndpointState(wasapi.ID, true));
            }
        }

        return true;
    }

    public bool PreFixUnknownDevices()
    {
        // An automagic fix for E_NUI_NOTPOWERED
        // This fix involves practically uninstalling all Unknown Kinect Drivers, using P/Invoke, then forcing a scan for hardware changes to trigger the drivers to re-scan

        // The method involved with fixing this is subject to change at this point
        return new DeviceTree(Host).DeviceNodes.Where(d => d.ClassGuid == DeviceClasses.Unknown).All(device =>
        {
            // Device is a Kinect 360 Device
            if (device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02B0&REV_0107" && // Kinect for Windows Device
                device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02BB&REV_0100&MI_00" && // Kinect for Windows Audio Array
                device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02BB&REV_0100&MI_01" && // Kinect for Windows Security Device
                device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02AE&REV_010;" && // Kinect for Windows Camera
                device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02BB&REV_0100&MI_02") return true; // Kinect USB Audio

            Host?.Log($"Found faulty Kinect device!  {{ Name: {device.Description} }}");
            Host?.Log($"Attempting to fix device {device.Description}...");
            return device.UninstallDevice();
        });
    }

    private async Task<bool> InstallDrivers(IProgress<InstallationProgress> progress)
    {
        progress.Report(new InstallationProgress
            { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallingDrivers") });

        var pathToDriversDirectory = Path.Join(Directory.GetParent(
            Assembly.GetExecutingAssembly().Location)!.FullName, "Assets", "Resources", "Dependencies", "Drivers");

        var driverTemp = (await ApplicationData.Current.TemporaryFolder.CreateFolderAsync(
            Guid.NewGuid().ToString().ToUpper(), CreationCollisionOption.OpenIfExists)).Path;

        Directory.CreateDirectory(driverTemp);

        // Device Driver
        {
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Device_cat"), Path.Combine(driverTemp, "kinect.cat"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Device_inf"), Path.Combine(driverTemp, "kinectdevice.inf"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Device_WdfCo"), Path.Combine(driverTemp, "WdfCoInstaller01009.dll"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Device_WinUsbCo"), Path.Combine(driverTemp, "WinUSBCoInstaller.dll"));

            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallDeviceDriver") });
            AssignDriverToDeviceId("USB\\VID_045E&PID_02B0&REV_0107", Path.Combine(driverTemp, "kinectdevice.inf"), Host);
            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallDeviceDriverSuccess") });

            File.Move(Path.Combine(driverTemp, "kinect.cat"), Path.Combine(pathToDriversDirectory, "Driver_Device_cat"));
            File.Move(Path.Combine(driverTemp, "kinectdevice.inf"), Path.Combine(pathToDriversDirectory, "Driver_Device_inf"));
            File.Move(Path.Combine(driverTemp, "WdfCoInstaller01009.dll"), Path.Combine(pathToDriversDirectory, "Driver_Device_WdfCo"));
            File.Move(Path.Combine(driverTemp, "WinUSBCoInstaller.dll"), Path.Combine(pathToDriversDirectory, "Driver_Device_WinUsbCo"));
        }

        // Audio Driver
        {
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Audio_cat"), Path.Combine(driverTemp, "kinect.cat"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Audio_inf"), Path.Combine(driverTemp, "kinectaudio.inf"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Audio_WdfCo"), Path.Combine(driverTemp, "WdfCoInstaller01009.dll"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Audio_WinUsbCo"), Path.Combine(driverTemp, "WinUSBCoInstaller.dll"));

            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallAudioDriver") });
            SetupApi.InstallDriverFromInf(Path.Combine(driverTemp, "kinectaudio.inf"));
            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallAudioDriverSuccess") });

            File.Move(Path.Combine(driverTemp, "kinect.cat"), Path.Combine(pathToDriversDirectory, "Driver_Audio_cat"));
            File.Move(Path.Combine(driverTemp, "kinectaudio.inf"), Path.Combine(pathToDriversDirectory, "Driver_Audio_inf"));
            File.Move(Path.Combine(driverTemp, "WdfCoInstaller01009.dll"), Path.Combine(pathToDriversDirectory, "Driver_Audio_WdfCo"));
            File.Move(Path.Combine(driverTemp, "WinUSBCoInstaller.dll"), Path.Combine(pathToDriversDirectory, "Driver_Audio_WinUsbCo"));
        }

        // Audio Array Driver
        {
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_AudioArray_cat"), Path.Combine(driverTemp, "kinect.cat"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_AudioArray_inf"), Path.Combine(driverTemp, "kinectaudioarray.inf"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_AudioArray_WdfCo"), Path.Combine(driverTemp, "WdfCoInstaller01009.dll"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_AudioArray_WinUsbCo"), Path.Combine(driverTemp, "WinUSBCoInstaller.dll"));

            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallAudioArrayDriver") });
            AssignDriverToDeviceId("USB\\VID_045E&PID_02BB&REV_0100&MI_00", Path.Combine(driverTemp, "kinectaudioarray.inf"), Host);
            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallAudioArrayDriverSuccess") });

            File.Move(Path.Combine(driverTemp, "kinect.cat"), Path.Combine(pathToDriversDirectory, "Driver_AudioArray_cat"));
            File.Move(Path.Combine(driverTemp, "kinectaudioarray.inf"), Path.Combine(pathToDriversDirectory, "Driver_AudioArray_inf"));
            File.Move(Path.Combine(driverTemp, "WdfCoInstaller01009.dll"), Path.Combine(pathToDriversDirectory, "Driver_AudioArray_WdfCo"));
            File.Move(Path.Combine(driverTemp, "WinUSBCoInstaller.dll"), Path.Combine(pathToDriversDirectory, "Driver_AudioArray_WinUsbCo"));
        }

        // Camera Driver
        {
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Camera_cat"), Path.Combine(driverTemp, "kinect.cat"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Camera_inf"), Path.Combine(driverTemp, "kinectcamera.inf"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Camera_sys"), Path.Combine(driverTemp, "kinectcamera.sys"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Camera_WdfCo"), Path.Combine(driverTemp, "WdfCoInstaller01009.dll"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Camera_WinUsbCo"), Path.Combine(driverTemp, "WinUSBCoInstaller.dll"));

            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallCameraDriver") });
            AssignDriverToDeviceId("USB\\VID_045E&PID_02AE&REV_010;", Path.Combine(driverTemp, "kinectcamera.inf"), Host);
            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallCameraDriverSuccess") });

            File.Move(Path.Combine(driverTemp, "kinect.cat"), Path.Combine(pathToDriversDirectory, "Driver_Camera_cat"));
            File.Move(Path.Combine(driverTemp, "kinectcamera.inf"), Path.Combine(pathToDriversDirectory, "Driver_Camera_inf"));
            File.Move(Path.Combine(driverTemp, "kinectcamera.sys"), Path.Combine(pathToDriversDirectory, "Driver_Camera_sys"));
            File.Move(Path.Combine(driverTemp, "WdfCoInstaller01009.dll"), Path.Combine(pathToDriversDirectory, "Driver_Camera_WdfCo"));
            File.Move(Path.Combine(driverTemp, "WinUSBCoInstaller.dll"), Path.Combine(pathToDriversDirectory, "Driver_Camera_WinUsbCo"));
        }

        // Security Driver
        {
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Security_cat"), Path.Combine(driverTemp, "kinect.cat"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Security_inf"), Path.Combine(driverTemp, "kinectsecurity.inf"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Security_WdfCo"), Path.Combine(driverTemp, "WdfCoInstaller01009.dll"));
            File.Move(Path.Combine(pathToDriversDirectory, "Driver_Security_WinUsbCo"), Path.Combine(driverTemp, "WinUSBCoInstaller.dll"));

            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallSecurityDriver") });
            AssignDriverToDeviceId("USB\\VID_045E&PID_02BB&REV_0100&MI_01", Path.Combine(driverTemp, "kinectsecurity.inf"), Host);
            progress.Report(new InstallationProgress
                { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/InstallSecurityDriverSuccess") });

            File.Move(Path.Combine(driverTemp, "kinect.cat"), Path.Combine(pathToDriversDirectory, "Driver_Security_cat"));
            File.Move(Path.Combine(driverTemp, "kinectsecurity.inf"), Path.Combine(pathToDriversDirectory, "Driver_Security_inf"));
            File.Move(Path.Combine(driverTemp, "WdfCoInstaller01009.dll"), Path.Combine(pathToDriversDirectory, "Driver_Security_WdfCo"));
            File.Move(Path.Combine(driverTemp, "WinUSBCoInstaller.dll"), Path.Combine(pathToDriversDirectory, "Driver_Security_WinUsbCo"));
        }

        // Microphone driver
        progress.Report(new InstallationProgress
            { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/AssignMicrophoneDriver") });
        AssignGenericAudioDriver(Host);
        progress.Report(new InstallationProgress
            { IsIndeterminate = true, StageTitle = Host.RequestLocalizedString("/Plugins/Kinect360/Fixes/NotReady/Stage/AssignMicrophoneDriverSuccess") });
        // We don't assign the endpoint driver because it's device ID is VERY GENERIC ( MMDEVAPI\AudioEndpoints )

        return true;
    }

    // Assigns a driver to a device programmatically
    public static bool AssignDriverToDeviceId(string deviceId, string infPath, IDependencyInstaller.ILocalizationHost host)
    {
        return SetupApi.AssignDriverToDeviceId(new DeviceTree(host).DeviceNodes.Where(device => device.GetProperty(DevRegProperty.HardwareId) == deviceId)
            .Select(device => device.GetInstanceId()).LastOrDefault(x => !string.IsNullOrEmpty(x), string.Empty), infPath, host);
    }

    // Assigns a driver to a device programmatically
    public static void AssignGenericAudioDriver(IDependencyInstaller.ILocalizationHost host)
    {
        SetupApi.AssignExistingDriverViaInfToDeviceId(new DeviceTree(host).DeviceNodes
                .Where(device => device.GetProperty(DevRegProperty.HardwareId) == "USB\\VID_045E&PID_02BB&REV_0100&MI_02")
                .Select(device => device.GetInstanceId()).LastOrDefault(x => !string.IsNullOrEmpty(x), string.Empty),
            "wdma_usb.inf", "(Generic USB Audio)", "USB Audio Device", host);
    }
}