using Amethyst.Plugins.Contract;
using Microsoft.VisualBasic;
using plugin_Kinect360.PInvoke;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Storage;

namespace plugin_Kinect360.Fixes;

internal class NotPoweredFix : IFix
{
    public IDependencyInstaller.ILocalizationHost Host { get; set; }
    public string Name { get; set; } = string.Empty; // Set in ListFixes()

    public bool IsMandatory => IsNecessary; // Runtime check (set both to 1 to auto-apply during setup)
    public bool IsNecessary => MustFixNotPowered(); // Runtime check (set both to 1 to auto-apply during setup)
    public string InstallerEula => string.Empty; // Don't show, check the KinectSdk one for reference

    public Task<bool> Apply(IProgress<InstallationProgress> progress,
        CancellationToken cancellationToken, object arg = null)
    {
        Host?.Log($"Received fix application arguments: {arg}");

        // An automagic fix for E_NUI_NOTPOWERED
        // This fix involves practically uninstalling all Unknown Kinect Drivers, using P/Invoke,
        // then forcing a scan for hardware changes to trigger the drivers to re-scan

        // The method involved with fixing this is subject to change at this point
        var deviceTree = new DeviceTree(Host);

        // Get Kinect Devices
        return Task.FromResult(deviceTree.DeviceNodes.Where(d => d.ClassGuid == DeviceClasses.Unknown).All(device =>
        {
            // Device is a Kinect 360 Device
            if (device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02B0&REV_0107" && // Kinect for Windows Device
                device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02BB&REV_0100&MI_00" && // Kinect for Windows Audio Array
                device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02BB&REV_0100&MI_01" && // Kinect for Windows Security Device
                device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02AE&REV_010;" && // Kinect for Windows Camera
                device.GetProperty(DevRegProperty.HardwareId) != "USB\\VID_045E&PID_02BB&REV_0100&MI_02") return true; // Kinect USB Audio

            Host?.Log($"Found faulty Kinect device! {{ Name: {device.Description} }}");
            Host?.Log($"Attempting to fix device {device.Description}...");

            return device.UninstallDevice();
        }) && deviceTree.RescanDevices());
    }

    private bool MustFixNotPowered()
    {
        try
        {
            // Load Kinect10.dll (if installed and check for E_NUI_NOTPOWERED)
            var kinectAssembly = Assembly.LoadFrom("KinectHandler.dll");
            var kinectHandler = kinectAssembly.GetType("KinectHandler.KinectHandler");

            if (kinectHandler is null)
            {
                Host?.Log("Could not find type KinectHandler.KinectHandler in KinectHandler.dll");
                return false;
            }

            dynamic handler = Activator.CreateInstance(kinectHandler);
            if (handler is not null) return handler.DeviceStatus is 6;

            Host?.Log("Activator.CreateInstance returned null for KinectHandler.KinectHandler");
            return false;
        }
        catch (Exception ex)
        {
            Host?.Log(ex);
            return false;
        }
    }
}