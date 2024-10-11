// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See LICENSE in the project root for license information.

using System;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Amethyst.Plugins.Contract;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Storage.Streams;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace plugin_Kinect360;

[Export(typeof(ITrackingDevice))]
[ExportMetadata("Name", "Xbox 360 Kinect")]
[ExportMetadata("Guid", "K2VRTEAM-AME2-APII-DVCE-DVCEKINECTV1")]
[ExportMetadata("Publisher", "K2VR Team")]
[ExportMetadata("Version", "1.0.0.0")]
[ExportMetadata("Website", "https://github.com/KinectToVR/plugin_Kinect360")]
[ExportMetadata("DependencyLink", "https://docs.k2vr.tech/{0}/360/setup/")]
[ExportMetadata("DependencySource", "https://download.microsoft.com/download/E/C/5/EC50686B-82F4-4DBF-A922-980183B214E6/KinectRuntime-v1.8-Setup.exe")]
[ExportMetadata("DependencyInstaller", typeof(SdkInstaller))]
[ExportMetadata("CoreSetupData", typeof(SetupData))]
public class Kinect360 : KinectHandler.KinectHandler, ITrackingDevice
{
    [Import(typeof(IAmethystHost))] private IAmethystHost Host { get; set; }

    private bool PluginLoaded { get; set; }
    private Page InterfaceRoot { get; set; }
    private NumberBox TiltNumberBox { get; set; }
    private TextBlock TiltTextBlock { get; set; }
    public WriteableBitmap CameraImage { get; set; }

    public bool IsPositionFilterBlockingEnabled => false;
    public bool IsPhysicsOverrideEnabled => false;
    public bool IsSelfUpdateEnabled => false;
    public bool IsFlipSupported => true;
    public bool IsAppOrientationSupported => true;
    public object SettingsInterfaceRoot => InterfaceRoot;

    public ObservableCollection<TrackedJoint> TrackedJoints { get; } =
        // Prepend all supported joints to the joints list
        new(Enum.GetValues<TrackedJointType>()
            .Where(x => x is not TrackedJointType.JointNeck and not TrackedJointType.JointManual and not
                TrackedJointType.JointHandTipLeft and not TrackedJointType.JointHandTipRight and not
                TrackedJointType.JointThumbLeft and not TrackedJointType.JointThumbRight)
            .Select(x => new TrackedJoint { Name = x.ToString(), Role = x }));

    public string DeviceStatusString => PluginLoaded
        ? DeviceStatus switch
        {
            0 => Host.RequestLocalizedString("/Plugins/Kinect360/Statuses/Success"),
            1 => Host.RequestLocalizedString("/Plugins/Kinect360/Statuses/Initializing"),
            2 => Host.RequestLocalizedString("/Plugins/Kinect360/Statuses/NotConnected"),
            3 => Host.RequestLocalizedString("/Plugins/Kinect360/Statuses/NotGenuine"),
            4 => Host.RequestLocalizedString("/Plugins/Kinect360/Statuses/NotSupported"),
            5 => Host.RequestLocalizedString("/Plugins/Kinect360/Statuses/InsufficientBandwidth"),
            6 => Host.RequestLocalizedString("/Plugins/Kinect360/Statuses/NotPowered"),
            7 => Host.RequestLocalizedString("/Plugins/Kinect360/Statuses/NotReady"),
            _ => $"Undefined: {DeviceStatus}\nE_UNDEFINED\nSomething weird has happened, though we can't tell what."
        }
        : $"Undefined: {DeviceStatus}\nE_UNDEFINED\nSomething weird has happened, though we can't tell what.";

    public Uri ErrorDocsUri => new(DeviceStatus switch
    {
        6 => $"https://docs.k2vr.tech/{Host?.DocsLanguageCode ?? "en"}/360/troubleshooting/notpowered/",
        7 => $"https://docs.k2vr.tech/{Host?.DocsLanguageCode ?? "en"}/360/troubleshooting/notready/",
        3 => $"https://docs.k2vr.tech/{Host?.DocsLanguageCode ?? "en"}/360/troubleshooting/notgenuine/",
        5 => $"https://docs.k2vr.tech/{Host?.DocsLanguageCode ?? "en"}/360/troubleshooting/insufficientbandwidth/",
        _ => $"https://docs.k2vr.tech/{Host?.DocsLanguageCode ?? "en"}/360/troubleshooting/"
    });

    public void OnLoad()
    {
        TiltNumberBox = new NumberBox
        {
            SpinButtonPlacementMode = NumberBoxSpinButtonPlacementMode.Inline,
            Margin = new Thickness { Left = 5, Top = 5, Right = 5, Bottom = 3 },
            VerticalAlignment = VerticalAlignment.Center,
            SmallChange = 1,
            LargeChange = 10
        };

        TiltNumberBox.ValueChanged += (sender, _) =>
        {
            if (!IsInitialized) return;
            if (double.IsNaN(sender.Value))
                sender.Value = ElevationAngle;

            sender.Value = Math.Clamp(sender.Value, -27, 27);
            ElevationAngle = (int)sender.Value; // Update
            Host?.PlayAppSound(SoundType.Invoke);
        };

        TiltTextBlock = new TextBlock
        {
            Text = Host.RequestLocalizedString("/Plugins/Kinect360/Settings/Labels/Angle"),
            Margin = new Thickness { Left = 3, Top = 3, Right = 5, Bottom = 3 },
            VerticalAlignment = VerticalAlignment.Center
        };

        CameraImage = new WriteableBitmap(CameraImageWidth, CameraImageHeight);
        InterfaceRoot = new Page
        {
            Content = new StackPanel
            {
                Children =
                {
                    new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        Children = { TiltTextBlock, TiltNumberBox }
                    }
                }
            }
        };

        PluginLoaded = true; // Mark as already-loaded
    }

    public void Initialize()
    {
        switch (InitializeKinect())
        {
            case 0:
                Host.Log($"Tried to initialize the Kinect sensor with status: {DeviceStatusString}");
                break;
            case 1:
                Host.Log($"Couldn't initialize the Kinect sensor! Status: {DeviceStatusString}", LogSeverity.Warning);
                break;
            default:
                Host.Log("Tried to initialize the Kinect, but a native exception occurred!", LogSeverity.Error);
                break;
        }

        if (IsInitialized && TiltNumberBox is not null)
            TiltNumberBox.Value = Math.Clamp(ElevationAngle, -27, 27);
    }

    public void Shutdown()
    {
        switch (ShutdownKinect())
        {
            case 0:
                Host.Log($"Tried to shutdown the Kinect sensor with status: {DeviceStatusString}");
                break;
            case 1:
                Host.Log($"Kinect sensor is already shut down! Status: {DeviceStatusString}", LogSeverity.Warning);
                break;
            case -2:
                Host.Log("Tried to shutdown the Kinect sensor, but a SEH exception occurred!", LogSeverity.Error);
                break;
            default:
                Host.Log("Tried to shutdown the Kinect sensor, but a native exception occurred!", LogSeverity.Error);
                break;
        }
    }

    public void Update()
    {
        // Update skeletal data
        var trackedJoints = GetTrackedKinectJoints();
        trackedJoints.ForEach(x =>
        {
            TrackedJoints[trackedJoints.IndexOf(x)].TrackingState =
                (TrackedJointState)x.TrackingState;

            TrackedJoints[trackedJoints.IndexOf(x)].Position = x.Position.Safe();
            TrackedJoints[trackedJoints.IndexOf(x)].Orientation = x.Orientation.Safe();
        });
        
        // Update camera feed
        if (!IsCameraEnabled) return;
        CameraImage.DispatcherQueue.TryEnqueue(async () =>
        {
            var buffer = GetImageBuffer(); // Read from Kinect
            if (buffer is null || buffer.Length <= 0) return;
            await CameraImage.PixelBuffer.AsStream().WriteAsync(buffer);
            CameraImage.Invalidate(); // Enqueue for preview refresh
        });
    }

    public void SignalJoint(int jointId)
    {
        // ignored
    }

    public override void StatusChangedHandler()
    {
        // The Kinect sensor requested a refresh
        InitializeKinect();

        if (IsInitialized) // Also refresh internal UI
            TiltNumberBox.DispatcherQueue.TryEnqueue(() =>
                TiltNumberBox.Value = Math.Clamp(ElevationAngle, -27, 27));

        // Request a refresh of the status UI
        Host?.RefreshStatusInterface();
    }

    public Func<BitmapSource> GetCameraImage => () => CameraImage;
    public Func<bool> GetIsCameraEnabled => () => IsCameraEnabled;
    public Action<bool> SetIsCameraEnabled => value => IsCameraEnabled = value;
    public Func<Vector3, Size> MapCoordinateDelegate => MapCoordinate;
}

internal static class Utils
{
    public static Quaternion Safe(this Quaternion q)
    {
        return (q.X is 0 && q.Y is 0 && q.Z is 0 && q.W is 0) ||
               float.IsNaN(q.X) || float.IsNaN(q.Y) || float.IsNaN(q.Z) || float.IsNaN(q.W)
            ? Quaternion.Identity // Return a placeholder quaternion
            : q; // If everything is fine, return the actual orientation
    }

    public static Vector3 Safe(this Vector3 v)
    {
        return float.IsNaN(v.X) || float.IsNaN(v.Y) || float.IsNaN(v.Z)
            ? Vector3.Zero // Return a placeholder position vector
            : v; // If everything is fine, return the actual orientation
    }
}