// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See LICENSE in the project root for license information.

using System;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Linq;
using Amethyst.Plugins.Contract;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace plugin_KinectV1;

[Export(typeof(ITrackingDevice))]
[ExportMetadata("Name", "Xbox 360 Kinect")]
[ExportMetadata("Guid", "K2VRTEAM-AME2-APII-DVCE-DVCEKINECTV1")]
[ExportMetadata("Publisher", "K2VR Team")]
[ExportMetadata("Version", "1.0.0.0")]
[ExportMetadata("Website", "https://github.com/KinectToVR/plugin_KinectV1")]
public class KinectV1 : KinectHandler.KinectHandler, ITrackingDevice
{
    [Import(typeof(IAmethystHost))] private IAmethystHost Host { get; set; }

    private bool PluginLoaded { get; set; }
    private Page InterfaceRoot { get; set; }
    private NumberBox TiltNumberBox { get; set; }
    private TextBlock TiltTextBlock { get; set; }

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
            0 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/Success"),
            1 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/Initializing"),
            2 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotConnected"),
            3 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotGenuine"),
            4 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotSupported"),
            5 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/InsufficientBandwidth"),
            6 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotPowered"),
            7 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotReady"),
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
            Text = Host.RequestLocalizedString("/Plugins/KinectV1/Settings/Labels/Angle"),
            Margin = new Thickness { Left = 3, Top = 3, Right = 5, Bottom = 3 },
            VerticalAlignment = VerticalAlignment.Center
        };

        InterfaceRoot = new Page
        {
            Content = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Children = { TiltTextBlock, TiltNumberBox }
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
        var trackedJoints = GetTrackedKinectJoints();
        trackedJoints.ForEach(x =>
        {
            TrackedJoints[trackedJoints.IndexOf(x)].TrackingState =
                (TrackedJointState)x.TrackingState;

            TrackedJoints[trackedJoints.IndexOf(x)].Position = x.Position;
            TrackedJoints[trackedJoints.IndexOf(x)].Orientation = x.Orientation;
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
}