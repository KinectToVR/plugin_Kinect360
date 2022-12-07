// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Reflection;
using Amethyst.Plugins.Contract;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace plugin_KinectV1;

public static class DeviceData
{
    public const string Name = "Xbox 360 Kinect";
    public const string Guid = "K2VRTEAM-AME2-APII-DVCE-DVCEKINECTV1";
}

[Export(typeof(ITrackingDevice))]
[ExportMetadata("Name", DeviceData.Name)]
[ExportMetadata("Guid", DeviceData.Guid)]
[ExportMetadata("Publisher", "K2VR Team")]
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

    public List<TrackedJoint> TrackedJoints { get; } =
        // Prepend all supported joints to the joints list
        Enum.GetValues<TrackedJointType>()
            .Where(x => x is not TrackedJointType.JointNeck and not TrackedJointType.JointManual and not
                TrackedJointType.JointHandTipLeft and not TrackedJointType.JointHandTipRight and not 
                TrackedJointType.JointThumbLeft and not TrackedJointType.JointThumbRight)
            .Select(x => new TrackedJoint { Name = x.ToString(), Role = x }).ToList();

    public string DeviceStatusString => PluginLoaded
        ? DeviceStatus switch
        {
            0 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/Success", DeviceData.Guid),
            1 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/Initializing", DeviceData.Guid),
            2 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotConnected", DeviceData.Guid),
            3 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotGenuine", DeviceData.Guid),
            4 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotSupported", DeviceData.Guid),
            5 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/InsufficientBandwidth", DeviceData.Guid),
            6 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotPowered", DeviceData.Guid),
            7 => Host.RequestLocalizedString("/Plugins/KinectV1/Statuses/NotReady", DeviceData.Guid),
            _ => $"Undefined: {DeviceStatus}\nE_UNDEFINED\nSomething weird has happened, though we can't tell what."
        }
        : $"Undefined: {DeviceStatus}\nE_UNDEFINED\nSomething weird has happened, though we can't tell what.";

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
        };

        TiltTextBlock = new TextBlock
        {
            Text = Host.RequestLocalizedString("/Plugins/KinectV1/Settings/Labels/Angle", DeviceData.Guid),
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

        PluginLoaded = true;
    }

    public void Initialize()
    {
        switch (InitializeKinect())
        {
            case 0:
                Host.Log($"[{DeviceData.Guid}] Tried to initialize the Kinect " +
                         $"sensor with status: {DeviceStatusString}", LogSeverity.Info);
                break;
            case 1:
                Host.Log($"[{DeviceData.Guid}] Couldn't initialize the Kinect " +
                         $"sensor! Status: {DeviceStatusString}", LogSeverity.Warning);
                break;
            default:
                Host.Log($"[{DeviceData.Guid}] Tried to initialize the Kinect, " +
                         "but a native exception occurred!", LogSeverity.Error);
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
                Host.Log($"[{DeviceData.Guid}] Tried to shutdown the Kinect " +
                         $"sensor with status: {DeviceStatusString}", LogSeverity.Info);
                break;
            case 1:
                Host.Log($"[{DeviceData.Guid}] Kinect sensor is already shut down! " +
                         $"Status: {DeviceStatusString}", LogSeverity.Warning);
                break;
            case -2:
                Host.Log($"[{DeviceData.Guid}] Tried to shutdown the Kinect sensor, " +
                         "but a SEH exception occurred!", LogSeverity.Error);
                break;
            default:
                Host.Log($"[{DeviceData.Guid}] Tried to shutdown the Kinect sensor, " +
                         "but a native exception occurred!", LogSeverity.Error);
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

    public override void StatusChangedHandler() 
    {
        // Request a refresh of the status UI
        Host?.RefreshStatusInterface();
    }

    public void SignalJoint(int jointId)
    {
        // ignored
    }
}