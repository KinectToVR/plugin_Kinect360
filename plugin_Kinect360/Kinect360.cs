// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Amethyst.Plugins.Contract;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
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
    public static IAmethystHost HostStatic { get; set; }

    private readonly GestureDetector
        _pauseDetectorLeft = new(),
        _pauseDetectorRight = new(),
        _pointDetectorLeft = new(),
        _pointDetectorRight = new();

    public ObservableCollection<TrackedJoint> TrackedJoints { get; } =
        // Prepend all supported joints to the joints list
        new(Enum.GetValues<TrackedJointType>()
            .Where(x => x is not TrackedJointType.JointNeck and not TrackedJointType.JointManual and not
                TrackedJointType.JointHandTipLeft and not TrackedJointType.JointHandTipRight and not
                TrackedJointType.JointThumbLeft and not TrackedJointType.JointThumbRight)
            .Select(x => new TrackedJoint
            {
                Name = x.ToString(), Role = x, SupportedInputActions = x switch
                {
                    TrackedJointType.JointHandLeft =>
                    [
                        new KeyInputAction<bool>
                        {
                            Name = "Left Pause", Description = "Left hand pause gesture",
                            Guid = "PauseLeft_360", GetHost = () => HostStatic
                        },
                        new KeyInputAction<bool>
                        {
                            Name = "Left Point", Description = "Left hand point gesture",
                            Guid = "PointLeft_360", GetHost = () => HostStatic
                        }
                    ],
                    TrackedJointType.JointHandRight =>
                    [
                        new KeyInputAction<bool>
                        {
                            Name = "Right Pause", Description = "Right hand pause gesture",
                            Guid = "PauseRight_360", GetHost = () => HostStatic
                        },
                        new KeyInputAction<bool>
                        {
                            Name = "Right Point", Description = "Right hand point gesture",
                            Guid = "PointRight_360", GetHost = () => HostStatic
                        }
                    ],
                    _ => []
                }
            })
        );

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
        // Backup the plugin host
        HostStatic = Host;

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

        try
        {
            // Re-generate joint names
            lock (Host.UpdateThreadLock)
            {
                for (var i = 0; i < TrackedJoints.Count; i++)
                {
                    TrackedJoints[i] = TrackedJoints[i].WithName(Host?.RequestLocalizedString(
                        $"/JointsEnum/{TrackedJoints[i].Role.ToString()}") ?? TrackedJoints[i].Role.ToString());

                    foreach (var action in TrackedJoints[i].SupportedInputActions)
                    {
                        action.Name = Host!.RequestLocalizedString($"/InputActions/Names/{action.Guid.Replace("_360", "")}");
                        action.Description = Host.RequestLocalizedString($"/InputActions/Descriptions/{action.Guid.Replace("_360", "")}");

                        action.Image = new Image
                        {
                            Source = new BitmapImage(new Uri($"ms-appx:///{Path.Join(
                                Directory.GetParent(Assembly.GetExecutingAssembly().Location)!.FullName,
                                "Assets", "Resources", "Icons", $"{(((dynamic)Host).IsDarkMode as bool? ?? false ? "D" : "W")}" +
                                                                $"_{action.Guid.Replace("_360", "")}.png")}"))
                        };
                    }
                }
            }
        }
        catch (Exception e)
        {
            Host?.Log($"Error setting joint names! Message: {e.Message}", LogSeverity.Error);
        }

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
        if (IsCameraEnabled)
            CameraImage.DispatcherQueue.TryEnqueue(async () =>
            {
                var buffer = GetImageBuffer(); // Read from Kinect
                if (buffer is null || buffer.Length <= 0) return;
                await CameraImage.PixelBuffer.AsStream().WriteAsync(buffer);
                CameraImage.Invalidate(); // Enqueue for preview refresh
            });

        // Update gestures
        if (trackedJoints.Count != 20) return;

        try
        {
            var shoulderLeft = trackedJoints[(int)NativeKinectJoints.JointShoulderLeft].Position;
            var shoulderRight = trackedJoints[(int)NativeKinectJoints.JointShoulderRight].Position;
            var elbowLeft = trackedJoints[(int)NativeKinectJoints.JointElbowLeft].Position;
            var elbowRight = trackedJoints[(int)NativeKinectJoints.JointElbowRight].Position;
            var handLeft = trackedJoints[(int)NativeKinectJoints.JointWristLeft].Position;
            var handRight = trackedJoints[(int)NativeKinectJoints.JointWristRight].Position;

            // >0.9f when elbow is not bent and the arm is straight : LEFT
            var armDotLeft = Vector3.Dot(
                Vector3.Normalize(elbowLeft - shoulderLeft),
                Vector3.Normalize(handLeft - elbowLeft));

            // >0.9f when the arm is pointing down : LEFT
            var armDownDotLeft = Vector3.Dot(
                new Vector3(0.0f, -1.0f, 0.0f),
                Vector3.Normalize(handLeft - elbowLeft));

            // >0.4f <0.6f when the arm is slightly tilted sideways : RIGHT
            var armTiltDotLeft = Vector3.Dot(
                Vector3.Normalize(shoulderLeft - shoulderRight),
                Vector3.Normalize(handLeft - elbowLeft));

            // >0.9f when elbow is not bent and the arm is straight : LEFT
            var armDotRight = Vector3.Dot(
                Vector3.Normalize(elbowRight - shoulderRight),
                Vector3.Normalize(handRight - elbowRight));

            // >0.9f when the arm is pointing down : RIGHT
            var armDownDotRight = Vector3.Dot(
                new Vector3(0.0f, -1.0f, 0.0f),
                Vector3.Normalize(handRight - elbowRight));

            // >0.4f <0.6f when the arm is slightly tilted sideways : RIGHT
            var armTiltDotRight = Vector3.Dot(
                Vector3.Normalize(shoulderRight - shoulderLeft),
                Vector3.Normalize(handRight - elbowRight));

            /* Trigger the detected gestures */

            if (TrackedJoints[(int)NativeKinectJoints.JointHandLeft].SupportedInputActions.IsUsed(0, out var pauseActionLeft))
                Host.ReceiveKeyInput(pauseActionLeft, _pauseDetectorLeft.Update(armDotLeft > 0.9f && armTiltDotLeft is > 0.4f and < 0.7f));

            if (TrackedJoints[(int)NativeKinectJoints.JointHandRight].SupportedInputActions.IsUsed(0, out var pauseActionRight))
                Host.ReceiveKeyInput(pauseActionRight, _pauseDetectorRight.Update(armDotRight > 0.9f && armTiltDotRight is > 0.4f and < 0.7f));

            if (TrackedJoints[(int)NativeKinectJoints.JointHandLeft].SupportedInputActions.IsUsed(1, out var pointActionLeft))
                Host.ReceiveKeyInput(pointActionLeft, _pointDetectorLeft
                    .Update(armDotLeft > 0.9f && armTiltDotLeft is > -0.5f and < 0.5f && armDownDotLeft is > -0.3f and < 0.7f));

            if (TrackedJoints[(int)NativeKinectJoints.JointHandRight].SupportedInputActions.IsUsed(1, out var pointActionRight))
                Host.ReceiveKeyInput(pointActionRight, _pointDetectorRight
                    .Update(armDotRight > 0.9f && armTiltDotRight is > -0.5f and < 0.5f && armDownDotRight is > -0.3f and < 0.7f));

            /* Trigger the detected gestures */
        }
        catch (Exception ex)
        {
            Host?.Log(ex);
        }
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

    public static T At<T>(this SortedSet<T> set, int at)
    {
        return set.ElementAt(at);
    }

    public static bool At<T>(this SortedSet<T> set, int at, out T result)
    {
        try
        {
            result = set.ElementAt(at);
        }
        catch
        {
            result = default;
            return false;
        }

        return true;
    }

    public static bool IsUsed(this SortedSet<IKeyInputAction> set, int at)
    {
        return set.At(at, out var action) && (Kinect360.HostStatic?.CheckInputActionIsUsed(action) ?? false);
    }

    public static bool IsUsed(this SortedSet<IKeyInputAction> set, int at, out IKeyInputAction action)
    {
        return set.At(at, out action) && (Kinect360.HostStatic?.CheckInputActionIsUsed(action) ?? false);
    }

    public static TrackedJoint WithName(this TrackedJoint joint, string name)
    {
        return new TrackedJoint
        {
            Name = name,
            Role = joint.Role,
            Acceleration = joint.Acceleration,
            AngularAcceleration = joint.AngularAcceleration,
            AngularVelocity = joint.AngularVelocity,
            Orientation = joint.Orientation,
            Position = joint.Position,
            SupportedInputActions = joint.SupportedInputActions,
            TrackingState = joint.TrackingState,
            Velocity = joint.Velocity
        };
    }
}

internal enum NativeKinectJoints
{
    JointHead,
    JointSpineShoulder,
    JointShoulderLeft,
    JointElbowLeft,
    JointWristLeft,
    JointHandLeft,
    JointShoulderRight,
    JointElbowRight,
    JointWristRight,
    JointHandRight,
    JointSpineMiddle,
    JointSpineWaist,
    JointHipLeft,
    JointKneeLeft,
    JointFootLeft,
    JointFootTipLeft,
    JointHipRight,
    JointKneeRight,
    JointFootRight,
    JointFootTipRight
}