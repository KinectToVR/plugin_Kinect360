#pragma once
#include <Windows.h>
#include <Ole2.h>
#include <NuiApi.h>
#include <NuiImageCamera.h>
#include <NuiSensor.h>
#include <NuiSkeleton.h>

#include <algorithm>
#include <iterator>
#include <memory>
#include <thread>
#include <array>
#include <map>
#include <functional>

inline void (*status_changed_event)();

class KinectWrapper
{
    inline static INuiSensor* kinectSensor = nullptr;
    HANDLE kinectRGBStream = nullptr;
    HANDLE kinectDepthStream = nullptr;
    NUI_SKELETON_FRAME skeletonFrame = {0};

    std::array<NUI_SKELETON_BONE_ORIENTATION, NUI_SKELETON_POSITION_COUNT> bone_orientations_;
    std::array<_Vector4, NUI_SKELETON_POSITION_COUNT> skeleton_positions_;
    std::array<int, NUI_SKELETON_POSITION_COUNT> tracking_states_;

    std::unique_ptr<std::thread> updater_thread_;

    inline static bool initialized_ = false;
    bool skeleton_tracked_ = false;
    bool rgb_stream_enabled_ = false;

    void updater()
    {
        // Auto-handles failures & etc
        while (true)update();
    }

    void updateSensorData()
    {
        NUI_IMAGE_FRAME imageFrame{};

        // Wait for a frame to arrive, give up if nothing
        if (rgb_stream_enabled_ && kinectSensor &&
            kinectSensor->NuiImageStreamGetNextFrame(kinectRGBStream, 0, &imageFrame) >= 0)
        {
            INuiFrameTexture* pTexture = imageFrame.pFrameTexture;

            // Lock the frame data so the Kinect knows not to modify it while we are reading it
            NUI_LOCKED_RECT lockedRect;
            pTexture->LockRect(0, &lockedRect, nullptr, 0);

            // Make sure we've received valid data
            size_in_bytes_last_ = 0;
            if (lockedRect.Pitch != 0)
            {
                ResetBuffer(CameraBufferSize()); // Allocate buffer for image, copy to buffer
                memcpy_s(color_buffer_, size_in_bytes_, lockedRect.pBits, lockedRect.size);
                size_in_bytes_last_ = size_in_bytes_; // Backup for the bitmap generator
            }

            pTexture->UnlockRect(0); // Unlock frame data and proceed
            kinectSensor->NuiImageStreamReleaseFrame(kinectRGBStream, &imageFrame);
        }

        // Wait for a frame to arrive, give up if nothing
        if (kinectSensor && kinectSensor->NuiSkeletonGetNextFrame(0, &skeletonFrame) >= 0)
            for (auto& i : skeletonFrame.SkeletonData)
            {
                if (NUI_SKELETON_TRACKED == i.eTrackingState)
                {
                    skeleton_tracked_ = true; // We've got it!

                    // Copy joint positions & orientations
                    std::copy(std::begin(i.SkeletonPositions), std::end(i.SkeletonPositions),
                              skeleton_positions_.begin());
                    std::copy(std::begin(i.eSkeletonPositionTrackingState), std::end(i.eSkeletonPositionTrackingState),
                              tracking_states_.begin());

                    (void)NuiSkeletonCalculateBoneOrientations(&i, bone_orientations_.data());

                    break; // Only the first skeleton
                }
                skeleton_tracked_ = false;
            }
    }

    bool initKinect()
    {
        // Register a StatusChanged event
        NuiSetDeviceStatusCallback(&statusCallbackEvent, nullptr);

        // Get a working Kinect Sensor
        int numSensors = 0;
        if (NuiGetSensorCount(&numSensors) < 0 || numSensors < 1)
        {
            return false;
        }
        if (NuiCreateSensorByIndex(0, &kinectSensor) < 0)
        {
            return false;
        }

        // Check the sensor (just in case)
        if (kinectSensor == nullptr) return false;

        // Initialize Sensor
        HRESULT hr = kinectSensor->NuiInitialize(NUI_INITIALIZE_FLAG_USES_SKELETON | NUI_INITIALIZE_FLAG_USES_COLOR);

        kinectSensor->NuiImageStreamOpen(NUI_IMAGE_TYPE_COLOR, resolution_,
                                         0, 2, nullptr, &kinectRGBStream);

        kinectSensor->NuiSkeletonTrackingEnable(nullptr, 0);

        return kinectSensor;
    }

    bool acquireKinectFrame(NUI_IMAGE_FRAME& imageFrame, HANDLE& rgbStream, INuiSensor*& sensor)
    {
        return (sensor->NuiImageStreamGetNextFrame(rgbStream, 1, &imageFrame) < 0);
    }

    void releaseKinectFrame(NUI_IMAGE_FRAME& imageFrame, HANDLE& rgbStream, INuiSensor*& sensor)
    {
        sensor->NuiImageStreamReleaseFrame(rgbStream, &imageFrame);
    }

    static void CALLBACK statusCallbackEvent(
        HRESULT hrStatus, const OLECHAR* instanceName,
        const OLECHAR* uniqueDeviceName, void* pUserData)
    {
        // The new status isn't "not connected"
        // Try opening the newly acquired sensor
        if (E_NUI_NOTCONNECTED != hrStatus)
            NuiCreateSensorById(instanceName, &kinectSensor);

        // Notify the CLR listener
        status_changed_event();
    }

    HRESULT kinect_status_result()
    {
        if (kinectSensor)
            return kinectSensor->NuiStatus();

        return E_NUI_NOTCONNECTED;
    }

    enum JointType
    {
        JointHead,
        JointNeck,
        JointSpineShoulder,
        JointShoulderLeft,
        JointElbowLeft,
        JointWristLeft,
        JointHandLeft,
        JointHandTipLeft,
        JointThumbLeft,
        JointShoulderRight,
        JointElbowRight,
        JointWristRight,
        JointHandRight,
        JointHandTipRight,
        JointThumbRight,
        JointSpineMiddle,
        JointSpineWaist,
        JointHipLeft,
        JointKneeLeft,
        JointFootLeft,
        JointFootTipLeft,
        JointHipRight,
        JointKneeRight,
        JointFootRight,
        JointFootTipRight,
        JointManual
    };

    std::map<JointType, _NUI_SKELETON_POSITION_INDEX> KinectJointTypeDictionary
    {
        {JointSpineWaist, NUI_SKELETON_POSITION_HIP_CENTER},
        {JointSpineMiddle, NUI_SKELETON_POSITION_SPINE},
        {JointSpineShoulder, NUI_SKELETON_POSITION_SHOULDER_CENTER},
        {JointHead, NUI_SKELETON_POSITION_HEAD},
        {JointShoulderLeft, NUI_SKELETON_POSITION_SHOULDER_LEFT},
        {JointElbowLeft, NUI_SKELETON_POSITION_ELBOW_LEFT},
        {JointWristLeft, NUI_SKELETON_POSITION_WRIST_LEFT},
        {JointHandLeft, NUI_SKELETON_POSITION_HAND_LEFT},
        {JointShoulderRight, NUI_SKELETON_POSITION_SHOULDER_RIGHT},
        {JointElbowRight, NUI_SKELETON_POSITION_ELBOW_RIGHT},
        {JointWristRight, NUI_SKELETON_POSITION_WRIST_RIGHT},
        {JointHandRight, NUI_SKELETON_POSITION_HAND_RIGHT},
        {JointHipLeft, NUI_SKELETON_POSITION_HIP_LEFT},
        {JointKneeLeft, NUI_SKELETON_POSITION_KNEE_LEFT},
        {JointFootLeft, NUI_SKELETON_POSITION_ANKLE_LEFT},
        {JointFootTipLeft, NUI_SKELETON_POSITION_FOOT_LEFT},
        {JointHipRight, NUI_SKELETON_POSITION_HIP_RIGHT},
        {JointKneeRight, NUI_SKELETON_POSITION_KNEE_RIGHT},
        {JointFootRight, NUI_SKELETON_POSITION_ANKLE_RIGHT},
        {JointFootTipRight, NUI_SKELETON_POSITION_FOOT_RIGHT}
    };

public:
    bool is_initialized()
    {
        return initialized_;
    }

    HRESULT status_result()
    {
        switch (kinect_status_result())
        {
        case S_OK: return 0;
        case S_NUI_INITIALIZING: return 1;
        case E_NUI_NOTCONNECTED: return 2;
        case E_NUI_NOTGENUINE: return 3;
        case E_NUI_NOTSUPPORTED: return 4;
        case E_NUI_INSUFFICIENTBANDWIDTH: return 5;
        case E_NUI_NOTPOWERED: return 6;
        case E_NUI_NOTREADY: return 7;
        default: return -1;
        }
    }

    int initialize()
    {
        try
        {
            initialized_ = initKinect();
            if (!initialized_) return 1;

            // Recreate the updater thread
            if (!updater_thread_)
                updater_thread_.reset(new std::thread(&KinectWrapper::updater, this));

            return 0; // OK
        }
        catch (...)
        {
            return -1;
        }
    }

    void update()
    {
        // Update (only if the sensor is on and its status is ok)
        if (initialized_ && kinectSensor &&
            kinectSensor->NuiStatus() == S_OK)
            updateSensorData();
    }

    int shutdown()
    {
        try
        {
            // Shut down the sensor (Only NUI API)
            if (kinectSensor) // Protect from null call
                return [&, this]
                {
                    __try
                    {
                        kinectSensor->NuiShutdown();

                        initialized_ = false;
                        kinectSensor = nullptr;

                        return 0;
                    }
                    __except (EXCEPTION_EXECUTE_HANDLER)
                    {
                        return -2;
                    }
                }();

            return 1;
        }
        catch (...)
        {
            return -1;
        }
    }

    std::array<NUI_SKELETON_BONE_ORIENTATION, NUI_SKELETON_POSITION_COUNT> bone_orientations()
    {
        return bone_orientations_;
    }

    std::array<_Vector4, NUI_SKELETON_POSITION_COUNT> skeleton_positions()
    {
        return skeleton_positions_;
    }

    std::array<int, NUI_SKELETON_POSITION_COUNT> tracking_states()
    {
        return tracking_states_;
    }

    std::tuple<BYTE*, int> color_buffer()
    {
        return std::make_tuple(color_buffer_, size_in_bytes_last_);
    }

    bool skeleton_tracked()
    {
        return skeleton_tracked_;
    }

    HRESULT elevation_angle(long angle)
    {
        return NuiCameraElevationSetAngle(angle);
    }

    long elevation_angle(void)
    {
        long angle = 0; // Placeholder

        (void)NuiCameraElevationGetAngle(&angle);
        return angle;
    }

    void camera_enabled(bool enabled)
    {
        rgb_stream_enabled_ = enabled;
    }

    bool camera_enabled(void)
    {
        return rgb_stream_enabled_;
    }

    int KinectJointType(int kinectJointType)
    {
        return KinectJointTypeDictionary.at(static_cast<JointType>(kinectJointType));
    }

    std::pair<int, int> CameraImageSize()
    {
        switch (resolution_)
        {
        case NUI_IMAGE_RESOLUTION_1280x960:
            return std::make_pair(1280, 760);
        case NUI_IMAGE_RESOLUTION_640x480:
            return std::make_pair(640, 480);
        case NUI_IMAGE_RESOLUTION_320x240:
            return std::make_pair(320, 240);
        case NUI_IMAGE_RESOLUTION_80x60:
            return std::make_pair(80, 60);
        default:
            return std::make_pair(1280, 760);
        }
    }

    unsigned long CameraBufferSize()
    {
        const auto& [width,height] = CameraImageSize();
        return width * height * 4;
    }


    std::pair<int, int> MapCoordinate(const _Vector4& skeletonPoint)
    {
        LONG x = 0, y = 0;
        LONG backupX = x, backupY = y;

        USHORT depthValue = 0;
        const auto& [width, height] = CameraImageSize();
        NuiTransformSkeletonToDepthImage(skeletonPoint, &x, &y, &depthValue, resolution_);

        if (FAILED(NuiImageGetColorPixelCoordinatesFromDepthPixelAtResolution(
            resolution_, resolution_, nullptr,
            x, y, depthValue, &x, &y)))
        {
            x = backupX;
            y = backupY;
        }

        // return std::make_pair(1000 * x / width, 1000 * y / width);
        return std::make_pair(x, y);
    }

private:
    DWORD size_in_bytes_ = 0;
    DWORD size_in_bytes_last_ = 0;
    BYTE* color_buffer_ = nullptr;

    _NUI_IMAGE_RESOLUTION resolution_ =
        NUI_IMAGE_RESOLUTION_640x480;

    BYTE* ResetBuffer(UINT size)
    {
        if (!color_buffer_ || size_in_bytes_ != size)
        {
            if (color_buffer_)
            {
                delete[] color_buffer_;
                color_buffer_ = nullptr;
            }

            if (0 != size)
            {
                color_buffer_ = new BYTE[size];
            }
            size_in_bytes_ = size;
        }

        return color_buffer_;
    }
};
