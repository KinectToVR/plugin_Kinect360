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

class KinectWrapper
{
	INuiSensor* kinectSensor = nullptr;
	HANDLE kinectRGBStream = nullptr;
	HANDLE kinectDepthStream = nullptr;
	NUI_SKELETON_FRAME skeletonFrame = {0};

	std::array<NUI_SKELETON_BONE_ORIENTATION, NUI_SKELETON_POSITION_COUNT> bone_orientations_;
	std::array<_Vector4, NUI_SKELETON_POSITION_COUNT> skeleton_positions_;
	std::array<int, NUI_SKELETON_POSITION_COUNT> tracking_states_;

	std::unique_ptr<std::thread> updater_thread_;

	bool initialized_ = false;
	bool skeleton_tracked_ = false;

	void updater()
	{
		// Auto-handles failures & etc
		while (true)update();
	}

	void updateSkeletalData()
	{
		// Wait for a frame to arrive, give up after >3s of nothing
		if (kinectSensor->NuiSkeletonGetNextFrame(3000, &skeletonFrame) >= 0)
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
		//Get a working Kinect Sensor
		int numSensors = 0;
		if (NuiGetSensorCount(&numSensors) < 0 || numSensors < 1)
		{
			return false;
		}
		if (NuiCreateSensorByIndex(0, &kinectSensor) < 0)
		{
			return false;
		}
		//Initialize Sensor
		HRESULT hr = kinectSensor->NuiInitialize(NUI_INITIALIZE_FLAG_USES_SKELETON);
		kinectSensor->NuiSkeletonTrackingEnable(nullptr, 0); //NUI_SKELETON_TRACKING_FLAG_ENABLE_IN_NEAR_RANGE

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
		JointAnkleLeft,
		JointFootLeft,
		JointHipRight,
		JointKneeRight,
		JointAnkleRight,
		JointFootRight,
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
		{JointAnkleLeft, NUI_SKELETON_POSITION_ANKLE_LEFT},
		{JointFootLeft, NUI_SKELETON_POSITION_FOOT_LEFT},
		{JointHipRight, NUI_SKELETON_POSITION_HIP_RIGHT},
		{JointKneeRight, NUI_SKELETON_POSITION_KNEE_RIGHT},
		{JointAnkleRight, NUI_SKELETON_POSITION_ANKLE_RIGHT},
		{JointFootRight, NUI_SKELETON_POSITION_FOOT_RIGHT}
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
		if (initialized_ && kinectSensor->NuiStatus() == S_OK)
			updateSkeletalData();
	}

	int shutdown()
	{
		try
		{
			// Shut down the sensor (Only NUI API)
			if (kinectSensor) // Protect from null call
				[&, this]
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

	int KinectJointType(int kinectJointType)
	{
		return KinectJointTypeDictionary.at(static_cast<JointType>(kinectJointType));
	}
};
