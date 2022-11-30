#pragma once

#pragma unmanaged
#include "KinectWrapper.h"
#pragma managed

using namespace System;
using namespace Numerics;
using namespace Collections::Generic;
using namespace ComponentModel::Composition;

using namespace Amethyst::Plugins::Contract;

namespace KinectHandler
{
	public ref class KinectJoint sealed
	{
	public:
		KinectJoint(const int role)
		{
			JointRole = role;
		}

		property Vector3 Position;
		property Quaternion Orientation;

		property int TrackingState;
		property int JointRole;
	};

	public ref class KinectHandler
	{
	private:
		KinectWrapper* kinect;

	public:
		List<KinectJoint^>^ GetTrackedKinectJoints()
		{
			if (!IsInitialized) return gcnew List<KinectJoint^>;

			const auto& positions = kinect->skeleton_positions();
			const auto& orientations = kinect->bone_orientations();
			const auto& states = kinect->tracking_states();

			auto trackedKinectJoints = gcnew List<KinectJoint^>;
			for each (auto v in Enum::GetValues<TrackedJointType>())
			{
				if (v == TrackedJointType::JointHandTipLeft ||
					v == TrackedJointType::JointHandTipRight ||
					v == TrackedJointType::JointThumbLeft ||
					v == TrackedJointType::JointThumbRight ||
					v == TrackedJointType::JointNeck ||
					v == TrackedJointType::JointManual)
					continue; // Skip unsupported joints

				auto joint = gcnew KinectJoint(static_cast<int>(v));

				joint->TrackingState =
					states[kinect->KinectJointType(static_cast<int>(v))];

				joint->Position = Vector3(
					positions[kinect->KinectJointType(static_cast<int>(v))].x,
					positions[kinect->KinectJointType(static_cast<int>(v))].y,
					positions[kinect->KinectJointType(static_cast<int>(v))].z);

				joint->Orientation = Quaternion(
					orientations[kinect->KinectJointType(static_cast<int>(v))].absoluteRotation.rotationQuaternion.x,
					orientations[kinect->KinectJointType(static_cast<int>(v))].absoluteRotation.rotationQuaternion.y,
					orientations[kinect->KinectJointType(static_cast<int>(v))].absoluteRotation.rotationQuaternion.z,
					orientations[kinect->KinectJointType(static_cast<int>(v))].absoluteRotation.rotationQuaternion.w);

				trackedKinectJoints->Add(joint);
			}

			return trackedKinectJoints;
		}

		property bool IsInitialized
		{
			bool get() { return kinect->is_initialized(); }
		}

		property bool IsSkeletonTracked
		{
			bool get() { return kinect->skeleton_tracked(); }
		}

		property int DeviceStatus
		{
			int get() { return kinect->status_result(); }
		}

		property int ElevationAngle
		{
			int get() { return kinect->elevation_angle(); }
			void set(const int value) { kinect->elevation_angle(value); }
		}

		property bool IsSettingsDaemonSupported
		{
			bool get() { return DeviceStatus == 0; }
		}

		KinectHandler() : kinect(new KinectWrapper())
		{
		}

		int InitializeKinect()
		{
			return kinect->initialize();
		}

		int ShutdownKinect()
		{
			return kinect->shutdown();
		}
	};
}
