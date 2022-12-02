#pragma once

#pragma unmanaged
#include "KinectWrapper.h"
#pragma managed

using namespace System;
using namespace Numerics;
using namespace Collections::Generic;
using namespace ComponentModel::Composition;
using namespace Runtime::InteropServices;

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

	delegate void FunctionToCallDelegate();

	public ref class KinectHandler
	{
	private:
		KinectWrapper* kinect_;
		FunctionToCallDelegate^ function_;

	public:
		KinectHandler() : kinect_(new KinectWrapper())
		{
			function_ = gcnew FunctionToCallDelegate(this, &KinectHandler::StatusChangedHandler);
			pin_ptr<FunctionToCallDelegate^> tmp = &function_; // Pin the function delegate

			status_changed_event = static_cast<void(__cdecl*)()>(
				Marshal::GetFunctionPointerForDelegate(function_).ToPointer());
		}
		
		virtual void StatusChangedHandler()
		{
			// implemented in the c# handler
		}

		List<KinectJoint^>^ GetTrackedKinectJoints()
		{
			if (!IsInitialized) return gcnew List<KinectJoint^>;

			const auto& positions = kinect_->skeleton_positions();
			const auto& orientations = kinect_->bone_orientations();
			const auto& states = kinect_->tracking_states();

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
					states[kinect_->KinectJointType(static_cast<int>(v))];

				joint->Position = Vector3(
					positions[kinect_->KinectJointType(static_cast<int>(v))].x,
					positions[kinect_->KinectJointType(static_cast<int>(v))].y,
					positions[kinect_->KinectJointType(static_cast<int>(v))].z);

				joint->Orientation = Quaternion(
					orientations[kinect_->KinectJointType(static_cast<int>(v))].absoluteRotation.rotationQuaternion.x,
					orientations[kinect_->KinectJointType(static_cast<int>(v))].absoluteRotation.rotationQuaternion.y,
					orientations[kinect_->KinectJointType(static_cast<int>(v))].absoluteRotation.rotationQuaternion.z,
					orientations[kinect_->KinectJointType(static_cast<int>(v))].absoluteRotation.rotationQuaternion.w);

				trackedKinectJoints->Add(joint);
			}

			return trackedKinectJoints;
		}

		property bool IsInitialized
		{
			bool get() { return kinect_->is_initialized(); }
		}

		property bool IsSkeletonTracked
		{
			bool get() { return kinect_->skeleton_tracked(); }
		}

		property int DeviceStatus
		{
			int get() { return kinect_->status_result(); }
		}

		property int ElevationAngle
		{
			int get() { return kinect_->elevation_angle(); }
			void set(const int value) { kinect_->elevation_angle(value); }
		}

		property bool IsSettingsDaemonSupported
		{
			bool get() { return DeviceStatus == 0; }
		}

		int InitializeKinect()
		{
			return kinect_->initialize();
		}

		int ShutdownKinect()
		{
			return kinect_->shutdown();
		}
	};
}
