/*
  LICENSE
  -------
  Copyright (C) 2007 Ray Molenkamp
  This source code is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this source code or the software it produces.
  Permission is granted to anyone to use this source code for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:
  1. The origin of this source code must not be misrepresented; you must not
     claim that you wrote the original source code.  If you use this source code
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original source code.
  3. This notice may not be removed or altered from any source distribution.
*/
// updated for use in NAudio
using System;
using System.Runtime.InteropServices;
using NAudio.CoreAudioApi.Interfaces;

namespace NAudio.CoreAudioApi {

    /// <summary>
    /// MM Device Enumerator
    /// </summary>
    public class MMDeviceEnumerator : IDisposable {
        private IMMDeviceEnumerator realEnumerator;

        /// <summary>
        /// Creates a new MM Device Enumerator
        /// </summary>
        public MMDeviceEnumerator() {
            if ( System.Environment.OSVersion.Version.Major < 6 ) {
                throw new NotSupportedException("This functionality is only supported on Windows Vista or newer.");
            }
            realEnumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        }

        /// <summary>
        /// Enumerate Audio Endpoints
        /// </summary>
        /// <param name="dataFlow">Desired DataFlow</param>
        /// <param name="dwStateMask">State Mask</param>
        /// <returns>Device Collection</returns>
        public MMDeviceCollection EnumerateAudioEndPoints(DataFlow dataFlow, DeviceState dwStateMask) {
            Marshal.ThrowExceptionForHR(realEnumerator.EnumAudioEndpoints(dataFlow, dwStateMask, out var result));
            return new MMDeviceCollection(result);
        }

        /// <inheritdoc/>
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Called to dispose/finalize contained objects.
        /// </summary>
        /// <param name="disposing">True if disposing, false if called from a finalizer.</param>
        protected virtual void Dispose(bool disposing) {
            if ( disposing ) {
                if ( realEnumerator != null ) {
                    // although GC would do this for us, we want it done now
                    Marshal.ReleaseComObject(realEnumerator);
                    realEnumerator = null;
                }
            }
        }
    }
}
