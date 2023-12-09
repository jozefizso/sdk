// Copyright (c) .NET Foundation and contributors. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ComAddin
{
    [ComVisible(true)]
    [Guid("D8494D38-D995-4670-AFF9-9425ED71D657")]
    [ProgId("ComAddin.AddinActivation")]
    public class AddinActivation : IDTExtensibility2
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
        }

        public void OnDisconnection([In] ext_DisconnectMode removeMode, [In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }

        public void OnAddInsUpdate([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }

        public void OnStartupComplete([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }

        public void OnBeginShutdown([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }
    }

    /// <summary>
    /// IDTExtensibility2 contains methods that act as interface between Microsoft Office applications and the add-in.
    /// Microsoft Office applications call these methods whenever an event that affects an add-in occurs,
    /// such as when it is loaded or unloaded.
    /// </summary>
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)]
    public interface IDTExtensibility2
    {
        /// <summary>
        /// Occurs whenever an add-in is loaded into Microsoft Office application.
        /// </summary>
        /// <param name="application">A reference to an instance of the office application</param>
        /// <param name="connectMode">An ext_ConnectMode enumeration value that indicates the way the add-in was loaded into MS-Office</param>
        /// <param name="addInInst">An AddIn reference to the add-in's own instance. This is stored for later use, such as determining the parent collection for the add-in</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        [DispId(1)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnConnection([MarshalAs(26)][In] object application, [In] ext_ConnectMode connectMode, [MarshalAs(26)][In] object addInInst, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        /// <summary>
        /// Occurs whenever an add-in is unloaded from Microsoft Office application.
        /// </summary>
        /// <param name="removeMode">An ext_DisconnectMode enumeration value that informs an add-in why it was unloaded.</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use after the add-in unloads</param>
        [DispId(2)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnDisconnection([In] ext_DisconnectMode removeMode, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        /// <summary>
        /// Occurs whenever an add-in is loaded or unloaded Microsoft Office.
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        [DispId(3)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnAddInsUpdate([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        /// <summary>
        ///  Occurs whenever an add-in, which is set to load when Microsoft Office application starts, loads.
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use when the add-in loads</param>
        [DispId(4)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnStartupComplete([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        /// <summary>
        /// Occurs whenever Microsoft Office application shuts down while an add-in is running.
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        [DispId(5)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnBeginShutdown([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);
    }

    [Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_ConnectMode
    {
        /// <summary>
        /// The add-in was loaded after Application started.
        /// </summary>
        ext_cm_AfterStartup = 0,

        /// <summary>
        /// The add-in was loaded when Application started.
        /// </summary>
        ext_cm_Startup = 1,

        /// <summary>
        /// The add-in was loaded by an external client.
        /// </summary>
        ext_cm_External = 2,

        /// <summary>
        /// The add-in was loaded from the command line.
        /// </summary>
        ext_cm_CommandLine = 3,

        /// <summary>
        /// The add-in was loaded with a solution.
        /// </summary>
        ext_cm_Solution = 4,

        /// <summary>
        /// The add-in was loaded for user interface setup.
        /// </summary>
        ext_cm_UISetup = 5
    }

    [Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_DisconnectMode
    {
        /// <summary>
        /// The add-in was unloaded when Application was shut down.
        /// </summary>
        ext_dm_HostShutdown = 0,

        /// <summary>
        /// The add-in was unloaded while Application was running.
        /// </summary>
        ext_dm_UserClosed = 1,

        /// <summary>
        /// The add-in was unloaded after the user interface was set up.
        /// </summary>
        ext_dm_UISetupComplete = 2,

        /// <summary>
        /// The add-in was unloaded when the solution was closed.
        /// </summary>
        ext_dm_SolutionClosed = 3
    }
}
