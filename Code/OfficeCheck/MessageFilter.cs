//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Runtime.InteropServices;

namespace SystemBannerOffice.Check
{
    //retries the automation calls an Office application refuses while it is busy (RPC_E_SERVERCALL_RETRYLATER),
    //the usual COM message filter of automation clients; the add-in runs inside Office and needs none
    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr interfaceInfo);
        [PreserveSig] int RetryRejectedCall(IntPtr callee, int tickCount, int rejectType);
        [PreserveSig] int MessagePending(IntPtr callee, int tickCount, int pendingType);
    }

    internal class MessageFilter : IOleMessageFilter
    {
        private const int ServerCallRejected = 1; //PowerPoint also rejects calls while it starts or closes
        private const int ServerCallRetryLater = 2;
        private const int GiveUpAfterMilliseconds = 30000;

        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        public static void Register()
        {
            IOleMessageFilter old;
            CoRegisterMessageFilter(new MessageFilter(), out old);
        }

        public static void Revoke()
        {
            IOleMessageFilter old;
            CoRegisterMessageFilter(null, out old);
        }

        public int HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr interfaceInfo)
        {
            return 0; //SERVERCALL_ISHANDLED
        }

        public int RetryRejectedCall(IntPtr callee, int tickCount, int rejectType)
        {
            if ((rejectType == ServerCallRetryLater || rejectType == ServerCallRejected) && tickCount < GiveUpAfterMilliseconds)
            {
                return 100; //retry after 100 ms
            }
            return -1; //give up
        }

        public int MessagePending(IntPtr callee, int tickCount, int pendingType)
        {
            return 2; //PENDINGMSG_WAITDEFPROCESS
        }
    }
}
