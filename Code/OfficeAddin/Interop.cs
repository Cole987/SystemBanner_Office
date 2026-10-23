//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SystemBannerOffice
{
    //the COM interfaces Office uses to load add-ins, declared here so the add-in needs no primary interop assembly:
    //Office 2000 has none, and one build must load in every version from Office 2000 to 2010
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode, [MarshalAs(UnmanagedType.IDispatch)] object addInInst, ref Array custom);
        [DispId(2)] void OnDisconnection(int removeMode, ref Array custom);
        [DispId(3)] void OnAddInsUpdate(ref Array custom);
        [DispId(4)] void OnStartupComplete(ref Array custom);
        [DispId(5)] void OnBeginShutdown(ref Array custom);
    }

    //Office 2007 and later ask the add-in for its ribbon through this interface
    [ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] string GetCustomUI(string ribbonId);
    }

    //Office 2000 to 2003 toolbar buttons raise their Click event through this dispinterface
    [ComImport, Guid("000C0351-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface CommandBarButtonEvents
    {
        [DispId(1)] void Click([MarshalAs(UnmanagedType.IDispatch)] object ctrl, ref bool cancelDefault);
    }

    //late binding helpers: every Office object is used through IDispatch, so the same calls work from Office 2000 to 2010
    public static class Late
    {
        public static object Get(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, args);
        }

        public static void Set(object target, string name, params object[] args)
        {
            target.GetType().InvokeMember(name, BindingFlags.SetProperty, null, target, args);
        }

        public static object Call(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args);
        }

        //an item of a collection: Item is a method in Word and PowerPoint and a property in Excel, both flags cover the two
        public static object Item(object collection, object index)
        {
            return collection.GetType().InvokeMember("Item", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, collection, new object[] { index });
        }

        //null instead of an exception when the property does not exist in this version of Office
        public static object TryGet(object target, string name, params object[] args)
        {
            try
            {
                return Get(target, name, args);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool TrySet(object target, string name, params object[] args)
        {
            try
            {
                Set(target, name, args);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static int Count(object collection)
        {
            return Convert.ToInt32(Get(collection, "Count"));
        }

        public static void Release(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }

        //Office colors are BGR integers
        public static int Rgb(int red, int green, int blue)
        {
            return red | (green << 8) | (blue << 16);
        }
    }

    //subscribes to a COM event source (a toolbar button, the Word application...) without an interop assembly
    public class ComEventLink : IDisposable
    {
        private System.Runtime.InteropServices.ComTypes.IConnectionPoint point;
        private int cookie;

        public ComEventLink(object source, Guid eventsInterface, object sink)
        {
            System.Runtime.InteropServices.ComTypes.IConnectionPointContainer container = (System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)source;
            container.FindConnectionPoint(ref eventsInterface, out point);
            point.Advise(sink, out cookie);
        }

        public void Dispose()
        {
            if (point != null)
            {
                try
                {
                    point.Unadvise(cookie);
                }
                catch (Exception)
                {
                    //the source is already gone when Office shuts down
                }
                point = null;
            }
        }
    }
}
