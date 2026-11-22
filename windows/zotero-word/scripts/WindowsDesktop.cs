using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

public static class WindowsDesktop {
    [StructLayout(LayoutKind.Sequential)]
    struct FileIdentity {
        public uint Attributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME Created, Accessed, Written;
        public uint Volume, SizeHigh, SizeLow, Links, IndexHigh, IndexLow;
    }
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern bool GetFileInformationByHandle(SafeFileHandle file, out FileIdentity id);
    public static string FileId(string path, bool independent) {
        using (var file = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
            FileIdentity id;
            if (!GetFileInformationByHandle(file.SafeFileHandle, out id)) throw new Win32Exception();
            if (independent && id.Links != 1) throw new IOException("Working copy has multiple hard links.");
            return String.Format("{0:X8}:{1:X8}:{2:X8}", id.Volume, id.IndexHigh, id.IndexLow);
        }
    }

    [StructLayout(LayoutKind.Sequential)] struct Point { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)] struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetClipboardSequenceNumber();
    [DllImport("user32.dll")] public static extern IntPtr GetClipboardOwner();
    [StructLayout(LayoutKind.Sequential)] struct LastInput { public uint Size, Tick; }
    [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LastInput input);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern bool SetProp(IntPtr h, string name, IntPtr value);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr GetProp(IntPtr h, string name);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr RemoveProp(IntPtr h, string name);
    const string Pending = "ZoteroWord.PendingEmptyEditor.v2";
    static IntPtr InputStamp() {
        var input = new LastInput { Size = (uint)Marshal.SizeOf(typeof(LastInput)) };
        return GetLastInputInfo(ref input) ? new IntPtr(unchecked((int)input.Tick)) : IntPtr.Zero;
    }
    public static void MarkFreshEditor(IntPtr h) {
        var stamp = InputStamp();
        if (stamp != IntPtr.Zero) SetProp(h, Pending, stamp);
    }
    public static bool WasFreshEditor(IntPtr h) { return GetProp(h, Pending) != IntPtr.Zero; }
    public static bool UntouchedEditor(IntPtr h) {
        var stamp = GetProp(h, Pending);
        return stamp != IntPtr.Zero && stamp == InputStamp();
    }
    public static void ReleaseEditor(IntPtr h) { RemoveProp(h, Pending); }
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint process);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool ShowWindowAsync(IntPtr h, int cmd);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr h, out Rect rect);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h, ref Point pt);
    [DllImport("user32.dll")] static extern bool ScreenToClient(IntPtr h, ref Point pt);
    [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point pt);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h, uint kind);
    [DllImport("user32.dll", SetLastError=true)] static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);

    public static uint ProcessOf(IntPtr h) {
        uint pid;
        GetWindowThreadProcessId(h, out pid);
        return pid;
    }
    public static void Activate(IntPtr h, int pid) {
        if (h == IntPtr.Zero || ProcessOf(h) != pid) throw new IOException("The Zotero window changed.");
        if (IsIconic(h)) ShowWindowAsync(h, 9);
        SetForegroundWindow(h);
    }
    public static void ClickPane(IntPtr h, int pid, double fraction) {
        if (ProcessOf(h) != pid || GetForegroundWindow() != h) throw new IOException("Zotero lost foreground focus.");
        IntPtr dpi = IntPtr.Zero;
        try {
            try { dpi = SetThreadDpiAwarenessContext(new IntPtr(-4)); } catch (EntryPointNotFoundException) { }
            Rect rect;
            if (!GetClientRect(h, out rect)) throw new IOException("Cannot read the Zotero editor bounds.");
            var pt = new Point { X = (int)(rect.Right * fraction), Y = (int)(rect.Bottom * 0.42) };
            if (!ClientToScreen(h, ref pt)) throw new IOException("Cannot locate the Zotero editor pane.");
            var child = WindowFromPoint(pt);
            if (child == IntPtr.Zero || GetAncestor(child, 2) != h || ProcessOf(child) != pid || GetForegroundWindow() != h)
                throw new IOException("The Zotero editor pane is covered by another window.");
            if (!ScreenToClient(child, ref pt)) throw new IOException("Cannot locate the Zotero editor pane.");
            var lp = new IntPtr((pt.X & 0xffff) | ((pt.Y & 0xffff) << 16));
            if (!PostMessage(child, 0x0201, new IntPtr(1), lp) || !PostMessage(child, 0x0202, IntPtr.Zero, lp))
                throw new Win32Exception();
        } finally {
            if (dpi != IntPtr.Zero) SetThreadDpiAwarenessContext(dpi);
        }
    }
}
