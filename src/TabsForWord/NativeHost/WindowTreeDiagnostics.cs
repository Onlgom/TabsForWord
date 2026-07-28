using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace TabsForWord.NativeHost
{
    /// <summary>
    /// Diagnostics of Word HWND hierarchy: a snapshot of the window tree from a
    /// given root, saved to a text file next to the logs.
    ///
    /// Used only in the in-window host mode (with dump=1, or once when locating the
    /// document area fails). Tolerant of windows dying during enumeration: every
    /// Win32 call is optional and missing data never aborts the walk.
    /// </summary>
    internal static class WindowTreeDiagnostics
    {
        private const int MaxNodes = 4096;      // guard against endless / enormous trees
        private const int MaxTextLength = 60;

        /// <summary>One node of the window-tree snapshot.</summary>
        internal sealed class Node
        {
            public IntPtr Hwnd;
            public IntPtr Parent;
            public string ClassName = string.Empty;
            public string Text = string.Empty;
            public uint ProcessId;
            public uint ThreadId;
            public NativeWin32.RECT WindowRect;
            public NativeWin32.RECT ClientRect;
            public bool Visible;
            public long Style;
            public long ExStyle;
        }

        /// <summary>
        /// Snapshot of a window subtree: the root plus every descendant (EnumChildWindows
        /// walks descendants recursively). Never throws.
        /// </summary>
        internal static List<Node> Snapshot(IntPtr root)
        {
            var nodes = new List<Node>();
            try
            {
                if (!NativeWin32.IsWindow(root)) return nodes;

                var handles = new List<IntPtr> { root };
                try
                {
                    NativeWin32.EnumChildWindows(root, delegate(IntPtr h, IntPtr l)
                    {
                        handles.Add(h);
                        return handles.Count < MaxNodes; // false stops the enumeration
                    }, IntPtr.Zero);
                }
                catch (Exception ex)
                {
                    LoggingService.Warn("EnumChildWindows failed: " + ex.Message);
                }

                foreach (var h in handles)
                {
                    var n = ReadNode(h);
                    if (n != null) nodes.Add(n);
                }
            }
            catch (Exception ex)
            {
                LoggingService.Error("WindowTreeDiagnostics.Snapshot failed", ex);
            }
            return nodes;
        }

        private static Node ReadNode(IntPtr h)
        {
            try
            {
                // The window may have died between enumeration and reading - that is normal.
                if (!NativeWin32.IsWindow(h)) return null;

                var n = new Node { Hwnd = h };
                n.Parent = NativeWin32.GetAncestor(h, NativeWin32.GA_PARENT);
                n.ClassName = NativeWin32.GetClassNameSafe(h);
                n.Text = Sanitize(NativeWin32.GetWindowTextSafe(h, MaxTextLength));
                uint pid;
                n.ThreadId = NativeWin32.GetWindowThreadProcessId(h, out pid);
                n.ProcessId = pid;
                NativeWin32.GetWindowRect(h, out n.WindowRect);
                NativeWin32.GetClientRect(h, out n.ClientRect);
                n.Visible = NativeWin32.IsWindowVisible(h);
                n.Style = NativeWin32.GetWindowLongPtrSafe(h, NativeWin32.GWL_STYLE).ToInt64();
                n.ExStyle = NativeWin32.GetWindowLongPtrSafe(h, NativeWin32.GWL_EXSTYLE).ToInt64();
                return n;
            }
            catch
            {
                return null;
            }
        }

        private static string Sanitize(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("\r", " ").Replace("\n", " ");
        }

        /// <summary>
        /// Saves the window tree to %LOCALAPPDATA%\TabsForWord\Logs\WindowTree-*.txt.
        /// Returns the file path, or null on failure.
        /// </summary>
        internal static string DumpToFile(IntPtr root, string context)
        {
            try
            {
                var nodes = Snapshot(root);
                var text = Format(root, nodes, context);

                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TabsForWord", "Logs");
                Directory.CreateDirectory(dir);
                var file = Path.Combine(dir, "WindowTree-" +
                    DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) +
                    "-" + root.ToInt64().ToString("X") + ".txt");
                File.WriteAllText(file, text, Encoding.UTF8);
                CleanupOldDumps(dir);
                LoggingService.Info("Window tree dumped: " + file);
                return file;
            }
            catch (Exception ex)
            {
                LoggingService.Error("WindowTreeDiagnostics.DumpToFile failed", ex);
                return null;
            }
        }

        /// <summary>Dumps must not pile up forever: at most 20 files are kept.</summary>
        private static void CleanupOldDumps(string dir)
        {
            try
            {
                var files = new DirectoryInfo(dir).GetFiles("WindowTree-*.txt");
                if (files.Length <= 20) return;
                Array.Sort(files, (a, b) => string.CompareOrdinal(b.Name, a.Name));
                for (int i = 20; i < files.Length; i++)
                {
                    try { files[i].Delete(); } catch { }
                }
            }
            catch { }
        }

        /// <summary>Formats the snapshot as text, indented by depth.</summary>
        internal static string Format(IntPtr root, List<Node> nodes, string context)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Word window tree dump");
            sb.AppendLine("Context : " + (context ?? string.Empty));
            sb.AppendLine("Time    : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            sb.AppendLine("Root    : 0x" + root.ToInt64().ToString("X"));
            sb.AppendLine("Nodes   : " + nodes.Count);
            sb.AppendLine(new string('-', 100));

            // Parent index -> children (in enumeration order)
            var byParent = new Dictionary<IntPtr, List<Node>>();
            Node rootNode = null;
            foreach (var n in nodes)
            {
                if (n.Hwnd == root) { rootNode = n; continue; }
                List<Node> list;
                if (!byParent.TryGetValue(n.Parent, out list))
                {
                    list = new List<Node>();
                    byParent[n.Parent] = list;
                }
                list.Add(n);
            }

            if (rootNode != null)
            {
                AppendNode(sb, rootNode, byParent, 0, new HashSet<IntPtr>());
            }
            else
            {
                // The root died during the snapshot - print a flat list instead.
                foreach (var n in nodes) AppendLineFor(sb, n, 0);
            }
            return sb.ToString();
        }

        private static void AppendNode(StringBuilder sb, Node n,
            Dictionary<IntPtr, List<Node>> byParent, int depth, HashSet<IntPtr> visited)
        {
            if (!visited.Add(n.Hwnd)) return; // cycle guard
            AppendLineFor(sb, n, depth);
            List<Node> children;
            if (byParent.TryGetValue(n.Hwnd, out children))
            {
                foreach (var c in children)
                    AppendNode(sb, c, byParent, depth + 1, visited);
            }
        }

        private static void AppendLineFor(StringBuilder sb, Node n, int depth)
        {
            sb.Append(new string(' ', depth * 2));
            sb.Append("0x").Append(n.Hwnd.ToInt64().ToString("X"));
            sb.Append(" \"").Append(n.ClassName).Append('"');
            if (n.Text.Length > 0) sb.Append(" text=\"").Append(n.Text).Append('"');
            sb.Append(n.Visible ? " visible" : " hidden");
            sb.Append(" rect=(").Append(n.WindowRect.Left).Append(',').Append(n.WindowRect.Top)
              .Append(")-(").Append(n.WindowRect.Right).Append(',').Append(n.WindowRect.Bottom)
              .Append(") ").Append(n.WindowRect.Width).Append('x').Append(n.WindowRect.Height);
            sb.Append(" client=").Append(n.ClientRect.Width).Append('x').Append(n.ClientRect.Height);
            sb.Append(" style=0x").Append(n.Style.ToString("X8"));
            sb.Append(" ex=0x").Append(n.ExStyle.ToString("X8"));
            sb.Append(" pid=").Append(n.ProcessId).Append(" tid=").Append(n.ThreadId);
            sb.AppendLine();
        }
    }
}
