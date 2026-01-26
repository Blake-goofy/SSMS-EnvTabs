using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        private static bool TryGetConnectionInfo(IVsWindowFrame frame, out string server, out string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            server = null;
            database = null;

            if (frame == null)
            {
                return false;
            }

            try
            {
                TryPopulateFromCaptions(frame, ref server, ref database);
                return !string.IsNullOrWhiteSpace(server) || !string.IsNullOrWhiteSpace(database);
            }
            catch
            {
                return false;
            }
        }

        private static void TryPopulateFromCaptions(IVsWindowFrame frame, ref string server, ref string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (frame == null) return;

            string caption = TryReadFrameCaption(frame);
            if (TryParseServerDatabaseFromCaption(caption, out string s1, out string d1))
            {
                if (string.IsNullOrWhiteSpace(server)) server = s1;
                if (string.IsNullOrWhiteSpace(database)) database = d1;
                return;
            }

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_EditorCaption, out object editorCaptionObj) == VSConstants.S_OK)
                {
                    string editorCaption = editorCaptionObj as string;
                    if (TryParseServerDatabaseFromCaption(editorCaption, out string s2, out string d2))
                    {
                        if (string.IsNullOrWhiteSpace(server)) server = s2;
                        if (string.IsNullOrWhiteSpace(database)) database = d2;
                    }
                }
            }
            catch
            {
                // ignore
            }
        }

        private static bool TryParseServerDatabaseFromCaption(string caption, out string server, out string database)
        {
            server = null;
            database = null;

            if (string.IsNullOrWhiteSpace(caption)) return false;

            try
            {
                int dash = caption.IndexOf(" - ", StringComparison.Ordinal);
                if (dash < 0)
                {
                    return false;
                }

                string tail = caption.Substring(dash + 3);
                int paren = tail.IndexOf(" (", StringComparison.Ordinal);
                if (paren >= 0)
                {
                    tail = tail.Substring(0, paren);
                }

                tail = tail.Trim();
                if (string.IsNullOrWhiteSpace(tail)) return false;

                int dot = tail.IndexOf('.');
                if (dot > 0 && dot < tail.Length - 1)
                {
                    server = tail.Substring(0, dot).Trim();
                    database = tail.Substring(dot + 1).Trim();
                }
                else
                {
                    server = tail;
                }

                return !string.IsNullOrWhiteSpace(server) || !string.IsNullOrWhiteSpace(database);
            }
            catch
            {
                return false;
            }
        }

        private List<OpenDocumentInfo> GetOpenDocumentsSnapshot()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var list = new List<OpenDocumentInfo>();

            rdt.GetRunningDocumentsEnum(out IEnumRunningDocuments enumDocs);
            if (enumDocs == null)
            {
                return list;
            }

            uint[] cookies = new uint[1];
            while (enumDocs.Next(1, cookies, out uint fetched) == VSConstants.S_OK && fetched == 1)
            {
                uint cookie = cookies[0];
                string moniker = TryGetMonikerFromCookie(cookie);
                if (string.IsNullOrWhiteSpace(moniker))
                {
                    continue;
                }

                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame == null)
                {
                    continue;
                }

                string caption = TryReadFrameCaption(frame);
                TryGetConnectionInfo(frame, out string server, out string database);
                list.Add(new OpenDocumentInfo
                {
                    Cookie = cookie,
                    Frame = frame,
                    Caption = caption,
                    Moniker = moniker,
                    Server = server,
                    Database = database
                });
            }

            return list;
        }

        private string TryGetMonikerFromCookie(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (docCookie == 0) return null;

            try
            {
                rdt.GetDocumentInfo(
                    docCookie,
                    out uint _,
                    out uint _,
                    out uint _,
                    out string moniker,
                    out IVsHierarchy _,
                    out uint _,
                    out IntPtr _);

                return moniker;
            }
            catch
            {
                return null;
            }
        }

        private static string TryReadFrameCaption(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null) return null;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out object caption) == VSConstants.S_OK)
                {
                    return caption as string;
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private static string TryReadFrameEditorCaption(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null) return null;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_EditorCaption, out object caption) == VSConstants.S_OK)
                {
                    return caption as string;
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private void LogFrameCaptions(IVsWindowFrame frame, string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null) return;

            string caption = TryReadFrameCaption(frame) ?? string.Empty;
            string editorCaption = TryReadFrameEditorCaption(frame) ?? string.Empty;
            EnvTabsLog.Info($"Frame caption ({reason}): '{caption}' | editor='{editorCaption}'");
        }

        private bool TryGetMonikerFromFrame(IVsWindowFrame frame, out string moniker)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            moniker = null;
            if (frame == null) return false;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out object mk) == VSConstants.S_OK)
                {
                    moniker = mk as string;
                    return !string.IsNullOrWhiteSpace(moniker);
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private bool TryGetCookieFromMoniker(string moniker, out uint cookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            cookie = 0;
            if (string.IsNullOrWhiteSpace(moniker)) return false;

            try
            {
                rdt.FindAndLockDocument((uint)_VSRDTFLAGS.RDT_NoLock, moniker, out IVsHierarchy _, out uint _, out IntPtr _, out cookie);
                return cookie != 0;
            }
            catch
            {
                return false;
            }
        }

        private IVsWindowFrame TryGetFrameFromMoniker(string moniker)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (shellOpenDoc == null || string.IsNullOrWhiteSpace(moniker)) return null;

            try
            {
                Guid logicalView = Guid.Empty;
                shellOpenDoc.OpenDocumentViaProject(
                    moniker,
                    ref logicalView,
                    out Microsoft.VisualStudio.OLE.Interop.IServiceProvider _,
                    out IVsUIHierarchy _,
                    out uint _,
                    out IVsWindowFrame frame);

                return frame;
            }
            catch
            {
                return null;
            }
        }

        private static bool IsRenameEligible(string caption)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(caption))
                {
                    return false;
                }

                return RenameEligibleRegex.IsMatch(caption);
            }
            catch
            {
                return false;
            }
        }
    }
}
