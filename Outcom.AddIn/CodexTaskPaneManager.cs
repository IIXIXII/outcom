using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outcom.AddIn
{
    /// <summary>
    /// Crée un volet et une conversation distincts par fenêtre Explorer Outlook.
    /// Aucun processus Codex n'est démarré avant que l'utilisateur ouvre le volet.
    /// </summary>
    internal sealed class CodexTaskPaneManager : IDisposable
    {
        private const int DefaultPaneWidth = 420;

        private readonly ThisAddIn _addIn;
        private readonly Dictionary<long, PaneEntry> _entries =
            new Dictionary<long, PaneEntry>();
        private int _disposeState;

        internal CodexTaskPaneManager(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        internal void Toggle(object ribbonContext)
        {
            ThrowIfDisposed();
            Outlook.Explorer explorer = ribbonContext as Outlook.Explorer;
            if (explorer == null)
            {
                explorer = _addIn.Application.ActiveExplorer();
            }

            if (explorer == null)
            {
                throw new InvalidOperationException(
                    "Aucune fenêtre principale Outlook n'est disponible.");
            }

            long identity = GetComIdentity(explorer);
            PaneEntry entry;
            if (!_entries.TryGetValue(identity, out entry))
            {
                entry = CreateEntry(identity, explorer);
                _entries.Add(identity, entry);
            }

            entry.Pane.Visible = !entry.Pane.Visible;
            if (entry.Pane.Visible)
            {
                entry.Control.Focus();
            }
        }

        internal void RefreshConnectionStatus()
        {
            ThrowIfDisposed();
            foreach (PaneEntry entry in new List<PaneEntry>(_entries.Values))
            {
                entry.Control.RefreshConnectionStatus();
            }
        }

        private PaneEntry CreateEntry(long identity, Outlook.Explorer explorer)
        {
            var control = new CodexTaskPaneControl(_addIn, explorer);
            CustomTaskPane pane = null;
            try
            {
                pane = _addIn.CustomTaskPanes.Add(
                    control,
                    "Outcom — Codex",
                    explorer);
                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.DockPositionRestrict =
                    Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
                pane.Width = DefaultPaneWidth;
                pane.Visible = false;
                return new PaneEntry(this, identity, explorer, pane, control);
            }
            catch
            {
                if (pane != null)
                {
                    try
                    {
                        _addIn.CustomTaskPanes.Remove(pane);
                    }
                    catch (Exception)
                    {
                    }
                }

                control.Dispose();
                throw;
            }
        }

        private void ExplorerClosed(PaneEntry entry)
        {
            if (entry == null || VolatileReadDisposeState() != 0)
            {
                return;
            }

            PaneEntry current;
            if (_entries.TryGetValue(entry.Identity, out current) &&
                ReferenceEquals(current, entry))
            {
                _entries.Remove(entry.Identity);
                RemoveAndDisposeEntry(entry);
            }
        }

        private void RemoveAndDisposeEntry(PaneEntry entry)
        {
            entry.DetachEvents();
            entry.Control.CancelActiveTurn();
            try
            {
                _addIn.CustomTaskPanes.Remove(entry.Pane);
            }
            catch (ArgumentException)
            {
                // Le volet peut déjà avoir été retiré pendant la fermeture Outlook.
            }
            catch (InvalidOperationException)
            {
                // La collection VSTO peut être en cours de destruction.
            }
            catch (COMException)
            {
                // La fenêtre Outlook peut avoir détruit son volet avant l'événement Close.
            }
            finally
            {
                entry.Control.Dispose();
            }
        }

        private static long GetComIdentity(object value)
        {
            IntPtr unknown = IntPtr.Zero;
            try
            {
                unknown = Marshal.GetIUnknownForObject(value);
                return unknown.ToInt64();
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "La fenêtre Outlook ne peut pas être identifiée.",
                    exception);
            }
            finally
            {
                if (unknown != IntPtr.Zero)
                {
                    Marshal.Release(unknown);
                }
            }
        }

        private int VolatileReadDisposeState()
        {
            return System.Threading.Volatile.Read(ref _disposeState);
        }

        private void ThrowIfDisposed()
        {
            if (VolatileReadDisposeState() != 0)
            {
                throw new ObjectDisposedException(nameof(CodexTaskPaneManager));
            }
        }

        public void Dispose()
        {
            if (System.Threading.Interlocked.Exchange(ref _disposeState, 1) != 0)
            {
                return;
            }

            var entries = new List<PaneEntry>(_entries.Values);
            _entries.Clear();
            foreach (PaneEntry entry in entries)
            {
                RemoveAndDisposeEntry(entry);
            }
        }

        private sealed class PaneEntry
        {
            private readonly CodexTaskPaneManager _manager;
            private readonly Outlook.ExplorerEvents_10_Event _explorerEvents;
            private bool _eventsDetached;

            internal PaneEntry(
                CodexTaskPaneManager manager,
                long identity,
                Outlook.Explorer explorer,
                CustomTaskPane pane,
                CodexTaskPaneControl control)
            {
                _manager = manager;
                Identity = identity;
                Explorer = explorer;
                Pane = pane;
                Control = control;
                _explorerEvents = (Outlook.ExplorerEvents_10_Event)explorer;
                _explorerEvents.Close += Explorer_Close;
            }

            internal long Identity { get; private set; }

            internal Outlook.Explorer Explorer { get; private set; }

            internal CustomTaskPane Pane { get; private set; }

            internal CodexTaskPaneControl Control { get; private set; }

            private void Explorer_Close()
            {
                _manager.ExplorerClosed(this);
            }

            internal void DetachEvents()
            {
                if (_eventsDetached)
                {
                    return;
                }

                _eventsDetached = true;
                try
                {
                    _explorerEvents.Close -= Explorer_Close;
                }
                catch (COMException)
                {
                }
            }
        }
    }
}
