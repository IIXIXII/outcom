using System;
using System.Collections.Generic;

namespace Outcom.AddIn
{
    internal enum CodexOperationState
    {
        Running,
        Canceling,
        Completed,
        Canceled,
        Skipped,
        Failed
    }

    internal sealed class CodexOperationInfo
    {
        internal Guid Id { get; set; }

        internal string Description { get; set; }

        internal CodexOperationState State { get; set; }

        internal string Detail { get; set; }

        internal DateTimeOffset StartedAt { get; set; }

        internal DateTimeOffset? CompletedAt { get; set; }

        internal bool CanCancel { get; set; }

        internal bool IsActive
        {
            get
            {
                return State == CodexOperationState.Running ||
                    State == CodexOperationState.Canceling;
            }
        }

        internal CodexOperationInfo Clone()
        {
            return new CodexOperationInfo
            {
                Id = Id,
                Description = Description,
                State = State,
                Detail = Detail,
                StartedAt = StartedAt,
                CompletedAt = CompletedAt,
                CanCancel = CanCancel
            };
        }

        public override string ToString()
        {
            string prefix;
            switch (State)
            {
                case CodexOperationState.Running:
                    prefix = "En cours";
                    break;
                case CodexOperationState.Canceling:
                    prefix = "Annulation";
                    break;
                case CodexOperationState.Completed:
                    prefix = "Terminé";
                    break;
                case CodexOperationState.Canceled:
                    prefix = "Annulé";
                    break;
                case CodexOperationState.Skipped:
                    prefix = "Sans insertion";
                    break;
                default:
                    prefix = "Échec";
                    break;
            }

            return prefix + " — " + Description +
                (string.IsNullOrWhiteSpace(Detail) ? string.Empty : " — " + Detail);
        }
    }

    internal sealed class CodexOperationChangedEventArgs : EventArgs
    {
        internal CodexOperationChangedEventArgs(IReadOnlyList<CodexOperationInfo> operations)
        {
            Operations = operations ?? new CodexOperationInfo[0];
        }

        internal IReadOnlyList<CodexOperationInfo> Operations { get; private set; }
    }

    internal sealed class CodexOperationTracker
    {
        private const int MaximumCompletedOperations = 8;
        private static readonly TimeSpan CompletedOperationLifetime = TimeSpan.FromMinutes(10);

        private readonly object _syncRoot = new object();
        private readonly Dictionary<Guid, Entry> _entries = new Dictionary<Guid, Entry>();

        internal event EventHandler<CodexOperationChangedEventArgs> Changed;

        internal CodexOperationHandle Begin(string description, Action cancelAction)
        {
            if (string.IsNullOrWhiteSpace(description))
            {
                throw new ArgumentException("La description de l'opération est vide.", nameof(description));
            }

            Guid id = Guid.NewGuid();
            lock (_syncRoot)
            {
                PruneCompletedOperations();
                _entries.Add(id, new Entry
                {
                    Info = new CodexOperationInfo
                    {
                        Id = id,
                        Description = description.Trim(),
                        State = CodexOperationState.Running,
                        StartedAt = DateTimeOffset.Now,
                        CanCancel = cancelAction != null
                    },
                    CancelAction = cancelAction
                });
            }

            RaiseChanged();
            return new CodexOperationHandle(this, id);
        }

        internal IReadOnlyList<CodexOperationInfo> GetSnapshot()
        {
            lock (_syncRoot)
            {
                PruneCompletedOperations();
                return CreateSnapshot();
            }
        }

        internal int ClearCompleted()
        {
            int removedCount = 0;
            lock (_syncRoot)
            {
                var completedIds = new List<Guid>();
                foreach (KeyValuePair<Guid, Entry> pair in _entries)
                {
                    if (!pair.Value.Info.IsActive)
                    {
                        completedIds.Add(pair.Key);
                    }
                }

                foreach (Guid id in completedIds)
                {
                    if (_entries.Remove(id))
                    {
                        removedCount++;
                    }
                }
            }

            if (removedCount > 0)
            {
                RaiseChanged();
            }

            return removedCount;
        }

        internal bool RequestCancellation(Guid id)
        {
            Action cancelAction;
            lock (_syncRoot)
            {
                Entry entry;
                if (!_entries.TryGetValue(id, out entry) ||
                    entry.Info.State != CodexOperationState.Running ||
                    entry.CancelAction == null)
                {
                    return false;
                }

                entry.Info.State = CodexOperationState.Canceling;
                entry.Info.CanCancel = false;
                cancelAction = entry.CancelAction;
            }

            RaiseChanged();
            try
            {
                cancelAction();
                return true;
            }
            catch (Exception exception)
            {
                Finish(id, CodexOperationState.Failed, "annulation impossible");
                LocalLogger.Error(
                    "Impossible d'annuler une demande Codex (" +
                    exception.GetType().Name + ").");
                return false;
            }
        }

        internal void Finish(
            Guid id,
            CodexOperationState state,
            string detail)
        {
            if (state == CodexOperationState.Running || state == CodexOperationState.Canceling)
            {
                throw new ArgumentException("L'état final de l'opération est invalide.", nameof(state));
            }

            lock (_syncRoot)
            {
                Entry entry;
                if (!_entries.TryGetValue(id, out entry) || !entry.Info.IsActive)
                {
                    return;
                }

                entry.Info.State = state;
                entry.Info.Detail = string.IsNullOrWhiteSpace(detail) ? null : detail.Trim();
                entry.Info.CompletedAt = DateTimeOffset.Now;
                entry.Info.CanCancel = false;
                entry.CancelAction = null;
                PruneCompletedOperations();
            }

            RaiseChanged();
        }

        private void PruneCompletedOperations()
        {
            DateTimeOffset cutoff = DateTimeOffset.Now - CompletedOperationLifetime;
            var completed = new List<Entry>();
            var expiredIds = new List<Guid>();
            foreach (KeyValuePair<Guid, Entry> pair in _entries)
            {
                if (!pair.Value.Info.IsActive)
                {
                    completed.Add(pair.Value);
                    if (pair.Value.Info.CompletedAt < cutoff)
                    {
                        expiredIds.Add(pair.Key);
                    }
                }
            }

            foreach (Guid id in expiredIds)
            {
                _entries.Remove(id);
            }

            completed.RemoveAll(entry => entry.Info.CompletedAt < cutoff);
            completed.Sort((left, right) => Nullable.Compare(
                right.Info.CompletedAt,
                left.Info.CompletedAt));
            for (int index = MaximumCompletedOperations; index < completed.Count; index++)
            {
                _entries.Remove(completed[index].Info.Id);
            }
        }

        private IReadOnlyList<CodexOperationInfo> CreateSnapshot()
        {
            var values = new List<CodexOperationInfo>();
            foreach (Entry entry in _entries.Values)
            {
                values.Add(entry.Info.Clone());
            }

            values.Sort((left, right) =>
            {
                if (left.IsActive != right.IsActive)
                {
                    return left.IsActive ? -1 : 1;
                }

                return right.StartedAt.CompareTo(left.StartedAt);
            });
            return values.AsReadOnly();
        }

        private void RaiseChanged()
        {
            EventHandler<CodexOperationChangedEventArgs> handler = Changed;
            if (handler == null)
            {
                return;
            }

            var eventArgs = new CodexOperationChangedEventArgs(GetSnapshot());
            foreach (EventHandler<CodexOperationChangedEventArgs> subscriber in
                handler.GetInvocationList())
            {
                try
                {
                    subscriber(this, eventArgs);
                }
                catch (Exception exception)
                {
                    LocalLogger.Error(
                        "Impossible de notifier l'activité Codex (" +
                        exception.GetType().Name + ").");
                }
            }
        }

        private sealed class Entry
        {
            internal CodexOperationInfo Info { get; set; }

            internal Action CancelAction { get; set; }
        }
    }

    internal sealed class CodexOperationHandle
    {
        private readonly CodexOperationTracker _tracker;
        private readonly Guid _id;

        internal CodexOperationHandle(CodexOperationTracker tracker, Guid id)
        {
            _tracker = tracker ?? throw new ArgumentNullException(nameof(tracker));
            _id = id;
        }

        internal void Complete(string detail)
        {
            _tracker.Finish(_id, CodexOperationState.Completed, detail);
        }

        internal void Cancel(string detail)
        {
            _tracker.Finish(_id, CodexOperationState.Canceled, detail);
        }

        internal void Skip(string detail)
        {
            _tracker.Finish(_id, CodexOperationState.Skipped, detail);
        }

        internal void Fail(string detail)
        {
            _tracker.Finish(_id, CodexOperationState.Failed, detail);
        }
    }
}
