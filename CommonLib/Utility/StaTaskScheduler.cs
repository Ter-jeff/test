using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CommonLib.Utility
{
    public sealed class StaTaskScheduler : TaskScheduler, IDisposable
    {
        private readonly List<Thread> _threads;
        private BlockingCollection<Task> _tasks;

        public StaTaskScheduler(int concurrencyLevel)
        {
            if (concurrencyLevel < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(concurrencyLevel));
            }

            _tasks = new BlockingCollection<Task>();
            _threads = Enumerable.Range(0, concurrencyLevel).Select(_ =>
            {
                var thread = new Thread(() =>
                {
                    foreach (Task t in _tasks.GetConsumingEnumerable())
                    {
                        TryExecuteTask(t);
                    }
                })
                {
                    IsBackground = true
                };
                thread.SetApartmentState(ApartmentState.STA);
                return thread;
            }).ToList();

            foreach (Thread thread in _threads)
            {
                thread.Start();
            }
        }

        public override int MaximumConcurrencyLevel => _threads.Count;

        public void Dispose()
        {
            if (_tasks != null)
            {
                _tasks.CompleteAdding();

                foreach (Thread thread in _threads)
                {
                    thread.Join();
                }

                _tasks.Dispose();
                _tasks = null;
            }
        }

        protected override void QueueTask(Task task)
        {
            _tasks.Add(task);
        }

        protected override IEnumerable<Task> GetScheduledTasks()
        {
            return _tasks.ToArray();
        }

        protected override bool TryExecuteTaskInline(Task task, bool taskWasPreviouslyQueued)
        {
            return Thread.CurrentThread.GetApartmentState() == ApartmentState.STA && TryExecuteTask(task);
        }
    }
}
