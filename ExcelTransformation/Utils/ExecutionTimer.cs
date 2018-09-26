using System;
using System.Diagnostics;

namespace ExcelTransformation.Utils
{
    public class ExecutionTimer : IDisposable
    {
        private Stopwatch _sw = null;
        private readonly string _title;

        public ExecutionTimer(string title)
        {
            _title = title;
        }

        public static ExecutionTimer StartNew(string title)
        {
            var timer = new ExecutionTimer(title);

            timer.Start();

            return timer;
        }

        public Stopwatch Start()
        {
            Console.WriteLine($"Starting {_title}...");

            if (_sw == null)
            {
                _sw = Stopwatch.StartNew();
            }
            else
            {
                _sw.Restart();
            }

            return _sw;
        }

        public Stopwatch Stop()
        {
            if (_sw != null)
            {
                _sw.Stop();

                Console.WriteLine($"Ending {_title}... Execution time: {_sw.ElapsedMilliseconds} ms.");
            }

            return _sw;
        }

        public void Dispose()
        {
            Stop();
        }
    }
}