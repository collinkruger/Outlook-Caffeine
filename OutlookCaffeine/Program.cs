using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Topshelf;

namespace OutlookCaffeine
{
    public class Core
    {
        readonly Timer _timer;

        void _timer_Elapsed(object _, ElapsedEventArgs __)
        {
            if (Process.GetProcessesByName("Outlook").Length == 0)
                Process.Start("Outlook");
        }

        public Core()
        {
            _timer = new Timer();
            _timer.Interval = 60 * 1000;
            _timer.AutoReset = true;
            _timer.Elapsed += _timer_Elapsed;
        }

        public void Start()
        {
            _timer.Start();
        }

        public void Stop()
        {
            _timer.Dispose();
        }

    }

    public static class Program
    {
        public static void Main()
        {
            HostFactory.Run(x =>
            {
                x.Service<Core>(s =>
                {
                    s.ConstructUsing(() => new Core());
                    s.WhenStarted(c => c.Start());
                    s.WhenStopped(c => c.Stop());
                });
                x.RunAsPrompt();

                x.SetDescription("Outlook Caffeine");
                x.SetDisplayName("Outlook Caffeine");
                x.SetDisplayName("Outlook Caffeine");
                x.StartAutomatically();
            });
        }
    }
}
