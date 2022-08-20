using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using SpeechBuilder;
//using StartLoader;

namespace DummyScreenReader
{
    static class Program
    {
        private static SpeechControl speaker;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            speaker = new SpeechControl();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1(speaker));
        }
    }
}
