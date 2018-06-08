using System;
using System.Diagnostics;


namespace DCRM_Utils
{
    public class MiscHelper
    {
        #region Properties
        private Stopwatch _stopwatch;
        private static int _CurrentIndex = 0;
        
        public delegate void MessageDisplayer(string message);
        public delegate void ProgressDisplayer(int maxIndex);

        public static MessageDisplayer WriteLine = DisplayMessageMethod;
        public static ProgressDisplayer IncrementProgressBar = IncrementProgressBarMethod;

        #endregion // Properties

        #region DisplayMessageMethod
        // Create a method for a delegate.
        public static void DisplayMessageMethod(string message)
        {
            Console.WriteLine(message);
        }
        #endregion // DisplayMessageMethod

        #region IncrementProgressBarMethod
        private static void IncrementProgressBarMethod(int maxIndex)
        {
            _CurrentIndex++;
            DrawTextProgressBar(_CurrentIndex, maxIndex);
        }
        #endregion // IncrementProgressBarMethod

        #region StartTimeWatch
        public void StartTimeWatch()
        {
            _stopwatch = new Stopwatch();
            _stopwatch.Start();
        }
        #endregion // StartTimeWatch

        #region StopTimeWatch
        public void StopTimeWatch()
        {
            if (_stopwatch != null)
            {
                _stopwatch.Stop();
            }
        }
        #endregion // StopTimeWatch

        #region GetDuration
        public string GetDuration()
        {
            var duration = string.Format($"{new DateTime(ticks: _stopwatch.ElapsedTicks).ToString("HH:mm:ss.fff")}");
            return (duration);
        }
        #endregion // GetDuration

        #region DrawTextProgressBar
        public static void DrawTextProgressBar(int progress, int total)
        {
            //draw empty progress bar
            Console.CursorLeft = 0;
            Console.Write("["); //start
            Console.CursorLeft = 32;
            Console.Write("]"); //end
            Console.CursorLeft = 1;
            float onechunk = 31.0f / total;

            //draw filled part
            int position = 1;
            for (int i = 0; i < onechunk * progress; i++)
            {
                Console.BackgroundColor = ConsoleColor.Gray;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw unfilled part
            for (int i = position; i <= 31; i++)
            {
                Console.BackgroundColor = ConsoleColor.Green;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw totals
            Console.CursorLeft = 35;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write(progress.ToString() + " of " + total.ToString() + "    "); //blanks at the end remove any excess
        }
        #endregion // DrawTextProgressBar

        #region PauseExecution
        public static void PauseExecution()
        {
            MiscHelper.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
        #endregion // PauseExecution
    }
}
