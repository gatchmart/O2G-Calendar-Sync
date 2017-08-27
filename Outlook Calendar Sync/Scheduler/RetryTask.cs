using System;

namespace Outlook_Calendar_Sync.Scheduler {

    public enum RetryAction
    {
        Add,
        Update,
        Delete,
        DeleteById
    }

    [Serializable]
    public class RetryTask {
        public CalendarItem CalendarItem;
        public string Calendar;
        public RetryAction Action;
        public int Delay;
        public DateTime LastRun;

        private readonly int m_retryCount;
        private int m_currentRetry;

        public RetryTask()
        {
            CalendarItem = null;
            Calendar = "";
            Action = 0;
            Delay = 0;
            LastRun = DateTime.MinValue;
            m_retryCount = 0;
            m_retryCount = 0;
        }

        public RetryTask( CalendarItem item, string calendar, RetryAction action, int delay = 1, int retryCount = 8 ) {
            CalendarItem = item;
            Calendar = calendar;
            Action = action;
            Delay = delay;
            LastRun = DateTime.Now;
            m_retryCount = retryCount;
            m_currentRetry = 0;
        }

        public void RetryFailed()
        {
            Delay = GetDelay();
            m_currentRetry++;
        }

        public void Successful()
        {
            Scheduler.Instance.RemoveRetry( this );
        }

        public bool Eligible()
        {
            return Delay <= 1440 && m_currentRetry <= m_retryCount;
        }

        private int GetDelay()
        {
            switch ( Delay )
            {
                case 1:
                    return 2;
                case 2:
                    return 5;
                case 5:
                    return 10;
                case 10:
                    return 30;
                case 30:
                    return 60;
                case 60:
                    return 720;
                case 720:
                    return 1440;
            }

            return int.MaxValue;
        }

    }
}
