using Microsoft.Practices.Prism.PubSubEvents;

namespace APLPX.Common
{
    public static class EventAgg
    {
        public static readonly IEventAggregator _eventAggregator = new EventAggregator();
    }

    public class StatusBarEvent : PubSubEvent<string> { }
    public class ErrorMessageEvent : PubSubEvent<string> { }

    /// <summary>
    /// Notifies subscribers that an operation is in progress.
    /// </summary>
    public class StatusBarMessage
    {
        public string Message { get; private set; }

        public StatusBarMessage(string message)
        {
            Message = message;
        }
    }
}
