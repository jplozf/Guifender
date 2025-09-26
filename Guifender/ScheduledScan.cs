using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Guifender
{
    [Flags]
    public enum DaysOfWeek
    {
        None = 0,
        Monday = 1,
        Tuesday = 2,
        Wednesday = 4,
        Thursday = 8,
        Friday = 16,
        Saturday = 32,
        Sunday = 64,
        All = Monday | Tuesday | Wednesday | Thursday | Friday | Saturday | Sunday
    }

    public class ScheduledScan : INotifyPropertyChanged
    {
        private Guid _id = Guid.NewGuid();
        public Guid Id { get => _id; set { _id = value; OnPropertyChanged(); } }

        private string _name;
        public string Name { get => _name; set { _name = value; OnPropertyChanged(); } }

        private string _scanType = "QuickScan";
        public string ScanType { get => _scanType; set { _scanType = value; OnPropertyChanged(); } }

        private string _path;
        public string Path { get => _path; set { _path = value; OnPropertyChanged(); } }

        private TimeSpan _timeOfDay = new TimeSpan(12, 0, 0);
        public TimeSpan TimeOfDay { get => _timeOfDay; set { _timeOfDay = value; OnPropertyChanged(); } }

        private DaysOfWeek _daysOfWeek = DaysOfWeek.All;
        public DaysOfWeek DaysOfWeek { get => _daysOfWeek; set { _daysOfWeek = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
