using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace StatusReportConverter.Models
{
    public class Risk : INotifyPropertyChanged
    {
        private string id;
        private string description;
        private string impact;
        private string mitigation;
        private string status;
        private DateTime dateIdentified;

        public string Id
        {
            get => id;
            set
            {
                id = value;
                OnPropertyChanged();
            }
        }

        public string Description
        {
            get => description;
            set
            {
                description = value;
                OnPropertyChanged();
            }
        }

        public string Impact
        {
            get => impact;
            set
            {
                impact = value;
                OnPropertyChanged();
            }
        }

        public string Mitigation
        {
            get => mitigation;
            set
            {
                mitigation = value;
                OnPropertyChanged();
            }
        }

        public string Status
        {
            get => status;
            set
            {
                status = value;
                OnPropertyChanged();
            }
        }

        public DateTime DateIdentified
        {
            get => dateIdentified;
            set
            {
                dateIdentified = value;
                OnPropertyChanged();
            }
        }

        public Risk()
        {
            id = Guid.NewGuid().ToString().Substring(0, 8);
            description = string.Empty;
            impact = string.Empty;
            mitigation = string.Empty;
            status = "Open";
            dateIdentified = DateTime.Now;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}