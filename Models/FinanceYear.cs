using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace bnipi_npv.Models
{
    public class FinanceYear : INotifyPropertyChanged
    {
        private decimal _income;
        private decimal _expenses;
        private decimal _netPresentValue;
        private int _year;
        public int Year
        {
            get => _year;
            set
            {
                _year = value;
                if (value < 0)
                    _year = 0;
                OnPropertyChanged(nameof(Year));
            }
        }
        public decimal Income
        {
            get => _income;
            set
            {
                _income = value;
                OnPropertyChanged(nameof(Income));
                OnPropertyChanged(nameof(Profit));
            }
        }
        public decimal Expenses
        {
            get => _expenses;
            set
            {
                _expenses = value;
                OnPropertyChanged(nameof(Expenses));
                OnPropertyChanged(nameof(Profit));
            }
        }
        public decimal NetPresentValue
        {
            get => _netPresentValue;
            set
            {
                _netPresentValue = value;
                OnPropertyChanged(nameof(NetPresentValue));
            }
        }
        public decimal Profit => Income - Expenses;

        public FinanceYear(int year, decimal income, decimal expenses)
        {
            Year = year;
            Income = income;
            Expenses = expenses;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
