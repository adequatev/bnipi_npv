using bnipi_npv.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace bnipi_npv.ViewModels
{
    public class Estimate : INotifyPropertyChanged
    {
        #region fields
        public ObservableCollection<FinanceYear> financeYears { get; set; }

        private int _targetYear = 2050;
        public int TargetYear
        {
            get => _targetYear;
            set
            {
                InfoTextLabel = string.Empty;
                _targetYear = value;
                if (value < 0)
                    _targetYear = 0;
            }
        }
        public decimal Coef { get; set; } = 0.2m;

        private decimal _targetYearNpv;
        public decimal TargetYearNpv
        {
            get => _targetYearNpv;
            private set
            {
                _targetYearNpv = value;
                OnPropertyChanged(nameof(TargetYearNpv));
            }
        }
        private string _infoText = string.Empty;
        private readonly string _errorText = "Некорретный год";
        private readonly string _errorOpenText = "Что-то пошло не так";
        public string InfoTextLabel
        {
            get => _infoText;
            set
            {
                _infoText = value;
                OnPropertyChanged(nameof(InfoTextLabel));
            }
        }

        #endregion

        public Estimate()
        {
            FillWithDefaultData();
            financeYears = new ObservableCollection<FinanceYear>(financeYears.OrderBy(x => x.Year));
        }

        #region commands
        private RelayCommand _countCommand;
        public ICommand BtnClickCountCommand
        {
            get
            {
                if (_countCommand == null)
                {
                    _countCommand = new RelayCommand(o =>
                    {
                        CountNPV();
                    });
                }
                return _countCommand;
            }
        }

        private RelayCommand _openExcelCommand;
        public ICommand BtnClickOpenExcelCommand
        {
            get
            {
                if (_openExcelCommand == null)
                {
                    _openExcelCommand = new RelayCommand(o =>
                    {
                        OpenExcel();
                    });
                }
                return _openExcelCommand;
            }
        }
        #endregion

        private void CountNPV()
        {
            financeYears = new ObservableCollection<FinanceYear>(financeYears.OrderBy(x => x.Year));
            FinanceYear target = null;
            try
            {
                target = financeYears.First(x => x.Year.Equals(TargetYear));
            }
            catch (Exception ex)
            {
                InfoTextLabel = _errorText;
                return;
            }

            financeYears[0].NetPresentValue =
                financeYears[0].Profit * (1 / (1 + Coef));
            for (int i = 1; i <= financeYears.IndexOf(target); i++)
            {
                financeYears[i].NetPresentValue =
                    financeYears[i].Profit * (decimal)(1 / Math.Pow((double)(1 + Coef), (i + 1))) +
                    financeYears[i - 1].NetPresentValue;
            }
            TargetYearNpv = target.NetPresentValue;

        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
        private void OpenExcel()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog.ShowDialog() == true)
            {
                financeYears.Clear();
                try
                {
                    ReadExcelFile(openFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    InfoTextLabel = _errorOpenText;
                    FillWithDefaultData();
                }
            }
        }
        private void ReadExcelFile(string filePath)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                foreach (Row row in sheetData.Elements<Row>().Skip(1))
                {
                    int i = 0;
                    decimal[] values = new decimal[3];

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string text = cell.CellValue.Text;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            int index = int.Parse(text);
                            text = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;
                        }
                        values[i++] = decimal.Parse(text);
                    }
                    financeYears.Add(new FinanceYear((int)values[0], values[1], values[2]));
                }
            }
        }

        private void FillWithDefaultData()
        {
            financeYears = new ObservableCollection<FinanceYear>
            {
                new FinanceYear(2020, 1000, 0),
                new FinanceYear(2021, 1000, 0),
                new FinanceYear(2022, 1000, 500),
                new FinanceYear(2023, 1000, 500),
                new FinanceYear(2024, 1000, 0),
                new FinanceYear(2025, 1000, 0),
                new FinanceYear(2026, 1000, 0),
                new FinanceYear(2027, 1000, 0),
                new FinanceYear(2028, 1000, 0),
                new FinanceYear(2029, 1000, 0),
                new FinanceYear(2030, 1000, 0),
                new FinanceYear(2031, 1000, 0),
                new FinanceYear(2032, 1000, 0),
                new FinanceYear(2033, 1000, 0),
                new FinanceYear(2034, 1000, 0),
                new FinanceYear(2035, 1000, 0),
                new FinanceYear(2036, 1000, 0),
                new FinanceYear(2037, 1000, 0),
                new FinanceYear(2038, 1000, 0),
                new FinanceYear(2039, 1000, 0),
                new FinanceYear(2040, 1000, 0),
                new FinanceYear(2041, 1000, 0),
                new FinanceYear(2042, 1000, 0),
                new FinanceYear(2043, 1000, 0),
                new FinanceYear(2044, 1000, 0),
                new FinanceYear(2045, 1000, 0),
                new FinanceYear(2046, 1000, 0),
                new FinanceYear(2047, 1000, 0),
                new FinanceYear(2048, 1000, 0),
                new FinanceYear(2049, 1000, 0),
                new FinanceYear(2050, 1000, 0),
            };

        }
    }
}
