using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using DataHelper;

namespace W_Opera
{

    public class Helper_ApproveSample : INotifyPropertyChanged
    {
        private int _ID;
        private int _idNumber;
        private string _model;
        private string _modelCode;
        private string _customer;
        private string _ver;
        private string _dateInput;
        private string _ma;
        private string _sx;
        private string _cl;
        private string _rd;
        private string _kd;
        private int _numberCheck;
        private int _IdColorRow;
        private string dateApprove;
        private PinValue _color;
        public int ID
        {
            get { return this._ID; }
            set
            {
                if (this._ID != value)
                {
                    this._ID = value;
                    NotifyPropertyChanged("ID");
                }
            }
        }

        public int IDNumber
        {
            get { return this._idNumber; }
            set
            {
                if (this._idNumber != value)
                {
                    this._idNumber = value;
                    NotifyPropertyChanged("IDNumber");
                }
            }
        }

        public string Model
        {
            get { return this._model; }
            set
            {
                if (this._model != value)
                {
                    this._model = value;
                    NotifyPropertyChanged("Model");
                }
            }
        }

        public string ModelCode
        {
            get { return this._modelCode; }
            set
            {
                if (this._modelCode != value)
                {
                    this._modelCode = value;
                    NotifyPropertyChanged("ModelCode");
                }
            }
        }

        public string Customer
        {
            get { return this._customer; }
            set
            {
                if (this._customer != value)
                {
                    this._customer = value;
                    NotifyPropertyChanged("Customer");
                }
            }
        }

        public string Ver
        {
            get { return this._ver; }
            set
            {
                if (this._ver != value)
                {
                    this._ver = value;
                    NotifyPropertyChanged("Ver");
                }
            }
        }

        public string DateInput
        {
            get { return this._dateInput; }
            set
            {
                if (this._dateInput != value)
                {
                    this._dateInput = value;
                    NotifyPropertyChanged("DateInput");
                }
            }
        }

        public string MA
        {
            get { return this._ma; }
            set
            {
                if (this._ma != value)
                {
                    this._ma = value;
                    NotifyPropertyChanged("MA");
                }
            }
        }

        public string SX
        {
            get { return this._sx; }
            set
            {
                if (this._sx != value)
                {
                    this._sx = value;
                    NotifyPropertyChanged("SX");
                }
            }
        }

        public string CL
        {
            get { return this._cl; }
            set
            {
                if (this._cl != value)
                {
                    this._cl = value;
                    NotifyPropertyChanged("CL");
                }
            }
        }

        public string RD
        {
            get { return this._rd; }
            set
            {
                if (this._rd != value)
                {
                    this._rd = value;
                    NotifyPropertyChanged("RD");
                }
            }
        }

        public string KD
        {
            get { return this._kd; }
            set
            {
                if (this._kd != value)
                {
                    this._kd = value;
                    NotifyPropertyChanged("KD");
                }
            }
        }

        public int NumberCheck
        {
            get { return this._numberCheck; }
            set
            {
                if (this._numberCheck != value)
                {
                    this._numberCheck = value;
                    NotifyPropertyChanged("NumberCheck");
                }
            }
        }

        public int IdColorRow
        {
            get { return this._IdColorRow; }
            set
            {
                if (this._IdColorRow != value)
                {
                    this._IdColorRow = value;
                    NotifyPropertyChanged("IdColorRow");
                }
            }
        }

        public string DateApprove
        {
            get { return this.dateApprove; }
            set
            {
                if (this.dateApprove != value)
                {
                    this.dateApprove = value;
                    NotifyPropertyChanged("DateApprove");
                }
            }
        }

        public PinValue RowColor
        {
            get { return this._color; }
            set
            {
                if (this._color != value)
                {
                    this._color = value;
                    NotifyPropertyChanged("RowColor");
                }
            }
        }

        private void NotifyPropertyChanged(string v)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(v));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
