using DataHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace W_Opera
{
    //public enum PinValue { ON,OFF,NotRun}
    public class StackItemControl : INotifyPropertyChanged
    {
        public int ID { get; set; }    
        
        public string _Name;      

       
        public string Name
        {
            get { return this._Name; }
            set
            {
                if (this._Name != value)
                {
                    this._Name = value;
                    NotifyPropertyChanged("Name");
                }
            }
        }        

        private void NotifyPropertyChanged(string v)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(v));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class ButtonIsnabled : INotifyPropertyChanged
    {
        public bool IsEnabledAddMan;
        public bool IsEnabledDelMan;
        public bool IsEnabledEditMan;
        public bool IsEnabledSaveMan;
        public bool IsEnabledRunMan;

        public bool IsEnabledAddBox;
        public bool IsEnabledDelBox;
        public bool IsEnabledEditBox;
        public bool IsEnabledSaveBox;
        public bool IsEnabledRunBox;

        private void NotifyPropertyChanged(string v)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(v));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class Helper_DataButton : INotifyPropertyChanged
    {
        public int ID { get; set; }
        public string ContentButton { get; set; }
        public string _ImageSource { get; set; }
        public PinValue _CheckCreatTab;
        public PinValue _BackGround { get; set; }
        public string _Visibility { get; set; }
        
        public string Visibility
        {
            get { return this._Visibility; }
            set
            {
                if (this._Visibility != value)
                {
                    this._Visibility = value;
                    NotifyPropertyChanged("Visibility");
                }

            }
        }
        public bool _IsEnabled { get; set; }
        public bool IsEnabled
        {
            get { return this._IsEnabled; }
            set
            {
                if (this._IsEnabled != value)
                {
                    this._IsEnabled = value;
                    NotifyPropertyChanged("IsEnabled");
                }

            }
        }

        public string ImageSource
        {
            get { return this._ImageSource; }
            set
            {
                if (this._ImageSource != value)
                {
                    this._ImageSource = value;
                    NotifyPropertyChanged("ImageSource");
                }
            }
        }

        public PinValue CheckCreatTab
        {
            get { return this._CheckCreatTab; }
            set
            {
                if (this._CheckCreatTab != value)
                {
                    this._CheckCreatTab = value;
                    NotifyPropertyChanged("CheckCreatTab");
                }
            }
        }

        public PinValue BackGroundColor
        {
            get { return this._BackGround; }
            set
            {
                if (this._BackGround != value)
                {
                    this._BackGround = value;
                    NotifyPropertyChanged("BackGroundColor");
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
