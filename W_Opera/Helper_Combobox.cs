using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace W_Opera
{
    public class Helper_Combobox
    {
        private int _ID;
        private string _code;
        private string _Nam_kor;
        private string _Name_eng;
        private string _Name_loc;
        private string _ChageChar1;


        public int ID { get { return _ID; } set { if (_ID != value) { _ID = value; NotifyPropertyChanged("ID"); } } }

        public string code { get { return _code; } set { if (_code != value) { _code = value; NotifyPropertyChanged("cmpcode"); } } }
        public string Nam_kor { get { return _Nam_kor; } set { if (_Nam_kor != value) { _Nam_kor = value; NotifyPropertyChanged("cmpcode"); } } }
        public string Name_eng { get { return _Name_eng; } set { if (_Name_eng != value) { _Name_eng = value; NotifyPropertyChanged("cmpcode"); } } }
        public string Name_loc { get { return _Name_loc; } set { if (_Name_loc != value) { _Name_loc = value; NotifyPropertyChanged("cmpcode"); } } }
        public string ChageChar1 { get { return _ChageChar1; } set { if (_ChageChar1 != value) { _ChageChar1 = value; NotifyPropertyChanged("cmpcode"); } } }

        private void NotifyPropertyChanged(string Name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(Name));
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }
}
