using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Net;
using System.Windows.Controls;
using System.Data.SqlClient;
using Newtonsoft.Json;
using W_Opera;

namespace DataHelper
{
    //Kiểm tra thời gian Login và Logout==========================================
    public class TimeRun : INotifyPropertyChanged
    {
        public int ID { get; set; }
        private string _timebegin;
        public string _timefinish;
        public string TimerBegin
        {
            get { return _timebegin; }
            set
            {
                if (_timebegin != value)
                {
                    _timebegin = value;
                    NotifyPropertyChanged("TimeBegin");
                }
            }
        }
        public string TimerFinish
        {
            get { return _timefinish; }
            set
            {
                if (_timefinish != value)
                {
                    _timefinish = value;
                    NotifyPropertyChanged("TimerFinish");
                }
            }
        }
        public string _MacAddres;
        public string MacAddres
        {
            get { return _MacAddres; }
            set
            {
                if (_MacAddres != value)
                {
                    _MacAddres = value;
                    NotifyPropertyChanged("MacAddres");
                }
            }
        }
        private void NotifyPropertyChanged(string Name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(Name));
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }

    //Dữ liệu Login===============================================================
    public class LoginData : INotifyPropertyChanged
    {
        #region Khai bao
        public string _User;
        public string _Pass;
        public string _MacAddres;
        public string _Remember;
        public string _Version;
        public int ID { get; set; }
        public string User
        {
            get { return _User; }
            set
            {
                if (_User != value)
                {
                    _User = value;
                    NotifyPropertyChanged("User");
                }
            }
        }
        public string Pass
        {
            get { return _Pass; }
            set
            {
                if (_Pass != value)
                {
                    _Pass = value;
                    NotifyPropertyChanged("Pass");
                }
            }
        }
        public string MacAddres
        {
            get { return _MacAddres; }
            set
            {
                if (_MacAddres != value)
                {
                    _MacAddres = value;
                    NotifyPropertyChanged("MacAddres");
                }
            }
        }

        public string Remember
        {
            get { return _Remember; }
            set
            {
                if (_Remember != value)
                {
                    _Remember = value;
                    NotifyPropertyChanged("Remember");
                }
            }
        }

        public string Version
        {
            get { return _Version; }
            set
            {
                if (_Version != value)
                {
                    _Version = value;
                    NotifyPropertyChanged("Version");
                }
            }
        }

        private void NotifyPropertyChanged(string Name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(Name));
        }

        public event PropertyChangedEventHandler PropertyChanged;
        #endregion          
    }

    public class DataFriend : INotifyPropertyChanged
    {
        #region Khai bao 
        public string _TimeBegin;
        public string _AdminSend;
        public string _GuestSend;
        public string _MessFriend;
        public string _FileFriend;
        public string _AdminIcon;
        public string _GuestIcon;
        public string _TimeEnd;
        public int ID { get; set; }
        public string TimeBegin
        {
            get { return _TimeBegin; }
            set
            {
                if (_TimeBegin != value)
                {
                    _TimeBegin = value;
                    NotifyPropertyChanged("TimeBegin");
                }
            }
        }
        public string AdminSend
        {
            get { return _AdminSend; }
            set
            {
                if (_AdminSend != value)
                {
                    _AdminSend = value;
                    NotifyPropertyChanged("AdminSend");
                }
            }
        }
        public string GuestSend
        {
            get { return _GuestSend; }
            set
            {
                if (_GuestSend != value)
                {
                    _GuestSend = value;
                    NotifyPropertyChanged("GuestSend");
                }
            }
        }
        public string MessFriend
        {
            get { return _MessFriend; }
            set
            {
                if (_MessFriend != value)
                {
                    _MessFriend = value;
                    NotifyPropertyChanged("MessFriend");
                }
            }
        }
        public string FileFriend
        {
            get { return _FileFriend; }
            set
            {
                if (_FileFriend != value)
                {
                    _FileFriend = value;
                    NotifyPropertyChanged("FileFriend");
                }
            }
        }
        public string AdminIcon
        {
            get { return _AdminIcon; }
            set
            {
                if (_AdminIcon != value)
                {
                    _AdminIcon = value;
                    NotifyPropertyChanged("AdminIcon");
                }
            }
        }
        public string GuestIcon
        {
            get { return _GuestIcon; }
            set
            {
                if (_GuestIcon != value)
                {
                    _GuestIcon = value;
                    NotifyPropertyChanged("GuestIcon");
                }
            }
        }
        

        public string TimeEnd
        {
            get { return _TimeEnd; }
            set
            {
                if (_TimeEnd != value)
                {
                    _TimeEnd = value;
                    NotifyPropertyChanged("_TimeEnd");
                }
            }
        }


        private void NotifyPropertyChanged(string Name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(Name));
        }

        public event PropertyChangedEventHandler PropertyChanged;
        #endregion
    }

    public class DanhSachKetBan : INotifyPropertyChanged
    {
        #region Khai bao   

        public string _AdminName;
        public string _GuestName;
        public string _AdminPass;
        public string _AdminIcon;
        public string _GuestPass;
        public string _GuestIcon;
        public string _AddFriendTime;
        public int ID { get; set; }
        public string AdminName
        {
            get { return _AdminName; }
            set
            {
                if (_AdminName != value)
                {
                    _AdminName = value;
                    NotifyPropertyChanged("AdminName");
                }
            }
        }
        public string GuestName
        {
            get { return _GuestName; }
            set
            {
                if (_GuestName != value)
                {
                    _GuestName = value;
                    NotifyPropertyChanged("GuestName");
                }
            }
        }
        public string AdminPass
        {
            get { return _AdminPass; }
            set
            {
                if (_AdminPass != value)
                {
                    _AdminPass = value;
                    NotifyPropertyChanged("AdminPass");
                }
            }
        }
        public string AdminIcon
        {
            get { return _AdminIcon; }
            set
            {
                if (_AdminIcon != value)
                {
                    _AdminIcon = value;
                    NotifyPropertyChanged("AdminIcon");
                }
            }
        }
        public string GuestPass
        {
            get { return _GuestPass; }
            set
            {
                if (_GuestPass != value)
                {
                    _GuestPass = value;
                    NotifyPropertyChanged("GuestPass");
                }
            }
        }
        public string GuestIcon
        {
            get { return _GuestIcon; }
            set
            {
                if (_GuestIcon != value)
                {
                    _GuestIcon = value;
                    NotifyPropertyChanged("GuestIcon");
                }
            }
        }
        public string AddFriendTime
        {
            get { return _AddFriendTime; }
            set
            {
                if (_AddFriendTime != value)
                {
                    _AddFriendTime = value;
                    NotifyPropertyChanged("AddFriendTime");
                }
            }
        }


        private void NotifyPropertyChanged(string Name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(Name));
        }

        public event PropertyChangedEventHandler PropertyChanged;
        #endregion
    }

    #region Khai báo Taxin_Db
    public enum PinValue { ON, OFF,NotRun };
    public class Helper_TaixinDB_Input : INotifyPropertyChanged
    {
        public int _ID;
        public PinValue _rowColor;
        public string _cmpcode;
        public string _bizdiv;
        public string _inputNo;
        public string _pccode;
        public string _seq;
        public string _whgubun;
        public string _positionFromName;
        public string _DateInput;
        public string _positionFromCode;
        public string _lotNo;
        public string _modelCodel;
        public string _modelName;
        public string _warehouseName;
        public string _custumerCode;
        public string _workedNo;
        public string _allTimeInput;
        public string _hhmmTime;
        public int _hourInput;
        public int _minuteInput;
        public string _qtyInput;
        public string _unitInput;
        public string _typeInput;
        public string _inputGb;
        public string _positionChangeDate;
        public string _positionToName;
        public string _positionToCode;
        public string _userInput;
        public string _Note;
        public string _customerOS;
        //=======================================
        public string _TMSTMODEL_ModelCode;
        public string _TMSTMODEL_ModelName;
        public string _TMSTMODEL_ApplyDate;
        public string _TMSTMODEL_Version;
        public string _TMSTMODEL_CustomerPartCode;
        public string _TMSTMODEL_CustomerPartVer;
        public string _TMSTMODEL_CustomerPartCodeVer;

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
        public PinValue RowColor
        {
            get { return this._rowColor; }
            set
            {
                if (this._rowColor != value)
                {
                    this._rowColor = value;
                    NotifyPropertyChanged("RowColor");
                }
            }
        }

        public string CmpCode
        {
            get { return this._cmpcode; }
            set
            {
                if(this._cmpcode != value)
                {
                    this._cmpcode = value;
                    NotifyPropertyChanged("CmpCode");
                }
            }
        }
        public string Bizdiv
        {
            get { return this._bizdiv;}
            set
            {
                if (this._bizdiv != value)
                {
                    this._bizdiv = value;
                    NotifyPropertyChanged("Bizdiv");
                }
            }
        }
        public string InputNo
        {
            get { return this._inputNo; }
            set
            {
                if (this._inputNo != value)
                {
                    this._inputNo = value;
                    NotifyPropertyChanged("InputNo");
                }
            }
        }
        public string PositonChangeCode
        {
            get { return this._pccode; }
            set
            {
                if (this._pccode != value)
                {
                    this._pccode = value;
                    NotifyPropertyChanged("PositonChangeCode");
                }
            }
        }
        public string SeqInput
        {
            get { return this._seq; }
            set
            {
                if (this._seq != value)
                {
                    this._seq = value;
                    NotifyPropertyChanged("SeqInput");
                }
            }
        }
        public string WareHouseCode
        {
            get { return this._whgubun; }
            set
            {
                if (this._whgubun != value)
                {
                    this._whgubun = value;
                    NotifyPropertyChanged("WareHouseCode");
                }
            }
        }
        public string PositionFromName
        {
            get { return this._positionFromName;  }
            set
            {
                if (this._positionFromName != value)
                {
                    this._positionFromName  = value;
                    NotifyPropertyChanged("PositionFromName");
                }
            }
        }
        public string DateInput
        {
            get { return this._DateInput; }
            set
            {
                if (this._DateInput != value)
                {
                    this._DateInput = value;
                    NotifyPropertyChanged("DateInput");
                }
            }
        }
        public string PositionFromCode
        {
            get { return this._positionFromCode; }
            set
            {
                if (this._positionFromCode != value)
                {
                    this._positionFromCode = value;
                    NotifyPropertyChanged("PositionFromCode");
                }
            }
        }
        public string LotNo
        {
            get { return this._lotNo; }
            set
            {
                if (this._lotNo != value)
                {
                    this._lotNo = value;
                    NotifyPropertyChanged("LotNo");
                }
            }
        }
        public string ModelCodel
        {
            get { return this._modelCodel; }
            set
            {
                if (this._modelCodel != value)
                {
                    this._modelCodel = value;
                    NotifyPropertyChanged("ModelCodel");
                }
            }
        }
        public string ModelName
        {
            get { return this._modelName; }
            set
            {
                if (this._modelName != value)
                {
                    this._modelName = value;
                    NotifyPropertyChanged("ModelName");
                }
            }
        }
        public string CustomerCode
        {
            get { return this._custumerCode; }
            set
            {
                if (this._custumerCode != value)
                {
                    this._custumerCode = value;
                    NotifyPropertyChanged("CustumerCode");
                }
            }
        }
        public string WorkedNo
        {
            get { return this._workedNo; }
            set
            {
                if (this._workedNo != value)
                {
                    this._workedNo = value;
                    NotifyPropertyChanged("_workerNo");
                }
            }
        }
        public string AllTimeInput
        {
            get { return this._allTimeInput; }
            set
            {
                if (this._allTimeInput != value)
                {
                    this._allTimeInput = value;
                    NotifyPropertyChanged("AllTimeInput");
                }
            }
        }

        public string HHmmTime
        {
            get { return this._hhmmTime; }
            set
            {
                if (this._hhmmTime != value)
                {
                    this._hhmmTime = value;
                    NotifyPropertyChanged("HHmmTime");
                }
            }
        }
        public int HourInput
        {
            get { return this._hourInput; }
            set
            {
                if (this._hourInput != value)
                {
                    this._hourInput = value;
                    NotifyPropertyChanged("HourInput");
                }
            }
        }
        public int MinuteInput
        {
            get { return this._minuteInput; }
            set
            {
                if (this._minuteInput != value)
                {
                    this._minuteInput = value;
                    NotifyPropertyChanged("MinuteInput");
                }
            }
        }
        public string QtyInput
        {
            get { return this._qtyInput;}
            set
            {
                if (this._qtyInput != value)
                {
                    this._qtyInput = value;
                    NotifyPropertyChanged("QtyInput");
                }
            }
        }
        public string UnitInput
        {
            get { return this._unitInput; }
            set
            {
                if (this._unitInput != value)
                {
                    this._unitInput = value;
                    NotifyPropertyChanged("UnitInput");
                }
            }
        }
        public string TypeInput
        {
            get { return this._typeInput; }
            set
            {
                if (this._typeInput != value)
                {
                    this._typeInput = value;
                    NotifyPropertyChanged("TypeInput");
                }
            }
        }
        public string PositionChangeDate
        {
            get { return this._positionChangeDate; }
            set
            {
                if (this._positionChangeDate != value)
                {
                    this._positionChangeDate = value;
                    NotifyPropertyChanged("PositionChangeDate");
                }
            }
        }
        public string PositionToName
        {
            get { return this._positionToName; }
            set
            {
                if (this._positionToName != value)
                {
                    this._positionToName = value;
                    NotifyPropertyChanged("PositionToName");
                }
            }
        }
        public string PositionToCode
        {
            get { return this._positionToCode; }
            set
            {
                if (this._positionToCode != value)
                {
                    this._positionToCode = value;
                    NotifyPropertyChanged("PositionToCode");
                }
            }
        }
        public string UserInput
        {
            get { return this._userInput; }
            set
            {
                if (this._userInput != value)
                {
                    this._userInput = value;
                    NotifyPropertyChanged("UserInput");
                }
            }
        }
        public string Note
        {
            get { return this._Note; }
            set
            {
                if (this._Note != value)
                {
                    this._Note = value;
                    NotifyPropertyChanged("Note");
                }
            }
        }
        public string WareHouseName
        {
            get { return this._warehouseName; }
            set
            {
                if (this._warehouseName != value)
                {
                    this._warehouseName = value;
                    NotifyPropertyChanged("WareHouseName");
                }
            }
        }
        public string CustomerOS
        {
            get { return this._customerOS; }
            set
            {
                if (this._customerOS != value)
                {
                    this._customerOS = value;
                    NotifyPropertyChanged("CustomerOS");
                }
            }
        }
        public string PLInputGb
        {
            get { return this._inputGb; }
            set
            {
                if (this._inputGb != value)
                {
                    this._inputGb = value;
                    NotifyPropertyChanged("PLInputGb");
                }
            }
        }
        public string _plInputGbName;
        public string PLInputGbName
        {
            get { return this._plInputGbName; }
            set
            {
                if (this._plInputGbName != value)
                {
                    this._plInputGbName = value;
                    NotifyPropertyChanged("PLInputGbName");
                }
            }
        }

        //==========================================

        public string TMSTMODEL_ModelCode
        {
            get { return this._TMSTMODEL_ModelCode; }
            set
            {
                if (this._TMSTMODEL_ModelCode != value)
                {
                    this._TMSTMODEL_ModelCode = value;
                    NotifyPropertyChanged("TMSTMODEL_ModelCode");
                }
            }
        }
        public string TMSTMODEL_ModelName
        {
            get { return this._TMSTMODEL_ModelName; }
            set
            {
                if (this._TMSTMODEL_ModelName != value)
                {
                    this._TMSTMODEL_ModelName = value;
                    NotifyPropertyChanged("TMSTMODEL_ModelName");
                }
            }
        }
        public string TMSTMODEL_ApplyDate
        {
            get { return this._TMSTMODEL_ApplyDate; }
            set
            {
                if (this._TMSTMODEL_ApplyDate != value)
                {
                    this._TMSTMODEL_ApplyDate = value;
                    NotifyPropertyChanged("TMSTMODEL_ApplyDate");
                }
            }
        }
        public string TMSTMODEL_Version
        {
            get { return this._TMSTMODEL_Version; }
            set
            {
                if (this._TMSTMODEL_Version != value)
                {
                    this._TMSTMODEL_Version = value;
                    NotifyPropertyChanged("TMSTMODEL_Version");
                }
            }
        }
        public string TMSTMODEL_CustomerPartCode
        {
            get { return this._TMSTMODEL_CustomerPartCode; }
            set
            {
                if (this._TMSTMODEL_CustomerPartCode != value)
                {
                    this._TMSTMODEL_CustomerPartCode = value;
                    NotifyPropertyChanged("TMSTMODEL_CustomerPartCode");
                }
            }
        }
        public string TMSTMODEL_CustomerPartVer
        {
            get { return this._TMSTMODEL_CustomerPartVer; }
            set
            {
                if (this._TMSTMODEL_CustomerPartVer != value)
                {
                    this._TMSTMODEL_CustomerPartVer = value;
                    NotifyPropertyChanged("TMSTMODEL_CustomerPartVer");
                }
            }
        }
        public string TMSTMODEL_CustomerPartCodeVer
        {
            get { return this._TMSTMODEL_CustomerPartCodeVer; }
            set
            {
                if (this._TMSTMODEL_CustomerPartCodeVer != value)
                {
                    this._TMSTMODEL_CustomerPartCodeVer = value;
                    NotifyPropertyChanged("TMSTMODEL_CustomerPartCodeVer");
                }
            }
        }
        //================================================

        private void NotifyPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }        
        public event PropertyChangedEventHandler PropertyChanged;
    }
    public class Helper_TaixinDB_Output : INotifyPropertyChanged
    {
        public int _ID;
        public PinValue _rowColor;
        public string _cmpcode;
        public string _bizdiv;
        public string _outputNo;
        public string _pccode;
        public string _seq;
        public string _whgubun;
        public string _positionFromName;
        public string _dateOutput;
        public string _positionFromCode;
        public string _lotNo;
        public string _modelCodel;
        public string _modelName;
        public string _warehouseName;
        public string _custumerCode;
        public string _workedNo;
        public string _allTimeInput;
        public string _hhmmTime;
        public int _hourInput;
        public int _minuteInput;
        public string _qtyInput;
        public string _unitInput;
        public string _typeInput;
        public string _inputGb;
        public string _positionChangeDate;
        public string _positionToName;
        public string _positionToCode;
        public string _userInput;
        public string _Note;
        public string _customerOS;
        //=======================================
        public string _TMSTMODEL_ModelCode;
        public string _TMSTMODEL_ModelName;
        public string _TMSTMODEL_ApplyDate;
        public string _TMSTMODEL_Version;
        public string _TMSTMODEL_CustomerPartCode;
        public string _TMSTMODEL_CustomerPartVer;
        public string _TMSTMODEL_CustomerPartCodeVer;

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
        public PinValue RowColor
        {
            get { return this._rowColor; }
            set
            {
                if (this._rowColor != value)
                {
                    this._rowColor = value;
                    NotifyPropertyChanged("RowColor");
                }
            }
        }

        public string CmpCode
        {
            get { return this._cmpcode; }
            set
            {
                if (this._cmpcode != value)
                {
                    this._cmpcode = value;
                    NotifyPropertyChanged("CmpCode");
                }
            }
        }
        public string Bizdiv
        {
            get { return this._bizdiv; }
            set
            {
                if (this._bizdiv != value)
                {
                    this._bizdiv = value;
                    NotifyPropertyChanged("Bizdiv");
                }
            }
        }
        public string OutputNo
        {
            get { return this._outputNo; }
            set
            {
                if (this._outputNo != value)
                {
                    this._outputNo = value;
                    NotifyPropertyChanged("OutputNo");
                }
            }
        }
        public string PositonChangeCode
        {
            get { return this._pccode; }
            set
            {
                if (this._pccode != value)
                {
                    this._pccode = value;
                    NotifyPropertyChanged("PositonChangeCode");
                }
            }
        }
        public string SeqInput
        {
            get { return this._seq; }
            set
            {
                if (this._seq != value)
                {
                    this._seq = value;
                    NotifyPropertyChanged("SeqInput");
                }
            }
        }
        public string WareHouseCode
        {
            get { return this._whgubun; }
            set
            {
                if (this._whgubun != value)
                {
                    this._whgubun = value;
                    NotifyPropertyChanged("WareHouseCode");
                }
            }
        }
        public string PositionFromName
        {
            get { return this._positionFromName; }
            set
            {
                if (this._positionFromName != value)
                {
                    this._positionFromName = value;
                    NotifyPropertyChanged("PositionFromName");
                }
            }
        }
        public string DateOutput
        {
            get { return this._dateOutput; }
            set
            {
                if (this._dateOutput != value)
                {
                    this._dateOutput = value;
                    NotifyPropertyChanged("DateOutput");
                }
            }
        }
        public string PositionFromCode
        {
            get { return this._positionFromCode; }
            set
            {
                if (this._positionFromCode != value)
                {
                    this._positionFromCode = value;
                    NotifyPropertyChanged("PositionFromCode");
                }
            }
        }
        public string LotNo
        {
            get { return this._lotNo; }
            set
            {
                if (this._lotNo != value)
                {
                    this._lotNo = value;
                    NotifyPropertyChanged("LotNo");
                }
            }
        }
        public string ModelCodel
        {
            get { return this._modelCodel; }
            set
            {
                if (this._modelCodel != value)
                {
                    this._modelCodel = value;
                    NotifyPropertyChanged("ModelCodel");
                }
            }
        }
        public string ModelName
        {
            get { return this._modelName; }
            set
            {
                if (this._modelName != value)
                {
                    this._modelName = value;
                    NotifyPropertyChanged("ModelName");
                }
            }
        }
        public string CustomerCode
        {
            get { return this._custumerCode; }
            set
            {
                if (this._custumerCode != value)
                {
                    this._custumerCode = value;
                    NotifyPropertyChanged("CustumerCode");
                }
            }
        }
        public string WorkedNo
        {
            get { return this._workedNo; }
            set
            {
                if (this._workedNo != value)
                {
                    this._workedNo = value;
                    NotifyPropertyChanged("_workerNo");
                }
            }
        }
        public string AllTimeInput
        {
            get { return this._allTimeInput; }
            set
            {
                if (this._allTimeInput != value)
                {
                    this._allTimeInput = value;
                    NotifyPropertyChanged("AllTimeInput");
                }
            }
        }

        public string HHmmTime
        {
            get { return this._hhmmTime; }
            set
            {
                if (this._hhmmTime != value)
                {
                    this._hhmmTime = value;
                    NotifyPropertyChanged("HHmmTime");
                }
            }
        }
        public int HourInput
        {
            get { return this._hourInput; }
            set
            {
                if (this._hourInput != value)
                {
                    this._hourInput = value;
                    NotifyPropertyChanged("HourInput");
                }
            }
        }
        public int MinuteInput
        {
            get { return this._minuteInput; }
            set
            {
                if (this._minuteInput != value)
                {
                    this._minuteInput = value;
                    NotifyPropertyChanged("MinuteInput");
                }
            }
        }
        public string QtyInput
        {
            get { return this._qtyInput; }
            set
            {
                if (this._qtyInput != value)
                {
                    this._qtyInput = value;
                    NotifyPropertyChanged("QtyInput");
                }
            }
        }
        public string UnitInput
        {
            get { return this._unitInput; }
            set
            {
                if (this._unitInput != value)
                {
                    this._unitInput = value;
                    NotifyPropertyChanged("UnitInput");
                }
            }
        }
        public string TypeInput
        {
            get { return this._typeInput; }
            set
            {
                if (this._typeInput != value)
                {
                    this._typeInput = value;
                    NotifyPropertyChanged("TypeInput");
                }
            }
        }
        public string PositionChangeDate
        {
            get { return this._positionChangeDate; }
            set
            {
                if (this._positionChangeDate != value)
                {
                    this._positionChangeDate = value;
                    NotifyPropertyChanged("PositionChangeDate");
                }
            }
        }
        public string PositionToName
        {
            get { return this._positionToName; }
            set
            {
                if (this._positionToName != value)
                {
                    this._positionToName = value;
                    NotifyPropertyChanged("PositionToName");
                }
            }
        }
        public string PositionToCode
        {
            get { return this._positionToCode; }
            set
            {
                if (this._positionToCode != value)
                {
                    this._positionToCode = value;
                    NotifyPropertyChanged("PositionToCode");
                }
            }
        }
        public string UserInput
        {
            get { return this._userInput; }
            set
            {
                if (this._userInput != value)
                {
                    this._userInput = value;
                    NotifyPropertyChanged("UserInput");
                }
            }
        }
        public string Note
        {
            get { return this._Note; }
            set
            {
                if (this._Note != value)
                {
                    this._Note = value;
                    NotifyPropertyChanged("Note");
                }
            }
        }
        public string WareHouseName
        {
            get { return this._warehouseName; }
            set
            {
                if (this._warehouseName != value)
                {
                    this._warehouseName = value;
                    NotifyPropertyChanged("WareHouseName");
                }
            }
        }
        public string CustomerOS
        {
            get { return this._customerOS; }
            set
            {
                if (this._customerOS != value)
                {
                    this._customerOS = value;
                    NotifyPropertyChanged("CustomerOS");
                }
            }
        }
        public string PLInputGb
        {
            get { return this._inputGb; }
            set
            {
                if (this._inputGb != value)
                {
                    this._inputGb = value;
                    NotifyPropertyChanged("PLInputGb");
                }
            }
        }
        public string _plInputGbName;
        public string PLInputGbName
        {
            get { return this._plInputGbName; }
            set
            {
                if (this._plInputGbName != value)
                {
                    this._plInputGbName = value;
                    NotifyPropertyChanged("PLInputGbName");
                }
            }
        }

        //==========================================

        public string TMSTMODEL_ModelCode
        {
            get { return this._TMSTMODEL_ModelCode; }
            set
            {
                if (this._TMSTMODEL_ModelCode != value)
                {
                    this._TMSTMODEL_ModelCode = value;
                    NotifyPropertyChanged("TMSTMODEL_ModelCode");
                }
            }
        }
        public string TMSTMODEL_ModelName
        {
            get { return this._TMSTMODEL_ModelName; }
            set
            {
                if (this._TMSTMODEL_ModelName != value)
                {
                    this._TMSTMODEL_ModelName = value;
                    NotifyPropertyChanged("TMSTMODEL_ModelName");
                }
            }
        }
        public string TMSTMODEL_ApplyDate
        {
            get { return this._TMSTMODEL_ApplyDate; }
            set
            {
                if (this._TMSTMODEL_ApplyDate != value)
                {
                    this._TMSTMODEL_ApplyDate = value;
                    NotifyPropertyChanged("TMSTMODEL_ApplyDate");
                }
            }
        }
        public string TMSTMODEL_Version
        {
            get { return this._TMSTMODEL_Version; }
            set
            {
                if (this._TMSTMODEL_Version != value)
                {
                    this._TMSTMODEL_Version = value;
                    NotifyPropertyChanged("TMSTMODEL_Version");
                }
            }
        }
        public string TMSTMODEL_CustomerPartCode
        {
            get { return this._TMSTMODEL_CustomerPartCode; }
            set
            {
                if (this._TMSTMODEL_CustomerPartCode != value)
                {
                    this._TMSTMODEL_CustomerPartCode = value;
                    NotifyPropertyChanged("TMSTMODEL_CustomerPartCode");
                }
            }
        }
        public string TMSTMODEL_CustomerPartVer
        {
            get { return this._TMSTMODEL_CustomerPartVer; }
            set
            {
                if (this._TMSTMODEL_CustomerPartVer != value)
                {
                    this._TMSTMODEL_CustomerPartVer = value;
                    NotifyPropertyChanged("TMSTMODEL_CustomerPartVer");
                }
            }
        }
        public string TMSTMODEL_CustomerPartCodeVer
        {
            get { return this._TMSTMODEL_CustomerPartCodeVer; }
            set
            {
                if (this._TMSTMODEL_CustomerPartCodeVer != value)
                {
                    this._TMSTMODEL_CustomerPartCodeVer = value;
                    NotifyPropertyChanged("TMSTMODEL_CustomerPartCodeVer");
                }
            }
        }
        //================================================

        private void NotifyPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }
    public class Helper_TaixinDB_Encoder : INotifyPropertyChanged
    {
        public string _gubunCode;
        public string _code;
        public string _name_Kr;
        public string _name_En;
        public string _name_Vn;
        public string _TmsTrack_CompanyCode;
        public string _TmsTrack_CompanyDiv;
        public string _TmsTrack_Department;
        public string _TmsTrack_PositionCode;
        public string _TmsTrack_PositionName;
        public string _TmsTrack_LineName;
        public string _TmsTrack_FloorName;
        public string _TmsTrack_Seq;      


        public string GubunCode
        {
            get { return this._gubunCode; }
            set
            {
                if (this._gubunCode != value)
                {
                    this._gubunCode = value;
                    NotifyPropertyChanged("GubunCode");
                }
            }
        }
        public string CodeEncode
        {
            get { return this._code; }
            set
            {
                if (this._code != value)
                {
                    this._code = value;
                    NotifyPropertyChanged("CodeEncode");
                }
            }
        }
        public string Name_Korea
        {
            get { return this._name_Kr; }
            set
            {
                if (this._name_Kr != value)
                {
                    this._name_Kr = value;
                    NotifyPropertyChanged("Name_Korea");
                }
            }
        }
        public string Name_English
        {
            get { return this._name_En; }
            set
            {
                if (this._name_En != value)
                {
                    this._name_En = value;
                    NotifyPropertyChanged("Name_English");
                }
            }
        }
        public string Name_VietNam
        {
            get { return this._name_Vn; }
            set
            {
                if (this._name_Vn != value)
                {
                    this._name_Vn = value;
                    NotifyPropertyChanged("Name_VietNam");
                }
            }
        }
        //Bảng TMSTRACK=================
        public string TmsTrack_CompanyCode
        {
            get { return this._TmsTrack_CompanyCode; }
            set
            {
                if (this._TmsTrack_CompanyCode != value)
                {
                    this._TmsTrack_CompanyCode = value;
                    NotifyPropertyChanged("TmsTrack_CompanyCode");
                }
            }
        }
        public string TmsTrack_CompanyDiv
        {
            get { return this._TmsTrack_CompanyDiv; }
            set
            {
                if (this._TmsTrack_CompanyDiv != value)
                {
                    this._TmsTrack_CompanyDiv = value;
                    NotifyPropertyChanged("TmsTrack_CompanyDiv");
                }
            }
        }
        public string TmsTrack_Department
        {
            get { return this._TmsTrack_Department; }
            set
            {
                if (this._TmsTrack_Department != value)
                {
                    this._TmsTrack_Department = value;
                    NotifyPropertyChanged("TmsTrack_Department");
                }
            }
        }        
        public string TmsTrack_PositionName
        {
            get { return this._TmsTrack_PositionName; }
            set
            {
                if (this._TmsTrack_PositionName != value)
                {
                    this._TmsTrack_PositionName = value;
                    NotifyPropertyChanged("TmsTrack_PositionName");
                }
            }
        }
        public string TmsTrack_LineName
        {
            get { return this._TmsTrack_LineName; }
            set
            {
                if (this._TmsTrack_LineName != value)
                {
                    this._TmsTrack_LineName = value;
                    NotifyPropertyChanged("TmsTrack_LineName");
                }
            }
        }
        public string TmsTrack_FloorName
        {
            get { return this._TmsTrack_FloorName; }
            set
            {
                if (this._TmsTrack_FloorName != value)
                {
                    this._TmsTrack_FloorName = value;
                    NotifyPropertyChanged("TmsTrack_FloorName");
                }
            }
        }
        public string TmsTrack_Seq
        {
            get { return this._TmsTrack_Seq; }
            set
            {
                if (this._TmsTrack_Seq != value)
                {
                    this._TmsTrack_Seq = value;
                    NotifyPropertyChanged("TmsTrack_Seq");
                }
            }
        }


        private void NotifyPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }
    public class Helper_TaixinDB_LotNo : INotifyPropertyChanged
    {
        public int _ID;
        public string _TMWPRD_ProductionNo;
        public string _TMWPRD_ModelCode;
        public string _TMWPRD_WorkerNo;
        public int ID
        {
            get { return this._ID; }
            set
            {
                if(this._ID != value)
                {
                    this._ID = value;
                    NotifyPropertyChanged("ID");
                }
            }
        }
        public string TMWPRD_ProductionNo
        {
            get { return this._TMWPRD_ProductionNo; }
            set
            {
                if (this._TMWPRD_ProductionNo != value)
                {
                    this._TMWPRD_ProductionNo = value;
                    NotifyPropertyChanged("TMWPRD_ProductionNo");
                }
            }
        }
        public string TMWPRD_ModelCode
        {
            get { return this._TMWPRD_ModelCode; }
            set
            {
                if (this._TMWPRD_ModelCode != value)
                {
                    this._TMWPRD_ModelCode = value;
                    NotifyPropertyChanged("TMWPRD_ModelCode");
                }
            }
        }
        public string TMWPRD_WorkerNo
        {
            get { return this._TMWPRD_WorkerNo; }
            set
            {
                if (this._TMWPRD_WorkerNo != value)
                {
                    this._TMWPRD_WorkerNo = value;
                    NotifyPropertyChanged("TMWPRD_WorkerNo");
                }
            }
        }

        private void NotifyPropertyChanged(string v)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(v));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
    public class Helper_TaixinDB_Model : INotifyPropertyChanged
    {

        public string _TMSTMODEL_ModelCode;
        public string _TMSTMODEL_ModelName;
        public string _TMSTMODEL_ApplyDate;
        public string _TMSTMODEL_Version;
        public string _TMSTMODEL_CustomerPartCode;
        public string _TMSTMODEL_CustomerPartVer;
        public string _TMSTMODEL_CustomerPartCodeVer;


        public string TMSTMODEL_ModelCode
        {
            get { return this._TMSTMODEL_ModelCode; }
            set
            {
                if (this._TMSTMODEL_ModelCode != value)
                {
                    this._TMSTMODEL_ModelCode = value;
                    NotifyPropertyChanged("TMSTMODEL_ModelCode");
                }
            }
        }
        public string TMSTMODEL_ModelName
        {
            get { return this._TMSTMODEL_ModelName; }
            set
            {
                if (this._TMSTMODEL_ModelName != value)
                {
                    this._TMSTMODEL_ModelName = value;
                    NotifyPropertyChanged("TMSTMODEL_ModelName");
                }
            }
        }
        public string TMSTMODEL_ApplyDate
        {
            get { return this._TMSTMODEL_ApplyDate; }
            set
            {
                if (this._TMSTMODEL_ApplyDate != value)
                {
                    this._TMSTMODEL_ApplyDate = value;
                    NotifyPropertyChanged("TMSTMODEL_ApplyDate");
                }
            }
        }
        public string TMSTMODEL_Version
        {
            get { return this._TMSTMODEL_Version; }
            set
            {
                if (this._TMSTMODEL_Version != value)
                {
                    this._TMSTMODEL_Version = value;
                    NotifyPropertyChanged("TMSTMODEL_Version");
                }
            }
        }
        public string TMSTMODEL_CustomerPartCode
        {
            get { return this._TMSTMODEL_CustomerPartCode; }
            set
            {
                if (this._TMSTMODEL_CustomerPartCode != value)
                {
                    this._TMSTMODEL_CustomerPartCode = value;
                    NotifyPropertyChanged("TMSTMODEL_CustomerPartCode");
                }
            }
        }
        public string TMSTMODEL_CustomerPartVer
        {
            get { return this._TMSTMODEL_CustomerPartVer; }
            set
            {
                if (this._TMSTMODEL_CustomerPartVer != value)
                {
                    this._TMSTMODEL_CustomerPartVer = value;
                    NotifyPropertyChanged("TMSTMODEL_CustomerPartVer");
                }
            }
        }
        public string TMSTMODEL_CustomerPartCodeVer
        {
            get { return this._TMSTMODEL_CustomerPartCodeVer; }
            set
            {
                if (this._TMSTMODEL_CustomerPartCodeVer != value)
                {
                    this._TMSTMODEL_CustomerPartCodeVer = value;
                    NotifyPropertyChanged("TMSTMODEL_CustomerPartCodeVer");
                }
            }
        }

        private int _id;
        private string _idNumber;
        private string _cmpcode;
        private string _bizdiv;
        private string _modelcode;
        private string _modelname;
        private string _applydt;
        private string _version;
        private string _custpartcode;
        private string _custpart_version;
        private string _custpartcode_ver;
        private string _custmodelcode;
        private string _repmodelcode;
        private string _useflag;
        private string _modelgroup;
        private string _modeldiv;
        private string _modeltype;
        private string _modelchild;
        private string _cust_gb;
        private string _ta;
        private string _buyer;
        private string _color;
        private string _modelspecIn;
        private string _modellengthIn;
        private string _modelwidthIn;
        private string _modelheightIn;
        private string _modelspecOut;
        private string _modellengtOut;
        private string _modelwidthOut;
        private string _modelheightOut;
        private string _modelunfoldlength;
        private string _modelunfoldwidth;
        private string _custcode;
        private string _custshortcode;
        private string _pagecnt;
        private string _pagetypeM;
        private string _pagetypeD;
        private string _insempcode;


        private string _seq;
        private string _type;
        private string _papergubunIn;
        private string _weightIn;
        private string _widthIn;
        private string _heightIn;
        private string _sidegubunIn;
        private string _frontccIn;
        private string _backccIn;
        private string _frontbcolorIn;
        private string _backbcolorIn;
        private string _bcolorcodeIn;
        private string _phcountIn;
        private string _totpageIn;
        private string _foldingpageIn;
        private string _tayIn;
        private string _taypageIn;



        private string _papergubunOut;
        private string _weightOut;
        private string _widthOut;
        private string _heightOut;
        private string _sidegubunOut;
        private string _frontccOut;
        private string _backccOut;
        private string _frontbcolorOut;
        private string _backbcolorOut;
        private string _bcolorcodeOut;
        private string _phcountOut;
        private string _totpageOut;
        private string _foldingpageOut;
        private string _tayOut;
        private string _taypageOut;

        private string _depCreat;
        private string _printed;
        private string _versionup;
        private string _papernameIn;
        private string _papernameFullIn;
        private string _papernameOut;
        private string _papernameFullOut;
        private string _ma;
        private string _sx;
        private string _cl;
        private string _rd;
        private string _kd;
        private string _reject;
        private string _dateInput;
        private string _dateApprove;
        private string _dateStartApprove;
        private string _dateFinishApprove;
        private string _note1;
        private string _note2;
        private string _etc1;
        private string _etc3;
        private string _qtyRequest;
        private string _insempcodeUp;
        private string _updt;




        private PinValue _colorApprove;

        public int ID
        {
            get { return this._id; }
            set
            {
                if (this._id != value)
                {
                    this._id = value;
                    NotifyPropertyChanged("ID");
                }
            }

        }

        public string IDNumber
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

        public PinValue RowColor
        {
            get { return this._colorApprove; }
            set
            {
                if (this._colorApprove != value)
                {
                    this._colorApprove = value;
                    NotifyPropertyChanged("RowColor");
                }
            }
        }

        private string _checkXLS;
        public string checkXLS { get { return _checkXLS; } set { if (_checkXLS != value) { _checkXLS = value; NotifyPropertyChanged("checkXLS"); } } }

        public string CMPCODE { get { return _cmpcode; } set { if (_cmpcode != value) { _cmpcode = value; NotifyPropertyChanged("CMPCODE"); } } }
        public string BIZDIV { get { return _bizdiv; } set { if (_bizdiv != value) { _bizdiv = value; NotifyPropertyChanged("BIZDIV"); } } }
        public string MODELCODE { get { return _modelcode; } set { if (_modelcode != value) { _modelcode = value; NotifyPropertyChanged("MODELCODE"); } } }
        public string MODELNAME { get { return _modelname; } set { if (_modelname != value) { _modelname = value; NotifyPropertyChanged("MODELNAME"); } } }
        public string APPLYDT { get { return _applydt; } set { if (_applydt != value) { _applydt = value; NotifyPropertyChanged("APPLYDT"); } } }
        public string VERSION { get { return _version; } set { if (_version != value) { _version = value; NotifyPropertyChanged("VERSION"); } } }
        public string CUSTPARTCODE { get { return _custpartcode; } set { if (_custpartcode != value) { _custpartcode = value; NotifyPropertyChanged("CUSTPARTCODE"); } } }
        public string CUSTPART_VERSION { get { return _custpart_version; } set { if (_custpart_version != value) { _custpart_version = value; NotifyPropertyChanged("CUSTPART_VERSION"); } } }
        public string CUSTPARTCODE_VER { get { return _custpartcode_ver; } set { if (_custpartcode_ver != value) { _custpartcode_ver = value; NotifyPropertyChanged("CUSTPARTCODE_VER"); } } }
        public string CUSTMODELCODE { get { return _custmodelcode; } set { if (_custmodelcode != value) { _custmodelcode = value; NotifyPropertyChanged("CUSTMODELCODE"); } } }
        public string REPMODELCODE { get { return _repmodelcode; } set { if (_repmodelcode != value) { _repmodelcode = value; NotifyPropertyChanged("REPMODELCODE"); } } }
        public string USEFLAG { get { return _useflag; } set { if (_useflag != value) { _useflag = value; NotifyPropertyChanged("USEFLAG"); } } }
        public string MODELGROUP { get { return _modelgroup; } set { if (_modelgroup != value) { _modelgroup = value; NotifyPropertyChanged("MODELGROUP"); } } }
        public string MODELDIV { get { return _modeldiv; } set { if (_modeldiv != value) { _modeldiv = value; NotifyPropertyChanged("MODELDIV"); } } }
        public string MODELTYPE { get { return _modeltype; } set { if (_modeltype != value) { _modeltype = value; NotifyPropertyChanged("MODELTYPE"); } } }
        public string MODELCHILD { get { return _modelchild; } set { if (_modelchild != value) { _modelchild = value; NotifyPropertyChanged("MODELCHILD"); } } }
        public string CUST_GB { get { return _cust_gb; } set { if (_cust_gb != value) { _cust_gb = value; NotifyPropertyChanged("CUST_GB"); } } }
        public string TA { get { return _ta; } set { if (_ta != value) { _ta = value; NotifyPropertyChanged("TA"); } } }
        public string BUYER { get { return _buyer; } set { if (_buyer != value) { _buyer = value; NotifyPropertyChanged("BUYER"); } } }
        public string COLOR { get { return _color; } set { if (_color != value) { _color = value; NotifyPropertyChanged("COLOR"); } } }

        public string MODELSPECIn { get { return _modelspecIn; } set { if (_modelspecIn != value) { _modelspecIn = value; NotifyPropertyChanged("MODELSPECIn"); } } }
        public string MODELLENGTHIn { get { return _modellengthIn; } set { if (_modellengthIn != value) { _modellengthIn = value; NotifyPropertyChanged("MODELLENGTHIn"); } } }
        public string MODELWIDTHIn { get { return _modelwidthIn; } set { if (_modelwidthIn != value) { _modelwidthIn = value; NotifyPropertyChanged("MODELWIDTHIn"); } } }
        public string MODELHEIGHTIn { get { return _modelheightIn; } set { if (_modelheightIn != value) { _modelheightIn = value; NotifyPropertyChanged("MODELHEIGHTIn"); } } }


        public string MODELSPECOut { get { return _modelspecOut; } set { if (_modelspecOut != value) { _modelspecOut = value; NotifyPropertyChanged("MODELSPECOut"); } } }
        public string MODELLENGTHOut { get { return _modellengtOut; } set { if (_modellengtOut != value) { _modellengtOut = value; NotifyPropertyChanged("MODELLENGTHOut"); } } }
        public string MODELWIDTHOut { get { return _modelwidthOut; } set { if (_modelwidthOut != value) { _modelwidthOut = value; NotifyPropertyChanged("MODELWIDTHOut"); } } }
        public string MODELHEIGHTOut { get { return _modelheightOut; } set { if (_modelheightOut != value) { _modelheightOut = value; NotifyPropertyChanged("MODELHEIGHTOut"); } } }

        public string MODELUNFOLDLENGTH { get { return _modelunfoldlength; } set { if (_modelunfoldlength != value) { _modelunfoldlength = value; NotifyPropertyChanged("MODELUNFOLDLENGTH"); } } }
        public string MODELUNFOLDWIDTH { get { return _modelunfoldwidth; } set { if (_modelunfoldwidth != value) { _modelunfoldwidth = value; NotifyPropertyChanged("MODELUNFOLDWIDTH"); } } }
        public string CUSTCODE { get { return _custcode; } set { if (_custcode != value) { _custcode = value; NotifyPropertyChanged("CUSTCODE"); } } }
        public string CUSTSHORTCODE { get { return _custshortcode; } set { if (_custshortcode != value) { _custshortcode = value; NotifyPropertyChanged("CUSTSHORTCODE"); } } }
        public string PAGECNT { get { return _pagecnt; } set { if (_pagecnt != value) { _pagecnt = value; NotifyPropertyChanged("PAGECNT"); } } }
        public string PAGETYPEM { get { return _pagetypeM; } set { if (_pagetypeM != value) { _pagetypeM = value; NotifyPropertyChanged("PAGETYPEM"); } } }
        public string PAGETYPED { get { return _pagetypeD; } set { if (_pagetypeD != value) { _pagetypeD = value; NotifyPropertyChanged("PAGETYPED"); } } }

        public string VERSIONUP { get { return _versionup; } set { if (_versionup != value) { _versionup = value; NotifyPropertyChanged("VERSIONUP"); } } }
        public string INSEMPCODE { get { return _insempcode; } set { if (_insempcode != value) { _insempcode = value; NotifyPropertyChanged("INSEMPCODE"); } } }

        public string SEQ { get { return _seq; } set { if (_seq != value) { _seq = value; NotifyPropertyChanged("SEQ"); } } }
        public string TYPE { get { return _type; } set { if (_type != value) { _type = value; NotifyPropertyChanged("TYPE"); } } }
        public string PAPERGUBUNIn { get { return _papergubunIn; } set { if (_papergubunIn != value) { _papergubunIn = value; NotifyPropertyChanged("PAPERGUBUNIn"); } } }
        public string WEIGHTIn { get { return _weightIn; } set { if (_weightIn != value) { _weightIn = value; NotifyPropertyChanged("WEIGHTIn"); } } }
        public string WIDTHIn { get { return _widthIn; } set { if (_widthIn != value) { _widthIn = value; NotifyPropertyChanged("WIDTHIn"); } } }
        public string HEIGHTIn { get { return _heightIn; } set { if (_heightIn != value) { _heightIn = value; NotifyPropertyChanged("HEIGHTIn"); } } }
        public string SIDEGUBUNIn { get { return _sidegubunIn; } set { if (_sidegubunIn != value) { _sidegubunIn = value; NotifyPropertyChanged("SIDEGUBUNIn"); } } }
        public string FRONTCCIn { get { return _frontccIn; } set { if (_frontccIn != value) { _frontccIn = value; NotifyPropertyChanged("FRONTCCIn"); } } }
        public string BACKCCIn { get { return _backccIn; } set { if (_backccIn != value) { _backccIn = value; NotifyPropertyChanged("BACKCCIn"); } } }
        public string FRONTBCOLORIn { get { return _frontbcolorIn; } set { if (_frontbcolorIn != value) { _frontbcolorIn = value; NotifyPropertyChanged("FRONTBCOLORIn"); } } }
        public string BACKBCOLORIn { get { return _backbcolorIn; } set { if (_backbcolorIn != value) { _backbcolorIn = value; NotifyPropertyChanged("BACKBCOLORIn"); } } }
        public string BCOLORCODEIn { get { return _bcolorcodeIn; } set { if (_bcolorcodeIn != value) { _bcolorcodeIn = value; NotifyPropertyChanged("BCOLORCODEIn"); } } }
        public string PHCOUNTIn { get { return _phcountIn; } set { if (_phcountIn != value) { _phcountIn = value; NotifyPropertyChanged("PHCOUNTIn"); } } }
        public string TOTALPAGEIN { get { return _totpageIn; } set { if (_totpageIn != value) { _totpageIn = value; NotifyPropertyChanged("TOTALPAGEIN"); } } }
        public string FOLDINGPAGEIN { get { return _foldingpageIn; } set { if (_foldingpageIn != value) { _foldingpageIn = value; NotifyPropertyChanged("FOLDINGPAGEIN"); } } }
        public string TAYIN { get { return _tayIn; } set { if (_tayIn != value) { _tayIn = value; NotifyPropertyChanged("TAYIN"); } } }
        public string TAYPAGEIN { get { return _taypageIn; } set { if (_taypageIn != value) { _taypageIn = value; NotifyPropertyChanged("TAYPAGEIN"); } } }





        public string PAPERGUBUNOut { get { return _papergubunOut; } set { if (_papergubunOut != value) { _papergubunOut = value; NotifyPropertyChanged("PAPERGUBUNOut"); } } }
        public string WEIGHTOut { get { return _weightOut; } set { if (_weightOut != value) { _weightOut = value; NotifyPropertyChanged("WEIGHTOut"); } } }
        public string WIDTHOut { get { return _widthOut; } set { if (_widthOut != value) { _widthOut = value; NotifyPropertyChanged("WIDTHOut"); } } }
        public string HEIGHTOut { get { return _heightOut; } set { if (_heightOut != value) { _heightOut = value; NotifyPropertyChanged("HEIGHTOut"); } } }
        public string SIDEGUBUNOut { get { return _sidegubunOut; } set { if (_sidegubunOut != value) { _sidegubunOut = value; NotifyPropertyChanged("SIDEGUBUNOut"); } } }
        public string FRONTCCOut { get { return _frontccOut; } set { if (_frontccOut != value) { _frontccOut = value; NotifyPropertyChanged("FRONTCCOut"); } } }
        public string BACKCCOut { get { return _backccOut; } set { if (_backccOut != value) { _backccOut = value; NotifyPropertyChanged("BACKCCOut"); } } }
        public string FRONTBCOLOROut { get { return _frontbcolorOut; } set { if (_frontbcolorOut != value) { _frontbcolorOut = value; NotifyPropertyChanged("FRONTBCOLOROut"); } } }
        public string BACKBCOLOROut { get { return _backbcolorOut; } set { if (_backbcolorOut != value) { _backbcolorOut = value; NotifyPropertyChanged("BACKBCOLOROut"); } } }
        public string BCOLORCODEOut { get { return _bcolorcodeOut; } set { if (_bcolorcodeOut != value) { _bcolorcodeOut = value; NotifyPropertyChanged("BCOLORCODEOut"); } } }
        public string PHCOUNTOut { get { return _phcountOut; } set { if (_phcountOut != value) { _phcountOut = value; NotifyPropertyChanged("PHCOUNTOut"); } } }
        public string TOTALPAGEOUT { get { return _totpageOut; } set { if (_totpageOut != value) { _totpageOut = value; NotifyPropertyChanged("TOTALPAGEOUT"); } } }
        public string FOLDINGPAGEOUT { get { return _foldingpageOut; } set { if (_foldingpageOut != value) { _foldingpageOut = value; NotifyPropertyChanged("FOLDINGPAGEOUT"); } } }
        public string TAYOUT { get { return _tayOut; } set { if (_tayOut != value) { _tayOut = value; NotifyPropertyChanged("TAYOUT"); } } }
        public string TAYPAGEOUT { get { return _taypageOut; } set { if (_taypageOut != value) { _taypageOut = value; NotifyPropertyChanged("TAYPAGEOUT"); } } }



        public string PAPERNAMEIn { get { return _papernameIn; } set { if (_papernameIn != value) { _papernameIn = value; NotifyPropertyChanged("PAPERNAMEIn"); } } }
        public string PAPERNAME_FullIn { get { return _papernameFullIn; } set { if (_papernameFullIn != value) { _papernameFullIn = value; NotifyPropertyChanged("PAPERNAME_FullIn"); } } }

        public string PAPERNAMEOut { get { return _papernameOut; } set { if (_papernameOut != value) { _papernameOut = value; NotifyPropertyChanged("PAPERNAMEOut"); } } }
        public string PAPERNAME_FullOut { get { return _papernameFullOut; } set { if (_papernameFullOut != value) { _papernameFullOut = value; NotifyPropertyChanged("PAPERNAME_FullOut"); } } }

        public string DATEINPUT { get { return _dateInput; } set { if (_dateInput != value) { _dateInput = value; NotifyPropertyChanged("DATEINPUT"); } } }
        public string MA { get { return _ma; } set { if (_ma != value) { _ma = value; NotifyPropertyChanged("MA"); } } }
        public string SX { get { return _sx; } set { if (_sx != value) { _sx = value; NotifyPropertyChanged("SX"); } } }
        public string CL { get { return _cl; } set { if (_cl != value) { _cl = value; NotifyPropertyChanged("CL"); } } }
        public string RD { get { return _rd; } set { if (_rd != value) { _rd = value; NotifyPropertyChanged("RD"); } } }
        public string KD { get { return _kd; } set { if (_kd != value) { _kd = value; NotifyPropertyChanged("KD"); } } }
        public string depCreat { get { return _depCreat; } set { if (_depCreat != value) { _depCreat = value; NotifyPropertyChanged("depCreat"); } } }
        public string printed { get { return _printed; } set { if (_printed != value) { _printed = value; NotifyPropertyChanged("printed"); } } }
        public string reject { get { return _reject; } set { if (_reject != value) { _reject = value; NotifyPropertyChanged("reject"); } } }
        public string DATEAPPROVE { get { return _dateApprove; } set { if (_dateApprove != value) { _dateApprove = value; NotifyPropertyChanged("DATEAPPROVE"); } } }
        public string DATESTARTAPPROVE { get { return _dateStartApprove; } set { if (_dateStartApprove != value) { _dateStartApprove = value; NotifyPropertyChanged("DATESTARTAPPROVE"); } } }
        public string DATEFINISHAPPROVE { get { return _dateFinishApprove; } set { if (_dateFinishApprove != value) { _dateFinishApprove = value; NotifyPropertyChanged("DATEFINISHAPPROVE"); } } }
        public string NOTE1 { get { return _note1; } set { if (_note1 != value) { _note1 = value; NotifyPropertyChanged("NOTE1"); } } }
        public string NOTE2 { get { return _note2; } set { if (_note2 != value) { _note2 = value; NotifyPropertyChanged("NOTE2"); } } }
        public string ETC1 { get { return _etc1; } set { if (_etc1 != value) { _etc1 = value; NotifyPropertyChanged("ETC1"); } } }
        public string ETC3 { get { return _etc3; } set { if (_etc3 != value) { _etc3 = value; NotifyPropertyChanged("ETC3"); } } }
        public string QTYREQUEST { get { return _qtyRequest; } set { if (_qtyRequest != value) { _qtyRequest = value; NotifyPropertyChanged("QTYREQUEST"); } } }
        public string INSEMPCODEUP { get { return _insempcodeUp; } set { if (_insempcodeUp != value) { _insempcodeUp = value; NotifyPropertyChanged("INSEMPCODEUP"); } } }
        public string UPDT { get { return _updt; } set { if (_updt != value) { _updt = value; NotifyPropertyChanged("UPDT"); } } }
        public PinValue COLORAPPROVE { get { return _colorApprove; } set { if (_colorApprove != value) { _colorApprove = value; NotifyPropertyChanged("COLORAPPROVE"); } } }





        private void NotifyPropertyChanged(string v)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(v));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
    public class Helper_TaixinDB_DO : INotifyPropertyChanged
    {
        public int _ID;
        public PinValue _rowColor;
        public string _pcode;       //0
        public string _bizdiv;      //1
        public string _DONOtx;      //2
        public string _seq;         //3 not us
        public string _PONOtx;      //4
        public string _PONOtx_seq;  //5 not us
        public string _DONOCust;    //6
        public string _DONoseq;     //7 not us
        public string _PONOCust;    //8
        public string _PONoseq;     //9 not us
        public string _Custpartcode;    //10
        public string _delivery_date;   //11
        public string _delivery_time;   //12
        public string _delivery_qty;    //13
        public string _out_qty;         //14 not us
        public string _remain_qty;      //15 not us
        public string _delivery_loc;    //16 gubucode = '413'
        public string _DOdt;            //17
        public string _Expriedt;        //18
        public string _delivdt_ss;      //19
        public string _delivtime_ss;    //20
        public string _GRsts;           //21 not us
        public string _delivflag;       //22 not us
        public string _outyn;           //23 not us
        public string _saleorderno;     //24 not us
        public string _upload_no;       //25 not us
        public string _Insempcode;      //26
        public string _insdt;           //27
        public string _UpdempCode;      //28
        public string _Upddt;           //29
        public string _qtyOutput;
        public bool _checkSumQty;

        //================================================

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
        public PinValue RowColor
        {
            get { return this._rowColor; }
            set
            {
                if (this._rowColor != value)
                {
                    this._rowColor = value;
                    NotifyPropertyChanged("RowColor");
                }
            }
        }
        public string CmpCode
        {
            get { return this._pcode; }
            set
            {
                if (this._pcode != value)
                {
                    this._pcode = value;
                    NotifyPropertyChanged("CmpCode");
                }
            }
        }
        public string Bizdiv
        {
            get { return this._bizdiv; }
            set
            {
                if (this._bizdiv != value)
                {
                    this._bizdiv = value;
                    NotifyPropertyChanged("Bizdiv");
                }
            }
        }
        public string DONOtx
        {
            get { return this._DONOtx; }
            set
            {
                if (this._DONOtx != value)
                {
                    this._DONOtx = value;
                    NotifyPropertyChanged("DONOtx");
                }
            }
        }
        public string PONOtx
        {
            get { return this._PONOtx; }
            set
            {
                if (this._PONOtx != value)
                {
                    this._PONOtx = value;
                    NotifyPropertyChanged("PONOtx");
                }
            }
        }
        public string SeqInput
        {
            get { return this._seq; }
            set
            {
                if (this._seq != value)
                {
                    this._seq = value;
                    NotifyPropertyChanged("SeqInput");
                }
            }
        }
        public string DONOCust
        {
            get { return this._DONOCust; }
            set
            {
                if (this._DONOCust != value)
                {
                    this._DONOCust = value;
                    NotifyPropertyChanged("DONOCust");
                }
            }
        }
        public string PONOCust
        {
            get { return this._PONOCust; }
            set
            {
                if (this._PONOCust != value)
                {
                    this._PONOCust = value;
                    NotifyPropertyChanged("PONOCust");
                }
            }
        }
        public string Custpartcode
        {
            get { return this._Custpartcode; }
            set
            {
                if (this._Custpartcode != value)
                {
                    this._Custpartcode = value;
                    NotifyPropertyChanged("Custpartcode");
                }
            }
        }
        public string Delivery_date
        {
            get { return this._delivery_date; }
            set
            {
                if (this._delivery_date != value)
                {
                    this._delivery_date = value;
                    NotifyPropertyChanged("Delivery_date");
                }
            }
        }
        public string Delivery_time
        {
            get { return this._delivery_time; }
            set
            {
                if (this._delivery_time != value)
                {
                    this._delivery_time = value;
                    NotifyPropertyChanged("Delivery_time");
                }
            }
        }
        public string Delivery_qty
        {
            get { return this._delivery_qty; }
            set
            {
                if (this._delivery_qty != value)
                {
                    this._delivery_qty = value;
                    NotifyPropertyChanged("Delivery_qty");
                }
            }
        }
        public string Delivery_loc
        {
            get { return this._delivery_loc; }
            set
            {
                if (this._delivery_loc != value)
                {
                    this._delivery_loc = value;
                    NotifyPropertyChanged("Delivery_loc");
                }
            }
        }
        public string DOdt
        {
            get { return this._DOdt; }
            set
            {
                if (this._DOdt != value)
                {
                    this._DOdt = value;
                    NotifyPropertyChanged("DOdt");
                }
            }
        }
        public string Expriedt
        {
            get { return this._Expriedt; }
            set
            {
                if (this._Expriedt != value)
                {
                    this._Expriedt = value;
                    NotifyPropertyChanged("Expriedt");
                }
            }
        }
        public string Delivdt_ss
        {
            get { return this._delivdt_ss; }
            set
            {
                if (this._delivdt_ss != value)
                {
                    this._delivdt_ss = value;
                    NotifyPropertyChanged("Delivdt_ss");
                }
            }
        }
        public string Delivtime_ss
        {
            get { return this._delivtime_ss; }
            set
            {
                if (this._delivtime_ss != value)
                {
                    this._delivtime_ss = value;
                    NotifyPropertyChanged("Delivtime_ss");
                }
            }
        }
        public string Insempcode
        {
            get { return this._Insempcode; }
            set
            {
                if (this._Insempcode != value)
                {
                    this._Insempcode = value;
                    NotifyPropertyChanged("Insempcode");
                }
            }
        }
        public string insdt
        {
            get { return this._insdt; }
            set
            {
                if (this._insdt != value)
                {
                    this._insdt = value;
                    NotifyPropertyChanged("insdt");
                }
            }
        }
        public string UpdempCode
        {
            get { return this._UpdempCode; }
            set
            {
                if (this._UpdempCode != value)
                {
                    this._UpdempCode = value;
                    NotifyPropertyChanged("UpdempCode");
                }
            }
        }
        public string Upddt
        {
            get { return this._Upddt; }
            set
            {
                if (this._Upddt != value)
                {
                    this._Upddt = value;
                    NotifyPropertyChanged("Upddt");
                }
            }
        }
        public string QtyOutput
        {
            get { return this._qtyOutput; }
            set
            {
                if (this._qtyOutput != value)
                {
                    this._qtyOutput = value;
                    NotifyPropertyChanged("QtyOutput");
                }
            }
        }
        public bool CheckTotalInput
        {
            get { return this._checkSumQty; }
            set
            {
                if (this._checkSumQty != value)
                {
                    this._checkSumQty = value;
                    NotifyPropertyChanged("SumQty");
                }
            }
        }

        //================================================

        private void NotifyPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }
    public class Helper_TaixinDB_InputNVL : INotifyPropertyChanged
    {
        private int _ID;
        private PinValue _rowColor;
        private string _cmpcode;
        private string _bizdiv;
        private string _ppono;
        private string _pposeq;
        private string _itemcode;
        private string _itemtypecode;
        private string _maker;
        private string _makername;
        private string _ppoqty;
        private string _unit;
        private string _purweightqty;
        private string _purunit;
        private string _ConvNum;
        private string _currencycode;
        private string _exchangecurrency;
        private string _exchangerate;
        private string _pricetype;
        private string _unitprice;
        private string _amt;
        private string _vat;
        private string _vatrate;
        private string _unitprice_Loc;
        private string _amt_Loc;
        private string _vat_Loc;
        private string _supplydt;
        private string _requestno;
        private string _requestseq;
        private string _insempcode;
        private string _insdt;
        private string _Updempcode;
        private string _Upddt;
        private string _vendor;
        private string _intqty;
        private string _remark;


        private string _rememberQty;
        private string _subQty;
        private string _nameCustomer;
        private string _nameNVL;
        private string _nameUnit;
        private bool _checkSumQty;

        //
        private string _intno;
        private string _intseq;

        //

        private string _intdt;
        private string _outdt;
        private string _outno;
        private string _outseq;
        private string _outqty;
        private string _outqtyr;

        public int ID { get { return _ID; } set { _ID = value; NotifyPropertyChanged("ID "); } }
        public PinValue rowColor  { get { return _rowColor; } set { _rowColor = value; NotifyPropertyChanged("rowColor"); } }
        public string cmpcode { get { return _cmpcode; } set { _cmpcode = value; NotifyPropertyChanged("cmpcode "); } }
        public string bizdiv { get { return _bizdiv; } set { _bizdiv = value; NotifyPropertyChanged("bizdiv"); } }
        public string ppono { get { return _ppono; } set { _ppono = value; NotifyPropertyChanged("ppono"); } }
        public string pposeq { get { return _pposeq; } set { _pposeq = value; NotifyPropertyChanged("pposeq"); } }
        public string itemcode { get { return _itemcode; } set { _itemcode = value; NotifyPropertyChanged("itemcode"); } }
        public string itemtypecode { get { return _itemtypecode; } set { _itemtypecode = value; NotifyPropertyChanged("itemtypecode"); } }
        public string maker { get { return _maker; } set { _maker = value; NotifyPropertyChanged("maker"); } }
        public string makername { get { return _makername; } set { _makername = value; NotifyPropertyChanged("makername"); } }
        public string ppoqty { get { return _ppoqty; } set { _ppoqty = value; NotifyPropertyChanged("ppoqty"); } }
        public string unit { get { return _unit; } set { _unit = value; NotifyPropertyChanged("unit"); } }
        public string purweightqty { get { return _purweightqty; } set { _purweightqty = value; NotifyPropertyChanged("purweightqty"); } }
        public string purunit { get { return _purunit; } set { _purunit = value; NotifyPropertyChanged("purunit"); } }
        public string ConvNum { get { return _ConvNum; } set { _ConvNum = value; NotifyPropertyChanged("ConvNum"); } }
        public string currencycode { get { return _currencycode; } set { _currencycode = value; NotifyPropertyChanged("currencycode"); } }
        public string exchangecurrency { get { return _exchangecurrency; } set { _exchangecurrency = value; NotifyPropertyChanged("exchangecurrency"); } }
        public string exchangerate { get { return _exchangerate; } set { _exchangerate = value; NotifyPropertyChanged("exchangerate"); } }
        public string pricetype { get { return _pricetype; } set { _pricetype = value; NotifyPropertyChanged("pricetype"); } }
        public string unitprice { get { return _unitprice; } set { _unitprice = value; NotifyPropertyChanged("unitprice"); } }
        public string amt { get { return _amt; } set { _amt = value; NotifyPropertyChanged("amt"); } }
        public string vat { get { return _vat; } set { _vat = value; NotifyPropertyChanged("vat"); } }
        public string vatrate { get { return _vatrate; } set { _vatrate = value; NotifyPropertyChanged("vatrate"); } }
        public string unitprice_Loc { get { return _unitprice_Loc; } set { _unitprice_Loc = value; NotifyPropertyChanged("unitprice_Loc"); } }
        public string amt_Loc { get { return _amt_Loc; } set { _amt_Loc = value; NotifyPropertyChanged("amt_Loc"); } }
        public string vat_Loc { get { return _vat_Loc; } set { _vat_Loc = value; NotifyPropertyChanged("vat_Loc"); } }
        public string supplydt { get { return _supplydt; } set { _supplydt = value; NotifyPropertyChanged("supplydt"); } }
        public string requestno { get { return _requestno; } set { _requestno = value; NotifyPropertyChanged("requestno"); } }
        public string requestseq { get { return _requestseq; } set { _requestseq = value; NotifyPropertyChanged("requestseq"); } }
        public string insempcode { get { return _insempcode; } set { _insempcode = value; NotifyPropertyChanged("insempcode"); } }
        public string insdt { get { return _insdt; } set { _insdt = value; NotifyPropertyChanged("insdt"); } }
        public string Updempcode { get { return _Updempcode; } set { _Updempcode = value; NotifyPropertyChanged("Updempcode"); } }
        public string Upddt { get { return _Upddt; } set { _Upddt = value; NotifyPropertyChanged("Upddt"); } }
        public string RememberQty { get { return _rememberQty; } set { _rememberQty = value; NotifyPropertyChanged("RememberQty"); } }
        public string SubQty { get { return _subQty; } set { _subQty = value; NotifyPropertyChanged("SubQty"); } }
        public string NameCustomer { get { return _nameCustomer; } set { _nameCustomer = value; NotifyPropertyChanged("NameCustomer"); } }
        public string NameNVL { get { return _nameNVL; } set { _nameNVL = value; NotifyPropertyChanged("NameItemA"); } }     
        public string NameUnit { get { return _nameUnit; } set { _nameUnit = value; NotifyPropertyChanged("NameUnit"); } }
        public string VendorCode { get { return _vendor; } set { _vendor = value; NotifyPropertyChanged("VendorCode"); } }
        public string intqty { get { return _intqty; } set { _intqty = value; NotifyPropertyChanged("VendorCode"); } }
        public string remark { get { return _remark; } set { _remark = value; NotifyPropertyChanged("VendorCode"); } }
        public bool CheckTotalInput 
        {
            get { return this._checkSumQty; }
            set
            {
                if (this._checkSumQty != value)
                {
                    this._checkSumQty = value;
                    NotifyPropertyChanged("CheckTotalInput");
                }
            }
        }

        public string intno { get { return _intno; } set { _intno = value; NotifyPropertyChanged("intno"); } }
        public string intseq { get { return _intseq; } set { _intseq = value; NotifyPropertyChanged("intseq"); } }

        //======
        public string intdt { get { return _intdt; } set { _intdt = value; NotifyPropertyChanged("intseq"); } }
        public string outdt { get { return _outdt; } set { _outdt = value; NotifyPropertyChanged("intseq"); } }
        public string outno { get { return _outno; } set { _outno = value; NotifyPropertyChanged("intseq"); } }
        public string outseq { get { return _outseq; } set { _outseq = value; NotifyPropertyChanged("intseq"); } }
        public string outqty { get { return _outqty; } set { _outqty = value; NotifyPropertyChanged("intseq"); } }
        public string outqtyr { get { return _outqtyr; } set { _outqtyr = value; NotifyPropertyChanged("intseq"); } }









        //================================================

        private void NotifyPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class Helper_TaixinDB_SampleBox : INotifyPropertyChanged
    {
        private int _id;
        private string _cmpcode;
        private string _bizdiv;
        private string _samno;
        private string _modelcode;
        private string _modelname;
        private string _applydt;
        private string _version;
        private string _custpartcode;
        private string _custpart_version;
        private string _custpartcode_ver;
        private string _custmodelcode;
        private string _repmodelcode;
        private string _useflag;
        private string _modelgroup;
        private string _modeldiv;
        private string _modeltype;
        private string _modelchild;
        private string _cust_gb;
        private string _information;
        private string _paper_name1;
        private string _paper_size1;
        private string _paper_scale1;
        private string _paper_name2;
        private string _paper_size2;
        private string _paper_scale2;
        private string _paper_name3;
        private string _paper_size3;
        private string _paper_scale3;
        private string _cover_up_name1;
        private string _cover_up_size1;
        private string _cover_up_scale1;
        private string _cover_up_name2;
        private string _cover_up_size2;
        private string _cover_up_scale2;
        private string _cover_up_name3;
        private string _cover_up_size3;
        private string _cover_up_scale3;
        private string _cover_up_name4;
        private string _cover_up_size4;
        private string _cover_up_scale4;
        private string _cover_up_name5;
        private string _cover_up_size5;
        private string _cover_up_scale5;
        private string _fullsize;
        private string _print_color1;
        private string _print_color2;
        private string _print_color3;
        private string _print_color4;
        private string _print_color5;
        private string _coating;
        private string _glossy_color;
        private string _glossy_detail;
        private string _holo_color1;
        private string _holo_detail1;
        private string _holo_color2;
        private string _holo_detail2;
        private string _holo_color3;
        private string _holo_detail3;
        private string _holo_color4;
        private string _holo_detail4;
        private string _debosing_detail;
        private string _imbosing_detail;
        private string _boi_detaiil;
        private string _boi_color;
        private string _qty;
        private string _forecast;
        private string _remark1;
        private string _remark2;
        private string _dep_mar;
        private string _dep_pro;
        private string _dep_qc;
        private string _dep_rnd;
        private string _dep_pur;
        private string _paperindex1;
        private string _paperindex2;
        private string _paperindex3;
        private string _printindex1;
        private string _printindex2;
        private string _printindex3;
        private string _reject;
        private string _qtyAttach;
        private string _depCreate;
        private string _printed;
        private string _etc2;
        private string _etc3;
        private string _etc4;
        private string _etc5;
        private string _etc6;
        private string _etc7;
        private string _etc8;
        private string _etc9;
        private string _etc10;
        private string _sadt;
        private string _fadt;
        private string _imsempcode;
        private string _insdt;
        private string _updempcode;
        private string _upddt;

        private string _checkXLS;
        public string checkXLS { get { return _checkXLS; } set { if (_checkXLS != value) { _checkXLS = value; NotifyPropertyChanged("checkXLS"); } } }

        public int ID { get { return _id; } set { if (_id != value) { _id = value; NotifyPropertyChanged("ID"); } } }
        public string cmpcode { get { return _cmpcode; } set { if (_cmpcode != value) { _cmpcode = value; NotifyPropertyChanged("cmpcode"); } } }
        public string bizdiv { get { return _bizdiv; } set { if (_bizdiv != value) { _bizdiv = value; NotifyPropertyChanged("bizdiv"); } } }
        public string samno { get { return _samno; } set { if (_samno != value) { _samno = value; NotifyPropertyChanged("samno"); } } }
        public string modelcode { get { return _modelcode; } set { if (_modelcode != value) { _modelcode = value; NotifyPropertyChanged("modelcode"); } } }
        public string modelname { get { return _modelname; } set { if (_modelname != value) { _modelname = value; NotifyPropertyChanged("modelname"); } } }
        public string applydt { get { return _applydt; } set { if (_applydt != value) { _applydt = value; NotifyPropertyChanged("applydt"); } } }
        public string version { get { return _version; } set { if (_version != value) { _version = value; NotifyPropertyChanged("version"); } } }
        public string custpartcode { get { return _custpartcode; } set { if (_custpartcode != value) { _custpartcode = value; NotifyPropertyChanged("custpartcode"); } } }
        public string custpart_version { get { return _custpart_version; } set { if (_custpart_version != value) { _custpart_version = value; NotifyPropertyChanged("custpart_version"); } } }
        public string custpartcode_ver { get { return _custpartcode_ver; } set { if (_custpartcode_ver != value) { _custpartcode_ver = value; NotifyPropertyChanged("custpartcode_ver"); } } }
        public string custmodelcode { get { return _custmodelcode; } set { if (_custmodelcode != value) { _custmodelcode = value; NotifyPropertyChanged("custmodelcode"); } } }
        public string repmodelcode { get { return _repmodelcode; } set { if (_repmodelcode != value) { _repmodelcode = value; NotifyPropertyChanged("repmodelcode"); } } }
        public string useflag { get { return _useflag; } set { if (_useflag != value) { _useflag = value; NotifyPropertyChanged("useflag"); } } }
        public string modelgroup { get { return _modelgroup; } set { if (_modelgroup != value) { _modelgroup = value; NotifyPropertyChanged("modelgroup"); } } }
        public string modeldiv { get { return _modeldiv; } set { if (_modeldiv != value) { _modeldiv = value; NotifyPropertyChanged("modeldiv"); } } }
        public string modeltype { get { return _modeltype; } set { if (_modeltype != value) { _modeltype = value; NotifyPropertyChanged("modeltype"); } } }
        public string modelchild { get { return _modelchild; } set { if (_modelchild != value) { _modelchild = value; NotifyPropertyChanged("modelchild"); } } }
        public string cust_gb { get { return _cust_gb; } set { if (_cust_gb != value) { _cust_gb = value; NotifyPropertyChanged("cust_gb"); } } }
        public string information { get { return _information; } set { if (_information != value) { _information = value; NotifyPropertyChanged("information"); } } }
        public string paper_name1 { get { return _paper_name1; } set { if (_paper_name1 != value) { _paper_name1 = value; NotifyPropertyChanged("paper_name1"); } } }
        public string paper_size1 { get { return _paper_size1; } set { if (_paper_size1 != value) { _paper_size1 = value; NotifyPropertyChanged("paper_size1"); } } }
        public string paper_scale1 { get { return _paper_scale1; } set { if (_paper_scale1 != value) { _paper_scale1 = value; NotifyPropertyChanged("paper_scale1"); } } }
        public string paper_name2 { get { return _paper_name2; } set { if (_paper_name2 != value) { _paper_name2 = value; NotifyPropertyChanged("paper_name2"); } } }
        public string paper_size2 { get { return _paper_size2; } set { if (_paper_size2 != value) { _paper_size2 = value; NotifyPropertyChanged("paper_size2"); } } }
        public string paper_scale2 { get { return _paper_scale2; } set { if (_paper_scale2 != value) { _paper_scale2 = value; NotifyPropertyChanged("paper_scale2"); } } }
        public string paper_name3 { get { return _paper_name3; } set { if (_paper_name3 != value) { _paper_name3 = value; NotifyPropertyChanged("paper_name3"); } } }
        public string paper_size3 { get { return _paper_size3; } set { if (_paper_size3 != value) { _paper_size3 = value; NotifyPropertyChanged("paper_size3"); } } }
        public string paper_scale3 { get { return _paper_scale3; } set { if (_paper_scale3 != value) { _paper_scale3 = value; NotifyPropertyChanged("paper_scale3"); } } }
        public string cover_up_name1 { get { return _cover_up_name1; } set { if (_cover_up_name1 != value) { _cover_up_name1 = value; NotifyPropertyChanged("cover_up_name1"); } } }
        public string cover_up_size1 { get { return _cover_up_size1; } set { if (_cover_up_size1 != value) { _cover_up_size1 = value; NotifyPropertyChanged("cover_up_size1"); } } }
        public string cover_up_scale1 { get { return _cover_up_scale1; } set { if (_cover_up_scale1 != value) { _cover_up_scale1 = value; NotifyPropertyChanged("cover_up_scale1"); } } }
        public string cover_up_name2 { get { return _cover_up_name2; } set { if (_cover_up_name2 != value) { _cover_up_name2 = value; NotifyPropertyChanged("cover_up_name2"); } } }
        public string cover_up_size2 { get { return _cover_up_size2; } set { if (_cover_up_size2 != value) { _cover_up_size2 = value; NotifyPropertyChanged("cover_up_size2"); } } }
        public string cover_up_scale2 { get { return _cover_up_scale2; } set { if (_cover_up_scale2 != value) { _cover_up_scale2 = value; NotifyPropertyChanged("cover_up_scale2"); } } }
        public string cover_up_name3 { get { return _cover_up_name3; } set { if (_cover_up_name3 != value) { _cover_up_name3 = value; NotifyPropertyChanged("cover_up_name3"); } } }
        public string cover_up_size3 { get { return _cover_up_size3; } set { if (_cover_up_size3 != value) { _cover_up_size3 = value; NotifyPropertyChanged("cover_up_size3"); } } }
        public string cover_up_scale3 { get { return _cover_up_scale3; } set { if (_cover_up_scale3 != value) { _cover_up_scale3 = value; NotifyPropertyChanged("cover_up_scale3"); } } }
        public string cover_up_name4 { get { return _cover_up_name4; } set { if (_cover_up_name4 != value) { _cover_up_name4 = value; NotifyPropertyChanged("cover_up_name4"); } } }
        public string cover_up_size4 { get { return _cover_up_size4; } set { if (_cover_up_size4 != value) { _cover_up_size4 = value; NotifyPropertyChanged("cover_up_size4"); } } }
        public string cover_up_scale4 { get { return _cover_up_scale4; } set { if (_cover_up_scale4 != value) { _cover_up_scale4 = value; NotifyPropertyChanged("cover_up_scale4"); } } }
        public string cover_up_name5 { get { return _cover_up_name5; } set { if (_cover_up_name5 != value) { _cover_up_name5 = value; NotifyPropertyChanged("cover_up_name5"); } } }
        public string cover_up_size5 { get { return _cover_up_size5; } set { if (_cover_up_size5 != value) { _cover_up_size5 = value; NotifyPropertyChanged("cover_up_size5"); } } }
        public string cover_up_scale5 { get { return _cover_up_scale5; } set { if (_cover_up_scale5 != value) { _cover_up_scale5 = value; NotifyPropertyChanged("cover_up_scale5"); } } }
        public string fullsize { get { return _fullsize; } set { if (_fullsize != value) { _fullsize = value; NotifyPropertyChanged("fullsize"); } } }
        public string print_color1 { get { return _print_color1; } set { if (_print_color1 != value) { _print_color1 = value; NotifyPropertyChanged("print_color1"); } } }
        public string print_color2 { get { return _print_color2; } set { if (_print_color2 != value) { _print_color2 = value; NotifyPropertyChanged("print_color2"); } } }
        public string print_color3 { get { return _print_color3; } set { if (_print_color3 != value) { _print_color3 = value; NotifyPropertyChanged("print_color3"); } } }
        public string print_color4 { get { return _print_color4; } set { if (_print_color4 != value) { _print_color4 = value; NotifyPropertyChanged("print_color4"); } } }
        public string print_color5 { get { return _print_color5; } set { if (_print_color5 != value) { _print_color5 = value; NotifyPropertyChanged("print_color5"); } } }
        public string coating { get { return _coating; } set { if (_coating != value) { _coating = value; NotifyPropertyChanged("coating"); } } }
        public string glossy_color { get { return _glossy_color; } set { if (_glossy_color != value) { _glossy_color = value; NotifyPropertyChanged("glossy_color"); } } }
        public string glossy_detail { get { return _glossy_detail; } set { if (_glossy_detail != value) { _glossy_detail = value; NotifyPropertyChanged("glossy_detail"); } } }
        public string holo_color1 { get { return _holo_color1; } set { if (_holo_color1 != value) { _holo_color1 = value; NotifyPropertyChanged("holo_color1"); } } }
        public string holo_detail1 { get { return _holo_detail1; } set { if (_holo_detail1 != value) { _holo_detail1 = value; NotifyPropertyChanged("holo_detail1"); } } }
        public string holo_color2 { get { return _holo_color2; } set { if (_holo_color2 != value) { _holo_color2 = value; NotifyPropertyChanged("holo_color2"); } } }
        public string holo_detail2 { get { return _holo_detail2; } set { if (_holo_detail2 != value) { _holo_detail2 = value; NotifyPropertyChanged("holo_detail2"); } } }
        public string holo_color3 { get { return _holo_color3; } set { if (_holo_color3 != value) { _holo_color3 = value; NotifyPropertyChanged("holo_color3"); } } }
        public string holo_color4 { get { return _holo_color4; } set { if (_holo_color4 != value) { _holo_color4 = value; NotifyPropertyChanged("holo_color4"); } } }
        public string holo_detail3 { get { return _holo_detail3; } set { if (_holo_detail3 != value) { _holo_detail3 = value; NotifyPropertyChanged("holo_detail3"); } } }
        public string holo_detail4 { get { return _holo_detail4; } set { if (_holo_detail4 != value) { _holo_detail4 = value; NotifyPropertyChanged("holo_detail4"); } } }
        public string debosing_detail { get { return _debosing_detail; } set { if (_debosing_detail != value) { _debosing_detail = value; NotifyPropertyChanged("debosing_detail"); } } }
        public string imbosing_detail { get { return _imbosing_detail; } set { if (_imbosing_detail != value) { _imbosing_detail = value; NotifyPropertyChanged("imbosing_detail"); } } }
        public string boi_detaiil { get { return _boi_detaiil; } set { if (_boi_detaiil != value) { _boi_detaiil = value; NotifyPropertyChanged("boi_detaiil"); } } }
        public string boi_color { get { return _boi_color; } set { if (_boi_color != value) { _boi_color = value; NotifyPropertyChanged("boi_color"); } } }
        public string qty { get { return _qty; } set { if (_qty != value) { _qty = value; NotifyPropertyChanged("qty"); } } }
        public string forecast { get { return _forecast; } set { if (_forecast != value) { _forecast = value; NotifyPropertyChanged("forecast"); } } }
        public string remark1 { get { return _remark1; } set { if (_remark1 != value) { _remark1 = value; NotifyPropertyChanged("remark1"); } } }
        public string remark2 { get { return _remark2; } set { if (_remark2 != value) { _remark2 = value; NotifyPropertyChanged("remark2"); } } }
        public string dep_mar { get { return _dep_mar; } set { if (_dep_mar != value) { _dep_mar = value; NotifyPropertyChanged("dep_mar"); } } }
        public string dep_pro { get { return _dep_pro; } set { if (_dep_pro != value) { _dep_pro = value; NotifyPropertyChanged("dep_pro"); } } }
        public string dep_qc { get { return _dep_qc; } set { if (_dep_qc != value) { _dep_qc = value; NotifyPropertyChanged("dep_qc"); } } }
        public string dep_rnd { get { return _dep_rnd; } set { if (_dep_rnd != value) { _dep_rnd = value; NotifyPropertyChanged("dep_rnd"); } } }
        public string dep_pur { get { return _dep_pur; } set { if (_dep_pur != value) { _dep_pur = value; NotifyPropertyChanged("dep_pur"); } } }
        public string paperindex1 { get { return _paperindex1; } set { if (_paperindex1 != value) { _paperindex1 = value; NotifyPropertyChanged("paperindex1"); } } }
        public string paperindex2 { get { return _paperindex2; } set { if (_paperindex2 != value) { _paperindex2 = value; NotifyPropertyChanged("paperindex2"); } } }
        public string paperindex3 { get { return _paperindex3; } set { if (_paperindex3 != value) { _paperindex3 = value; NotifyPropertyChanged("paperindex3"); } } }
        public string printindex1 { get { return _printindex1; } set { if (_printindex1 != value) { _printindex1 = value; NotifyPropertyChanged("printindex1"); } } }
        public string printindex2 { get { return _printindex2; } set { if (_printindex2 != value) { _printindex2 = value; NotifyPropertyChanged("printindex2"); } } }
        public string printindex3 { get { return _printindex3; } set { if (_printindex3 != value) { _printindex3 = value; NotifyPropertyChanged("printindex3"); } } }
        public string reject { get { return _reject; } set { if (_reject != value) { _reject = value; NotifyPropertyChanged("reject"); } } }
        public string qtyAttach { get { return _qtyAttach; } set { if (_qtyAttach != value) { _qtyAttach = value; NotifyPropertyChanged("qtyAttach"); } } }
        public string depCreate { get { return _depCreate; } set { if (_depCreate != value) { _depCreate = value; NotifyPropertyChanged("depCreate"); } } }
        public string printed { get { return _printed; } set { if (_printed != value) { _printed = value; NotifyPropertyChanged("printed"); } } }
        public string etc2 { get { return _etc2; } set { if (_etc2 != value) { _etc2 = value; NotifyPropertyChanged("etc2"); } } }
        public string etc3 { get { return _etc3; } set { if (_etc3 != value) { _etc3 = value; NotifyPropertyChanged("etc3"); } } }
        public string etc4 { get { return _etc4; } set { if (_etc4 != value) { _etc4 = value; NotifyPropertyChanged("etc4"); } } }
        public string etc5 { get { return _etc5; } set { if (_etc5 != value) { _etc5 = value; NotifyPropertyChanged("etc5"); } } }
        public string etc6 { get { return _etc6; } set { if (_etc6 != value) { _etc6 = value; NotifyPropertyChanged("etc6"); } } }
        public string etc7 { get { return _etc7; } set { if (_etc7 != value) { _etc7 = value; NotifyPropertyChanged("etc7"); } } }
        public string etc8 { get { return _etc8; } set { if (_etc8 != value) { _etc8 = value; NotifyPropertyChanged("etc8"); } } }
        public string etc9 { get { return _etc9; } set { if (_etc9 != value) { _etc9 = value; NotifyPropertyChanged("etc9"); } } }
        public string etc10 { get { return _etc10; } set { if (_etc10 != value) { _etc10 = value; NotifyPropertyChanged("etc10"); } } }
        public string sadt { get { return _sadt; } set { if (_sadt != value) { _sadt = value; NotifyPropertyChanged("sadt"); } } }
        public string fadt { get { return _fadt; } set { if (_fadt != value) { _fadt = value; NotifyPropertyChanged("fadt"); } } }
        public string imsempcode { get { return _imsempcode; } set { if (_imsempcode != value) { _imsempcode = value; NotifyPropertyChanged("imsempcode"); } } }
        public string insdt { get { return _insdt; } set { if (_insdt != value) { _insdt = value; NotifyPropertyChanged("insdt"); } } }
        public string updempcode { get { return _updempcode; } set { if (_updempcode != value) { _updempcode = value; NotifyPropertyChanged("updempcode"); } } }
        public string upddt { get { return _upddt; } set { if (_upddt != value) { _upddt = value; NotifyPropertyChanged("upddt"); } } }


        private void NotifyPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        public event PropertyChangedEventHandler PropertyChanged;

    }

    public class Helper_TaixinDB_RejectSample : INotifyPropertyChanged
    {
        private string _cmpcode;
        private string _bizdiv;
        private string _samno;
        private string _seq;
        private string _typeSample;
        private string _modelcode;
        private string _filename;
        private string _filedata;
        private string _etc1;
        private string _etc2;
        private string _etc3;
        private string _etc4;
        private string _etc5;
        private string _etc6;
        private string _etc7;
        private string _etc8;
        private string _etc9;
        private string _etc10;
        private string _remark;
        private string _insempcode;
        private string _insdt;
        private string _updempcode;
        private string _upddt;
        public string cmpcode { get { return _cmpcode; } set { if (_cmpcode != value) { _cmpcode = value; NotifyPropertyChanged("cmpcode"); } } }
        public string bizdiv { get { return _bizdiv; } set { if (_bizdiv != value) { _bizdiv = value; NotifyPropertyChanged("bizdiv"); } } }
        public string samno { get { return _samno; } set { if (_samno != value) { _samno = value; NotifyPropertyChanged("samno"); } } }
        public string seq { get { return _seq; } set { if (_seq != value) { _seq = value; NotifyPropertyChanged("seq"); } } }
        public string typeSample { get { return _typeSample; } set { if (_typeSample != value) { _typeSample = value; NotifyPropertyChanged("typeSample"); } } }
        public string modelcode { get { return _modelcode; } set { if (_modelcode != value) { _modelcode = value; NotifyPropertyChanged("modelcode"); } } }
        public string filename { get { return _filename; } set { if (_filename != value) { _filename = value; NotifyPropertyChanged("filename"); } } }
        public string filedata { get { return _filedata; } set { if (_filedata != value) { _filedata = value; NotifyPropertyChanged("filedata"); } } }
        public string etc1 { get { return _etc1; } set { if (_etc1 != value) { _etc1 = value; NotifyPropertyChanged("etc1"); } } }
        public string etc2 { get { return _etc2; } set { if (_etc2 != value) { _etc2 = value; NotifyPropertyChanged("etc2"); } } }
        public string etc3 { get { return _etc3; } set { if (_etc3 != value) { _etc3 = value; NotifyPropertyChanged("etc3"); } } }
        public string etc4 { get { return _etc4; } set { if (_etc4 != value) { _etc4 = value; NotifyPropertyChanged("etc4"); } } }
        public string etc5 { get { return _etc5; } set { if (_etc5 != value) { _etc5 = value; NotifyPropertyChanged("etc5"); } } }
        public string etc6 { get { return _etc6; } set { if (_etc6 != value) { _etc6 = value; NotifyPropertyChanged("etc6"); } } }
        public string etc7 { get { return _etc7; } set { if (_etc7 != value) { _etc7 = value; NotifyPropertyChanged("etc7"); } } }
        public string etc8 { get { return _etc8; } set { if (_etc8 != value) { _etc8 = value; NotifyPropertyChanged("etc8"); } } }
        public string etc9 { get { return _etc9; } set { if (_etc9 != value) { _etc9 = value; NotifyPropertyChanged("etc9"); } } }
        public string etc10 { get { return _etc10; } set { if (_etc10 != value) { _etc10 = value; NotifyPropertyChanged("etc10"); } } }
        public string remark { get { return _remark; } set { if (_remark != value) { _remark = value; NotifyPropertyChanged("remark"); } } }
        public string insempcode { get { return _insempcode; } set { if (_insempcode != value) { _insempcode = value; NotifyPropertyChanged("insempcode"); } } }
        public string insdt { get { return _insdt; } set { if (_insdt != value) { _insdt = value; NotifyPropertyChanged("insdt"); } } }
        public string updempcode { get { return _updempcode; } set { if (_updempcode != value) { _updempcode = value; NotifyPropertyChanged("updempcode"); } } }
        public string upddt { get { return _upddt; } set { if (_upddt != value) { _upddt = value; NotifyPropertyChanged("upddt"); } } }

        private void NotifyPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        public event PropertyChangedEventHandler PropertyChanged;
    }
    #endregion


    //Các hàm xử lý dữ liệu thêm xóa sửa Database=======================================================================
    public class DataBaseHelper
    {
        //public static string nameDb = "MyDataBase";//Đặt tên cho Database
        //public static string nameTable = "Login";//Đặt tên cho bảng dữ liệu
        //Tạo đường dẫn để tạo mới 1 DATABASE============================================================
        //string pathCreatDb = "Data Source = (local);User ID =sa;Password=123456;Integrated Security = True;";

        //Đường dẫn sau khi đã tạo được DATABASE dùng để tạo mới bảng dữ liệu============================

        //Kiểm tra bảng và tạo bảng dữ liệu
        #region Kiểm tra bảng và tạo bảng dữ liệu
        public void CreatDatabase(string nameDb)
        {
            try
            {
                string pathCreatDb = "Data Source = (local);Integrated Security = True;";
                using (SqlConnection conn = new SqlConnection(pathCreatDb))
                {
                    conn.Open();

                    //Truy vấn xem trong toàn bộ DATABASE đã tồn tại Database định tạo chưa
                    //Nếu chưa tồn tại thì tạo mới nếu tồn tại rồi thì xuất ra thông báo
                    var command = "SELECT * FROM master.dbo.sysdatabases WHERE name= '" + nameDb + "'";
                    SqlCommand cmdCheckDb = new SqlCommand(command, conn);
                    bool checkDatabase = false;
                    using (SqlDataReader dr = cmdCheckDb.ExecuteReader())
                    {
                        if (dr.HasRows == true)
                        {
                            checkDatabase = true;
                        }
                        else if (dr.HasRows == false)
                        {
                            checkDatabase = false;
                        }
                    }                    
                    if (checkDatabase == false)
                    {
                        string comm = "CREATE DATABASE " + nameDb + "";
                        SqlCommand cmdCreatDb = new SqlCommand(comm, conn);
                        cmdCreatDb.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "DataBase.CreatTable", "CreatTable", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        //Tạo bảng dữ liệu
        #region Tạo bảng dữ liệu

        public void CreatTable(string nameDb, string nameTable, int numberColum, string[] listColum)
        {
            string pathCreatTable = "Data Source=(local);Initial Catalog='" + nameDb + "';Integrated Security=True";
            try
            {
                
                using (SqlConnection conn = new SqlConnection(pathCreatTable))
                {
                    conn.Open();
                    string nameColums = "";
                    for (int i = 0; i < numberColum; i++)
                    {
                        if (i < numberColum - 1)
                        {
                            nameColums = nameColums + listColum[i] + ",";
                        }
                        if (i == numberColum - 1)
                        {
                            nameColums = nameColums + listColum[i];
                        }
                    }
                    //var command = "CREATE TABLE IF NOT EXISTS " + nameTable + "(" + nameColums + ")";
                    var command = " If not exists (select name from sysobjects where name = '" + nameTable + "') CREATE TABLE " + nameTable + "("+nameColums+")";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "DataBase.CreatTable", "CreatTable", MessageBoxButton.OK, MessageBoxImage.Error);
            }       
        }

        #endregion

        //Đọc dữ liệu từ bảng TaixinDb_Input ra==========================================================================
        string namePage = "Database Infomation";       
        public ObservableCollection<Helper_TaixinDB_Input> Read_TaxinDb_Input(string path_sql, string nameTable, bool search, string WareHouseCode, string inputgb,
            string date_start, string date_finish)//,string posName,string plInputGb,string uintName)//, 
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    ObservableCollection<Helper_TaixinDB_Input> list_TaixnDb = new ObservableCollection<Helper_TaixinDB_Input>();
                    string command = "";
                    if (search == true)
                    {
                        command = "SELECT * from " + nameTable + " WHERE whgubun = '" + WareHouseCode + "' and inputgb ='" + inputgb + "' and insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";
                    }
                    else
                    {
                        command = "SELECT * from " + nameTable + "  WHERE inputgb ='" + inputgb + "' and insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";
                    }

                    //command = "SELECT A.cmpcode,A.bizdiv,A.intno,A.pccode,A.seq,A.whgubun,A.pstcode,A.intdt," +
                    //        "A.pstgb,A.lotno,A.modelcode,A.custpartcode,A.wono,A.inputtime,A.inthh," +
                    //        "A.intmm,A.inputqty,A.unit,A.inputgb,A.custcode_os,A.etc,A.pcgb,A.useyn," +
                    //        "A.pcdt,A.prepstcode,A.prepccode,A.preoutno,A.prepstgb,A.intno_m," +
                    //        "A.remark,A.insempcode,A.insdt,A.upsempcode,A.upddt," +
                    //        "B.modelcode,B.modelname,B.applydt,B.version,B.custpartcode,B.custpart_version," +
                    //        "B.custpartcode_ver,C.Name_loc " +
                    //        "FROM " + nameTable + " A full outer join TMSTMODEL B ON A.modelcode = B.modelcode FULL outer join TEMSTETC C ON A.inputgb = C.Code WHERE C.GubunCode = '402' and A.insdt <= DATEADD(hour,1,'" + date_finish + "') ORDER BY A.insdt desc";



                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            //int coutRow = 0;
                            try
                            {
                                while (dr.Read())
                                {
                                    Helper_TaixinDB_Input taixinDB = new Helper_TaixinDB_Input();

                                    if (dr[0] != null)
                                    {
                                        var formatDate = new JsonSerializerSettings { DateFormatString = "yyyy/MM/dd HH:mm:ss" };
                                        var json = JsonConvert.SerializeObject(dr[33], formatDate);
                                        //taixinDB.ID = coutRow;
                                        taixinDB.CmpCode = dr[0].ToString();
                                        taixinDB.Bizdiv = dr[1].ToString();
                                        taixinDB.InputNo = dr[2].ToString();
                                        taixinDB.PositonChangeCode = dr[3].ToString();
                                        taixinDB.SeqInput = dr[4].ToString();
                                        taixinDB.WareHouseCode = dr[5].ToString();
                                        switch (dr[5].ToString())
                                        {
                                            case "01":
                                                {
                                                    taixinDB.WareHouseName = "Kho Thành Phẩm";
                                                    break;
                                                }
                                            case "02":
                                                {
                                                    taixinDB.WareHouseName = "Kho BTP Box Sabari";
                                                    break;
                                                }
                                            case "03":
                                                {
                                                    taixinDB.WareHouseName = "Kho BTP Manual";
                                                    break;
                                                }
                                            case "04":
                                                {
                                                    taixinDB.WareHouseName = "Kho TP Cushion(T2)";
                                                    break;
                                                }
                                            case "05":
                                                {
                                                    taixinDB.WareHouseName = "Kho BTP Box Thường";
                                                    break;
                                                }
                                            case "06":
                                                {
                                                    taixinDB.WareHouseName = "Kho BTP Box Gia Công";
                                                    break;
                                                }

                                            case "09":
                                                {
                                                    taixinDB.WareHouseName = "Kho W";
                                                    break;
                                                }
                                        }

                                        taixinDB.PositionToName = dr[6].ToString();
                                        string dateInput = dr[7].ToString();
                                        if (dateInput.Length > 6)
                                        {
                                            taixinDB.DateInput = dateInput.Substring(0, 4) + "/" + dateInput.Substring(4, 2) + "/" + dateInput.Substring(6, 2);
                                        }                                            
                                        taixinDB.PositionToCode = dr[8].ToString();
                                        taixinDB.LotNo = dr[9].ToString();
                                        taixinDB.ModelCodel = dr[10].ToString();
                                        //taixinDB.ModelName = Read_TaxinDb_1ModelCode(path_sql, "tmstmodel", dr[10].ToString());
                                        taixinDB.CustomerCode = dr[11].ToString();
                                        taixinDB.WorkedNo = dr[12].ToString();
                                        
                                        if (json.Length>18)
                                        {
                                            
                                            taixinDB.HHmmTime = json.Substring(12, 5);
                                            taixinDB.AllTimeInput = json.Substring(1, json.Length - 2);
                                            taixinDB.HourInput = int.Parse(json.Substring(12, 2));
                                            taixinDB.MinuteInput = int.Parse(json.Substring(15, 2));
                                        }                                       
                                        taixinDB.QtyInput = dr[16].ToString();
                                        if (dr[17].ToString() == "03")
                                        {
                                            taixinDB.UnitInput = "Cái";
                                        }
                                        taixinDB.PLInputGb = dr[18].ToString();
                                        taixinDB.CustomerOS = dr[19].ToString();
                                        taixinDB.PositionFromName = dr[26].ToString();
                                        taixinDB.PositionFromCode = dr[29].ToString();                                       
                                        taixinDB.Note = dr[31].ToString();
                                        taixinDB.UserInput = dr[32].ToString();
                                        //taixinDB.TMSTMODEL_ModelCode = dr[34].ToString();
                                        //taixinDB.TMSTMODEL_ModelName = dr[35].ToString();
                                        //taixinDB.TMSTMODEL_ApplyDate = dr[36].ToString();
                                        //taixinDB.TMSTMODEL_Version = dr[37].ToString();
                                        //taixinDB.TMSTMODEL_CustomerPartCode = dr[38].ToString();
                                        //taixinDB.TMSTMODEL_CustomerPartVer = dr[39].ToString();
                                        //taixinDB.TMSTMODEL_CustomerPartCodeVer = dr[40].ToString();
                                        //taixinDB.PLInputGbName = dr[41].ToString();                                      
                                        list_TaixnDb.Add(taixinDB);
                                    }
                                }
                            }
                            catch (Exception )
                            {
                                throw;
                            }
                            
                        }
                    }
                    conn.Close();
                  
                    return list_TaixnDb;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Read_TaxinDb_Input", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<Helper_TaixinDB_Output> Read_TaxinDb_Output(string path_sql, string nameTable, bool search, string WareHouseCode, string outputGb,
            string date_start, string date_finish)//,string posName,string plInputGb,string uintName)//, 
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    List<Helper_TaixinDB_Output> list_TaixnDb = new List<Helper_TaixinDB_Output>();
                    string command = "";
                    if (search == true)
                    {
                        command = "SELECT * from " + nameTable + " WHERE whgubun = '" + WareHouseCode + "' and outgb ='" + outputGb + "' and insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "') ORDER BY insdt DESC";
                    }
                    else
                    {
                        command = "SELECT * from " + nameTable + "  WHERE outgb ='" + outputGb + "' and insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "') ORDER BY insdt DESC";
                    }                    

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {                           
                            try
                            {
                                while (dr.Read())
                                {
                                    Helper_TaixinDB_Output taixinDB = new Helper_TaixinDB_Output();

                                    if (dr[0] != null)
                                    {
                                        var formatDate = new JsonSerializerSettings { DateFormatString = "yyyy/MM/dd HH:mm:ss" };
                                        var json = JsonConvert.SerializeObject(dr[36], formatDate);                                       
                                        taixinDB.CmpCode = dr[0].ToString();
                                        taixinDB.Bizdiv = dr[1].ToString();
                                        taixinDB.OutputNo = dr[2].ToString();
                                        taixinDB.PositonChangeCode = dr[3].ToString();
                                        taixinDB.SeqInput = dr[4].ToString();
                                        taixinDB.WareHouseCode = dr[5].ToString();
                                        switch (dr[5].ToString())
                                        {
                                            case "01":
                                                {
                                                    taixinDB.WareHouseName = "Kho Thành Phẩm";
                                                    break;
                                                }
                                            case "02":
                                                {
                                                    taixinDB.WareHouseName = "Kho BTP Box Sabari";
                                                    break;
                                                }
                                            case "03":
                                                {
                                                    taixinDB.WareHouseName = "Kho BTP Manual";
                                                    break;
                                                }
                                            case "04":
                                                {
                                                    taixinDB.WareHouseName = "Kho TP Cushion(T2)";
                                                    break;
                                                }
                                            case "05":
                                                {
                                                    taixinDB.WareHouseName = "Kho BTP Box Thường";
                                                    break;
                                                }
                                            case "06":
                                                {
                                                    taixinDB.WareHouseName = "Kho BTP Box Gia Công";
                                                    break;
                                                }

                                            case "09":
                                                {
                                                    taixinDB.WareHouseName = "Kho W";
                                                    break;
                                                }
                                        }
                                        taixinDB.PositionToName = dr[6].ToString();
                                        string dateOutput = dr[7].ToString();
                                        if(dateOutput.Length>6)
                                        taixinDB.DateOutput = dateOutput.Substring(0, 4) + "/" + dateOutput.Substring(4, 2) + "/" + dateOutput.Substring(6,2);
                                        taixinDB.PositionToCode = dr[8].ToString();
                                        taixinDB.LotNo = dr[9].ToString();
                                        taixinDB.ModelCodel = dr[10].ToString();
                                        //taixinDB.ModelName = Read_TaxinDb_1ModelCode(path_sql, "tmstmodel", dr[10].ToString());
                                        taixinDB.CustomerCode = dr[11].ToString();
                                        taixinDB.WorkedNo = dr[14].ToString();
                                        if(json.Length>18)
                                        {
                                            taixinDB.HHmmTime = json.Substring(12, 5);
                                            taixinDB.AllTimeInput = json.Substring(1, json.Length - 2);
                                            taixinDB.HourInput = int.Parse(json.Substring(12, 2));
                                            taixinDB.MinuteInput = int.Parse(json.Substring(15, 2));
                                        }                                       
                                        taixinDB.QtyInput = dr[18].ToString();
                                        if (dr[19].ToString() == "03")
                                        {
                                            taixinDB.UnitInput = "Cái";
                                        }
                                        taixinDB.PLInputGb = dr[20].ToString();
                                        //taixinDB.PLInputGb = dr[18].ToString();
                                        //taixinDB.CustomerOS = dr[19].ToString();
                                        //taixinDB.PositionFromCode = dr[27].ToString();
                                        //taixinDB.PositionFromName = dr[24].ToString();
                                        taixinDB.Note = dr[34].ToString();
                                        taixinDB.UserInput = dr[35].ToString();                                       
                                        list_TaixnDb.Add(taixinDB);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }

                        }
                    }
                    conn.Close();                   
                    return list_TaixnDb;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper Database/Read_TaxinDb_Output", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }


        public List<Helper_TaixinDB_Encoder> Read_TaxinDb_Encoder(string path_sql, string nameTable)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    List<Helper_TaixinDB_Encoder> listEncoder= new List<Helper_TaixinDB_Encoder>();
                    var command = "SELECT * from " + nameTable + "";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                Helper_TaixinDB_Encoder _encoder = new Helper_TaixinDB_Encoder();
                                if (dr[0] != null)
                                {
                                    _encoder._gubunCode = dr[0].ToString();
                                    _encoder._code = dr[1].ToString();
                                    _encoder._name_Kr = dr[2].ToString();
                                    _encoder._name_En = dr[3].ToString();
                                    _encoder._name_Vn = dr[4].ToString();                                   
                                    listEncoder.Add(_encoder);
                                }
                            }
                        }
                    }
                    conn.Close();                  
                    return listEncoder;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper Database/Read_TaxinDb_Encoder", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<Helper_TaixinDB_Encoder> Read_TaxinDb_RACK(string path_sql, string nameTable)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    List<Helper_TaixinDB_Encoder> listCodeRack = new List<Helper_TaixinDB_Encoder>();
                    var command = "SELECT * from " + nameTable + "";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                Helper_TaixinDB_Encoder taixinDB = new Helper_TaixinDB_Encoder();
                                if (dr[0] != null)
                                {
                                    taixinDB.TmsTrack_CompanyCode = dr[0].ToString();
                                    taixinDB.TmsTrack_CompanyDiv = dr[1].ToString();                                   
                                    taixinDB.TmsTrack_Department = dr[2].ToString();
                                    taixinDB.TmsTrack_PositionName = dr[3].ToString();
                                    taixinDB.TmsTrack_LineName = dr[4].ToString();
                                    taixinDB.TmsTrack_FloorName = dr[5].ToString();
                                    taixinDB.TmsTrack_Seq = dr[6].ToString();                                   
                                    listCodeRack.Add(taixinDB);
                                }
                            }
                        }
                    }
                    conn.Close();                   
                    return listCodeRack;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaxinDb_RACK", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public Helper_TaixinDB_LotNo Read_TaxinDb_LotNo(string path_sql, string nameTable,string nameLotNo)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    
                    Helper_TaixinDB_LotNo taixinDB = new Helper_TaixinDB_LotNo();
                    var command = "SELECT * from " + nameTable + " WHERE PRDNO = '" + nameLotNo + "'";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {                                
                                if (dr[0] != null)
                                {
                                    taixinDB.TMWPRD_ProductionNo = dr[2].ToString();
                                    taixinDB.TMWPRD_ModelCode = dr[4].ToString();
                                    taixinDB.TMWPRD_WorkerNo = dr[12].ToString(); 
                                }
                            }
                        }
                    }
                    conn.Close();                  
                    return taixinDB;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaxinDb_LotNo", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<Helper_TaixinDB_Model> Read_TaxinDb_ModelCode(string path_sql, string nameTable,string nameColumnSort,string strCustomerSort)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    List<Helper_TaixinDB_Model> listModelCode = new List<Helper_TaixinDB_Model>();
                    var command = "SELECT * from " + nameTable + " a Full Outer Join tmstmodelpaper b " +
                        "on a.modelcode = b.modelcode Full Outer Join tmstpaper c on b.papergubun = c.papercode  WHERE a." + nameColumnSort+" like '%" + strCustomerSort+ "%' ORDER BY a.Insdt DESC";
                    //var command = "SELECT * from " + nameTable + "";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                Helper_TaixinDB_Model taixinDB = new Helper_TaixinDB_Model();
                                if (dr[0] != null)
                                {
                                    taixinDB.TMSTMODEL_ModelCode = dr[2].ToString();
                                    taixinDB.TMSTMODEL_ModelName = dr[3].ToString();
                                    taixinDB.TMSTMODEL_ApplyDate = dr[4].ToString();
                                    taixinDB.TMSTMODEL_Version = dr[5].ToString();
                                    taixinDB.TMSTMODEL_CustomerPartCode = dr[6].ToString();
                                    taixinDB.TMSTMODEL_CustomerPartVer = dr[7].ToString();
                                    taixinDB._TMSTMODEL_CustomerPartCodeVer = dr[8].ToString();
                                    taixinDB.SEQ = dr[66].ToString();
                                   
                                    if (taixinDB.SEQ=="1" || taixinDB.SEQ =="")
                                    {
                                        //==========================
                                        taixinDB.CMPCODE = dr[0].ToString();
                                        taixinDB.BIZDIV = dr[1].ToString();
                                        taixinDB.MODELCODE = dr[2].ToString();
                                        taixinDB.MODELNAME = dr[3].ToString();
                                        taixinDB.APPLYDT = dr[4].ToString();
                                        taixinDB.VERSION = dr[5].ToString();
                                        taixinDB.CUSTPARTCODE = dr[6].ToString();
                                        taixinDB.CUSTPART_VERSION = dr[7].ToString();
                                        taixinDB.CUSTPARTCODE_VER = dr[8].ToString();
                                        taixinDB.CUSTMODELCODE = dr[9].ToString();
                                        taixinDB.REPMODELCODE = dr[10].ToString();
                                        taixinDB.USEFLAG = dr[11].ToString();
                                        taixinDB.MODELGROUP = dr[12].ToString();
                                        taixinDB.MODELDIV = dr[13].ToString();
                                        taixinDB.MODELTYPE = dr[14].ToString();
                                        taixinDB.MODELCHILD = dr[15].ToString();
                                        taixinDB.CUST_GB = dr[16].ToString();
                                        taixinDB.TA = dr[17].ToString();
                                        taixinDB.BUYER = dr[18].ToString();
                                        taixinDB.COLOR = dr[19].ToString();
                                        taixinDB.MODELSPECOut = dr[20].ToString();
                                        taixinDB.MODELLENGTHOut = dr[21].ToString();
                                        taixinDB.MODELWIDTHOut = dr[22].ToString();
                                        taixinDB.MODELHEIGHTOut = dr[23].ToString();
                                        taixinDB.MODELUNFOLDLENGTH = dr[24].ToString();
                                        taixinDB.MODELUNFOLDWIDTH = dr[25].ToString();
                                        taixinDB.CUSTCODE = dr[26].ToString();
                                        taixinDB.CUSTSHORTCODE = dr[27].ToString();
                                        taixinDB.PAGECNT = dr[28].ToString();
                                        taixinDB.VERSIONUP = dr[44].ToString();
                                        taixinDB.INSEMPCODE = dr[59].ToString();

                                        //===========================

                                        taixinDB.SEQ = dr[66].ToString();
                                        taixinDB.TYPE = dr[67].ToString();
                                        taixinDB.PAPERGUBUNOut = dr[68].ToString();
                                        taixinDB.WEIGHTOut = dr[69].ToString();
                                        taixinDB.WIDTHOut = dr[70].ToString();
                                        taixinDB.HEIGHTOut = dr[71].ToString();
                                        taixinDB.SIDEGUBUNOut = dr[72].ToString();
                                        taixinDB.FRONTCCOut = dr[73].ToString();
                                        taixinDB.BACKCCOut = dr[74].ToString();
                                        taixinDB.FRONTBCOLOROut = dr[75].ToString();
                                        taixinDB.BACKBCOLOROut = dr[76].ToString();
                                        taixinDB.BCOLORCODEOut = dr[77].ToString();
                                        taixinDB.PHCOUNTOut = dr[78].ToString();
                                        taixinDB.TOTALPAGEOUT = dr[81].ToString();
                                        taixinDB.FOLDINGPAGEOUT = dr[82].ToString();
                                        taixinDB.TAYOUT = dr[83].ToString();
                                        taixinDB.TAYPAGEOUT = dr[84].ToString();
                                        taixinDB.PAPERNAMEOut = dr[98].ToString();
                                        taixinDB.PAPERNAME_FullOut = dr[99].ToString();
                                    }
                                    else if(taixinDB.SEQ=="2")
                                    {
                                        taixinDB.CMPCODE = dr[0].ToString();
                                        taixinDB.BIZDIV = dr[1].ToString();
                                        taixinDB.MODELCODE = dr[2].ToString();
                                        taixinDB.MODELNAME = dr[3].ToString();
                                        taixinDB.APPLYDT = dr[4].ToString();
                                        taixinDB.VERSION = dr[5].ToString();
                                        taixinDB.CUSTPARTCODE = dr[6].ToString();
                                        taixinDB.CUSTPART_VERSION = dr[7].ToString();
                                        taixinDB.CUSTPARTCODE_VER = dr[8].ToString();
                                        taixinDB.CUSTMODELCODE = dr[9].ToString();
                                        taixinDB.REPMODELCODE = dr[10].ToString();
                                        taixinDB.USEFLAG = dr[11].ToString();
                                        taixinDB.MODELGROUP = dr[12].ToString();
                                        taixinDB.MODELDIV = dr[13].ToString();
                                        taixinDB.MODELTYPE = dr[14].ToString();
                                        taixinDB.MODELCHILD = dr[15].ToString();
                                        taixinDB.CUST_GB = dr[16].ToString();
                                        taixinDB.TA = dr[17].ToString();
                                        taixinDB.BUYER = dr[18].ToString();
                                        taixinDB.COLOR = dr[19].ToString();
                                        taixinDB.MODELSPECIn = dr[20].ToString();
                                        taixinDB.MODELLENGTHIn = dr[21].ToString();
                                        taixinDB.MODELWIDTHIn = dr[22].ToString();
                                        taixinDB.MODELHEIGHTIn = dr[23].ToString();

                                        taixinDB.SEQ = dr[66].ToString();
                                        taixinDB.TYPE = dr[67].ToString();
                                        taixinDB.PAPERGUBUNIn = dr[68].ToString();
                                        taixinDB.WEIGHTIn = dr[69].ToString();
                                        taixinDB.WIDTHIn = dr[70].ToString();
                                        taixinDB.HEIGHTIn = dr[71].ToString();
                                        taixinDB.SIDEGUBUNIn = dr[72].ToString();
                                        taixinDB.FRONTCCIn = dr[73].ToString();
                                        taixinDB.BACKCCIn = dr[74].ToString();
                                        taixinDB.FRONTBCOLORIn = dr[75].ToString();
                                        taixinDB.BACKBCOLORIn = dr[76].ToString();
                                        taixinDB.BCOLORCODEIn = dr[77].ToString();
                                        taixinDB.PHCOUNTIn = dr[78].ToString();
                                        taixinDB.TOTALPAGEIN = dr[81].ToString();
                                        taixinDB.FOLDINGPAGEIN = dr[82].ToString();
                                        taixinDB.TAYIN = dr[83].ToString();
                                        taixinDB.TAYPAGEIN = dr[84].ToString();
                                        taixinDB.PAPERNAMEIn = dr[98].ToString();
                                        taixinDB.PAPERNAME_FullIn = dr[99].ToString();
                                    }                

                                    listModelCode.Add(taixinDB);
                                }
                            }
                        }
                    }
                    conn.Close();                  
                    return listModelCode;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaxinDb_ModelCode", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public string Read_TaxinDb_1ModelCode(string path_sql, string nameTable,string model)
        {
            try
            {
                string modelName = "";
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();                   
                    var command = "SELECT * from " + nameTable + "";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                Helper_TaixinDB_Model taixinDB = new Helper_TaixinDB_Model();
                                if (dr[0] != null)
                                {
                                    if(taixinDB.TMSTMODEL_ModelCode == model)
                                    {
                                        modelName = taixinDB.TMSTMODEL_ModelName;                                      
                                    }
                                    
                                }
                            }
                        }
                    }
                    conn.Close();                  
                    return modelName;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaxinDb_1ModelCode", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }
        

        //public List<Helper_TaixinDB_SampleBox> Read_TaxinDb_SamplePrint(string path_sql,string typeSample)
        //{
        //    List<Helper_TaixinDB_SampleBox> listPrint = new List<Helper_TaixinDB_SampleBox>();
        //    using (SqlConnection conn = new SqlConnection(path_sql))
        //    {
        //        conn.Open();                
        //        string commandPrint = "Select * From tbSamplePrint where imsempcode ='"+MainWindow.UserLogin+"'";
        //        using (SqlCommand cmd1 = new SqlCommand(commandPrint, conn))
        //        {
        //            using (SqlDataReader dr = cmd1.ExecuteReader())
        //            {
        //                while (dr.Read())
        //                {
        //                    Helper_TaixinDB_SampleBox sample = new Helper_TaixinDB_SampleBox();
        //                    if(typeSample=="Box")
        //                    {
        //                        if (dr[0] != null && dr[1].ToString() == "Box")
        //                        {
        //                            sample.samno = dr[0].ToString();
        //                        }
        //                    }
        //                    else if(typeSample =="Manual")
        //                    {
        //                        if (dr[0] != null && dr[1].ToString() == "Manual")
        //                        {
        //                            sample.samno = dr[0].ToString();
        //                        }
        //                    }                    
        //                    listPrint.Add(sample);
        //                }
        //            }
        //        }
        //        return listPrint;
        //    }
        //}

        public List<Helper_TaixinDB_SampleBox> Read_TaxinDb_SampleBox(string path_sql, string command)
        {
            try
            {
                
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();                       
                    List<Helper_TaixinDB_SampleBox> list_sample = new List<Helper_TaixinDB_SampleBox>();
                    //List<Helper_TaixinDB_SampleBox> list_print = Read_TaxinDb_SamplePrint(path_sql,"Box");
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                Helper_TaixinDB_SampleBox sample = new Helper_TaixinDB_SampleBox();
                                if (dr[0] != null)
                                {
                                    sample.cmpcode = dr[0].ToString();
                                    sample.bizdiv = dr[1].ToString();
                                    sample.samno = dr[2].ToString();
                                    sample.modelcode = dr[3].ToString();
                                    sample.modelname = dr[4].ToString();
                                    sample.applydt = dr[5].ToString();
                                    sample.version = dr[6].ToString();
                                    sample.custpartcode = dr[7].ToString();
                                    sample.custpart_version = dr[8].ToString();
                                    sample.custpartcode_ver = dr[9].ToString();
                                    sample.custmodelcode = dr[10].ToString();
                                    sample.repmodelcode = dr[11].ToString();
                                    sample.useflag = dr[12].ToString();
                                    sample.modelgroup = dr[13].ToString();
                                    sample.modeldiv = dr[14].ToString();
                                    sample.modeltype = dr[15].ToString();
                                    sample.modelchild = dr[16].ToString();
                                    sample.cust_gb = dr[17].ToString();
                                    sample.information = dr[18].ToString();
                                    sample.paper_name1 = dr[19].ToString();
                                    sample.paper_size1 = dr[20].ToString();
                                    sample.paper_scale1 = dr[21].ToString();
                                    sample.paper_name2 = dr[22].ToString();
                                    sample.paper_size2 = dr[23].ToString();
                                    sample.paper_scale2 = dr[24].ToString();
                                    sample.paper_name3 = dr[25].ToString();
                                    sample.paper_size3 = dr[26].ToString();
                                    sample.paper_scale3 = dr[27].ToString();
                                    sample.cover_up_name1 = dr[28].ToString();
                                    sample.cover_up_size1 = dr[29].ToString();
                                    sample.cover_up_scale1 = dr[30].ToString();
                                    sample.cover_up_name2 = dr[31].ToString();
                                    sample.cover_up_size2 = dr[32].ToString();
                                    sample.cover_up_scale2 = dr[33].ToString();
                                    sample.cover_up_name3 = dr[34].ToString();
                                    sample.cover_up_size3 = dr[35].ToString();
                                    sample.cover_up_scale3 = dr[36].ToString();
                                    sample.cover_up_name4 = dr[37].ToString();
                                    sample.cover_up_size4 = dr[38].ToString();
                                    sample.cover_up_scale4 = dr[39].ToString();
                                    sample.cover_up_name5 = dr[40].ToString();
                                    sample.cover_up_size5 = dr[41].ToString();
                                    sample.cover_up_scale5 = dr[42].ToString();
                                    sample.fullsize = dr[43].ToString();
                                    sample.print_color1 = dr[44].ToString();
                                    sample.print_color2 = dr[45].ToString();
                                    sample.print_color3 = dr[46].ToString();
                                    sample.print_color4 = dr[47].ToString();
                                    sample.print_color5 = dr[48].ToString();
                                    sample.coating = dr[49].ToString();
                                    sample.glossy_color = dr[50].ToString();
                                    sample.glossy_detail = dr[51].ToString();
                                    sample.holo_color1 = dr[52].ToString();
                                    sample.holo_detail1 = dr[53].ToString();
                                    sample.holo_color2 = dr[54].ToString();
                                    sample.holo_detail2 = dr[55].ToString();
                                    sample.holo_color3 = dr[56].ToString();
                                    sample.holo_detail3 = dr[57].ToString();
                                    sample.debosing_detail = dr[58].ToString();
                                    sample.imbosing_detail = dr[59].ToString();
                                    sample.boi_detaiil = dr[60].ToString();
                                    sample.boi_color = dr[61].ToString();
                                    sample.qty = dr[62].ToString();
                                    sample.forecast = dr[63].ToString();
                                    sample.remark1 = dr[64].ToString();
                                    sample.remark2 = dr[65].ToString();
                                    sample.dep_mar = dr[66].ToString();
                                    sample.dep_pro = dr[67].ToString();
                                    sample.dep_qc = dr[68].ToString();
                                    sample.dep_rnd = dr[69].ToString();
                                    sample.dep_pur = dr[70].ToString();
                                    sample.paperindex1 = dr[71].ToString();
                                    sample.paperindex2 = dr[72].ToString();
                                    sample.paperindex3 = dr[73].ToString();
                                    sample.printindex1 = dr[74].ToString();
                                    sample.printindex2 = dr[75].ToString();
                                    sample.printindex3 = dr[76].ToString();
                                    sample.reject = dr[77].ToString();
                                    sample.qtyAttach = dr[78].ToString();
                                    sample.depCreate = dr[79].ToString();
                                    sample.printed = dr[80].ToString();                                   
                                    sample.etc2 = dr[81].ToString();
                                    sample.etc3 = dr[82].ToString();
                                    sample.etc4 = dr[83].ToString();
                                    sample.etc5 = dr[84].ToString();
                                    sample.etc6 = dr[85].ToString();
                                    sample.etc7 = dr[86].ToString();
                                    sample.etc8 = dr[87].ToString();
                                    sample.etc9 = dr[88].ToString();
                                    sample.etc10 = dr[89].ToString();
                                    sample.sadt = dr[90].ToString();
                                    sample.fadt = dr[91].ToString();
                                    sample.imsempcode = dr[92].ToString();
                                    sample.insdt = dr[93].ToString();
                                    sample.updempcode = dr[94].ToString();
                                    sample.upddt = dr[95].ToString();
                                    if (sample.dep_mar == "Y")
                                    {
                                        sample.dep_mar = "DodgerBlue";
                                    }
                                    else if (sample.dep_mar == "N")
                                    {
                                        sample.dep_mar = "Red";
                                    }
                                    else
                                    {
                                        sample.dep_mar = "LightGray";
                                    }
                                    //7
                                    if (sample.dep_pro == "Y")
                                    {
                                        sample.dep_pro = "DodgerBlue";
                                    }
                                    else if (sample.dep_pro == "N")
                                    {
                                        sample.dep_pro = "Red";
                                    }
                                    else
                                    {
                                        sample.dep_pro = "LightGray";
                                    }
                                    //8
                                    if (sample.dep_qc == "Y")
                                    {
                                        sample.dep_qc = "DodgerBlue";
                                    }
                                    else if (sample.dep_qc == "N")
                                    {
                                        sample.dep_qc = "Red";
                                    }
                                    else
                                    {
                                        sample.dep_qc = "LightGray";
                                    }
                                    //9
                                    if (sample.dep_rnd == "Y")
                                    {
                                        sample.dep_rnd = "DodgerBlue";
                                    }
                                    else if (sample.dep_rnd == "N")
                                    {
                                        sample.dep_rnd = "Red";
                                    }
                                    else
                                    {
                                        sample.dep_rnd = "LightGray";
                                    }
                                    //10
                                    if (sample.dep_pur == "Y")
                                    {
                                        sample.dep_pur = "DodgerBlue";
                                    }
                                    else if (sample.dep_pur == "N")
                                    {
                                        sample.dep_pur = "Red";
                                    }
                                    else
                                    {
                                        sample.dep_pur = "LightGray";
                                    }                                    

                                }
                                list_sample.Add(sample);
                            }
                        }
                    }
                    conn.Close();
                    return list_sample;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaxinDb_SampleBox", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<Helper_TaixinDB_DO> Read_TaxinDb_DO(string path_sql, string nameTable, bool search, string AreaOutput, string outputGb,
            string date_start, string date_finish)//,string posName,string plInputGb,string uintName)//, 
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    List<Helper_TaixinDB_DO> list_TaixinDb = new List<Helper_TaixinDB_DO>();
                    string command = "";
                    if (search == true)
                    {
                        command = "SELECT * from " + nameTable + " WHERE  Delivery_loc = '" + AreaOutput + "' and  insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";
                    }
                    else
                    {
                        command = "SELECT * from " + nameTable + "  WHERE insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";
                    }

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {                            
                            try
                            {
                                while (dr.Read())
                                {
                                    Helper_TaixinDB_DO taixinDB = new Helper_TaixinDB_DO();

                                    if (dr[0] != null)
                                    {
                                        var formatDate = new JsonSerializerSettings { DateFormatString = "yyyy/MM/dd HH:mm:ss" };
                                        var json = JsonConvert.SerializeObject(dr[27], formatDate);                                       
                                        taixinDB.CmpCode        = dr[0].ToString();
                                        taixinDB.Bizdiv         = dr[1].ToString();
                                        taixinDB.DONOtx         = dr[2].ToString();
                                        taixinDB.PONOtx         = dr[4].ToString();
                                        taixinDB.DONOCust       = dr[6].ToString();
                                        taixinDB.PONOCust       = dr[8].ToString();
                                        taixinDB.Custpartcode   = dr[10].ToString();
                                        string date  = dr[11].ToString();  
                                        if(date.Length>6)
                                        {
                                            taixinDB.Delivery_date = date.Substring(0, 4) + "/" + date.Substring(4, 2) + "/" + date.Substring(6, 2);
                                        }
                                        if (dr[12].ToString().Length>6)
                                        {
                                            taixinDB.Delivery_time = dr[12].ToString().Substring(0, 5);
                                        }                                        
                                        taixinDB.Delivery_qty   = dr[13].ToString();
                                        taixinDB.Delivery_loc   = dr[16].ToString();
                                        taixinDB.DOdt           = dr[17].ToString();
                                        taixinDB.Expriedt       = dr[18].ToString();
                                        taixinDB.Delivdt_ss     = dr[19].ToString();
                                        taixinDB.Delivtime_ss   = dr[20].ToString();
                                        taixinDB.Insempcode     = dr[26].ToString();
                                        taixinDB.insdt          = dr[27].ToString();
                                        taixinDB.UpdempCode     = dr[28].ToString();
                                        taixinDB.Upddt          = dr[29].ToString();
                                        taixinDB.QtyOutput = "";
                                        taixinDB.CheckTotalInput = false;                                        
                                        list_TaixinDb.Add(taixinDB);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                    }
                    conn.Close();                   
                    return list_TaixinDb;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaxinDb_DO", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<Helper_TaixinDB_InputNVL> Read_TaxinDb_InputNVL(string path_sql, string nameTable, string _vendor,
            string date_start, string date_finish)//,string posName,string plInputGb,string uintName)//, 
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    List<Helper_TaixinDB_InputNVL> list_TaixinDb = new List<Helper_TaixinDB_InputNVL>();
                    string command = "";

                    //command = "SELECT * from " + nameTable + " WHERE  Delivery_loc = '" + AreaOutput + "' and  insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";

                    command = "select b.cmpcode ,b.bizdiv,b.ppono,b.pposeq,b.itemcode,b.itemtypecode,b.maker," +
                        "b.makername,b.ppoqty,b.unit,b.purweightqty,b.purunit,b.ConvNum,b.currencycode," +
                        "b.exchangecurrency,b.exchangerate,b.pricetype,b.unitprice,b.amt,b.vat," +
                        "b.vatrate,b.unitprice_Loc,b.amt_Loc,b.vat_Loc,b.supplydt,b.requestno," +
                        "b.requestseq,b.insempcode,b.insdt,b.Updempcode,b.Upddt,c.Tong,d.custnm_loc,e.itemfullname,f.itemfullname,g.Name_loc,a.vendor " +
                        "from tpPo_Head a left outer join tpPo_Detail b  " +
                        "on a.ppono = b.ppono " +
                        "left outer join (select ppono,pposeq,sum(intqty) as Tong from tqiint_Detail group by ppono,pposeq) c " +
                        "on b.ppono = c.ppono and b.pposeq = c.pposeq " +
                        "left outer join (select* from tmstcust) d " +
                        "on a.vendor = d.custcode " +
                        "left outer join (select * from tmstitemA) e "+
                        "on b.itemcode = e.itemCode and b.itemtypecode = 'M' "+
                        "left outer join "+
                        "(select * from tmstitemB) f "+
                        "on b.itemcode = f.itemCode and b.itemtypecode <> 'M' "+
                        "left outer join (select * from temstetc) g on b.unit = g.Code "+
                        "where b.ppoqty <> c.Tong and  (a.vendor like '%" + _vendor + "%' or d.custnm_loc like '%"+_vendor+"%') " +
                        "and  a.insdt >= '" + date_start + "' and a.insdt <= DATEADD(hour,1,'" + date_finish + "') and g.GubunCode='017' ORDER BY b.ppono DESC";


                    //if (search == true)
                    //{
                    //    command = "SELECT * from " + nameTable + " WHERE  Delivery_loc = '" + AreaOutput + "' and  insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";
                    //}
                    //else
                    //{
                    //    command = "SELECT * from " + nameTable + "  WHERE insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";
                    //}

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            try
                            {
                                while (dr.Read())
                                {
                                    Helper_TaixinDB_InputNVL taixinDB = new Helper_TaixinDB_InputNVL();

                                    if (dr[0] != null)
                                    {
                                        var formatDate = new JsonSerializerSettings { DateFormatString = "yyyy/MM/dd HH:mm:ss" };
                                        var json = JsonConvert.SerializeObject(dr[27], formatDate);
                                        taixinDB.cmpcode = dr[0].ToString();
                                        taixinDB.bizdiv = dr[1].ToString();
                                        taixinDB.ppono = dr[2].ToString();
                                        taixinDB.pposeq = dr[3].ToString();
                                        taixinDB.itemcode = dr[4].ToString();
                                        taixinDB.itemtypecode = dr[5].ToString();
                                        taixinDB.maker = dr[6].ToString();
                                        taixinDB.makername = dr[7].ToString();
                                        taixinDB.ppoqty = double.Parse(dr[8].ToString()).ToString();
                                        taixinDB.unit = dr[9].ToString();
                                        taixinDB.purweightqty = dr[10].ToString();
                                        taixinDB.purunit = dr[11].ToString();
                                        taixinDB.ConvNum = dr[12].ToString();
                                        taixinDB.currencycode = dr[13].ToString();
                                        taixinDB.exchangecurrency = dr[14].ToString();
                                        taixinDB.exchangerate = dr[15].ToString();
                                        taixinDB.pricetype = dr[16].ToString();
                                        taixinDB.unitprice = dr[17].ToString();
                                        taixinDB.amt = dr[18].ToString();
                                        taixinDB.vat = dr[19].ToString();
                                        taixinDB.vatrate = dr[20].ToString();
                                        taixinDB.unitprice_Loc = dr[21].ToString();
                                        taixinDB.amt_Loc = dr[22].ToString();
                                        taixinDB.vat_Loc = dr[23].ToString();
                                        taixinDB.supplydt = dr[24].ToString();
                                        taixinDB.requestno = dr[25].ToString();
                                        taixinDB.requestseq = dr[26].ToString();
                                        taixinDB.insempcode = dr[27].ToString();
                                        taixinDB.insdt = dr[28].ToString();
                                        taixinDB.Updempcode = dr[29].ToString();
                                        taixinDB.Upddt = dr[30].ToString();                                       
                                        taixinDB.RememberQty = double.Parse(dr[31].ToString()).ToString();
                                        taixinDB.SubQty = (double.Parse(dr[8].ToString()) - double.Parse(dr[31].ToString())).ToString();
                                        taixinDB.NameCustomer = dr[32].ToString();
                                        taixinDB.NameNVL = dr[33].ToString() + dr[34].ToString();
                                        taixinDB.NameUnit = dr[35].ToString();
                                        taixinDB.VendorCode = dr[36].ToString();
                                        list_TaixinDb.Add(taixinDB);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                    }
                    conn.Close();
                    return list_TaixinDb;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaxinDb_InputNVL", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<Helper_TaixinDB_InputNVL> Read_TaxinDb_OutputNVL(string path_sql, string nameTable, string _vendor,
           string date_start, string date_finish)//,string posName,string plInputGb,string uintName)//, 
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    List<Helper_TaixinDB_InputNVL> list_TaixinDb = new List<Helper_TaixinDB_InputNVL>();
                    string command = "";

                    //command = "SELECT * from " + nameTable + " WHERE  Delivery_loc = '" + AreaOutput + "' and  insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";

                    //command = "select b.cmpcode ,b.bizdiv,b.ppono,b.pposeq,b.itemcode,b.itemtypecode,b.maker," 

                    command = "SELECT b.cmpcode,	b.bizdiv,	b.intno,	b.intseq,	b.itemcode,	b.itemname, " +
                            "b.itemtypecode,	b.vendor,	b.maker,	b.makername,	b.unitprice, " +
                            "b.unitprice_Loc,	b.currencycode,	b.unit,	b.purunit,	b.intdt, " +
                            "b.intqty,	b.intqtyr,	b.exrate,	b.amt,	b.amt_Loc,	b.vat,	b.vat_Loc, " +
                            "b.ppono,	b.pposeq,	b.insempcode,	b.insdt,b.updempcode,	b.upddt, " +
                            "a.vendor,g.Name_loc,d.custnm_loc,e.itemfullname,f.itemfullname,h.TongOutput " +
                            "from TQIINT_DETAIL b left outer join TPPO_Head a " +
                            "on a.ppono = b.ppono " +
                            "left outer join(select intno, intseq, sum(outqty) as TongOutput from tqoout_Detail group by intno, intseq) h " +
                            "on h.intno = b.intno and h.intseq = b.intseq " +
                            "left outer join(select * from temstetc) g " +
                            "on b.unit = g.Code " +
                            "left outer join(select * from tmstcust) d " +
                            "on a.vendor = d.custcode " +
                            "left outer join(select * from tmstitemA) e " +
                            "on b.itemcode = e.itemCode and b.itemtypecode = 'M' " +
                            "left outer join(select * from tmstitemB) f on b.itemcode = f.itemCode and b.itemtypecode <> 'M' " +
                            "where(a.vendor like '%" + _vendor + "%' or d.custnm_loc like '%" + _vendor + "%') and(b.intqty <> h.TongOutput or h.TongOutput is null) " +
                            "and  a.insdt >= '" + date_start + "' and a.insdt <= DATEADD(hour,1,'" + date_finish + "') and g.GubunCode='017' ORDER BY b.ppono DESC ";


                    //if (search == true)
                    //{
                    //    command = "SELECT * from " + nameTable + " WHERE  Delivery_loc = '" + AreaOutput + "' and  insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";
                    //}
                    //else
                    //{
                    //    command = "SELECT * from " + nameTable + "  WHERE insdt >= '" + date_start + "' and insdt <= DATEADD(hour,1,'" + date_finish + "')   ORDER BY insdt DESC";
                    //}

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            try
                            {
                                while (dr.Read())
                                {
                                    Helper_TaixinDB_InputNVL taixinDB = new Helper_TaixinDB_InputNVL();

                                    if (dr[0] != null)
                                    {
                                        var formatDate = new JsonSerializerSettings { DateFormatString = "yyyy/MM/dd HH:mm:ss" };
                                        var json = JsonConvert.SerializeObject(dr[27], formatDate);
                                        taixinDB.cmpcode = dr[0].ToString();
                                        taixinDB.bizdiv = dr[1].ToString();
                                        taixinDB.intno = dr[2].ToString();
                                        taixinDB.intseq = dr[3].ToString();
                                        taixinDB.itemcode = dr[4].ToString();
                                        taixinDB.NameNVL = dr[5].ToString();
                                        taixinDB.itemtypecode = dr[6].ToString();
                                        taixinDB.VendorCode = dr[7].ToString();
                                        taixinDB.maker = dr[8].ToString();
                                        taixinDB.makername = dr[9].ToString();
                                        taixinDB.unitprice = dr[10].ToString();
                                        taixinDB.unitprice_Loc = dr[11].ToString();
                                        taixinDB.currencycode = dr[12].ToString();
                                        taixinDB.unit = dr[13].ToString();
                                        taixinDB.purunit = dr[14].ToString();
                                        taixinDB.intdt = dr[15].ToString();
                                        taixinDB.intqty = double.Parse(dr[16].ToString()).ToString();
                                        taixinDB.purweightqty = dr[17].ToString();
                                        taixinDB.exchangerate = dr[18].ToString();
                                        taixinDB.amt = dr[19].ToString();
                                        taixinDB.amt_Loc = dr[20].ToString();
                                        taixinDB.vat = dr[21].ToString();
                                        taixinDB.vat_Loc = dr[22].ToString();
                                        taixinDB.ppono = dr[23].ToString();
                                        taixinDB.pposeq = dr[24].ToString();
                                        taixinDB.insempcode = dr[25].ToString();
                                        taixinDB.insdt = dr[26].ToString();
                                        taixinDB.Updempcode = dr[27].ToString();
                                        taixinDB.Upddt = dr[28].ToString();
                                        taixinDB.VendorCode = dr[29].ToString();
                                        taixinDB.NameUnit = dr[30].ToString();
                                        taixinDB.NameCustomer = dr[31].ToString();
                                        taixinDB.NameNVL = dr[32].ToString() + dr[33].ToString();
                                        if (dr[34].ToString()=="")
                                        {
                                            taixinDB.RememberQty = "0";
                                        }
                                        else
                                        {
                                            taixinDB.RememberQty = double.Parse(dr[34].ToString()).ToString();
                                        }
                                        taixinDB.SubQty =(double.Parse(dr[16].ToString())- double.Parse(taixinDB.RememberQty)).ToString();
                                        list_TaixinDb.Add(taixinDB);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                    }
                    conn.Close();
                    return list_TaixinDb;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,"Helper DataBase/Read_TaxinDb_OutputNVL", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<DataFriend> ReadMessageFriend(string path_sql)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    int count = 0;
                    List<DataFriend> list = new List<DataFriend>();
                    var command = "SELECT * from DataFriend";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                DataFriend item = new DataFriend();
                                if (dr[0] != null)
                                {
                                    item.ID = int.Parse(dr[0].ToString());
                                    item.TimeBegin = dr[1].ToString();
                                    item.AdminSend = dr[2].ToString();
                                    item.GuestSend = dr[3].ToString();
                                    item.MessFriend = dr[4].ToString();
                                    item.FileFriend = dr[5].ToString();
                                    item.AdminIcon = dr[6].ToString();
                                    item.GuestIcon = dr[7].ToString();
                                    item.TimeEnd = dr[8].ToString();
                                    list.Add(item);
                                    count++;
                                }
                            }
                        }
                    }
                    conn.Close();
                    return list;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/ReadMessageFriend", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<DataFriend> ReadNewMessage(string path_sql)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    int count = 0;
                    List<DataFriend> list = new List<DataFriend>();
                    var command = "SELECT TOP(1)* from DataFriend ORDER BY TimeBegin desc ";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                DataFriend item = new DataFriend();
                                if (dr[0] != null)
                                {
                                    item.ID = int.Parse(dr[0].ToString());
                                    item.TimeBegin = dr[1].ToString();
                                    item.AdminSend = dr[2].ToString();
                                    item.GuestSend = dr[3].ToString();
                                    item.MessFriend = dr[4].ToString();
                                    item.FileFriend = dr[5].ToString();
                                    item.AdminIcon = dr[6].ToString();
                                    item.GuestIcon = dr[7].ToString();
                                    item.TimeEnd = dr[8].ToString();
                                    list.Add(item);
                                    count++;
                                }
                            }
                        }
                    }
                    conn.Close();
                    return list;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/ReadNewMessage", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<LoginData> ReadTable(string path_sql)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    int count = 0;
                    List<LoginData> list = new List<LoginData>();
                    var command = "SELECT * from Logins";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                LoginData item = new LoginData();
                                if (dr[0] != null)
                                {
                                    item.ID = int.Parse(dr[0].ToString());
                                    item.User = dr[1].ToString();
                                    item.Pass = dr[2].ToString();
                                    item.MacAddres = dr[3].ToString();
                                    item.Remember = dr[4].ToString();
                                    item.Version = dr[5].ToString();
                                    list.Add(item);
                                    count++;
                                }
                            }
                        }
                    }
                    conn.Close();
                    return list;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/ReadTable", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public string Read_TaixinIdNo(string path_sql,string sortColumnName,string nameTable,string whereName,string whereValue)
        {
            try
            {
                string ID_No= "";
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    //var command = "SELECT TOP(1) * from " + nameTable + " WHERE INTDT = '" + Date + "' ORDER BY INSDT";
                    var command = "SELECT MAX("+sortColumnName+") from " + nameTable + " WHERE "+ whereName + " = '"+ whereValue + "'";                   
                    //var command = "SELECT MAX(intno) FROM twhint WHERE intdt = '20200204'";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        ID_No = cmd.ExecuteScalar().ToString();
                    }
                    conn.Close();
                    return ID_No;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaixinIdNo", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public string Read_TaixinMaxSeq(string path_sql, string sortColumnName, string nameTable, string whgubun, string pstcode,string date,string dateValue)
        {
            try
            {
                string ID_No = "";
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    //var command = "SELECT TOP(1) * from " + nameTable + " WHERE INTDT = '" + Date + "' ORDER BY INSDT";
                    var command = "SELECT MAX(" + sortColumnName + ") from " + nameTable + " WHERE whgubun ='" + whgubun + "' and pstcode = '" + pstcode + "' and "+date+" ='" + dateValue + "'";
                    //var command = "SELECT MAX(intno) FROM twhint WHERE intdt = '20200204'";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        ID_No = cmd.ExecuteScalar().ToString();
                    }
                    conn.Close();
                    return ID_No;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaixinMaxSeq", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        //Thêm dữ liệu vào bảng dữ liệu=================================================================================
        public void InsertTable(string path_sql, string nameTable, int numberInsert, string[] listInsert, object[] valueInsert)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    string nameInsert = "";
                    string value = "";
                    for (int i = 0; i < numberInsert; i++)
                    {
                        if(i<numberInsert-1)
                        {
                            nameInsert += listInsert[i] + ",";
                            value += "@" + listInsert[i].ToLower()+",";
                        }
                        if(i== numberInsert-1)
                        {
                            nameInsert += listInsert[i];
                            value += "@" + listInsert[i].ToLower();
                        }
                      
                    }
                    var command = "INSERT INTO " + nameTable + "(" + nameInsert + ") Values(" + value + ")";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        for (int i = 0; i < numberInsert; i++)
                        {
                            if (i < numberInsert - 1)
                            {
                                cmd.Parameters.Add(new SqlParameter("@" + listInsert[i].ToLower(), valueInsert[i]));                                
                            }
                            if (i == numberInsert - 1)
                            {
                                cmd.Parameters.Add(new SqlParameter("@" + listInsert[i].ToLower(), valueInsert[i]));
                            }
                        }
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/InsertTable", MessageBoxButton.OK, MessageBoxImage.Error);
              
            }
        }

        //Thêm xóa bảng dữ liệu=========================================================================================
        public void DeleteTable(string path_sql, string nameTable)
        {
            try
            {
                if (MessageBoxResult.Yes == MessageBox.Show("You want DELETE ALL DATA?", "DELETE DATA", MessageBoxButton.YesNo, MessageBoxImage.Asterisk))
                {
                    using (SqlConnection conn = new SqlConnection(path_sql))
                    {
                        conn.Open();
                        var command = "DROP TABLE " + nameTable + "";
                        using (SqlCommand cmd = new SqlCommand(command, conn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                        conn.Close();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/DeleteTable", MessageBoxButton.OK, MessageBoxImage.Error);
               
            }
        }

        //Nhớ Tên và mật khẩu đăng nhập=================================================================================
        public void RememberLogin(string path_sql, string nameTalbe, string mac, string remember)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();

                    var command = "UPDATE "+nameTalbe+" SET Remember = '" + remember + "'  where MacAddres ='" + mac + "'";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/RememberLogin", MessageBoxButton.OK, MessageBoxImage.Error);
               
            }
        }

        //Sửa dữ liệu DataBase==========================================================================================
        public void EditLogin(string path_sql,string nameTable, string user, string pass,string mac, string value)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var command = "UPDATE "+nameTable+" SET UserLogin = '" + user + "', PassLogin = '" + pass + "', MacAddres = '" + mac + "' where ID = " + value + "";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/EditLogin", MessageBoxButton.OK, MessageBoxImage.Error);
               
            }
        }

        public LoginData CheckRemember(string path_sql)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var command = "SELECT * from Logins";
                    LoginData item = new LoginData();
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {

                                if (dr[0] != null)
                                {
                                    //if (dr[3].ToString() == CreatUser.macAddr)
                                    //{
                                    //    item.User = dr[1].ToString();
                                    //    item.Pass = dr[2].ToString();
                                    //    item.Remember = dr[4].ToString();
                                    //}
                                }
                            }
                        }
                    }
                    conn.Close();
                    return item;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/CheckRemember", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public string Login_Check(string path_sql,string nameTable, string user, string pass,string ItemCheck)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    string strcheck = "";
                    int count = 0;
                    var command = "SELECT * from "+nameTable+"";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                if(ItemCheck == "Login")
                                {
                                    count++;
                                    if (user == "it" && pass == "it")
                                    {
                                        count = 0;
                                        strcheck = "admin";
                                        break;
                                    }
                                    if (dr[1].ToString() == user && dr[2].ToString() == pass)
                                    {
                                        count = 0;
                                        strcheck = "user";
                                        break;
                                    }
                                }                                

                                if (ItemCheck == "Ketban")
                                {
                                    count++;
                                    if (dr[1].ToString() == user)
                                    {
                                        count = 0;
                                        strcheck = "Tontai";
                                        break;
                                    }
                                }

                            }
                            if (count > 0 && ItemCheck == "Login")
                            {
                                if (MessageBoxResult.Yes == MessageBox.Show("The USER or PASSWORD is incorrect, Please input again", "User Login", MessageBoxButton.YesNo, MessageBoxImage.Error))
                                {
                                    return "Failse";
                                }
                                else
                                {
                                    Application.Current.Shutdown();
                                }
                            }
                            if (count > 0 && ItemCheck == "Ketban")
                            {
                                MessageBox.Show("Không tồn tại người bạn này trên hệ thống", "Kết bạn", MessageBoxButton.YesNo, MessageBoxImage.Error);
                            }
                        }
                    }
                    conn.Close();
                    return strcheck;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Login_Check", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        //Xóa từng dòng dữ liệu========================================================================================
        public void DeleteRow(string path_sql,string nameTable, string nameRow, string value)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();

                    var command = "DELETE from "+nameTable+" where " + nameRow + " = '" + value + "'";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/DeleteRow", MessageBoxButton.OK, MessageBoxImage.Error);                
            }
        }

        //Kiểm tra Framework nếu chưa có thì thông báo update==========================================================
        string VerFramework = "";
        public void ProcessVersionFramework()
        {
            using (RegistryKey ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey("SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full\\"))
            {
                int releaseKey = Convert.ToInt32(ndpKey.GetValue("Release"));
                if (true)
                {
                    VerFramework = CheckVersionFramework(releaseKey);
                    if (VerFramework == "OLD")
                    {
                        MessageBox.Show("Update Framework up to Version : 4.61");
                        //Process.Start(@".\Framework4.61.exe");
                    }
                }
            }
        }
        public string CheckVersionFramework(int releaseKey)
        {
            string ver = "";
            if (releaseKey > 394254)//4.6.1
            {
                ver = "NEW";
            }
            else
            {
                ver = "OLD";
            }
            return ver;
        }

        public string Rejetc_MaxSeq(string path_sql, string nameTable, string samNo, string typeSample)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();

                    var command = "SELECT max(seq) FROM " + nameTable + " Where samno='" + samNo + "' and typeSample='" + typeSample + "'";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        string seq = cmd.ExecuteScalar().ToString();
                        if (seq != "")
                        {
                            seq = (int.Parse(seq) + 1).ToString("0000");
                        }
                        else
                        {
                            seq = ("0001");
                        }
                        conn.Close();
                        return seq;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Rejetc_MaxSeq", MessageBoxButton.OK, MessageBoxImage.Error);
                return "null";
            }
        }



    }
}
