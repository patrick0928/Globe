
using Globe_Script.Helper;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace Globe_Script.ViewModel
{
    public class ViewModelMain : ViewModelBase
    {
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("User32")]
        private static extern int ShowWindow(int hwnd, int nCmdShow);
        private const int SW_MAXIMIZE = 3;

        #region #global variables

        ServiceController serviceController = new ServiceController();
        PnrMain pnrMain = new PnrMain();
        public RelayCommand RetreiveCommand { get; set; }
        public RelayCommand SubmitCommand { get; set; }
        public RelayCommand UndoCommand { get; set; }
        public RelayCommand SaveCommand { get; set; }
        string action;
        string _reLoc;
        string _name;
        string _nameText;
        string _itenerary;
        string _remarks;
        string _remarksText;
        string _cost;
        string _workID;
        string _trip;
        string _rqID;
        string _agentName;
        string _recvFrom;
        Visibility _isVisible;
        bool _isEnable;
        bool _isEditMode;
        int index = 0, poIndex = 0, inv = 0;
        int workIndex = 0;
        #endregion
        #region #properties
        public string ReLoc
        {
            get { return _reLoc; }
            set
            {
                if (_reLoc != value)
                {
                    _reLoc = value;
                    OnPropertyChange("ReLoc");
                }
            }
        }

        public string Name
        {
            get { return _name; }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChange("Name");
                }
            }
        }

        public string NameText
        {
            get { return _nameText; }
            set
            {
                if (_nameText != value)
                {
                    _nameText = value;
                    OnPropertyChange("NameText");
                }
            }
        }

        public string Itenerary
        {
            get { return _itenerary; }
            set
            {
                if (_itenerary != value)
                {
                    _itenerary = value;
                    OnPropertyChange("Itenerary");
                }
            }
        }
        public string Remarks
        {
            get { return _remarks; }
            set
            {
                if (_remarks != value)
                {
                    _remarks = value;
                    OnPropertyChange("Remarks");
                }
            }
        }

        public string RemarksText
        {
            get { return _remarksText; }
            set
            {
                if (_remarksText != value)
                {
                    _remarksText = value;
                    OnPropertyChange("RemarksText");
                }
            }
        }

        public string Cost
        {
            get { return _cost; }
            set
            {
                if (_cost != value)
                {
                    _cost = value;
                    OnPropertyChange("Cost");
                }
            }
        }

        public string WorkID
        {
            get { return _workID; }
            set
            {
                if (_workID != value)
                {
                    _workID = value;
                    OnPropertyChange("WorkID");
                }
            }
        }

        public string Trip
        {
            get { return _trip; }
            set
            {
                if (_trip != value)
                {
                    _trip = value;
                    OnPropertyChange("Trip");
                }
            }
        }

        public string RqID
        {
            get { return _rqID; }
            set
            {
                if (_rqID != value)
                {
                    _rqID = value;
                    OnPropertyChange("RqID");
                }
            }
        }

        public string AgentName
        {
            get { return _agentName; }
            set
            {
                if (_agentName != value)
                {
                    _agentName = value;
                    OnPropertyChange("AgentName");
                }
            }
        }

        public string ReciveFrom
        {
            get { return _recvFrom; }
            set
            {
                if (_recvFrom != value)
                {
                    _recvFrom = value;
                    OnPropertyChange("ReciveFrom");
                }
            }
        }

        public bool IsEnable
        {
            get { return _isEnable; }
            set
            {
                if (_isEnable != value)
                {
                    _isEnable = value;
                    OnPropertyChange("IsEnable");
                }
            }
        }

        public Visibility IsVisible
        {
            get { return _isVisible; }
            set
            {
                _isVisible = value;
                OnPropertyChange("IsVisible");
            }
        }
        #endregion

        public string getRemarks(string remarks)
        {
            string result = "";
            foreach (PnrRemarks _remarks in pnrMain.Remarks)
            {
                RemarksText = "Remarks :";
                string[] split = _remarks.Text.Split('/');

                if (split[0] == remarks)
                    result = split[1];
            }

            return result;
        }

        public int getRemarks(string remarks, int index)
        {
            foreach (PnrRemarks _remarks in pnrMain.Remarks)
            {
                RemarksText = "Remarks :";
                string[] split = _remarks.Text.Split('/');

                if (split[0] == remarks)
                    break;

                index++;
            }
            return index;
        }

        public void loadData()
        {
            clearPnr();
            IsVisible = Visibility.Visible;
            try
            {
                pnrMain = serviceController.RetrievePNR(ReLoc);

                NameText = "Name:";
                foreach (PassengerNames _name in pnrMain.Passengers)
                {
                    Name += $"     {_name.GivenName} {_name.SurName} \n";
                }

                string[] recveFrom = pnrMain.ReceivedFrom.Split('/');
                ReciveFrom = $"Received From : {pnrMain.ReceivedFrom}";
                AgentName = recveFrom[0];

                try
                {
                    Remarks += string.IsNullOrWhiteSpace(getRemarks("FF34")) ? "" : $"     Cost Center  :  {getRemarks("FF34")} \n";
                    Remarks += string.IsNullOrWhiteSpace(getRemarks("FF35")) ? "" : $"     ID Work Order  :  {getRemarks("FF35")} \n";
                    Remarks += string.IsNullOrWhiteSpace(getRemarks("FF36")) ? "" : $"     Purpose of Trip  :  {getRemarks("FF36")} \n";
                    Remarks += string.IsNullOrWhiteSpace(getRemarks("FF37")) ? "" : $"     Request ID  :  {getRemarks("FF37")} \n";
                    Remarks += string.IsNullOrWhiteSpace(getRemarks("FF37")) ? "" : $"     PO  :  {getRemarks("PO")} \n";
                    Remarks += string.IsNullOrWhiteSpace(getRemarks("INV")) ? "" : $"     Invoice  :  {getRemarks("INV")} \n";
                }
                catch (Exception e) { }

                IsEnable = true;

                if (!string.IsNullOrWhiteSpace(Remarks))
                {
                    Cost = getRemarks("FF34");
                    WorkID = getRemarks("FF35");
                    Trip = getRemarks("FF36");
                    RqID = getRemarks("FF37");

                    index = getRemarks("FF34", 0);
                    workIndex = getRemarks("FF35", 0);
                    poIndex = getRemarks("PO", 0);
                    //inv = getRemarks("INV", 0);
                    _isEditMode = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No Record Found" + ex);
            }

            try
            {
                foreach (AirSegmentDetails _details in pnrMain.Segments)
                {
                    string[] time = _details.DepartureDateTime.Split('T');
                    string dateTime = time[0].Substring(5) + " " + time[1];

                    if (_details.Key == "1")
                        Itenerary = $"Itinerary:  {_details.Origin} / {_details.Destination} / {dateTime}";
                    else
                        Itenerary += $"\n \t  {_details.Origin} / {_details.Destination} / {dateTime}";
                }
            }
            catch (Exception e){ }
            IsVisible = Visibility.Collapsed;
        }

        private async void retrieve()
        {
            await Task.Run(() => loadData());
        }

        public void IgnoreRetrieve()
        {
            sendKey(" (%{BACKSPACE}) ");
            sendKey("(%{BACKSPACE})");
            sendKey("I");
            sendKey("{ENTER}");

            sendKey("=");
            sendKey(ReLoc);
            sendKey("{ENTER}");
        }

        public void SavePnr()
        {
            sendKey("E");            
            sendKey("R");
            sendKey("{ENTER}");

            clearPnr();
            IsVisible = Visibility.Collapsed;
        }

        public void sendKey(string command)
        {
            Process[] process = Process.GetProcessesByName("abacusworkspace");
            if (process.Count() > 0)
            {
                IntPtr ipHwnd = process[0].MainWindowHandle;
                bool success = SetForegroundWindow(ipHwnd);

                if (success)
                {
                    int hwnd = process[0].MainWindowHandle.ToInt32();
                    ShowWindow(hwnd, SW_MAXIMIZE);

                    System.Windows.Forms.SendKeys.SendWait(command);
                }
            }
            System.Threading.Thread.Sleep(500);
        }

        public void clearPnr()
        {
            NameText = "";
            Name = "";
            Itenerary = "";
            Remarks = "";
            RemarksText = "";

            Cost = "";
            WorkID = "";
            RqID = "";
            Trip = "";
            AgentName = "";
            ReciveFrom = "";

            IsEnable = false;
            _isEditMode = false;
            IsEnable = false;
        }

        public void submitButton()
        {
            if (_isEditMode == true)
            {
                MessageBoxResult result = MessageBox.Show("Data already exist. Do you want to overwrite ?", "Overrite Data", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                    edit();
            }
            else
                submit();             
        }
        
        #region #method submit command
        public void submit()
        {
            IsVisible = Visibility.Visible;            
            sendKey("5'");             
            sendKey("FF34/" + Cost);             
            sendKey("{ENTER}");

            if(!string.IsNullOrWhiteSpace(WorkID))
            {
                sendKey("5'");
                sendKey("FF35/" + WorkID);
                sendKey("{ENTER}");
            }
           
             
            sendKey("5'");             
            sendKey("FF36/" + Trip);             
            sendKey("{ENTER}");
             
            sendKey("5'");             
            sendKey("FF37/" + RqID);             
            sendKey("{ENTER}");
             
            sendKey("5'");             
            sendKey("PO/" + RqID);             
            sendKey("{ENTER}");
            
            sendKey("5'");             
            sendKey("INV/DEFAULT");             
            sendKey("{ENTER}");
             
            sendKey("6");             
            sendKey(AgentName);
            sendKey("{ENTER}");

            evaluateForError("submit");
        }
        #endregion

        #region #method edit command
        public void edit()
        {             
            sendKey("5");             
            sendKey($"{index + 1}");             
            sendKey("[");
            sendKey("'");
            sendKey("FF34/" + Cost);             
            sendKey("{ENTER}");


            if (!string.IsNullOrWhiteSpace(WorkID))
            {
                sendKey("5");
                if (Remarks.Contains("Work"))
                {
                    sendKey($"{workIndex + 1}");
                    sendKey("[");
                }

                sendKey("'");
                sendKey("FF35/" + WorkID);
                sendKey("{ENTER}");
            }
             
            sendKey("5");            
            sendKey($"{index + 2}");             
            sendKey("[");            
            sendKey("'");             
            sendKey("FF36/" + Trip);            
            sendKey("{ENTER}");
             
            sendKey("5");             
            sendKey($"{index + 3}");             
            sendKey("[");             
            sendKey("'");             
            sendKey("FF37/" + RqID);             
            sendKey("{ENTER}");
             
            sendKey("5");             
            sendKey($"{poIndex + 1}");             
            sendKey("[");             
            sendKey("'");             
            sendKey("PO/" + RqID);             
            sendKey("{ENTER}");
             
            sendKey("6");             
            sendKey(AgentName);             
            sendKey("{ENTER}");

            evaluateForError("edit");

        }
        #endregion

        #region #can Execute
        public bool Command_CanExecute()
        {
            if (string.IsNullOrWhiteSpace(ReLoc))
            {
                clearPnr();
                return false;
            }
            else
                return true;
        }

        public bool CommandInsert_CanExecute()
        {
            if (string.IsNullOrWhiteSpace(Cost) || string.IsNullOrWhiteSpace(Trip) || string.IsNullOrWhiteSpace(RqID) || string.IsNullOrWhiteSpace(AgentName))
                return false;
            else
                return true;
        }

        #endregion

        #region #Evaluate Input

        public void getClipboard_Text()
        {
            string clipboardText = Clipboard.GetText();
            string[] data = clipboardText.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            bool hasError = false;
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i].Contains("FORMAT"))
                    hasError = true;
                else if (data[i].Contains(".ITEM"))
                    hasError = true;
                else if (data[i].Contains("ITEM."))
                    hasError = true;
                else if (data[i].Contains("ITEM"))
                    hasError = true;
                else if (data[i].Contains("SUCH"))
                    hasError = true;
                else if (data[i].Contains(".NO"))
                    hasError = true;
                else if (data[i].Contains("BGNG"))
                    hasError = true;
                else if (data[i].Contains("CHECK"))
                    hasError = true;
                else if (data[i].Contains("ENTRY"))
                    hasError = true;
                else if (data[i].Contains("INVALID"))
                    hasError = true;
                else if (data[i].Contains(".FRMT"))
                    hasError = true;
                else if (data[i].Contains("FRMT"))
                    hasError = true;
                else if (data[i].Contains("FRMT."))
                    hasError = true;
                else if (data[i].Contains(".NOT"))
                    hasError = true;
                else if (data[i].Contains("CODE"))
                    hasError = true;
                else if (data[i].Contains("ACTION"))
                    hasError = true;
                else if (data[i].Contains("‡CARRIER"))
                    hasError = true;
                else if (data[i].Contains("‡FORMAT"))
                    hasError = true;
                else if (data[i].Contains(".FRMT."))
                    hasError = true;
            }

            if (hasError == true)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                {
                    Process process = Process.GetProcessesByName("Globe Script").FirstOrDefault();
                    bool success = SetForegroundWindow(process.MainWindowHandle);
                    
                    MessageBoxResult result = MessageBox.Show(Application.Current.MainWindow, "Error Found ! trying to Re-input...", "Error", MessageBoxButton.YesNo, MessageBoxImage.Error);
                    if (result == MessageBoxResult.Yes)
                    {
                        IgnoreRetrieve();
                        if (action == "submit")
                            submit();
                        else
                            edit();
                    }
                    IsVisible = Visibility.Collapsed;
                }));
            }
            else
                Task.Run(() => SavePnr());
        }

        public async void evaluateForError(string action)
        {
            this.action = action;

            InputChecker checker = new InputChecker();
            await Task.Run(() => checker.RunChecker());

            await Task.Run(() =>
            {
                Task.Delay(1000).Wait();

                Thread thread = new Thread(new System.Threading.ThreadStart(getClipboard_Text));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
            });
        }
        #endregion

        public ViewModelMain()
        {
            RetreiveCommand = new RelayCommand(retrieve, Command_CanExecute);
            SubmitCommand = new RelayCommand(submitButton, CommandInsert_CanExecute);
            SaveCommand = new RelayCommand(SavePnr, CommandInsert_CanExecute);
            //ReLoc = "HDDXFX";
            ReLoc = "Record Locator";

            IsVisible = Visibility.Collapsed;
        }
    }
}
