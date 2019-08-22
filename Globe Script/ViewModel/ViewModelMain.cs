using System;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Collections.Generic;
using Globe_Script.ViewModel;
using System.Windows.Threading;

namespace Globe_Script.ViewModel
{
    public class ViewModelMain : ViewModelBase
    {
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        #region #global variables

        ServiceController serviceController = new ServiceController();
        PnrMain pnrMain = new PnrMain();
        public RelayCommand RetreiveCommand { get; set; }
        public RelayCommand SubmitCommand { get; set; }
        public RelayCommand UndoCommand { get; set; }
        public RelayCommand SaveCommand { get; set; }
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

        internal void loadingScreen()
        {
            IsVisible = Visibility.Visible;

            System.Threading.Thread.Sleep(51);
        }

        internal void load()
        {
            System.Threading.Thread.Sleep(300);
        }
        internal void loadSabreWindow()
        {
            System.Threading.Thread.Sleep(800);
        }
        internal void loadChecker()
        {
            System.Threading.Thread.Sleep(180);
        }

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

                Remarks += string.IsNullOrWhiteSpace(getRemarks("FF34")) ? "" : $"     Cost Center  :  {getRemarks("FF34")} \n";
                Remarks += string.IsNullOrWhiteSpace(getRemarks("FF35")) ? "" : $"     ID Work Order  :  {getRemarks("FF35")} \n";
                Remarks += string.IsNullOrWhiteSpace(getRemarks("FF36")) ? "" : $"     Purpose of Trip  :  {getRemarks("FF36")} \n";
                Remarks += string.IsNullOrWhiteSpace(getRemarks("FF37")) ? "" : $"     Request ID  :  {getRemarks("FF37")} \n";
                Remarks += string.IsNullOrWhiteSpace(getRemarks("FF37")) ? "" : $"     PO  :  {getRemarks("PO")} \n";
                Remarks += string.IsNullOrWhiteSpace(getRemarks("INV")) ? "" : $"     Invoice  :  {getRemarks("INV")} \n";

                IsEnable = true;
                if (!string.IsNullOrWhiteSpace(Remarks))
                {
                    Cost = getRemarks("FF34");
                    WorkID = getRemarks("FF35");
                    Trip = getRemarks("FF36");
                    RqID = getRemarks("FF37");

                    index = getRemarks("FF34", 0);
                    poIndex = getRemarks("PO", 0);
                    //inv = getRemarks("INV", 0);
                    _isEditMode = true;

                    //MessageBoxResult result = MessageBox.Show("Remarks Field already exist", "", MessageBoxButton.OK, MessageBoxImage.Information);
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
            catch (Exception e)
            {
                //MessageBox.Show("No Itenerary Found");
            }
            IsVisible = Visibility.Collapsed;
        }

        private async void retrieve()
        {
            await Task.Run(() => loadData());
        }

        public void IgnoreRetrieve()
        {
            sendKey("I");
            sendKey("{ENTER}");

            sendKey("=");
            sendKey(ReLoc);
            sendKey("{ENTER}");
        }

        public async void SavePnr()
        {
            await Task.Run(() => load());
            sendKey("E");

            await Task.Run(() => load());
            sendKey("R");

            await Task.Run(() => load());
            sendKey("{ENTER}");

            clearPnr();

            IsVisible = Visibility.Collapsed;
        }

        public async void sendKey(string command)
        {
            Process[] process = Process.GetProcessesByName("abacusworkspace");
            if (process.Count() > 0)
            {
                IntPtr ipHwnd = process[0].MainWindowHandle;
                bool success = SetForegroundWindow(ipHwnd);

                if (success)
                {
                    await Task.Run(() => load());
                    System.Windows.Forms.SendKeys.SendWait(command);
                }
                else
                    await Task.Run(() => loadSabreWindow());
            }
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
        public async void submit()
        {

            IsVisible = Visibility.Visible;

            await Task.Run(() => load());
            sendKey("5'");

            await Task.Run(() => load());
            sendKey("FF34/" + Cost);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("5'");

            await Task.Run(() => load());
            sendKey("FF35/" + WorkID);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("5'");

            await Task.Run(() => load());
            sendKey("FF36/" + Trip);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("5'");

            await Task.Run(() => load());
            sendKey("FF37/" + RqID);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("5'");

            await Task.Run(() => load());
            sendKey("PO/" + RqID);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("5'");

            await Task.Run(() => load());
            sendKey("INV/DEFAULT");

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("6");

            await Task.Run(() => load());
            sendKey(AgentName);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            evaluateForError("submit");
        }
        #endregion

        #region #method edit command
        public async void edit()
        {
            await Task.Run(() => load());
            sendKey("5");

            await Task.Run(() => load());
            sendKey($"{index + 1}");

            await Task.Run(() => load());
            sendKey("[");

            await Task.Run(() => load());
            sendKey("'");

            await Task.Run(() => load());
            sendKey("FF34/" + Cost);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("5");

            await Task.Run(() => load());
            sendKey($"{index + 2}");

            await Task.Run(() => load());
            sendKey("[");

            await Task.Run(() => load());
            sendKey("'");

            await Task.Run(() => load());
            sendKey("FF35/" + WorkID);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("5");

            await Task.Run(() => load());
            sendKey($"{index + 3}");

            await Task.Run(() => load());
            sendKey("[");

            await Task.Run(() => load());
            sendKey("'");

            await Task.Run(() => load());
            sendKey("FF36/" + Trip);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("5");

            await Task.Run(() => load());
            sendKey($"{index + 4}");

            await Task.Run(() => load());
            sendKey("[");

            await Task.Run(() => load());
            sendKey("'");

            await Task.Run(() => load());
            sendKey("FF37/" + RqID);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("5");

            await Task.Run(() => load());
            sendKey($"{poIndex + 1}");

            await Task.Run(() => load());
            sendKey("[");

            await Task.Run(() => load());
            sendKey("'");

            await Task.Run(() => load());
            sendKey("PO/" + RqID);

            await Task.Run(() => load());
            sendKey("{ENTER}");

            await Task.Run(() => load());
            sendKey("6");

            await Task.Run(() => load());
            sendKey(AgentName);

            await Task.Run(() => load());
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
            if (string.IsNullOrWhiteSpace(Cost) || string.IsNullOrWhiteSpace(WorkID) || string.IsNullOrWhiteSpace(Trip) || string.IsNullOrWhiteSpace(RqID) || string.IsNullOrWhiteSpace(AgentName))
                return false;
            else
                return true;
        }

        #endregion

        #region #Evaluate Input

        public async void evaluateForError(string action)
        {
            await Task.Run(() => loadChecker());
            sendKey("^{down}");

            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");
            await Task.Run(() => loadChecker());
            sendKey("^{up}");

            await Task.Run(() => loadChecker());
            sendKey("^c");

            await Task.Run(() => loadChecker());
            sendKey("^c");

            await Task.Run(() => loadChecker());

            if (Clipboard.GetData(DataFormats.Text).ToString().Contains("FORMAT") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("ITEM") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("SUCH") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("NO") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("BGNG") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("CHECK") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("ENTRY") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("INVALID") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("FRMT") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("CODE") ||
                 Clipboard.GetData(DataFormats.Text).ToString().Contains("ACTION"))
            {


                Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() => {
                    MessageBox.Show(Application.Current.MainWindow, "Error Found ! trying to re-input..");
                }));

                await Task.Run(() => loadChecker());
                sendKey("{ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC}");

                await Task.Run(() => loadChecker());
                IgnoreRetrieve();

                if (action == "submit")
                    submit();
                else
                    edit();
            }
            else
            {
                await Task.Run(() => loadChecker());
                sendKey("{ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC} {ESC}");

                SavePnr();
            }
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
