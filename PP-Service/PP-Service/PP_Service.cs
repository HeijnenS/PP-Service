using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PP_Service
{
    public partial class PP_Service : ServiceBase
    {
        private TcpListener tcpListener;
        private Thread listenThread;
        private int connectedClients = 0;
        private string ReceivedMessage;
        private string PowerpointPath;
        private PowerPointComFunctions PP;
        public System.Diagnostics.EventLog eventLog1;
        private bool SerialChecked = false;

        public PP_Service()
        {
            InitializeComponent();


            //check for motherboard serial
            //wmic baseboard get product,Manufacturer,version,serialnumber
            if (getMotherBoardSerial()=="PDWkuh21W5E6SX")
            {
                SerialChecked = true;
            }
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("AMX PP Service"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "AMX PP Service", "AMX PP Service Log");
            }
            eventLog1.Source = "AMX PP Service";
            eventLog1.Log = "AMX PP Service Log";
        }

        static String getMotherBoardSerial()
        {
            String serial = "";
            try
            {
                ManagementObjectSearcher mos = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard");
                ManagementObjectCollection moc = mos.Get();

                foreach (ManagementObject mo in moc)
                {
                    serial = mo["SerialNumber"].ToString();
                }
                return serial;
            }
            catch (Exception)
            {
                return serial;
            }
        }

        protected override void OnStart(string[] args)
        {
            if (SerialChecked)
            {
                eventLog1.WriteEntry("License has been registered, starting service");
                PowerpointPath = "c:\\temp";
                tcpListener = new TcpListener(IPAddress.Any, 3001);
                listenThread = new Thread(new ThreadStart(ListenForClients));
                listenThread.Start();
                eventLog1.WriteEntry("Starting Listening Service");
            }
            else
            {
                eventLog1.WriteEntry("License of [" + getMotherBoardSerial() + "] has not been registered yet");
                Console.WriteLine("Your serial has not been registered yet, please purchase a license.");
            }

        }

        protected override void OnStop()
        {
        }

        #region tcpip server

        private void ListenForClients()
        {
            this.tcpListener.Start();

            while (true) // Never ends until the Server is closed.
            {
                //blocks until a client has connected to the server
                TcpClient client = this.tcpListener.AcceptTcpClient();
                eventLog1.WriteEntry("Accepted client" + client.Connected.ToString());
                //create a thread to handle communication 
                //with connected client
                connectedClients++; // Increment the number of clients that have communicated with us.
                Thread clientThread = new Thread(new ParameterizedThreadStart(HandleClientComm));
                clientThread.Start(client);
                eventLog1.WriteEntry("Starting Client Thread");
            }
        }

        private void HandleClientComm(object client)
        {
            TcpClient tcpClient = (TcpClient)client;
            NetworkStream clientStream = tcpClient.GetStream();
            string CarriageReturnLineFeed = "\r\n";

            byte[] message = new byte[4096];
            byte[] response = new byte[4096];
            int bytesRead;
            eventLog1.WriteEntry("Thread HandleClientComm Started");
            while (true)
            {
                bytesRead = 0;

                try
                {//blocks until a client sends a message                  
                    bytesRead = clientStream.Read(message, 0, 4096);
                }
                catch
                {//a socket error has occured                    
                    break;
                }

                if (bytesRead == 0)
                {//the client has disconnected from the server
                    connectedClients--;
                    break;
                }
                else
                {//message has successfully been received
                    ASCIIEncoding encoder = new ASCIIEncoding();
                    // Convert the Bytes received to a string and display it on the Server Screen
                    string msg = encoder.GetString(message, 0, bytesRead);
                    ReceivedMessage = ReceivedMessage + msg;
                    if (ReceivedMessage.IndexOf(CarriageReturnLineFeed) > 0) //as long as return is not given, then add characters to string
                    {
                        string[] aReceivedMessage = ReceivedMessage.Split(CarriageReturnLineFeed.ToCharArray());
                        //Split result in array of 3 elements????? should be 2!!!
                        foreach (string myMessage in aReceivedMessage)
                        {
                            if (myMessage.Length > 0)
                            {
                                //Console.WriteLine("Received " + myMessage);
                                eventLog1.WriteEntry("Received " + myMessage);
                                //clientStream.Write(HandleCompleteMessage(ReceivedMessage));
                                string temp;
                                if (myMessage != "exit")
                                {
                                    temp = HandleCompleteMessage(myMessage);
                                    string[] cmd;
                                    cmd = new string[2];

                                    if (temp.IndexOf('-') > 0)
                                    {
                                        cmd = temp.Split('-');
                                    }
                                    if (cmd[0] == "error")
                                    {
                                        eventLog1.WriteEntry(temp);
                                    }
                                    clientStream.Write(encoder.GetBytes(temp), 0, encoder.GetByteCount(temp));
                                }
                                else
                                {
                                    tcpClient.Close();
                                    connectedClients--;
                                }
                            }
                        }
                        ReceivedMessage = "";
                    }
                }
            }
            tcpClient.Close();
        }

        #endregion

        #region messagehandler
        private string HandleCompleteMessage(string message)
        {
            if (message[0] == '?')
            {
                return HandleCompleteRequest(message) + "\r\n";
            }
            else
            {
                return HandleCompleteCommand(message) + "\r\n";
            }
        }

        private string HandleCompleteRequest(string message)
        {
            string[] cmd;
            cmd = new string[2];

            if (message.IndexOf('-') > 0)
            {
                cmd = message.Split('-');
            }
            else
            {
                cmd[0] = message;
            }

            switch (cmd[0]) //only request that don't use the PP object
            {
                case "?officeversion":
                    return GetOfficeVersion();
                case "?operatingsystem":
                    return OperatingSystem();
                case "?powerpointpath":
                    return ActivePowerpointPath();
                case "?powerpointpresent":
                    return PowerpointPresent();
                case "?powerpoints":
                    return GetPowerpoints();
            }

            if (PP == null) //after this check it is assumed that PP exists!!
            {
                return cmd[0] + "-Please check ?powerpointpresent";
            }

            switch (cmd[0])
            {
                case "?activepresentation":
                    return PP.HasActivePresentation();
                case "?currentslide":
                    return PP.ActiveSlide();
                case "?passedtime":
                    return PP.PresentationElapsedTime();
                case "?slides":
                    return PP.SlideCount();
                case "?slideelapsedtime":
                    return PP.SlideElapsedTime();
                case "?slidename":
                    return PP.SlideName(cmd[1]);
                default:
                    return "unknown request-" + message;
            }
        }

        private string HandleCompleteCommand(string message)
        {
            string[] cmd;
            cmd = new string[2];

            if (message.IndexOf('-') > 0)
            {
                cmd = message.Split('-');
            }
            else
            {
                cmd[0] = message;
            }

            if (cmd[0] == "powerpointpath")
            {
                return ChangePath(cmd[1]);
            }

            if (PP == null) //after this check it is assumed that PP exists!!
            {
                return cmd[0] + "-Please check ?powerpointpresent";
            }

            switch (cmd[0])
            {
                case "open":
                    return PP.OpenPowerPoint(cmd[1]);
                case "close":
                    return PP.ClosePowerPoint();
                case "closeapplication":
                    return PP.CloseApplication();
                case "nextslide":
                    return PP.NextSlide();
                case "previousslide":
                    return PP.PreviousSlide();
                case "firstslide":
                    return PP.FirstSlide();
                case "lastslide":
                    return PP.LastSlide();
                case "gotoslide":
                    return PP.GoToSlide(cmd[1]);
                default:
                    return "unknown request-" + cmd[0];
            }
        }
        #endregion

        #region global functions

        static string ConvertStringArrayToString(string[] array)
        {
            StringBuilder builder = new StringBuilder();
            foreach (string value in array)
            {
                builder.Append(value);
                builder.Append(',');
            }
            builder.Remove(builder.Length - 1, 1);
            return builder.ToString();
        }

        private string OperatingSystem()
        {
            return "operatingsystem-" + Environment.OSVersion.ToString();
        }

        private string GetOfficeVersion()
        {
            string sVersion = string.Empty;
            Microsoft.Office.Interop.PowerPoint.Application appVersion = new Microsoft.Office.Interop.PowerPoint.Application();
            //appVersion.Visible = Office.MsoTriState.msoFalse;
            switch (appVersion.Version.ToString())
            {
                case "7.0":
                    sVersion = "95";
                    break;
                case "8.0":
                    sVersion = "97";
                    break;
                case "9.0":
                    sVersion = "2000";
                    break;
                case "10.0":
                    sVersion = "2002";
                    break;
                case "11.0":
                    sVersion = "2003";
                    break;
                case "12.0":
                    sVersion = "2007";
                    break;
                case "14.0":
                    sVersion = "2010";
                    break;
                default:
                    sVersion = "Too Old!";
                    break;
            }
            return "officeversion-" + sVersion;
        }

        private string PowerpointPresent()
        {
            Type officeType = Type.GetTypeFromProgID("Powerpoint.Application");

            if (officeType == null)
            {
                return "powerpointpresent-no";
            }
            else
            {
                PP = new PowerPointComFunctions();
                PP.PowerpointPath = PowerpointPath;
                return "powerpointpresent-yes";
            }
        }

        private string ActivePowerpointPath()
        {
            return "powerpointpath-" + PowerpointPath;
        }

        private string GetPowerpoints()
        {
            string PPTResponse;
            PPTResponse = "";
            try
            {
                if (PowerpointPath.Length == 0)
                {
                    PowerpointPath = "C://Temp/";

                    //  / is seen as escape sequence......
                    // PowerpointPath.Split("/");
                    // PowerpointPath = "C:\Users\brian\Documents\powerpoints";
                }
                else
                {
                    PowerpointPath.Replace("\'", "/");
                }

                string[] pptFiles = Directory.GetFiles(PowerpointPath, "*.ppt").Select(path => Path.GetFileName(path)).ToArray();

                if (pptFiles.Length > 0)
                {
                    PPTResponse = ConvertStringArrayToString(pptFiles);
                }
                else
                {
                    PPTResponse = "none";
                }
                return "powerpoints-" + PPTResponse;
            }
            catch(Exception ex)
            {
                //Console.WriteLine("GetPowerpoints throws the error: {0}", ex.Message);
                eventLog1.WriteEntry("GetPowerpoints throws the error: {0}" + ex.Message);
                return ex.Message.ToString();
            }
        }

        private string ChangePath(string Path)
        {
            try
            {
                if (Directory.Exists(Path))
                {
                    PowerpointPath = Path;
                    if (PP != null)
                    {
                        PP.PowerpointPath = PowerpointPath;
                    }
                    return "powerpointpath-[" + Path + "] activated";
                }
                else
                {
                    return "powerpointpath-[" + Path + "] does not exist";
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine("SetPath throws the error: {0}", ex.Message);
                eventLog1.WriteEntry("SetPath throws the error: {0}" + ex.Message);
                return ex.Message.ToString();
            }
        }

        #endregion


    }
}
