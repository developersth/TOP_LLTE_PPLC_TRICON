using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Threading;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.ComponentModel;
using System.Windows.Forms;
using System.Diagnostics;
using System.Collections;
using System.Runtime.InteropServices;

using Modbus.Data;
using Modbus.Device;
using Modbus.Utility;

using Kepware.ClientAce.OpcDaClient;
using Kepware.ClientAce.OpcCmn;

namespace PPLC
{
    public class PLCProcess : IDisposable
    {
        #region Constant, Struct and Enum
        const int offsetFC01 = 1;
        const int offsetFC02 = 100001;
        const int offsetFC03 = 400001;
        const int offsetFC04 = 300001;
        const int offsetFC05 = 1;
        const int offsetFC06 = 400001;
        const ushort offsetFC15 = 1;
        const int offsetFC16 = 400001;

        struct _PLCBuffer
        {
            public string PLCName;
            public int Id;
            public string AddressIO;
            public string TagName;
            public int DataType;
            public string DataTypeDesc;
            public bool ReadEnable;
            public DateTime TimeStamp;
            public int ItemQuailty;
            public string ItemQualityDesc;
            public object ReadNumber;
            public string TagGroup;
            public double WriteNo;
            public string WriteNumber;
        }
                
        enum _PLCStepProcess :int
        {
            InitialDataBuffer=0,
            InitialAddressModbus=10,
            InitialReadModbus=11,
            ReadDataModbus=12,
            InitialKepwareOPCServer=20,
            ConnectKepwareOPCServer=21,
            InitialKepwareOPCItem=22,
            BroweOPCItem=23,
            ReadOPCItem=24,
            ChangeProcess=30
        } 
        #endregion
        
        Comport plcPort;
        frmMain fMain;

        _PLCBuffer[] plcBuffer;
        
        _PLCStepProcess plcStepProcess;

        int processId=0;
        //int atgAddress;
        DateTime chkResponse;
        DateTime refreshDisplay;
        bool opcConnectStatus;
        int timeOutChangeProcess=0;
        int updateRateIPS = 350;

        #region Construct and Deconstruct
        private bool IsDisposed = false;
        
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
            //mRunn = false;
        }
        
        protected void Dispose(bool Diposing)
        {
            if (!IsDisposed)
            {
                if (Diposing)
                {
                    //Clean Up managed resources
                    thrShutdown = true;
                    //if (atgModbus != null)
                    //    atgModbus.Dispose();

                    //if (daServerMgt != null)
                    //{
                    //    DisconnectOPCServer();
                    //    daServerMgt.Dispose();
                    //}
                }
                //Clean up unmanaged resources
            }
            IsDisposed = true;
        }
        
        public PLCProcess(frmMain pFrm, ref Comport pPort)
        {
            fMain = pFrm;
            plcPort = pPort;
        }

        public PLCProcess(frmMain pFrm,int pAddr,int pId, string pScanName,string pOPCServerName,string pOPCChannelName,string pOPCDeviceName)
        {
            fMain = pFrm;
            processId = pId;
            clientName = pScanName;
            endPointURL = pOPCServerName;
            opcChannelName = pOPCChannelName;
            opcDeviceName = pOPCDeviceName;
        }

        public PLCProcess(frmMain pFrm, int pAddr, int pId, string pScanName, string pOPCServerName, string pOPCChannelName, string pOPCDeviceName,string pPrefixTag)
        {
            fMain = pFrm;
            processId = pId;
            clientName = pScanName;
            endPointURL = pOPCServerName;
            opcChannelName = pOPCChannelName;
            opcDeviceName = pOPCDeviceName;
            opcPrefixTag = pPrefixTag;
        }

        ~PLCProcess()
        { }
        #endregion

        #region Class Events
        public delegate void ATGEventsHaneler(object pSender, string pEventMsg);
        public event ATGEventsHaneler OnATGEvents;
        string logFileName;

        void RaiseEvents(string pSender, string pMsg)
        {
            string vMsg = DateTime.Now + ">[" + pSender + "]" + pMsg;
            logFileName = "ATG";
            try
            {
                fMain.AddListBox = "" + logFileName + ">" + vMsg;
                //fMain.LogFile.WriteLog(logFileName, vMsg);
            }
            catch (Exception exp)
            { }
        }

        void RaiseEvents(string pMsg)
        {
            logFileName = "ATG";
            try
            {
                fMain.AddListBox = DateTime.Now + "> " + pMsg;
                //fMain.LogFile.WriteLog(logFileName, vMsg);
            }
            catch (Exception exp)
            { }
        }
        #endregion

        #region Thread
        bool thrConnect;
        bool thrShutdown;
        bool thrRunning;
        object thrLock = new object();

        Thread thrMain;
        Thread thrWritePLC;
        Thread thrReWirePLC;
        Thread thrGUI;
        public void StartThread()
        {
            System.Threading.Thread.Sleep(1000);
            thrMain = new Thread(this.RunProcess);
            thrMain.Name = processId.ToString() + "thrPLC";
            thrMain.Start();

            thrWritePLC = new Thread(this.ProcessReWritePLC);
            thrWritePLC.Name = "thrReWritePLC";
            thrWritePLC.Start();
            Thread.Sleep(1000);

            thrWritePLC = new Thread(this.ProcessWritePLC);
            thrWritePLC.Name = "thrWritePLC";
            thrWritePLC.Start();

            

            thrGUI = new Thread(this.DisplayMessageGUI);
            thrGUI.Name = "thrGUI";
            thrGUI.Start();
        }

        public void StopThread()
        {
            thrShutdown = true;
            UpdatePLCConnect(0);

            if (atgModbus != null)
                atgModbus.Dispose();

            if (daServerMgt != null)
            {
                DisconnectOPCServer();
                daServerMgt.Dispose();
            }
        }
        
        private void RunProcess()
        {
            thrRunning = true;
            thrShutdown = false;
            plcStepProcess = _PLCStepProcess.InitialDataBuffer;
            chkResponse = DateTime.Now;
            refreshDisplay = DateTime.Now;
            while (thrRunning)
            {
                try
                {
                    if (thrShutdown)
                        return;
                    switch (plcStepProcess)
                    {
                        case _PLCStepProcess.InitialDataBuffer:
                            if (InitialDataBuffer())
                            {
                                InitialTimeoutChangeProcess();
                                //SetComport();     //N/A for connect to OPC server
                                //atgStepProcess = _ATGStepProcess.InitialAddressModbus;        //connect via NModbus
                                plcStepProcess = _PLCStepProcess.InitialKepwareOPCServer;       //connect via Kepware OPC Server
                            }
                            break;
                        case _PLCStepProcess.InitialAddressModbus:
                            plcStepProcess = _PLCStepProcess.InitialReadModbus;
                            break;
                        case _PLCStepProcess.InitialReadModbus:
                            if (InitialReadModbus())
                            {
                                plcStepProcess = _PLCStepProcess.ReadDataModbus;
                            }
                            break;
                        case _PLCStepProcess.InitialKepwareOPCServer:
                            if (KepwareInitialOPCServer())
                            {
                                //atgStepProcess = _ATGStepProcess.ConnectKepwareOPCServer;
                                //SubscribeToOPCDAServerEvents();
                                plcStepProcess = _PLCStepProcess.InitialKepwareOPCItem;
                            }
                            break;
                        case _PLCStepProcess.ConnectKepwareOPCServer:
                            if (KepwareConnectOPCServer())
                            {
                                //SubscribeToOPCDAServerEvents();
                                //atgStepProcess = _ATGStepProcess.InitialKepwareOPCItem;

                                //Subsctibe();
                                plcStepProcess = _PLCStepProcess.ReadOPCItem;
                            }
                            else
                            {
                                //atgStepProcess = _ATGStepProcess.InitialKepwareOPCServer;
                                ChangeProcess();
                            }
                            break;
                        case _PLCStepProcess.InitialKepwareOPCItem:
                            if (KepwareInitialOPCItem())
                            {
                                //atgStepProcess = _ATGStepProcess.ReadOPCItem;
                                plcStepProcess = _PLCStepProcess.ConnectKepwareOPCServer;
                                SubscribeToOPCDAServerEvents();
                            }
                            break;
                        case _PLCStepProcess.ReadOPCItem:
                            KepwareReadItem();
                            //Thread.Sleep(500);
                            if (daServerMgt.IsConnected)
                            {
                                if (daServerMgt.ServerState != ServerState.CONNECTED)
                                {
                                    //DisconnectOPCServer();
                                    //atgStepProcess = _ATGStepProcess.ConnectKepwareOPCServer;

                                    //Unsubscribe();
                                    //Subsctibe();
                                }
                            }
                            //WritePLC();
                            plcStepProcess = _PLCStepProcess.ChangeProcess;;
                            break;
                        case _PLCStepProcess.ChangeProcess:
                            plcStepProcess = _PLCStepProcess.ReadOPCItem;
                            ChangeProcess();
                            //bool vRet = false;
                            //vRet = CheckItemBadValueAll();
                            //var vDiff = (DateTime.Now - chkResponse).TotalMinutes;
                            //if ((vRet) && (vDiff>=1))
                            //{
                            //    chkResponse = DateTime.Now;
                            //    ChangeProcess();
                            //    Thread.Sleep(3000);
                            //}
                            break;
                    }
                }
                catch (Exception exp)
                { 
                    RaiseEvents(exp.Message + "[source>" + exp.Source + "]");
                    //Thread.Sleep(3000);
                }
                //finally
                //{
                //    mShutdown = true;
                //    mRunning = false;
                //}
                //Thread.Sleep(updateRateIPS);
            }
            UpdatePLCConnect(0);
        }

        #endregion

        #region PLC TRICONEX
        void InitialTimeoutChangeProcess()
        {
            DataSet vDataset = new DataSet();
            DataTable dt;
            string vSql = "select tas.F_TIMEOUT_IPS as timeout from dual";
            if (fMain.OraDb.OpenDyns(vSql, "TableName", ref vDataset))
            {
                dt = vDataset.Tables[0];
                timeOutChangeProcess = Convert.ToInt16(dt.Rows[0]["timeout"].ToString());
            }
        }

        bool InitialDataBuffer()
        {
            bool vRet = false;
            string vMsg;
            string strSQL = "select" +
                            " t.IPS_ID, t.OPC_TAG,t.TAG_GROUP" +
                            " from steqi.view_initial_ips_opc_tag t order by t.ips_id";

            DataSet vDataset = null;
            DataTable dt;
            if (fMain.OraDb.OpenDyns(strSQL, "TableName", ref vDataset))
            {
                dt = vDataset.Tables["TableName"];
                plcBuffer = new _PLCBuffer[dt.Rows.Count];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    plcBuffer[i].AddressIO = dt.Rows[i]["IPS_ID"].ToString();
                    plcBuffer[i].TagName = dt.Rows[i]["OPC_TAG"].ToString();
                    if (dt.Rows[i]["OPC_TAG"].ToString() != "")
                        plcBuffer[i].TagGroup = dt.Rows[i]["TAG_GROUP"].ToString();
                    else
                        plcBuffer[i].TagGroup = "";

                    //plcBuffer[i].DataType = Convert.ToInt16(dt.Rows[i]["DATA_TYPE"].ToString());
                    //   plcBuffer[i].DataTypeDesc = dt.Rows[i]["DATA_TYPE"].ToString();

                    vMsg = string.Format("Initial {0}, {1}.",
                                          (i + 1).ToString(),
                                          plcBuffer[i].TagName);

                    RaiseEvents(vMsg);
                    vRet = true;
                }
                RaiseEvents(string.Format("Total tag={0}", plcBuffer.Length - 1));
            }
            return vRet;
        }
        void UpdateToDatabase(_PLCBuffer pPLCBuffer)
        {
            //string vSQL = "begin update steqi.PLC_AB_DATA t set ";
            string vSQL = "";
            if (pPLCBuffer.ReadNumber != null)
            {
                vSQL = "begin steqi.P_UPDATE_PLC_IPS_VALUE(" +
                    "'" + pPLCBuffer.TagName + "'," +
                    pPLCBuffer.ReadNumber + "," + pPLCBuffer.ItemQuailty + ",'" + pPLCBuffer.TimeStamp + "'" +
                    ");end;";
                fMain.OraDb.ExecuteSQL(vSQL);
            }
        }
        void UpdateToDatebase()
        {
            string vSql = "";
            string vStatement = "";
            for (int i = 0; i < plcBuffer.Length; i++)
            {
                if (plcBuffer[i].ReadNumber != null)
                {
                    vStatement = string.Format("t.value={0}, t.connect_status={1}, t.time_stamp=to_date('{2}','DD/MM/YYYY HH24:MI:SS'),t.update_date=sysdate"
                        , plcBuffer[i].ReadNumber, plcBuffer[i].ItemQuailty, plcBuffer[i].TimeStamp);
                    vSql += string.Format("update steqi.IPS_OPC_TAG t set {0} where t.opc_tag='{1}';", vStatement,plcBuffer[i].TagName);
                }
                if (i % 20 == 0 && i > 0 && vSql.Length > 0)
                {
                    fMain.OraDb.ExecuteSQL("begin " + Environment.NewLine + vSql + "end;");
                    vSql = "";
                }
            }

            if(vSql.Length > 0)
                fMain.OraDb.ExecuteSQL("begin " + Environment.NewLine + vSql + "end;");
        }
        void UpdatePLCConnect()
        {
            string strSQL = "begin tas.P_UPDATE_IPS_CONNECT(" +
                             "'" + processId.ToString() + "'," + Convert.ToInt16(daServerMgt.IsConnected) +
                             ");end;";

            fMain.OraDb.ExecuteSQL(strSQL);
        }

        void UpdatePLCConnect(int pConnect)
        {
            string strSQL = "begin tas.P_UPDATE_IPS_CONNECT(" +
                             "'" + processId.ToString() + "'," + pConnect +
                             ");end;";

            fMain.OraDb.ExecuteSQL(strSQL);
        }
        void ProcessReWritePLC()
        {
            RaiseEvents("Start re-write plc.");
            string vMsg = "";
            while (thrRunning)
            {
                try
                {
                    if (thrShutdown)
                        return;
                    if (daServerMgt == null || daServerMgt.IsConnected == false)
                        goto Nextt;

                    string strSQL = "select" +
                                " t.IPS_TAG,t.REWRITE_VALUE,t.OPC_TAG" +
                                " from steqi.VIEW_IPS_OPC_REWRITE_TAG t " +
                                " order by t.IPS_TAG ";
                    DataSet vDataset = null;
                    DataTable dt;
                    if (fMain.OraDb.OpenDyns(strSQL, "TableName", ref vDataset))
                    {
                        dt = vDataset.Tables["TableName"];
                        if (dt.Rows.Count > 0)
                        {
                            lock (thrLock)
                            {
                                RaiseEvents("Begin re-write plc.");
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    KepwareInitialWriteItem(dt.Rows[i]["OPC_TAG"].ToString(), (object)dt.Rows[i]["REWRITE_VALUE"].ToString());
                                    Thread.Sleep(updateRateIPS);
                                }
                                RaiseEvents("End re-write plc.");
                            }
                        }
                        else
                            break;
                        //iniRead = false;
                    }
                Nextt:
                    Thread.Sleep(1000);
                }
                catch (Exception exp)
                {
                    fMain.LogFile.WriteErrLog(exp.Message); 
                }
            }
        }
        void ProcessWritePLC()
        {
            RaiseEvents("Start write plc.");
            string vMsg = "";
            while (thrRunning)
            {
                try
                {
                    if (thrShutdown)
                        return;
                    if (daServerMgt == null || daServerMgt.IsConnected == false)
                        goto Nextt;

                    string strSQL = "select" +
                                " t.WRITE_NO,t.IPS_NAME,t.IPS_ID,t.WRITE_NUMBER, t.DESCRIPTION,t.OPC_TAG,t.READ_TAG" +
                                " from steqi.VIEW_IPS_OPC_WRITE t" +
                                " order by t.WRITE_NO";
                    DataSet vDataset = null;
                    DataTable dt;
                    if (fMain.OraDb.OpenDyns(strSQL, "TableName", ref vDataset))
                    {
                        dt = vDataset.Tables["TableName"];
                        if (dt.Rows.Count > 0)
                        {
                            lock (thrLock)
                            {
                                RaiseEvents("Begin write plc.");
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    KepwareInitialWriteItem(dt.Rows[i]["OPC_TAG"].ToString(), (object)dt.Rows[i]["WRITE_NUMBER"].ToString(), dt.Rows[i]["READ_TAG"].ToString());
                                    DeleteWritePLC(Convert.ToDouble(dt.Rows[i]["WRITE_NO"].ToString()));
                                    vMsg = string.Format("Delete write no.={0} OPC_TAG:{1} TAG_NAME:{2}", dt.Rows[i]["WRITE_NUMBER"].ToString()
                                                                                                        , dt.Rows[i]["OPC_TAG"]
                                                                                                        , dt.Rows[i]["IPS_NAME"]);
                                    Thread.Sleep(updateRateIPS);
                                }
                                RaiseEvents("End write plc.");
                            }
                        }
                        //iniRead = false;
                    }
                Nextt:
                    Thread.Sleep(1000);
                }
                catch (Exception exp)
                {
                    fMain.LogFile.WriteErrLog(exp.Message);
                }
            }
        }
        void DeleteWritePLC(double pWriteNo)
        {
            string strSQL = "begin steqi.P_UPDATE_DELETE_WRITE_IPS_OPC(" +
                              pWriteNo + 
                             ");end;";

            fMain.OraDb.ExecuteSQL(strSQL);
        }
        int FindItemIndex(string pTagName)
        {
            int vIndex = -1;
            vIndex = Array.FindIndex(plcBuffer, x => x.TagName.Contains(pTagName));

            return vIndex;
        }
        #endregion

        #region define Enum String Value
        enum _DataType
        {
            [EnumValue("BOOLEAN")]
            BOOLEAN = 1,
            [EnumValue("BYTE")]
            BYTE = 2,
            [EnumValue("SHORT INT")]
            SHORT_INT = 3,
            [EnumValue("LONG INT")]
            LONG_INT = 4,
            [EnumValue("FLOAT")]
            FLOAT = 5,
            [EnumValue("STRING")]
            STRING = 6,
            [EnumValue("DIGI")]
            DIGI=7,
            [EnumValue("LONG")]
            LONG=8
        }
        class EnumValue : System.Attribute
        {
            private string _value;
            public EnumValue(string value)
            {
                _value = value;
            }
            public string Value
            {
                get { return _value; }
            }
        }
        static class EnumString
        {
            public static string GetStringValue(Enum value)
            {
                string output = null;
                Type type = value.GetType();
                System.Reflection.FieldInfo fi = type.GetField(value.ToString());
                EnumValue[] attrs = fi.GetCustomAttributes(typeof(EnumValue), false) as EnumValue[];
                if (attrs.Length > 0)
                {
                    output = attrs[0].Value;
                }
                return output;
            }
        } 
        #endregion

        #region NModbus
        ModbusSerialMaster atgModbus;
        //const ushort offsetFC01 = 1;
        //const ushort offsetFC02 = 10001;
        //const ushort offsetFC03 = 40001;
        //const ushort offsetFC04 = 30001;
        //const ushort offsetFC05 = 1;
        //const ushort offsetFC06 = 40001;
        //const ushort offsetFC15 = 1;
        //const ushort offsetFC16 = 40001;

        bool InitialReadModbus()
        {
            bool vRet = false;
            try
            {
                //SetComport();
                atgModbus = ModbusSerialMaster.CreateRtu(plcPort.Sp);
                atgModbus.Transport.ReadTimeout = 1000;
                atgModbus.Transport.Retries = 3;
                vRet = true;
            }
            catch (Exception exp)
            {
                RaiseEvents(exp.Message + "[source>" + exp.Source + "]");
                Thread.Sleep(3000);
            }
            return vRet;
        }

        #endregion

        #region Kepware CilentAce
        #region Constant, Struct and Enum
        const string opcItemName = "ns=2;s=$val1.$val2.$val3";  //$val1,$val2,$val3 are Channel Name,Device Name and Tag Name respectively.
        const string opcServerName = "";
        const string nodeName = "localhost";
        //const string endPointURL = "opc.tcp://localhost:49320";
        //const string clientName = "ATG";
        #endregion
        
        string opcChannelName = "";
        string opcDeviceName = "";
        string endPointURL = "";
        string clientName = "";
        string opcPrefixTag = "";

        DaServerMgt daServerMgt;
        int clientHandle = 1;
        ConnectInfo connectInfo;
        bool connectFailed = false;

        ServerIdentifier[] opcServerIdentifier;
        ItemIdentifier[] itemIdentifiers;
        ItemIdentifier[] diagIdentifiers;
        ItemValue[] diagValues;
        ItemValue[] opcItemValue;

        OpcServerEnum opcServerNum = new OpcServerEnum();

        // When we create a subscription, the server will give us a handle that
        // we must use when referencing that subscription in the server. For
        // example, we would use this handle to modify or cancel the subscription.
        // In a real world application, we would probably want to have the ability
        // to create multiple subscriptions. For example, you could create a subscription
        // with a high update rate for items that are time critical, and a second 
        // subscription with a lower update rate for the less time critical items.
        // This design would help you achieve the response you need for the time
        // critical items with minimum stress on the server, network, and devices.
        // However, we will be working with a single subscription in this example 
        // to keep things simple.
        int activeServerSubscriptionHandle;

        // Where the server subscription handle refers to a subscription within the
        // server, you can specify a client subscription handle to refer to that same
        // subscription here. This handle will be included in each data change
        // notification, along with latest values for one or more items enrolled
        // in that subscription. You should choose your client handle such that
        // it will be easy to "look up" that subscription and its associated
        // items. For example, the client handle might be a collection key, where
        // each object in the collection describes the subscription and the items
        // associated with it. Again, we will be working with a single subscription
        // in this example to keep things simple.
        int activeClientSubscriptionHandle;

        bool KepwareInitialOPCServer()
        {
            bool vRet = false;

            try
            {
                // The connectInfo structure defines a number of connection parameters.
                connectInfo = new ConnectInfo();

                // The LocalID member allows you to specify possible language options the server may support
                connectInfo.LocalId = "th";

                // The KeepAliveTime member is the time interval, in ms, in
                // which the connection to the server is checked by the API.
                connectInfo.KeepAliveTime = 60000;

                // The RetryAfterConnectionError tells the API to automatically
                // try to reconnect after a connection loss. This is nice, so 
                connectInfo.RetryAfterConnectionError = true;

                // The RetryInitialConnection tells the API to continue to try to
                // establish an initial connection. This is good as long as we
                // know for sure the server is really present and will likely allow
                // a connection at some point. if not, we could be here for a while
                // so best to set to false:
                connectInfo.RetryInitialConnection = false;

                // The Client Name allows the unique identification of client applications.
                // This allows the ability to easily identify multiple
                // connected clients to a server.
                connectInfo.ClientName = clientName;

                // Create an OPC DA Server Management object for each server you will connect with.
                daServerMgt = new DaServerMgt();
                string vMsg = string.Format("Initial OPC Server endpoint URL={0} Channel={1} Device={2}.",
                                            endPointURL,opcChannelName,opcDeviceName);
                RaiseEvents(vMsg);
                
                vRet = true;
            }
            catch (Exception exp)
            {
                RaiseEvents(exp.Message + "[source>" + exp.Source + "]"); 
            }
            return vRet;
        }

        bool KepwareConnectOPCServer()
        {
            bool vRet = false;

            string vMsg = string.Format("Connect OPC Server endpoint URL={0} Channel={1} Device={2}.",
                                            endPointURL, opcChannelName, opcDeviceName);
            RaiseEvents(vMsg);
            try
            {
                daServerMgt.Connect(endPointURL, clientHandle, ref connectInfo, out connectFailed);
            }
            catch (Exception exp)
            {
                RaiseEvents("Handled Connect exception. Reason: " + exp.Message + "[source>" + exp.Source + "]."); 
            }
            
            //vRet = !connectFailed;
            opcConnectStatus = daServerMgt.IsConnected;
            vRet = daServerMgt.IsConnected;
            vMsg = string.Format("OPC connect={0}.",vRet.ToString());
            RaiseEvents(vMsg);

            return vRet;
        }

        bool KepwareInitialOPCItem()
        {
            bool vRet = false;
            string vItemName = "";
            //List<ItemIdentifier> vItemList = new List<ItemIdentifier>();
            try
            {
                itemIdentifiers = new ItemIdentifier[plcBuffer.Length];
                opcItemValue = new ItemValue[itemIdentifiers.Length];
                for (int i = 0; i < plcBuffer.Length; i++)
                {
                    itemIdentifiers[i] = new ItemIdentifier();
                    opcItemValue[i] = new ItemValue();
                    if (opcPrefixTag == "")
                    {
                        vItemName = opcItemName.Replace("$val1", opcChannelName)
                                                .Replace("$val2", opcDeviceName)
                                                .Replace("$val3", OPCTagName(plcBuffer[i]));
                        itemIdentifiers[i].ItemName = vItemName;
                        itemIdentifiers[i].ClientHandle = i;
                    }
                    else
                    {
                        vItemName = opcItemName.Replace("$val1", opcPrefixTag + "." + opcChannelName)
                                                .Replace("$val2", opcDeviceName)
                                                .Replace("$val3", OPCTagName(plcBuffer[i]));
                        itemIdentifiers[i].ItemName = vItemName;
                        itemIdentifiers[i].ClientHandle = i;
                    }
                    //if (atgMember[i].DataType == (int)_DataType.BOOLEAN)
                    //{
                    //    itemIdentifiers[i].DataType = Type.GetType("System.Boolean[]");
                    //}
                }
                vRet = true;
            }
            catch (Exception exp)
            {
                RaiseEvents(exp.Message + "[source>" + exp.Source + "]."); 
            }

            return vRet;
        }

        void KepwareInitialDiagOPCServer()
        {
            diagIdentifiers = new ItemIdentifier[9];
            diagValues = new ItemValue[diagIdentifiers.Length];

            for (int i = 0; i < diagIdentifiers.Length; i++)
            {
                diagIdentifiers[i] = new ItemIdentifier();
                diagValues[i] = new ItemValue();

                diagIdentifiers[i].ClientHandle = i;
            }

            diagIdentifiers[0].ItemName = opcItemName.Replace("$val1",opcChannelName)
                                                        .Replace("$val2", "_Statistics")
                                                        .Replace("$val3","_LastResponseTime");

            diagIdentifiers[1].ItemName = opcItemName.Replace("$val1", opcChannelName)
                                                        .Replace("$val2", "_Statistics")
                                                        .Replace("$val3", "_LateData");

            diagIdentifiers[2].ItemName = opcItemName.Replace("$val1", opcChannelName)
                                                        .Replace("$val2", "_Statistics")
                                                        .Replace("$val3", "_MsgSent");
 
            diagIdentifiers[3].ItemName = opcItemName.Replace("$val1", opcChannelName)
                                                        .Replace("$val2", "_Statistics")
                                                        .Replace("$val3", "_MsgTotal");

            diagIdentifiers[4].ItemName = opcItemName.Replace("$val1", opcChannelName)
                                                        .Replace("$val2", "_Statistics")
                                                        .Replace("$val3", "_RxBytes");

            diagIdentifiers[5].ItemName = opcItemName.Replace("$val1", opcChannelName)
                                                        .Replace("$val2", "_Statistics")
                                                        .Replace("$val3", "_SuccessfulReads");

            diagIdentifiers[6].ItemName = opcItemName.Replace("$val1", opcChannelName)
                                                        .Replace("$val2", "_Statistics")
                                                        .Replace("$val3", "_SuccessfulWrites");

            diagIdentifiers[7].ItemName = opcItemName.Replace("$val1", opcChannelName)
                                                        .Replace("$val2", "_Statistics")
                                                        .Replace("$val3", "_TotalResponses");

            diagIdentifiers[8].ItemName = opcItemName.Replace("$val1", opcChannelName)
                                                        .Replace("$val2", "_Statistics")
                                                        .Replace("$val3", "_TxBytes");
        }

        void KepwareDiagOPCServer()
        {
            try
            {
                int maxAge = 100;
                int TransID = RandomNumber(65535, 1);
                ReturnCode returnCode = daServerMgt.Read(maxAge, ref diagIdentifiers, out diagValues);
            }
            catch (Exception exp)
            { }
        }

        void KepwareInitialWriteItem(string pItemName, object pValue,string pReadTag)
        {
            string vItemName = "";
            int vCount=0;
            int vIndex =0;
            ItemIdentifier[] itemIdentifier = new ItemIdentifier[1];
            ItemValue[] itemValue = new ItemValue[1];
            ReturnCode vReturnCode;

            itemIdentifier[0] = new ItemIdentifier();
            itemValue[0] = new ItemValue();

            if (opcPrefixTag == "")
            {
                vItemName = opcItemName.Replace("$val1", opcChannelName)
                                                .Replace("$val2", opcDeviceName)
                                                .Replace("$val3", OPCTagName(pItemName));
                itemIdentifier[0].ItemName = vItemName;
                itemIdentifier[0].ClientHandle = itemIdentifiers[vIndex].ClientHandle;
            }
            else
            {
                vItemName = opcItemName.Replace("$val1", opcPrefixTag + "." + opcChannelName)
                                                .Replace("$val2", opcDeviceName)
                                                .Replace("$val3", OPCTagName(pItemName));
                itemIdentifier[0].ItemName = vItemName;
                itemIdentifier[0].ClientHandle = itemIdentifiers[vIndex].ClientHandle;
            }

            itemValue[0].Value = pValue;
            lock (thrLock)
            {
                KepwareWriteItemOPC(ref itemIdentifier, ref itemValue, pItemName);

                Thread.Sleep(updateRateIPS);
                itemValue[0] = new ItemValue();

                vReturnCode = daServerMgt.Read(0, ref itemIdentifier, out itemValue);

                if (vReturnCode == ReturnCode.SUCCEEDED)
                {
                    vIndex = FindItemIndex(pItemName);
                    ReadOPCItem(ref plcBuffer[vIndex], itemValue[0]);
                    opcItemValue[vIndex] = itemValue[0];
                    plcBuffer[vIndex].ItemQualityDesc = itemValue[0].Quality.FullCode + "[" + itemValue[0].Quality.Name + "]"; //"Unknow";
                    plcBuffer[vIndex].TimeStamp = itemValue[0].TimeStamp;
                    plcBuffer[vIndex].ItemQuailty = itemValue[0].Quality.FullCode;
                    UpdateToDatabase(plcBuffer[vIndex]);

                    RaiseEvents(string.Format("DataChange Tag Name= {0,-20} Read value={1,-20} Quality={2,-10}.",
                                                  plcBuffer[vIndex].TagName,
                                                  plcBuffer[vIndex].ReadNumber,
                                                  plcBuffer[vIndex].ItemQualityDesc));
                    if(pItemName != pReadTag)
                    {
                        vIndex = FindItemIndex(pReadTag);
                        if (opcPrefixTag == "")
                        {
                            vItemName = opcItemName.Replace("$val1", opcChannelName)
                                                            .Replace("$val2", opcDeviceName)
                                                            .Replace("$val3", OPCTagName(pItemName));
                            itemIdentifier[0].ItemName = vItemName;
                            itemIdentifier[0].ClientHandle = itemIdentifiers[vIndex].ClientHandle;
                        }
                        else
                        {
                            vItemName = opcItemName.Replace("$val1", opcPrefixTag + "." + opcChannelName)
                                                            .Replace("$val2", opcDeviceName)
                                                            .Replace("$val3", OPCTagName(pReadTag));
                            itemIdentifier[0].ItemName = vItemName;
                            itemIdentifier[0].ClientHandle = itemIdentifiers[vIndex].ClientHandle;
                        }

                        vReturnCode = daServerMgt.Read(0, ref itemIdentifier, out itemValue);
                        if (vReturnCode == ReturnCode.SUCCEEDED)
                        {
                            ReadOPCItem(ref plcBuffer[vIndex], itemValue[0]);
                            opcItemValue[vIndex] = itemValue[0];
                            plcBuffer[vIndex].ItemQualityDesc = itemValue[0].Quality.FullCode + "[" + itemValue[0].Quality.Name + "]"; //"Unknow";
                            plcBuffer[vIndex].TimeStamp = itemValue[0].TimeStamp;
                            plcBuffer[vIndex].ItemQuailty = itemValue[0].Quality.FullCode;
                            UpdateToDatabase(plcBuffer[vIndex]);

                            RaiseEvents(string.Format("DataChange Tag Name= {0,-20} Read value={1,-20} Quality={2,-10}.",
                                                          plcBuffer[vIndex].TagName,
                                                          plcBuffer[vIndex].ReadNumber,
                                                          plcBuffer[vIndex].ItemQualityDesc));
                        }
                    }
                }
            }
        }

        void KepwareInitialWriteItem(string pItemName, object pValue)
        {
            string vItemName = "";
            int vCount = 0;
            int vIndex = FindItemIndex(pItemName);
            ItemIdentifier[] itemIdentifier = new ItemIdentifier[1];
            ItemValue[] itemValue = new ItemValue[1];
            ReturnCode vReturnCode;

            itemIdentifier[0] = new ItemIdentifier();
            itemValue[0] = new ItemValue();

            if (opcPrefixTag == "")
            {
                vItemName = opcItemName.Replace("$val1", opcChannelName)
                                                .Replace("$val2", opcDeviceName)
                                                .Replace("$val3", OPCTagName(pItemName));
                itemIdentifier[0].ItemName = vItemName;
                itemIdentifier[0].ClientHandle = itemIdentifiers[vIndex].ClientHandle;
            }
            else
            {
                vItemName = opcItemName.Replace("$val1", opcPrefixTag + "." + opcChannelName)
                                                .Replace("$val2", opcDeviceName)
                                                .Replace("$val3", OPCTagName(pItemName));
                itemIdentifier[0].ItemName = vItemName;
                itemIdentifier[0].ClientHandle = itemIdentifiers[vIndex].ClientHandle;
            }

            itemValue[0].Value = pValue;
            lock (thrLock)
            {
                KepwareWriteItemOPC(ref itemIdentifier, ref itemValue, pItemName);

                Thread.Sleep(updateRateIPS);
                itemValue[0] = new ItemValue();

                vReturnCode = daServerMgt.Read(0, ref itemIdentifier, out itemValue);

                if (vReturnCode == ReturnCode.SUCCEEDED)
                {
                    ReadOPCItem(ref plcBuffer[vIndex], itemValue[0]);
                    opcItemValue[vIndex] = itemValue[0];
                    plcBuffer[vIndex].ItemQualityDesc = itemValue[0].Quality.FullCode + "[" + itemValue[0].Quality.Name + "]"; //"Unknow";
                    plcBuffer[vIndex].TimeStamp = itemValue[0].TimeStamp;
                    plcBuffer[vIndex].ItemQuailty = itemValue[0].Quality.FullCode;
                    UpdateToDatabase(plcBuffer[vIndex]);

                    RaiseEvents(string.Format("DataChange Tag Name= {0,-20} Read value={1,-20} Quality={2,-10}.",
                                                  plcBuffer[vIndex].TagName,
                                                  plcBuffer[vIndex].ReadNumber,
                                                  plcBuffer[vIndex].ItemQualityDesc));
                }
            }
        }

        void KepwareWriteItemOPC(ref ItemIdentifier[] pItemIdentifier, ref ItemValue[] pItemValue, string pItemName)
        {
            ReturnCode returnCode;
            try
            {
                int vTransID = RandomNumber(65535, 1);
                // Call the Write API method:
                returnCode = daServerMgt.WriteAsync(vTransID, ref pItemIdentifier, pItemValue);
                if (returnCode != ReturnCode.SUCCEEDED)
                {
                    for (int i = 0; i < pItemIdentifier.Length; i++)
                    {
                        //string[] vTag = pItemIdentifier[i].ItemName.Split('.');
                        RaiseEvents(pItemName + "=" + pItemValue[i].Value.ToString() + "-> Write failed with a result of " + System.Convert.ToString(pItemIdentifier[i].ResultID.Code) + "\r\n" +
                                    " Description: " + pItemIdentifier[i].ResultID.Description);
                    }
                }
                else
                {
                    for (int i = 0; i < pItemIdentifier.Length; i++)
                    {
                        //string[] vTag = pItemIdentifier[i].ItemName.Split('.');
                        RaiseEvents(pItemName + "=" + pItemValue[i].Value.ToString() + "-> Write " + Enum.GetName(typeof(ReturnCode), returnCode));
                    }
                }
            }
            catch (Exception exp)
            {
                RaiseEvents("Handled Async Write exception. Reason: " + exp.Message);
            }
        }

        int transID=0;
        bool iniRead = false;
        void KepwareReadItem()
        {
            //return;
            //if (transID == 0)
            //{
            //    transID = RandomNumber(65535, 1);
            //    if (opcConnectStatus)
            //    {
            //        //ReturnCode returnCode = daServerMgt.Read(itemIdentifiers.Length + 1, ref itemIdentifiers, out opcItemValue);
            //        ReturnCode returnCode = daServerMgt.ReadAsync(transID, 0, ref itemIdentifiers);
            //    }
            //}
            //if (transID != -1)
            //    return;

            if (iniRead == true)
                goto Nextt;

            iniRead = false;
            lock (thrLock)
            {
                if (opcConnectStatus)
                {
                    ReturnCode returnCode = daServerMgt.Read(itemIdentifiers.Length + 1, ref itemIdentifiers, out opcItemValue);
                }
            }
                //Thread.Sleep(updateRateIPS);
                transID = 0;
            Nextt:
                for (int i = 0; i < plcBuffer.Length; i++)
                {
                    try
                    {
                        if (opcItemValue[i].Value == null)
                        {
                            //plcBuffer[i].ItemQualityDesc = opcItemValue[i].Quality.FullCode + "[" + opcItemValue[i].Quality.Name + "]"; //"Unknow";
                            //plcBuffer[i].TimeStamp = opcItemValue[i].TimeStamp;
                            //plcBuffer[i].ItemQuailty = opcItemValue[i].Quality.FullCode;
                            //UpdateToDatabase(plcBuffer[i]);
                        }
                        else
                        {
                            if (opcItemValue[i].Value != null)
                            {
                                if (opcItemValue[i].Quality.IsGood)
                                {
                                    ReadOPCItem(ref plcBuffer[i], opcItemValue[i]);
                                    plcBuffer[i].ItemQualityDesc = opcItemValue[i].Quality.FullCode + "[" + opcItemValue[i].Quality.Name + "]"; //"Unknow";
                                    plcBuffer[i].TimeStamp = opcItemValue[i].TimeStamp;
                                    plcBuffer[i].ItemQuailty = opcItemValue[i].Quality.FullCode;
                                    UpdateToDatabase(plcBuffer[i]);
                                }
                                else
                                {
                                    ReadOPCItem(ref plcBuffer[i], opcItemValue[i]);
                                    plcBuffer[i].ItemQualityDesc = opcItemValue[i].Quality.FullCode + "[" + opcItemValue[i].Quality.Name + "]"; //"Unknow";
                                    plcBuffer[i].TimeStamp = opcItemValue[i].TimeStamp;
                                    plcBuffer[i].ItemQuailty = opcItemValue[i].Quality.FullCode;
                                    UpdateToDatabase(plcBuffer[i]);
                                }
                                //if (plcBuffer[i].DataTypeDesc == "DIGI")
                                //{
                                //    RaiseEvents(DataChangeMessage(plcBuffer[i]));
                                //}
                            }
                        }
                    }
                    catch (Exception exp)
                    { }
                }
                //Nextt:
                transID = 0;
                //UpdateToDatebase();
                DateTime vDate = DateTime.Now;
                if (plcDiag || (vDate - refreshDisplay).Minutes >= 1)
                {
                    refreshDisplay = DateTime.Now;
                    RaiseEvents("Random display tag value.");
                    int vIndex = RandomNumber(plcBuffer.Length, 0);
                    for (int i = 0; i < 20; i++)
                    {
                        if(i < plcBuffer.Length)
                            DataChangeMessage(plcBuffer[i + vIndex]);
                    }
                }
        }

        bool CheckItemBadValueAll()
        {
            bool vRet = false;
            double[] vSum;
            double vAvg=0;

            vSum = new double[plcBuffer.Length];
            for (int i = 0; i < plcBuffer.Length; i++)
            {
                vSum[i] = plcBuffer[i].ItemQuailty;
            }
            vAvg = vSum.Average();
            if (vAvg == 0)  //bad quality all item
            {
                vRet = true;
                //vCheck=vSum/((Convert.ToInt16(QualityID.OPC_QUALITY_BAD.FullCode)*(atgMember.Length+1)));   //bad quality = 0
            }
            return vRet;
        }

        string OPCTagName(_PLCBuffer pPLCMember)
        {
            string vRet="";

            if (pPLCMember.TagGroup == "")
                vRet = pPLCMember.TagName;
            else
                vRet = pPLCMember.TagGroup + "." + pPLCMember.TagName;

            return vRet;
            //double vRet = 0;
            //switch (pATGMember.FncNo)
            //{
            //    case 1:
            //        vRet = (offsetFC01 + pATGMember.StartAddr) - 1;
            //        break;
            //    case 2:
            //        vRet = (offsetFC02 + pATGMember.StartAddr) - 1;
            //        break;
            //    case 3:
            //        vRet = (offsetFC03 + pATGMember.StartAddr) - 1;
            //        break;
            //    case 4:
            //        vRet = (offsetFC04 + pATGMember.StartAddr) - 1;
            //        break;
            //}
            //return vRet.ToString();
        }

        string OPCTagName(string pTagName)
        {
            string vRet="";
            int vIndex = FindItemIndex(pTagName);

            if (plcBuffer[vIndex].TagGroup == "")
                vRet = pTagName;
            else
                vRet = plcBuffer[vIndex].TagGroup + "." + pTagName;

            return vRet;
        }

        string DataChangeMessage(_PLCBuffer pPLCBuffer)
        {
            string vMsg = "";
            lock (thrLock)
            {
                //string vMsg = string.Format("DataChange Tag Name= {0} Address={1} Read Number={2} Quality={3}.",
                //                              pPLCBuffer.TagName,
                //                              ,pPLCBuffer.AddressIO,
                //                              pPLCBuffer.ReadNumber,
                //                              pPLCBuffer.ItemQualityDesc);
                vMsg = string.Format("DataChange Tag Name= {0,-20} Read value={1,-20} Quality={2,-10}.",
                                              pPLCBuffer.TagName,
                                              pPLCBuffer.ReadNumber,
                                              pPLCBuffer.ItemQualityDesc);
                queueMsg.Enqueue(vMsg);
            }
            return vMsg;
        }

        void RaiseEventMessage(string pMsg)
        {
            queueMsg.Enqueue(pMsg);
        }

        void ReadOPCItem(ref _PLCBuffer pPLCBuffer, ItemValueCallback pValue)
        {
            //cast pValue to array type
            //object[] vObj = ((IEnumerable)pValue.Value).Cast<object>()
            //                                          .Select(x => (object)x)
            //                                          .ToArray();           
            try
            {
                switch (pPLCBuffer.DataTypeDesc)
                {
                    //case (int)_DataType.BOOLEAN:
                    case "DIGI":
                        //pPLCBuffer.ValueBool = vObj.Cast<Boolean>().ToArray();
                        //Array.Reverse(pATGMember.ValueBool);
                        //BitArray arr = new BitArray(pPLCBuffer.ValueBool);
                        //var result = new int[1];
                        //arr.CopyTo(result, 0);
                        //pPLCBuffer.ValueUshort[0] = (ushort)result[0];
                        pPLCBuffer.ReadNumber = Convert.ToInt16(pValue.Value);
                        break;
                    //case (int)_DataType.SHORT_INT:
                    case "FLOAT":
                        //pPLCBuffer.ValueUshort = Array.ConvertAll<object, UInt16>(vObj, (x) => Convert.ToUInt16(x));
                        //pATGMember.ValueUshort = Array.ConvertAll(vObj, x => Convert.ToUInt16(x));
                        pPLCBuffer.ReadNumber = Convert.ToSingle(pValue.Value);
                        break;
                    case "LONG":
                        pPLCBuffer.ReadNumber = Convert.ToInt16(pValue.Value);
                        break;
                    default:
                        if (pValue.Value.ToString().ToLower() == "true" || pValue.Value.ToString().ToLower() == "false")
                        {
                            pPLCBuffer.ReadNumber = Convert.ToInt16(pValue.Value);
                        }
                        else
                        {
                            pPLCBuffer.ReadNumber = (object)pValue.Value;
                        }
                        break;
                }
            }
            catch (Exception exp)
            {
                RaiseEvents(exp.Message + "[source>" + exp.Source + "]."); 
            }
        }

        void ReadOPCItem(ref _PLCBuffer pPLCBuffer, ItemValue pValue)
        {
            //cast pValue to array type
            //object[] vObj = ((IEnumerable)pValue.Value).Cast<object>()
            //                                          .Select(x => (object)x)
            //                                          .ToArray();      
            lock (thrLock)
            {
                try
                {
                    switch (pPLCBuffer.DataTypeDesc)
                    {
                        //case (int)_DataType.BOOLEAN:
                        case "DIGI":
                            //pPLCBuffer.ValueBool = vObj.Cast<Boolean>().ToArray();
                            //Array.Reverse(pATGMember.ValueBool);
                            //BitArray arr = new BitArray(pPLCBuffer.ValueBool);
                            //var result = new int[1];
                            //arr.CopyTo(result, 0);
                            //pPLCBuffer.ValueUshort[0] = (ushort)result[0];
                            pPLCBuffer.ReadNumber = Convert.ToInt16(pValue.Value);
                            break;
                        //case (int)_DataType.SHORT_INT:
                        case "FLOAT":
                            //pPLCBuffer.ValueUshort = Array.ConvertAll<object, UInt16>(vObj, (x) => Convert.ToUInt16(x));
                            //pATGMember.ValueUshort = Array.ConvertAll(vObj, x => Convert.ToUInt16(x));
                            pPLCBuffer.ReadNumber = Convert.ToSingle(pValue.Value);
                            break;
                        case "LONG":
                            pPLCBuffer.ReadNumber = Convert.ToInt32(pValue.Value);
                            break;
                        default:
                            if (pValue.Value.ToString().ToLower() == "true" || pValue.Value.ToString().ToLower() == "false")
                            {
                                pPLCBuffer.ReadNumber = Convert.ToInt16(pValue.Value);
                            }
                            else
                            {
                                pPLCBuffer.ReadNumber = (object)pValue.Value;
                            }
                            break;
                    }
                }
                catch (Exception exp)
                {
                    RaiseEvents(exp.Message + "[source>" + exp.Source + "].");
                }
            }
        }

        void DisconnectOPCServer()
        {
            try
            {
                //Unsubscribe();
                if (daServerMgt.IsConnected)
                {
                    daServerMgt.Disconnect();
                    string vMsg = string.Format("Disconnect OPC Server endpoint URL={0} Channel={1} Device={2}.",
                                            endPointURL, opcChannelName, opcDeviceName);
                    RaiseEvents(vMsg);
                }
            }
            catch (Exception exp)
            {
                RaiseEvents("Handled Disconnect exception. Reason: " + exp.Message + "[source>" + exp.Source + "]");
            }
        }

        private int RandomNumber(int pMaxNumber, int pMinNumber)
        {
            //initialize random number generator
            Random r = new Random(System.DateTime.Now.Millisecond);

            //if passed incorrect arguments, swap them
            //can also throw exception or return 0

            if (pMinNumber > pMaxNumber)
            {
                int t = pMinNumber;
                pMinNumber = pMaxNumber;
                pMaxNumber = t;
            }

            return r.Next(pMinNumber, pMaxNumber);
        }

        private void SubscribeToOPCDAServerEvents()
        {
            daServerMgt.ReadCompleted += new DaServerMgt.ReadCompletedEventHandler(daServerMgt_ReadCompleted);
            daServerMgt.WriteCompleted += new DaServerMgt.WriteCompletedEventHandler(daServerMgt_WriteCompleted);
            daServerMgt.DataChanged += new DaServerMgt.DataChangedEventHandler(daServerMgt_DataChanged);
            daServerMgt.ServerStateChanged += new DaServerMgt.ServerStateChangedEventHandler(daServerMgt_ServerStateChanged);
        }

        #region DaServerMgt Event handler
        // ***********************************************************************
        // (Asynchronous) ReadCompleted event handler
        //
        // Check result of read request and update form with returned data.
        // ***********************************************************************
        public void daServerMgt_ReadCompleted(int pTransactionHandle, bool pAllQualitiesGood, bool pNoErrors, ItemValueCallback[] pItemValues)
        {
            //Debug.WriteLine("daServerMgt_ReadCompleted enter");
            //RaiseEvents("daServerMgt_ReadCompleted enter.");

            // We need to forward the callback to the main thread of the application if
            // we access the GUI directly from the callback. It is recommended to do this
            // even if the application is running in the back ground.
            //
            // See Control.Invoke Method (Delegate, Object[]) for a good explanation. Note, we are using BeginInvoke()
            // instead of Invoke() so that the delegate is called asynchronously.

            // Create an instance of the delegate and assign its address of method with same signature as ReadCompletedEventHandler.
            DaServerMgt.ReadCompletedEventHandler RCevHndlr = new DaServerMgt.ReadCompletedEventHandler(ReadCompleted);
            IAsyncResult returnValue;
            object[] RCevHndlrArray = new object[4];
            RCevHndlrArray[0] = pTransactionHandle;
            RCevHndlrArray[1] = pAllQualitiesGood;
            RCevHndlrArray[2] = pNoErrors;
            RCevHndlrArray[3] = pItemValues;
            //returnValue = fMain.BeginInvoke(RCevHndlr, RCevHndlrArray);
            ReadCompleted(pTransactionHandle, pAllQualitiesGood, pNoErrors, pItemValues);
            //Debug.WriteLine("daServerMgt_ReadCompleted exit");
            //RaiseEvents("daServerMgt_ReadCompleted exit.");
        }

        // ***********************************************************************
        // (Asynchronous) ReadCompleted event delegate
        // ***********************************************************************
        public void ReadCompleted(int pTransactionHandle, bool pAllQualitiesGood, bool pNoErrors, ItemValueCallback[] pItemValues)
        {
            //Debug.WriteLine("ReadCompleted enter");
            //RaiseEvents("ReadCompleted enter");
            // Since we are only reading one item in the read we know that there will only be
            // once response. Normally you would want to create a loop and walk through each
            // returned value and update the form accordingly.
            if (pTransactionHandle == transID)
            {
                opcItemValue = pItemValues;
                transID = -1;
            }
            else
            {
                int itemIndex = (int)pItemValues[0].ClientHandle;

                try
                {
                    // Update item values on read success:
                    // (We would loop over items if reading more than one.)
                    if (pItemValues[0].ResultID.Succeeded)
                    {
                        // Set item value (could be NULL if quality goes bad):
                        if (pItemValues[0].Value == null)
                        {
                            //OPCItemValueTextBoxes[itemIndex].Text = "Unknown";
                            object obj = pItemValues[0].Value;

                        }
                        else
                        {
                            //OPCItemValueTextBoxes[itemIndex].Text = itemValues[0].Value.ToString();
                            diagValues[itemIndex] = pItemValues[0];
                        }
                    }

                    // Set item quality (will be bad if ResultID.Succeeded = false):
                    //OPCItemQualityTextBoxes[itemIndex].Text = itemValues[0].Quality.Name;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Handled Async Read Complete exception. Reason: " + ex.Message);

                    // Set item quality to bad:
                    //OPCItemQualityTextBoxes[itemIndex].Text = "OPC_QUALITY_BAD";
                }
            }
            //Debug.WriteLine("ReadCompleted exit");
            //RaiseEvents("ReadCompleted exit");
        }
       
        public void daServerMgt_WriteCompleted(int pTransaction, bool pNoErrors, ItemResultCallback[] pItemResults)
        {
            //Debug.WriteLine("daServerMgt_WriteCompleted enter");
            //RaiseEvents("daServerMgt_WriteCompleted enter");
            // We need to forward the callback to the main thread of the
            // application if we access the GUI directly from the callback. It is
            // recommended to do this even if the application is running in the back ground.
            object[] WCevHndlrArray = new object[3];
            WCevHndlrArray[0] = pTransaction;
            WCevHndlrArray[1] = pNoErrors;
            WCevHndlrArray[2] = pItemResults;
            //fMain.BeginInvoke(new DaServerMgt.WriteCompletedEventHandler(WriteCompleted), WCevHndlrArray);
            WriteCompleted(pTransaction, pNoErrors, pItemResults);
            //Debug.WriteLine("daServerMgt_WriteCompleted exit");
            //RaiseEvents("daServerMgt_WriteCompleted exit");
        }

        // ***********************************************************************
        // (Asynchronous) WriteCompleted event delegate
        // ***********************************************************************
        
        public void WriteCompleted(int pTransaction, bool pNoErrors, ItemResultCallback[] pItemResults)
        {
            //Debug.WriteLine("WriteCompleted enter");
            //RaiseEvents("WriteCompleted enter");
            try
            {
                //~~ Process the call back infomration here.
                if (!pItemResults[0].ResultID.Succeeded)
                {
                    //MessageBox.Show("Async Write Complete failed with error: " + System.Convert.ToString(pItemResults[0].ResultID.Code) + "\r\n" + "Description: " + pItemResults[0].ResultID.Description);
                }
            }
            catch (Exception ex)
            {
                //~~ Handle any exception errors here.
                //MessageBox.Show("Handled Async Write Complete exception. Reason: " + ex.Message);
            }

            //Debug.WriteLine("WriteCompleted exit");
            //RaiseEvents("WriteCompleted exit");
        }

        // ***********************************************************************
        // DataChanged Event handler
        //
        // Update value and quality text boxes. Do not make any other calls into
        // the OPC server from this call back.
        // ***********************************************************************
        public void daServerMgt_DataChanged(int pClientSubscription, bool pAllQualitiesGood, bool pNoErrors, ItemValueCallback[] pItemValues)
        {
            lock (thrLock)
            {
                //Debug.WriteLine("daServerMgt_DataChanged enter");
                //RaiseEvents("daServerMgt_DataChanged enter");
                // We need to forward the callback to the main thread of the application if we access the GUI directly from the callback. 
                //It is recommended to do this even if the application is running in the back ground.
                object[] DCevHndlrArray = new object[4];
                DCevHndlrArray[0] = pClientSubscription;
                DCevHndlrArray[1] = pAllQualitiesGood;
                DCevHndlrArray[2] = pNoErrors;
                DCevHndlrArray[3] = pItemValues;
                //fMain.BeginInvoke(new DaServerMgt.DataChangedEventHandler(DataChanged), DCevHndlrArray);
                DataChanged(pClientSubscription, pAllQualitiesGood, pNoErrors, pItemValues);
                //Debug.WriteLine("daServerMgt_DataChanged exit");
                //RaiseEvents("daServerMgt_DataChanged exit");
            }
        }

        // ***********************************************************************
        // DataChanged event delegate
        // ***********************************************************************

        public void DataChanged(int pClientSubscription, bool pAllQualitiesGood, bool pNoErrors, ItemValueCallback[] pItemValues)
        {
            //RaiseEvents("DataChanged enter");
            DateTime vDate = DateTime.Now;
            int itemIndex=0;
            try
            {
                // if we were dealing with multiple subscriptions, we would want to use
                // the clientSubscripion parameter to determine which subscription this
                // event pertains to. In this simple example, we will simply validate the
                // handle.
                if (activeClientSubscriptionHandle == pClientSubscription)
                {
                    // Loop over values returned. You should not assume that data
                    // for all items enrolled in a subscription will be included
                    // in every data changed event. In actuality, the number of
                    // values will likely vary each time.
                    foreach (ItemValueCallback itemValue in pItemValues)
                    {
                        
                        // Get the item handle. We used the item's index into the
                        // control arrays in this example.
                        itemIndex = (int)itemValue.ClientHandle;

                        // Update value control (Could be NULL if quaulity goes bad):
                        if (opcItemValue[itemIndex].Quality.IsGood)
                        {
                            ReadOPCItem(ref plcBuffer[itemIndex], opcItemValue[itemIndex]);
                            plcBuffer[itemIndex].ItemQualityDesc = opcItemValue[itemIndex].Quality.FullCode + "[" + opcItemValue[itemIndex].Quality.Name + "]"; //"Unknow";
                            plcBuffer[itemIndex].TimeStamp = opcItemValue[itemIndex].TimeStamp;
                            plcBuffer[itemIndex].ItemQuailty = opcItemValue[itemIndex].Quality.FullCode;
                            //UpdateToDatabase(plcBuffer[itemIndex]);
                        }
                        else
                        {
                            ReadOPCItem(ref plcBuffer[itemIndex], opcItemValue[itemIndex]);
                            plcBuffer[itemIndex].ItemQualityDesc = opcItemValue[itemIndex].Quality.FullCode + "[" + opcItemValue[itemIndex].Quality.Name + "]"; //"Unknow";
                            plcBuffer[itemIndex].TimeStamp = opcItemValue[itemIndex].TimeStamp;
                            plcBuffer[itemIndex].ItemQuailty = opcItemValue[itemIndex].Quality.FullCode;
                            //UpdateToDatabase(plcBuffer[itemIndex]);
                        }

                        if (plcDiag || (vDate - refreshDisplay).Minutes >= 1)
                        {
                            //refreshDisplay = DateTime.Now;
                            DataChangeMessage(plcBuffer[itemIndex]);
                        }
                        // Update quality control:
                        //OPCItemQualityTextBoxes[itemIndex].Text = itemValue.Quality.Name;
                    }
                }
                if (plcDiag || (vDate - refreshDisplay).Minutes >= 1)
                {
                    refreshDisplay = DateTime.Now;
                    //DataChangeMessage(plcBuffer[itemIndex]);
                }
            }
            catch (Exception ex)
            {
                //~~ Handle any exception errors here.
                RaiseEvents("Handled Data Changed exception. Reason: " + ex.Message + Environment.NewLine +
                            DataChangeMessage(plcBuffer[itemIndex]));
            }

            //RaiseEvents("DataChanged exit.");
        }

        // ***********************************************************************
        // ServerStateChanged event handler
        //
        // Monitor the connection state of OPC server. We will show a message box
        // when a problem with the connection is detected. In a real world
        // application, you would probably use an event log instead.
        // ***********************************************************************
        public void daServerMgt_ServerStateChanged(int pClientHandle, ServerState pServerState)
        {
            RaiseEvents("daServerMgt_ServerStateChanged enter");

            // We need to forward the callback to the main thread of the
            // application if we access the GUI directly from the callback. It is
            // recommended to do this even if the application is running in the back ground.
            object[] SSCevHndlrArray = new object[2];
            SSCevHndlrArray[0] = pClientHandle;
            SSCevHndlrArray[1] = pServerState;
            //fMain.BeginInvoke(new DaServerMgt.ServerStateChangedEventHandler(ServerStateChanged), SSCevHndlrArray);
            ServerStateChanged(pClientHandle, pServerState);

            RaiseEvents("daServerMgt_ServerStateChanged exit");
        }

        // ***********************************************************************
        // ServerStateChanged event delegate
        // ***********************************************************************
        public void ServerStateChanged(int pClientHandle, ServerState pServerState)
        {
            RaiseEvents("ServerStateChanged enter");

            try
            {
                //~~ Process the call back infomration here.
                switch (pServerState)
                {
                    case ServerState.ERRORSHUTDOWN:
                        //Unsubscribe();
                        //DisconnectOPCServer();
                        RaiseEvents("The server is shutting down. The client has automatically disconnected.");
                        break;

                    case ServerState.ERRORWATCHDOG:
                        // server connection has failed. ClientAce will attempt to reconnect to the server 
                        // because connectInfo.RetryAfterConnectionError was set true when the Connect method was called.
                        RaiseEvents("Server connection has been lost. Client will keep attempting to reconnect.");
                        break;

                    case ServerState.CONNECTED:
                        RaiseEvents("ServerStateChanged, connected.");
                        UpdatePLCConnect(1);
                        //Subsctibe();
                        break;

                    case ServerState.DISCONNECTED:
                        RaiseEvents("ServerStateChanged, disconnected.");
                        UpdatePLCConnect(0);
                        break;

                    default:
                        RaiseEvents("ServerStateChanged, undefined state found.");
                        break;
                }
            }
            catch (Exception ex)
            {
                //~~ Handle any exception errors here.
                RaiseEvents("Handled Server State Changed exception. Reason: " + ex.Message);
            }

            RaiseEvents("ServerStateChanged exit");
        }

        #endregion

        private void Subsctibe()
        {
            // Define parameters for Subscribe method:

            // The client subscription handle is described above (see global
            // activeServerSubscriptionHandle.) We can use an arbitrary value
            // in this example since we will be dealing with only one subscription.
            // if we were managing multiple subscriptions, we would want to use
            // unique and meaningful handles.
            int clientSubscriptionHandle = 1;

            // The revisedUpdateRate parameter is the actual update rate that the
            // server will be using.
            int revisedUpdateRate;

            // The updateRate parameter is used to tell the server how fast we
            // would like to see data updates. This translates roughly into how
            // fast the server should poll the items enrolled in this subscription.
            // This is a REQUESTED rate. The server may not be able to honor this
            // request. This number is measured in milliseconds.
            int updateRate = updateRateIPS;

            // The deadband parameter specifies the minimum deviation needed
            // to be considered a change of value. It is expressed as a percentage
            // (0 - 100). In a real world application, you should validate text
            // first.
            Single deadBand = 0;

            try
            {
                // Save the active client subscription handle for use in 
                // DataChanged events:
                activeClientSubscriptionHandle = clientSubscriptionHandle;

                daServerMgt.Subscribe(clientSubscriptionHandle, true, updateRate, out revisedUpdateRate, deadBand,
                                        ref itemIdentifiers, out activeServerSubscriptionHandle);
            }
            catch (Exception exp)
            {
                RaiseEvents("Handled Subscribe exception. Reason: " + exp.Message);
            }
        }

        private void Unsubscribe()
        {
            // Call SubscriptionCancel API method:
            // (Note, we are using the server subscription handle here.)
            try
            {
                daServerMgt.SubscriptionCancel(activeServerSubscriptionHandle);
            }
            catch (Exception e)
            {
                RaiseEvents("Handled SubscriptionCancel exception. Reason: " + e.Message);
            }
        }

        #endregion

        int GetReadLength(int pDataType, int pLength)
        {
            int vReadLength=0;
            switch (pDataType)
            {
                case (int)_DataType.BOOLEAN:
                    vReadLength = pLength * 8;
                    break;
                default:
                    vReadLength = pLength;
                    break;
            }
            //if (pDataType == EnumString.GetStringValue(_DataType.BOOLEAN))
            //    vReadLength = pLength * 8;
            //else
            //    vReadLength = pLength;

            return vReadLength;
        }
        
        public void SetComport()
        {
            //atgPort = fMain.ATGComport[processId];
            //atgPort.InitialPort();
            //atgPort.StartThread();
        }

        #region Diagnostic Process
        bool plcDiag;
        Thread thrDiag;
        public void DiagnasticPLC(bool pDiag)
        {
            plcDiag = pDiag;
            if (plcDiag == true)
            {
                KepwareInitialDiagOPCServer();
                thrDiag = new Thread(this.DiagProcess);
                thrDiag.Start();
                Thread.Sleep(1000);
            }
        }

        void DiagProcess()
        {
            while (plcDiag && !thrShutdown)
            {
                KepwareDiagOPCServer();
                //string[] v = DiagnosticReceive();
                Thread.Sleep(1000);
            }
        }

        public string[] DiagnosticSend()
        {
            string[] vRet;
            //DiagKepwareOPCServer();
            //vRet =diagValues[0].TimeStamp + ">" + diagValues[0].Value.ToString();
            try
            {
                vRet = new string[diagIdentifiers.Length + 1];
                vRet[0] = DateTime.Now.ToString() + "> Send";
                for (int i = 1; i < diagIdentifiers.Length; i++)
                {
                    if(diagValues[i] != null)
                        vRet[i] = diagIdentifiers[i].ItemName + "=" + diagValues[i].Value.ToString();
                }
            }
            catch (Exception exp)
            {
                vRet = new string[3];
                vRet[0] = DateTime.Now.ToString() + "> Send";
                vRet[1] = exp.Data.ToString();
                vRet[2] = exp.Message;
            }
            return vRet;
        }

        public string[] DiagnosticReceive()
        {
            string[] vRet;
            try
            {
                vRet = new string[plcBuffer.Length + 1];
                vRet[0] = DateTime.Now.ToString() + "> Receive";
                for (int i = 1; i < plcBuffer.Length; i++)
                {
                    //vRet[i] = plcBuffer[i].TimeStamp.ToString() + ">" + itemIdentifiers[i].ItemName + "".PadRight(3, ' ') + plcBuffer[i].ItemQualityDesc + "".PadRight(3, ' ') + "".PadRight(plcBuffer[i].TimeStamp.ToString().Length, ' ');
                    //vRet[i] += "Value[" + plcBuffer[i].ReadNumber + "]";
                    vRet[i] = string.Format("{0,-20} > {1,-60} Value={2,-15} Quality={3,-20}"
                        , plcBuffer[i].TimeStamp.ToString(), itemIdentifiers[i].ItemName, plcBuffer[i].ReadNumber, plcBuffer[i].ItemQualityDesc);
                }
            }
            catch (Exception exp)
            {
                vRet = new string[3];
                vRet[0] = DateTime.Now.ToString() + "> Receive";
                vRet[1] = exp.Data.ToString();
                vRet[2] = exp.Message;
            }
            return vRet;
        }
        #endregion

        #region Change process or comport
        void ChangeProcess()
        {
            //return;
            //Thread.Sleep(3000);
            bool vRet = false;
            vRet = CheckItemBadValueAll();
            var vDiff = (DateTime.Now - chkResponse).TotalSeconds;
            if ((vRet) && (vDiff >= timeOutChangeProcess))
            {
                chkResponse = DateTime.Now;
                thrShutdown = true;
                DisconnectOPCServer();
                daServerMgt.Dispose();
                //ChangeComport();
                RaiseEvents("Communication changed.");
                fMain.ChangeProcess();
            }   
        }

        void ChangeComport()
        {
            return;
            int vControlComport;
            if (processId == 1)
                vControlComport = 0;
            else
                vControlComport = 1;

            string strSQL = "begin steqi.P_CHANGE_COMPORT_ATG(" +
                             vControlComport.ToString() +
                             ");end;";

            fMain.OraDb.ExecuteSQL(strSQL);
        }
        
        public bool IsThreadAlive
        {
            get { return thrMain.IsAlive; }
        }
        #endregion

        #region display envent/message to GUI
        Queue<string> queueMsg = new Queue<string>();
        void DisplayMessageGUI()
        {
            string vMsg = "";
            while (true)
            {
                if (thrShutdown)
                    break;

                while (queueMsg.Count > 0)
                {
                    if (thrShutdown)
                        break;
                    try
                    {
                        vMsg = queueMsg.Dequeue();
                        RaiseEvents(vMsg);
                    }
                    catch (Exception exp)
                    { }
                    //Thread.Sleep(100);
                }
            }
        }
        #endregion
    }
}
