using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.IO.Ports;

namespace PPLC
{
    public partial class frmMain : Form
    {
        #region Constant, Struct and Enum
        enum _StepProcess : int
        {
            InitialDatabase = 0
                ,
            InitialPLC = 1
                ,
            InitialComportPLC = 2
                ,
            InitialDataGrid = 3
                ,
            InitialClassEvent = 4
                ,
            ChangeProcess = 5
        }

        #endregion
        
        public Comport[] ATGComport;
        SerialPort[] atgSerialPort;
        public Logfile LogFile = new Logfile();

        private PLCProcess[] plcProcess;
        public Database OraDb;
        public IniLib.CINI iniFile = new IniLib.CINI();
        string arg;
        string[] args;
        string scanID;
        string logFileName;
        int processID;

        _StepProcess mainStepProcess;

        private bool IsSingleInstance()
        {
            foreach (Process process in Process.GetProcesses())
            {
                if (process.MainWindowTitle == this.Text)
                    return false;
            }
            return true;
        }

        public frmMain()
        {
            InitializeComponent();

            try
            {
                args = Environment.GetCommandLineArgs();
                arg = args[1];
            }
            catch (Exception exp)
            {
                arg = "1";
                //AddListBoxItem(mBayNo);
            }
            this.Text = iniFile.INIRead(Directory.GetCurrentDirectory() + "\\AppStartup.ini", arg, "TITLE");
            scanID = iniFile.INIRead(Directory.GetCurrentDirectory() + "\\AppStartup.ini", arg, "SCANID");
            logFileName = this.Text;
            lblVersion.Text = "PLC Triconex v1.0";
            if (!IsSingleInstance())
            {
                MessageBox.Show("Another instance of this app is running.", this.Text);
                //Application.Exit();
                System.Environment.Exit(1);
            }
            //StartThread();
            AddListBox = DateTime.Now + "><------------Application Start------------->";

        }

        #region Thread
        bool thrConnect;
        bool thrShutdown;
        bool thrRunning;

        Thread thrMain;

        private void StartThread()
        {
            //System.Threading.Thread.Sleep(1000);
            thrMain = new Thread(this.RunProcess);
            thrMain.Name = this.Text;
            thrMain.Start();
        }

        private void RunProcess()
        {
            thrRunning = true;
            System.Threading.Thread.Sleep(1000);
            if (mainStepProcess != _StepProcess.ChangeProcess)
            {
                InitialDataBase();
                mainStepProcess = _StepProcess.InitialDatabase;
            }
            while (thrRunning)
            {
                try
                {
                    if (thrShutdown)
                        return;
                    switch (mainStepProcess)
                    {
                        case _StepProcess.InitialDatabase:
                            if (OraDb.ConnectStatus())
                            {
                                mainStepProcess = _StepProcess.InitialPLC;
                                //AddListBoxItem(mStepProcess.ToString());
                            }
                            Thread.Sleep(500);
                            break;
                        case _StepProcess.InitialPLC:
                            //thrRunning = true;
                            //Thread.Sleep(500);
                            if (InitialPLC())
                            {
                                //if (InitialComportATG())
                                //{
                                InitialCurrentIPS_ID();
                                    mainStepProcess = _StepProcess.InitialDataGrid;
                                    plcProcess[processID].StartThread();
                                //}
                                //AddListBoxItem(mStepProcess.ToString());
                            }

                            break;
                        case _StepProcess.InitialDataGrid:
                            //thrRunning = false;
                            InitialDataGrid();
                            //AddListBoxItem(mStepProcess.ToString());
                            mainStepProcess = _StepProcess.InitialClassEvent;
                            break;
                        case _StepProcess.InitialClassEvent:
                            //thrRunning = false;
                            thrShutdown = true;
                            InitialClassEvent();
                            break;
                        case _StepProcess.ChangeProcess:
                            thrShutdown = true;
                            Thread.Sleep(300);
                            plcProcess[processID].StartThread();
                            break;
                    }
                    DisplayDateTime();
                    Thread.Sleep(300);
                }
                catch (Exception exp)
                { AddListBoxItem(DateTime.Now + ">" + exp.Message + "[" + exp.Source + "-" + mainStepProcess.ToString() + "]"); }
                //finally
                //{
                //    mShutdown = true;
                //    mRunning = false;
                //}
            }
        }
        #endregion

        #region ListboxItem

        public void DisplayMessage(string pFileName, string pMsg)
        {
            if (this.lstMain.InvokeRequired)
            {
                // This is a worker thread so delegate the task.
                if (lstMain.Items.Count > 1000)
                {
                    //lstMain.Items.Clear();
                    this.Invoke((Action)(() => lstMain.Items.Clear()));
                }

                this.lstMain.Invoke(new DisplayMessageEventHandler(this.DisplayMessage), pFileName, pMsg);
            }
            else
            {
                // This is the UI thread so perform the task.
                if (pMsg != null)
                {
                    if (lstMain.Items.Count > 1000)
                    {
                        //lstMain.Items.Clear();
                        this.Invoke((Action)(() => lstMain.Items.Clear()));
                    }

                    this.lstMain.Items.Insert(0, pMsg);
                    //logfile.WriteLog("System", item.ToString());
                    //PLog.WriteLog(pFileName, iMsg);
                }
            }
        }

        private delegate void DisplayMessageEventHandler(string pFileName, string pMsg);

        public object AddListBox
        {
            set
            {
                AddListBoxItem(value);
            }
        }

        private delegate void AddListBoxItemEventHandler(object pItem);

        private void AddListBoxItem(object pItem)
        {
            if (this.lstMain.InvokeRequired)
            {
                // This is a worker thread so delegate the task.
                if (lstMain.Items.Count > 1000)
                {
                    //lstMain.Items.Clear();
                    this.Invoke((Action)(() => lstMain.Items.Clear()));
                }

                this.lstMain.Invoke(new AddListBoxItemEventHandler(this.AddListBoxItem), pItem);
            }
            else
            {
                // This is the UI thread so perform the task.
                if (pItem != null)
                {
                    if (lstMain.Items.Count > 1000)
                    {
                        //lstMain.Items.Clear();
                        this.Invoke((Action)(() => lstMain.Items.Clear()));
                    }

                    this.lstMain.Items.Insert(0, (processID + 1).ToString() + "-" + logFileName + ">" + pItem);
                    //logfile.WriteLog("System", item.ToString());
                    LogFile.WriteLog(logFileName, (processID + 1).ToString() + "-" + logFileName + ">" + pItem.ToString());
                }
            }
        }

        private delegate void ClearListBoxItemDelegate();
        private void ClearListBoxItem()
        {
            if (this.lstMain.InvokeRequired)
            {
                this.lstMain.Invoke(new ClearListBoxItemDelegate(ClearListBoxItem));

            }
            else
            {
                this.lstMain.Items.Clear();
            }

        }
        #endregion

        #region DataGrid
        private delegate void AddDataGridItemEventHandler(int pRow, int pCol, object pValue);
        private delegate void AddDataGridRowsEventHandler(int pRows);

        private void AddDataGridRows(int pRows)
        {
            if (this.dataGridView1.InvokeRequired)
            {
                // This is a worker thread so delegate the task.

                this.dataGridView1.Invoke(new AddDataGridRowsEventHandler(this.AddDataGridRows), pRows);
            }
            else
            {
                // This is the UI thread so perform the task.
                if (pRows != 0)
                {
                    dataGridView1.Rows.Add(pRows);
                }
            }
        }
        private void AddDataGridItem(int pRow, int pCol, object pValue)
        {
            if (this.dataGridView1.InvokeRequired)
            {
                // This is a worker thread so delegate the task.

                this.dataGridView1.Invoke(new AddDataGridItemEventHandler(this.AddDataGridItem), pRow, pCol, pValue);
            }
            else
            {
                // This is the UI thread so perform the task.
                if (pRow >= 0)
                {
                    //dataGridView1.Rows[1].Cells[0].Value = "bay_no";
                    dataGridView1.Rows[pRow].Cells[pCol].Value = pValue;
                }
            }
        }
        #endregion

        #region Combobox
        private delegate void AddComboboxItemEvenHandler(object pItem);

        public void AddComboboxItem(object pItem)
        {
            if (this.cboComport.InvokeRequired)
            {
                // This is a worker thread so delegate the task.

                this.cboComport.Invoke(new AddComboboxItemEvenHandler(this.AddComboboxItem), pItem);
            }
            else
            {
                // This is the UI thread so perform the task.
                cboComport.Items.Add(pItem);
            }
        }
        #endregion

        #region MynotifyIcon
        private void FormResize()
        {
            //if (this.WindowState == FormWindowState.Minimized)
            //{
            //    mynotifyIcon1.Icon = this.Icon;
            //    mynotifyIcon1.Visible = true;
            //    mynotifyIcon1.BalloonTipText = this.Text;
            //    mynotifyIcon1.ShowBalloonTip(500);
            //    this.Hide();
            //}
        }
        private void mynotifyIcon1_Click(object sender, MouseEventArgs e) 
        {
            //mynotifyIcon1.Visible = false;
            //this.Show();
            //this.WindowState = FormWindowState.Normal;
        }
        #endregion

        #region Class Events
        void InitialATGEventHandler()
        {
            //ATGProcess.ATGEventsHaneler handler1 = new ATGProcess.ATGEventsHaneler(WriteEventsHandler);
            //atgProcess[0].OnATGEvents += handler1;
        }

        void InitialComportEventHandler()
        {
            Comport.ComportEventsHandler hander1 = new Comport.ComportEventsHandler(WriteEventsHandler);
            for (int i = 0; i < ATGComport.Length; i++)
            {
                ATGComport[i].OnComportEvents += hander1;
            }
        }

        void WriteEventsHandler(object pSender, string pMessage)
        {
            AddListBoxItem(pMessage);
            //mLog.WriteLog(mLogFileName, message);
        }
        #endregion

        #region Main Step Process
        private void InitialDataBase()
        {
            OraDb = new Database(this);
        }

        private bool InitialPLC()
        {
            //string strSQL = "select t.plc_id,t.plc_name,t.opc_server_name,t.opc_channel_name,t.opc_device_name," +
            //                "'Addr=' || t.plc_name || ', ' || t.comp_id || ':' || t.comport_no  as Description" +
            //                " from tas.VIEW_OPC_CONFIG_INITAIL t " +
            //                " order by t.plc_id";
            string strSQL = "select t.* from tas.VIEW_IPS_CONFIG_OPC t";
            DataSet vDataset = null;
            DataTable dt;
            bool vRet = false;
            try
            {
                if (OraDb.OpenDyns(strSQL, "TableName", ref vDataset))
                {
                    dt = vDataset.Tables["TableName"];
                    plcProcess = new PLCProcess[dt.Rows.Count];
                    //dataGridView1.Rows.Add(dt.Rows.Count - 1);
                    //AddDataGridRows(dt.Rows.Count - 1);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        plcProcess[i] = new PLCProcess(this
                                                    ,Convert.ToInt16(dt.Rows[i]["SCAN_ID"].ToString())
                                                    ,Convert.ToInt16(dt.Rows[i]["SCAN_ID"].ToString())
                                                    ,dt.Rows[i]["SCAN_NAME"].ToString()
                                                    ,dt.Rows[i]["ENDPOINT_URL"].ToString()
                                                    ,dt.Rows[i]["OPC_CHANNEL_NAME"].ToString()
                                                    ,dt.Rows[i]["OPC_DEVICE_NAME"].ToString()
                                                    ,dt.Rows[i]["PREFIX_TAG"].ToString()
                                                    );
                        AddComboboxItem(dt.Rows[i]["SCAN_ID"].ToString() + "-" + dt.Rows[i]["SCAN_NAME"].ToString());
                    }
                    vRet = true;
                }
                //vRet= true;
            }
            catch (Exception exp)
            { LogFile.WriteErrLog(exp.Message); }
            vDataset = null;
            dt = null;
            return vRet;
        }
        
        bool InitialComportATG()
        {
            return true;
            string strSQL = "select" +
                            " t.comp_id,t.comport_no,t.comport_setting" +
                            " from tas.VIEW_ATG_CONFIG_COMPORT t " +
                            " order by t.atg_id";
            DataSet vDataset = null;
            DataTable dt;
            bool vRet = false;
            try
            {
                if (OraDb.OpenDyns(strSQL, "TableName", ref vDataset))
                {
                    dt = vDataset.Tables["TableName"];
                    ATGComport = new Comport[dt.Rows.Count];
                    atgSerialPort = new SerialPort[dt.Rows.Count];
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        atgSerialPort[i]=new SerialPort();
                        ATGComport[i] = new Comport(this, ref atgSerialPort[i],Convert.ToDouble(dt.Rows[i]["comp_id"].ToString()));
                        ATGComport[i].InitialPort();
                    }
                }
                vRet = true;
            }
            catch (Exception exp)
            { LogFile.WriteErrLog(exp.Message); }
            vDataset = null;
            dt = null;
            return vRet;
        }

        private bool InitialDataGrid()
        {
            string strSQL = "select t.* from tas.VIEW_IPS_CONFIG_OPC t";
            DataSet vDataset = null;
            DataTable dt;
            bool vRet = false;
            try
            {
                if (OraDb.OpenDyns(strSQL, "TableName", ref vDataset))
                {
                    dt = vDataset.Tables["TableName"];
                    //dataGridView1.Rows.Add(dt.Rows.Count - 1);
                    AddDataGridRows(dt.Rows.Count - 1);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        //dataGridView1.Rows[i].Cells[0].Value = dt.Rows[i]["bay_no"];
                        //dataGridView1.Rows[i].Cells[1].Value = dt.Rows[i]["card_reader_name"];
                        //dataGridView1.Rows[i].Cells[2].Value = dt.Rows[i]["Description"];
                        AddDataGridItem(i, 0, i + 1);
                        AddDataGridItem(i, 1, "IPS");
                        AddDataGridItem(i, 2, dt.Rows[i]["SCAN_NAME"]);
                    }
                }
                vRet = true;
                dataGridView1.ClearSelection();
            }
            catch (Exception exp)
            { }
            vDataset = null;
            dt = null;
            return vRet;
        }

        private void InitialClassEvent()
        {
            //InitialATGEventHandler();
            //InitialComportEventHandler();
        }

        public void ChangeProcess()
        {
            processID += 1;
            //if (processID >= cboComport.Items.Count)
            if(processID >= plcProcess.Length)
            {
                processID = 0;
            }
            //plcProcess[processID].StartThread();
            thrShutdown = false;
            ChangeIPSProcessID(processID + 1);
            mainStepProcess = _StepProcess.ChangeProcess;
            StartThread();
        }

        void InitialCurrentIPS_ID()
        {
            DataSet vDataset = new DataSet();
            DataTable dt;
            string vSql = "select tas.F_CURRENT_IPS_ID as id from dual";
            if (OraDb.OpenDyns(vSql, "TableName", ref vDataset))
            {
                dt = vDataset.Tables[0];
                processID = Convert.ToInt16(dt.Rows[0]["id"].ToString());
            }
        }

        public void ChangeIPSProcessID(int pProcessId)
        {
            string vSQL = "";

            vSQL = "begin tas.P_SET_IPS_ACTIVE(" +
                pProcessId +
                ");end;";
            OraDb.ExecuteSQL(vSQL);
        }
        #endregion

        private void DisplayDateTime()
        {
            toolStripStatusLabel1.Text = "Database connect = " + OraDb.ConnectStatus().ToString() +
                                                "[" + OraDb.ConnectServiceName() + "]" +
                                                "   [Date Time : " + DateTime.Now + "]";
        }

        private void DiagnosticComport(bool pDiag)
        {
            try
            {
                if (pDiag)
                {
                    // to do some thing
                    txtSend.Lines = plcProcess[cboComport.SelectedIndex].DiagnosticSend();
                    //txtSend.WordWrap = false;
                    txtRecv.Lines = plcProcess[cboComport.SelectedIndex].DiagnosticReceive();
                    //txtRecv.WordWrap = false;
                    LogFile.WriteComportLog(cboComport.Text, txtSend.Lines);
                    LogFile.WriteComportLog(cboComport.Text, txtRecv.Lines);
                }
                else
                {
                    //txtRecv.Text = txtRecv.Text;
                    //txtSend.Text = txtSend.Text;
                }
            }
            catch (Exception exp)
            { }
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            FormResize();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                toolStripStatusLabel1.Text = "Database connect = " + OraDb.ConnectStatus().ToString() +
                                                "[" + OraDb.ConnectServiceName() + "]" +
                                                "   [Date Time : " + DateTime.Now + "]";

            }
            catch (Exception exp)
            { }
            //DisplayMessageCardReader();
            //DisplayCardReader();
            DiagnosticComport(chkDiag.Checked);
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            AddListBox = DateTime.Now + "><------------Application Stop-------------->";
            thrShutdown = true;
            iniFile = null;
            //VisibleObject(true);
            timer1.Enabled = false;
            Application.DoEvents();
            this.Cursor = Cursors.WaitCursor;
            //System.Threading.Thread.Sleep(500);
            if(ATGComport!=null)
            {
                foreach (Comport p in ATGComport)
                {
                    p.Dispose();
                }
            }
            
            System.Threading.Thread.Sleep(300);
            if (plcProcess != null)
            {
                for (int i = 0; i < plcProcess.Length; i++)
                {
                    if (plcProcess[i] != null)
                    {
                        plcProcess[i].StopThread();
                        Thread.Sleep(100);
                        plcProcess[i].Dispose();
                    }
                }
            }
            OraDb.Close();
            OraDb.Dispose();         
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            StartThread();
            timer1.Enabled = true;
        }

        private void lstMain_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            MessageBox.Show(lstMain.SelectedItem.ToString(),this.Text,MessageBoxButtons.OK);
        }

        private void chkDiag_CheckedChanged(object sender, EventArgs e)
        {
            foreach (PLCProcess p in plcProcess)
            {
                p.DiagnasticPLC(false);
            }
            if (cboComport.SelectedIndex > -1)
            {
                plcProcess[cboComport.SelectedIndex].DiagnasticPLC(chkDiag.Checked);
            }
        }

        private void cboComport_SelectedIndexChanged(object sender, EventArgs e)
        {
            chkDiag_CheckedChanged(null, null);
        }

    }
}
