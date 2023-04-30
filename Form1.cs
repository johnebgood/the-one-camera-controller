/****************************************************************************
While the underlying libraries are covered by LGPL, this sample is released 
as public domain.  It is distributed in the hope that it will be useful, but 
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY 
or FITNESS FOR A PARTICULAR PURPOSE.  

Written by oohansen@gmail.com
*****************************************************************************/

using DirectShowLib;
using Microsoft.Win32;
using Newtonsoft.Json;
using SharpOSC;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TheOneCameraControler
{
    public class Form1 : System.Windows.Forms.Form
    {
        private Button buttonDump;
        private Button buttonStopServer;
        private Button buttonStartServer;
        private Button buttonUp;
        private Button buttonDown;
        private Button buttonLeft;
        private Button buttonRight;
        private Button buttonCenter;
        private Button buttonRightLimit;
        private Button buttonLeftLimit;
        private Button buttonUpLimit;
        private Button buttonDownLimit;
        private Button buttonZoomOutFull;
        private Button buttonZoomInFull;
        private Button buttonZoomOut;
        private Button buttonZoomIn;
        private Button buttonFastLeft;
        private Button buttonFastRight;
        private Button buttonFastDown;
        private Button buttonFastUp;
        private TextBox textPort;
        private Label label1;
        private Label label2;
        public TextBox textLastMessage;
        private DataGridView cameraGridView;
        private DataGridViewTextBoxColumn CameraNumber;
        private DataGridViewTextBoxColumn SystemCameraNumber;
        private DataGridViewTextBoxColumn Camera;
        private DataGridViewTextBoxColumn DevicePath;
        private Button PurgeRedItemsButton;

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        //A (modified) definition of OleCreatePropertyFrame found here: http://groups.google.no/group/microsoft.public.dotnet.languages.csharp/browse_thread/thread/db794e9779144a46/55dbed2bab4cd772?lnk=st&q=[DllImport(%22olepro32.dll%22)]&rnum=1&hl=no#55dbed2bab4cd772
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern int OleCreatePropertyFrame(
            IntPtr hwndOwner,
            int x,
            int y,
            [MarshalAs(UnmanagedType.LPWStr)] string lpszCaption,
            int cObjects,
            [MarshalAs(UnmanagedType.Interface, ArraySubType=UnmanagedType.IUnknown)]
            ref object ppUnk,
            int cPages,
            IntPtr lpPageClsID,
            int lcid,
            int dwReserved,
            IntPtr lpvReserved);

        public IBaseFilter theDevice = null;
        public string theDevicePath = "";
        OSCReceiver oscReceiver = null;

        public Form1()
        {
            InitializeComponent();

            if (IsDarkModeEnabled())
            {
                ApplyDarkTheme(this);
            }

            cameraGridView.CellMouseDown += CameraGridView_CellMouseDown;
            cameraGridView.MouseDown += CameraGridView_MouseDown;
            cameraGridView.MouseMove += CameraGridView_MouseMove;
            cameraGridView.DragOver += CameraGridView_DragOver;
            cameraGridView.DragDrop += CameraGridView_DragDrop;

            int i = 1;
            //enumerate Video Input filters and add them to cameraGridView
            foreach (DsDevice device in DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice))

            {
                object source = null;

                try
                {
                    Guid iid = typeof(IBaseFilter).GUID;
                    device.Mon.BindToObject(null, null, ref iid, out source);
                }
                catch (Exception ex)
                {
                    continue;
                }

                cameraGridView.Rows.Add(new string[] { i.ToString(), i.ToString(), device.Name, device.DevicePath });

                theDevice = (IBaseFilter)source;
                theDevicePath = device.DevicePath;
                i++;
                //break;
            }

            //LoadDataGridViewFromJson(cameraGridView, "config.json");
            LoadAndCompareData(cameraGridView, "config.json");
            UpdateSystemCameraNumber(cameraGridView);
            SelectDeviceByCameraNumber(1);
            StartServer();
        }
        private void ApplyDarkTheme(Control parentControl)
        {
            // Set the dark theme colors for the parent control
            parentControl.BackColor = Color.FromArgb(30, 30, 30);
            parentControl.ForeColor = Color.White;

            // Apply the dark theme to all child controls
            foreach (Control childControl in parentControl.Controls)
            {
                if (childControl is DataGridView dataGridView)
                {
                    ApplyDarkThemeToDataGridView(dataGridView);
                }
                else
                {
                    ApplyDarkTheme(childControl);
                }
            }
        }

        private void ApplyDarkThemeToDataGridView(DataGridView dataGridView)
        {
            dataGridView.EnableHeadersVisualStyles = false;

            // Set DataGridView general colors
            dataGridView.BackgroundColor = Color.FromArgb(30, 30, 30);
            dataGridView.ForeColor = Color.White;
            dataGridView.BorderStyle = BorderStyle.None;

            // Set DataGridView header colors
            DataGridViewCellStyle headerStyle = new DataGridViewCellStyle();
            headerStyle.BackColor = Color.FromArgb(50, 50, 50);
            headerStyle.ForeColor = Color.White;
            headerStyle.Font = dataGridView.Font;
            dataGridView.ColumnHeadersDefaultCellStyle = headerStyle;
            dataGridView.RowHeadersDefaultCellStyle = headerStyle;

            // Set DataGridView row and cell colors
            DataGridViewCellStyle cellStyle = new DataGridViewCellStyle();
            cellStyle.BackColor = Color.FromArgb(60, 60, 60);
            cellStyle.ForeColor = Color.White;
            cellStyle.SelectionBackColor = Color.FromArgb(80, 80, 80);
            cellStyle.SelectionForeColor = Color.White;
            dataGridView.DefaultCellStyle = cellStyle;

            // Set DataGridView grid and row separator colors
            dataGridView.GridColor = Color.FromArgb(80, 80, 80);
            dataGridView.RowHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(80, 80, 80);
            dataGridView.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(80, 80, 80);

            // Set DataGridView alternating row colors
            DataGridViewCellStyle alternatingCellStyle = new DataGridViewCellStyle();
            alternatingCellStyle.BackColor = Color.FromArgb(50, 50, 50);
            alternatingCellStyle.ForeColor = Color.White;
            alternatingCellStyle.SelectionBackColor = Color.FromArgb(80, 80, 80);
            alternatingCellStyle.SelectionForeColor = Color.White;
            dataGridView.AlternatingRowsDefaultCellStyle = alternatingCellStyle;
        }



        private bool IsDarkModeEnabled()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    if (key != null)
                    {
                        object registryValueObject = key.GetValue("AppsUseLightTheme");
                        if (registryValueObject != null)
                        {
                            int registryValue = (int)registryValueObject;
                            return registryValue == 0; // Dark mode is enabled if the value is 0
                        }
                    }
                }
            }
            catch
            {
                // In case of any error, fall back to the default light theme
            }

            return false;
        }


        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.buttonDump = new System.Windows.Forms.Button();
            this.buttonStopServer = new System.Windows.Forms.Button();
            this.buttonStartServer = new System.Windows.Forms.Button();
            this.buttonUp = new System.Windows.Forms.Button();
            this.buttonDown = new System.Windows.Forms.Button();
            this.buttonLeft = new System.Windows.Forms.Button();
            this.buttonRight = new System.Windows.Forms.Button();
            this.buttonCenter = new System.Windows.Forms.Button();
            this.buttonRightLimit = new System.Windows.Forms.Button();
            this.buttonLeftLimit = new System.Windows.Forms.Button();
            this.buttonUpLimit = new System.Windows.Forms.Button();
            this.buttonDownLimit = new System.Windows.Forms.Button();
            this.buttonZoomOutFull = new System.Windows.Forms.Button();
            this.buttonZoomInFull = new System.Windows.Forms.Button();
            this.buttonZoomOut = new System.Windows.Forms.Button();
            this.buttonZoomIn = new System.Windows.Forms.Button();
            this.buttonFastLeft = new System.Windows.Forms.Button();
            this.buttonFastRight = new System.Windows.Forms.Button();
            this.buttonFastDown = new System.Windows.Forms.Button();
            this.buttonFastUp = new System.Windows.Forms.Button();
            this.textPort = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textLastMessage = new System.Windows.Forms.TextBox();
            this.cameraGridView = new System.Windows.Forms.DataGridView();
            this.CameraNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SystemCameraNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Camera = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DevicePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PurgeRedItemsButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.cameraGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonDump
            // 
            this.buttonDump.Location = new System.Drawing.Point(579, 12);
            this.buttonDump.Name = "buttonDump";
            this.buttonDump.Size = new System.Drawing.Size(120, 24);
            this.buttonDump.TabIndex = 2;
            this.buttonDump.Text = "Dump Settings";
            this.buttonDump.Click += new System.EventHandler(this.buttonDump_Click);
            // 
            // buttonStopServer
            // 
            this.buttonStopServer.Location = new System.Drawing.Point(138, 760);
            this.buttonStopServer.Name = "buttonStopServer";
            this.buttonStopServer.Size = new System.Drawing.Size(110, 23);
            this.buttonStopServer.TabIndex = 4;
            this.buttonStopServer.Text = "Stop OSC Server";
            this.buttonStopServer.Click += new System.EventHandler(this.buttonStopServer_Click);
            // 
            // buttonStartServer
            // 
            this.buttonStartServer.Location = new System.Drawing.Point(138, 727);
            this.buttonStartServer.Name = "buttonStartServer";
            this.buttonStartServer.Size = new System.Drawing.Size(111, 23);
            this.buttonStartServer.TabIndex = 5;
            this.buttonStartServer.Text = "Start OSC Server";
            this.buttonStartServer.Click += new System.EventHandler(this.buttonStartServer_Click);
            // 
            // buttonUp
            // 
            this.buttonUp.Location = new System.Drawing.Point(284, 532);
            this.buttonUp.Name = "buttonUp";
            this.buttonUp.Size = new System.Drawing.Size(120, 24);
            this.buttonUp.TabIndex = 7;
            this.buttonUp.Text = "Up";
            this.buttonUp.Click += new System.EventHandler(this.buttonUp_Click);
            // 
            // buttonDown
            // 
            this.buttonDown.Location = new System.Drawing.Point(284, 592);
            this.buttonDown.Name = "buttonDown";
            this.buttonDown.Size = new System.Drawing.Size(120, 24);
            this.buttonDown.TabIndex = 8;
            this.buttonDown.Text = "Down";
            this.buttonDown.Click += new System.EventHandler(this.buttonDown_Click);
            // 
            // buttonLeft
            // 
            this.buttonLeft.Location = new System.Drawing.Point(237, 562);
            this.buttonLeft.Name = "buttonLeft";
            this.buttonLeft.Size = new System.Drawing.Size(44, 24);
            this.buttonLeft.TabIndex = 9;
            this.buttonLeft.Text = "Left";
            this.buttonLeft.Click += new System.EventHandler(this.buttonLeft_Click);
            // 
            // buttonRight
            // 
            this.buttonRight.Location = new System.Drawing.Point(401, 562);
            this.buttonRight.Name = "buttonRight";
            this.buttonRight.Size = new System.Drawing.Size(40, 24);
            this.buttonRight.TabIndex = 10;
            this.buttonRight.Text = "Right";
            this.buttonRight.Click += new System.EventHandler(this.buttonRight_Click);
            // 
            // buttonCenter
            // 
            this.buttonCenter.Location = new System.Drawing.Point(315, 562);
            this.buttonCenter.Name = "buttonCenter";
            this.buttonCenter.Size = new System.Drawing.Size(60, 24);
            this.buttonCenter.TabIndex = 13;
            this.buttonCenter.Text = "Center";
            this.buttonCenter.Click += new System.EventHandler(this.buttonCenter_Click);
            // 
            // buttonRightLimit
            // 
            this.buttonRightLimit.Location = new System.Drawing.Point(451, 562);
            this.buttonRightLimit.Name = "buttonRightLimit";
            this.buttonRightLimit.Size = new System.Drawing.Size(70, 24);
            this.buttonRightLimit.TabIndex = 14;
            this.buttonRightLimit.Text = "Right Limit";
            this.buttonRightLimit.Click += new System.EventHandler(this.buttonRightLimit_Click);
            // 
            // buttonLeftLimit
            // 
            this.buttonLeftLimit.Location = new System.Drawing.Point(157, 562);
            this.buttonLeftLimit.Name = "buttonLeftLimit";
            this.buttonLeftLimit.Size = new System.Drawing.Size(70, 24);
            this.buttonLeftLimit.TabIndex = 15;
            this.buttonLeftLimit.Text = "Left Limit";
            this.buttonLeftLimit.Click += new System.EventHandler(this.buttonLeftLimit_Click);
            // 
            // buttonUpLimit
            // 
            this.buttonUpLimit.Location = new System.Drawing.Point(284, 504);
            this.buttonUpLimit.Name = "buttonUpLimit";
            this.buttonUpLimit.Size = new System.Drawing.Size(120, 24);
            this.buttonUpLimit.TabIndex = 16;
            this.buttonUpLimit.Text = "Up Limit";
            this.buttonUpLimit.Click += new System.EventHandler(this.buttonUpLimit_Click);
            // 
            // buttonDownLimit
            // 
            this.buttonDownLimit.Location = new System.Drawing.Point(284, 620);
            this.buttonDownLimit.Name = "buttonDownLimit";
            this.buttonDownLimit.Size = new System.Drawing.Size(120, 24);
            this.buttonDownLimit.TabIndex = 17;
            this.buttonDownLimit.Text = "Down Limit";
            this.buttonDownLimit.Click += new System.EventHandler(this.buttonDownLimit_Click);
            // 
            // buttonZoomOutFull
            // 
            this.buttonZoomOutFull.Location = new System.Drawing.Point(129, 687);
            this.buttonZoomOutFull.Name = "buttonZoomOutFull";
            this.buttonZoomOutFull.Size = new System.Drawing.Size(120, 24);
            this.buttonZoomOutFull.TabIndex = 23;
            this.buttonZoomOutFull.Text = "Zoom Out Full";
            this.buttonZoomOutFull.Click += new System.EventHandler(this.buttonZoomOutFull_Click);
            // 
            // buttonZoomInFull
            // 
            this.buttonZoomInFull.Location = new System.Drawing.Point(129, 601);
            this.buttonZoomInFull.Name = "buttonZoomInFull";
            this.buttonZoomInFull.Size = new System.Drawing.Size(120, 23);
            this.buttonZoomInFull.TabIndex = 22;
            this.buttonZoomInFull.Text = "Zoom In Full";
            this.buttonZoomInFull.Click += new System.EventHandler(this.buttonZoomInFull_Click);
            // 
            // buttonZoomOut
            // 
            this.buttonZoomOut.Location = new System.Drawing.Point(129, 659);
            this.buttonZoomOut.Name = "buttonZoomOut";
            this.buttonZoomOut.Size = new System.Drawing.Size(120, 24);
            this.buttonZoomOut.TabIndex = 21;
            this.buttonZoomOut.Text = "Zoom Out";
            this.buttonZoomOut.Click += new System.EventHandler(this.buttonZoomOut_Click);
            // 
            // buttonZoomIn
            // 
            this.buttonZoomIn.Location = new System.Drawing.Point(129, 629);
            this.buttonZoomIn.Name = "buttonZoomIn";
            this.buttonZoomIn.Size = new System.Drawing.Size(120, 24);
            this.buttonZoomIn.TabIndex = 20;
            this.buttonZoomIn.Text = "Zoom In";
            this.buttonZoomIn.Click += new System.EventHandler(this.buttonZoomIn_Click);
            // 
            // buttonFastLeft
            // 
            this.buttonFastLeft.Location = new System.Drawing.Point(436, 698);
            this.buttonFastLeft.Name = "buttonFastLeft";
            this.buttonFastLeft.Size = new System.Drawing.Size(103, 24);
            this.buttonFastLeft.TabIndex = 24;
            this.buttonFastLeft.Text = "Fast 10 (-160) Left";
            this.buttonFastLeft.Click += new System.EventHandler(this.buttonFastLeft_Click);
            // 
            // buttonFastRight
            // 
            this.buttonFastRight.Location = new System.Drawing.Point(429, 670);
            this.buttonFastRight.Name = "buttonFastRight";
            this.buttonFastRight.Size = new System.Drawing.Size(110, 24);
            this.buttonFastRight.TabIndex = 25;
            this.buttonFastRight.Text = "Fast 10 (160) Right";
            this.buttonFastRight.Click += new System.EventHandler(this.buttonFastRight_Click);
            // 
            // buttonFastDown
            // 
            this.buttonFastDown.Location = new System.Drawing.Point(429, 601);
            this.buttonFastDown.Name = "buttonFastDown";
            this.buttonFastDown.Size = new System.Drawing.Size(110, 24);
            this.buttonFastDown.TabIndex = 27;
            this.buttonFastDown.Text = "Fast 11 (120) Up";
            this.buttonFastDown.Click += new System.EventHandler(this.button1_Click);
            // 
            // buttonFastUp
            // 
            this.buttonFastUp.Location = new System.Drawing.Point(429, 629);
            this.buttonFastUp.Name = "buttonFastUp";
            this.buttonFastUp.Size = new System.Drawing.Size(103, 24);
            this.buttonFastUp.TabIndex = 26;
            this.buttonFastUp.Text = "Fast 11 (-120) Down";
            this.buttonFastUp.Click += new System.EventHandler(this.button2_Click);
            // 
            // textPort
            // 
            this.textPort.Location = new System.Drawing.Point(290, 730);
            this.textPort.Name = "textPort";
            this.textPort.Size = new System.Drawing.Size(62, 20);
            this.textPort.TabIndex = 28;
            this.textPort.Text = "33333";
            this.textPort.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(260, 732);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 13);
            this.label1.TabIndex = 29;
            this.label1.Text = "Port:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(290, 767);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 30;
            this.label2.Text = "Last Message:";
            // 
            // textLastMessage
            // 
            this.textLastMessage.Location = new System.Drawing.Point(364, 765);
            this.textLastMessage.Name = "textLastMessage";
            this.textLastMessage.Size = new System.Drawing.Size(175, 20);
            this.textLastMessage.TabIndex = 31;
            // 
            // cameraGridView
            // 
            this.cameraGridView.AllowDrop = true;
            this.cameraGridView.AllowUserToAddRows = false;
            this.cameraGridView.AllowUserToDeleteRows = false;
            this.cameraGridView.AllowUserToResizeRows = false;
            this.cameraGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.cameraGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CameraNumber,
            this.SystemCameraNumber,
            this.Camera,
            this.DevicePath});
            this.cameraGridView.Location = new System.Drawing.Point(12, 50);
            this.cameraGridView.MultiSelect = false;
            this.cameraGridView.Name = "cameraGridView";
            this.cameraGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.cameraGridView.Size = new System.Drawing.Size(781, 448);
            this.cameraGridView.TabIndex = 32;
            // 
            // CameraNumber
            // 
            this.CameraNumber.HeaderText = "Camera #";
            this.CameraNumber.Name = "CameraNumber";
            this.CameraNumber.ReadOnly = true;
            this.CameraNumber.Width = 60;
            // 
            // SystemCameraNumber
            // 
            this.SystemCameraNumber.HeaderText = "System Camera #";
            this.SystemCameraNumber.Name = "SystemCameraNumber";
            this.SystemCameraNumber.ReadOnly = true;
            this.SystemCameraNumber.Width = 60;
            // 
            // Camera
            // 
            this.Camera.HeaderText = "Camera";
            this.Camera.Name = "Camera";
            this.Camera.Width = 140;
            // 
            // DevicePath
            // 
            this.DevicePath.HeaderText = "Device Path";
            this.DevicePath.Name = "DevicePath";
            this.DevicePath.Width = 400;
            // 
            // PurgeRedItemsButton
            // 
            this.PurgeRedItemsButton.Location = new System.Drawing.Point(401, 12);
            this.PurgeRedItemsButton.Name = "PurgeRedItemsButton";
            this.PurgeRedItemsButton.Size = new System.Drawing.Size(170, 24);
            this.PurgeRedItemsButton.TabIndex = 33;
            this.PurgeRedItemsButton.Text = "Purge Missing Devices (Red)";
            this.PurgeRedItemsButton.Click += new System.EventHandler(this.PurgeRedItemsButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(805, 798);
            this.Controls.Add(this.PurgeRedItemsButton);
            this.Controls.Add(this.cameraGridView);
            this.Controls.Add(this.textLastMessage);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textPort);
            this.Controls.Add(this.buttonFastDown);
            this.Controls.Add(this.buttonFastUp);
            this.Controls.Add(this.buttonFastRight);
            this.Controls.Add(this.buttonFastLeft);
            this.Controls.Add(this.buttonZoomOutFull);
            this.Controls.Add(this.buttonZoomInFull);
            this.Controls.Add(this.buttonZoomOut);
            this.Controls.Add(this.buttonZoomIn);
            this.Controls.Add(this.buttonDownLimit);
            this.Controls.Add(this.buttonUpLimit);
            this.Controls.Add(this.buttonLeftLimit);
            this.Controls.Add(this.buttonRightLimit);
            this.Controls.Add(this.buttonCenter);
            this.Controls.Add(this.buttonRight);
            this.Controls.Add(this.buttonLeft);
            this.Controls.Add(this.buttonDown);
            this.Controls.Add(this.buttonUp);
            this.Controls.Add(this.buttonStartServer);
            this.Controls.Add(this.buttonStopServer);
            this.Controls.Add(this.buttonDump);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "The One Camera Controller";
            ((System.ComponentModel.ISupportInitialize)(this.cameraGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {


            Application.Run(new Form1());


        }

        public void StopServer()
        {
            buttonStartServer.Enabled = true;
            buttonStopServer.Enabled = false;
            textPort.Enabled = true;
            if (oscReceiver != null)
                oscReceiver.stop();
            return;
        }

        public void StartServer()
        {
            buttonStartServer.Enabled = false;
            buttonStopServer.Enabled = true;
            textPort.Enabled = false;
            oscReceiver = new OSCReceiver(int.Parse(this.textPort.Text), this);

            return;
        }

        // Function to load and compare data and highlight rows accordingly
        private void LoadAndCompareData(DataGridView dataGridView, string filePath)
        {
            // Load the data from the JSON file
            List<Dictionary<string, object>> loadedData = LoadDataFromJson(filePath);

            // Get the current system data
            List<Dictionary<string, object>> currentSystemData = DataGridViewToDictList(dataGridView);

            // Find missing devices and new devices
            List<Dictionary<string, object>> missingDevices = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> newDevices = new List<Dictionary<string, object>>();

            // Assuming the DevicePath column is the unique identifier for each device
            HashSet<string> loadedDevicePaths = new HashSet<string>();
            HashSet<string> currentDevicePaths = new HashSet<string>();

            foreach (var device in loadedData)
            {
                loadedDevicePaths.Add(device["DevicePath"].ToString());
            }

            foreach (var device in currentSystemData)
            {
                currentDevicePaths.Add(device["DevicePath"].ToString());
            }

            // Find missing devices
            foreach (var device in loadedData)
            {
                if (!currentDevicePaths.Contains(device["DevicePath"].ToString()))
                {
                    missingDevices.Add(device);
                }
            }

            // Find new devices
            foreach (var device in currentSystemData)
            {
                if (!loadedDevicePaths.Contains(device["DevicePath"].ToString()))
                {
                    newDevices.Add(device);
                }
            }

            // Combine the data and load it into the DataGridView
            List<Dictionary<string, object>> finalData = new List<Dictionary<string, object>>(loadedData);
            finalData.AddRange(newDevices);

            dataGridView.Rows.Clear();

            foreach (var rowData in finalData)
            {
                int rowIndex = dataGridView.Rows.Add();
                DataGridViewRow row = dataGridView.Rows[rowIndex];

                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    if (rowData.ContainsKey(column.Name))
                    {
                        row.Cells[column.Name].Value = rowData[column.Name];
                    }
                }

                // Highlight the row based on its status
                if (missingDevices.Contains(rowData))
                {
                    row.DefaultCellStyle.BackColor = Color.Red;
                }
                else if (newDevices.Contains(rowData))
                {
                    row.DefaultCellStyle.BackColor = Color.Green;
                }
            }

            UpdateSystemCameraNumber(cameraGridView);
        }

        private void UpdateSystemCameraNumber(DataGridView dataGridView)
        {
            // Iterate through all rows in the DataGridView
            foreach (DataGridViewRow row in dataGridView.Rows)
            {

                // Retrieve the DevicePath value (or any other value you need for your criteria)
                string devicePath = row.Cells["DevicePath"].Value.ToString();

                // Update the SystemCameraNumber based on your criteria
                // Replace this with your actual implementation to get the new SystemCameraNumber
                int newSystemCameraNumber = GetSystemCameraNumberByDevicePath(devicePath);

                // Set the new SystemCameraNumber value for the row
                row.Cells["SystemCameraNumber"].Value = newSystemCameraNumber;

            }
        }

        // Function to load data from a JSON file and return a list of dictionaries
        private List<Dictionary<string, object>> LoadDataFromJson(string filePath)
        {
            List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();

            if (File.Exists(filePath))
            {
                var json = File.ReadAllText(filePath);
                data = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);
            }

            return data;
        }

        private IBaseFilter CreateFilter(Guid category, string dpath)
        {
            object source = null;
            Guid iid = typeof(IBaseFilter).GUID;
            foreach (DsDevice device in DsDevice.GetDevicesOfCat(category))
            {
                if (device.DevicePath.CompareTo(dpath) == 0)
                {
                    device.Mon.BindToObject(null, null, ref iid, out source);
                    break;
                }
            }

            return (IBaseFilter)source;
        }

        // Function to convert DataGridView to a list of dictionaries
        private List<Dictionary<string, object>> DataGridViewToDictList(DataGridView dataGridView)
        {
            var dictList = new List<Dictionary<string, object>>();

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.IsNewRow) continue;
                var rowData = new Dictionary<string, object>();
                foreach (DataGridViewCell cell in row.Cells)
                    rowData[dataGridView.Columns[cell.ColumnIndex].Name] = cell.Value;
                dictList.Add(rowData);
            }

            return dictList;
        }

        // Function to save DataGridView data to a JSON file
        private void SaveDataGridViewToJson(DataGridView dataGridView, string filePath)
        {
            var data = DataGridViewToDictList(dataGridView);
            var json = JsonConvert.SerializeObject(data, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }

        // Function to load DataGridView data from a JSON file
        private void LoadDataGridViewFromJson(DataGridView dataGridView, string filePath)
        {
            if (File.Exists(filePath))
            {
                var json = File.ReadAllText(filePath);
                var data = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);

                dataGridView.Rows.Clear();

                foreach (var rowData in data)
                {
                    var row = new DataGridViewRow();
                    row.CreateCells(dataGridView);

                    foreach (var cellData in rowData)
                    {
                        if (dataGridView.Columns.Contains(cellData.Key))
                        {
                            // Use the column index instead of the column name
                            int columnIndex = dataGridView.Columns[cellData.Key].Index;
                            row.Cells[columnIndex].Value = cellData.Value;
                        }
                        else
                        {
                            // You can either ignore the missing column or add some error handling logic here
                        }
                    }

                    dataGridView.Rows.Add(row);
                }
            }
        }



        public static async Task PostDataAsync(string url, object data)
        {
            HttpClient client = new HttpClient();

            // Convert the data object to a JSON string
            string jsonData = JsonConvert.SerializeObject(data);

            StringContent content = new StringContent(jsonData, Encoding.UTF8, "application/json");

            HttpResponseMessage response = await client.PostAsync(url, content);

            if (response.IsSuccessStatusCode)
            {
                Console.WriteLine("Data successfully posted to the server.");
                string responseBody = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"Server response: {responseBody}");
            }
            else
            {
                Console.WriteLine($"Error posting data: {response.StatusCode}");
            }
        }

        private async void dumpAll()
        {

            var cameraList = new List<object>();

            object source = null;
            IBaseFilter idevice = null;
            //var cameraControl = null;
            Guid iid = typeof(IBaseFilter).GUID;
            foreach (DsDevice device in DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice))
            {
                try
                {
                    device.Mon.BindToObject(null, null, ref iid, out source);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"device failed: {device.Name}:");
                    //Console.WriteLine(ex.ToString());
                    continue;
                }

                idevice = (IBaseFilter)source;
                var cameraControl = idevice as IAMCameraControl;
                if (cameraControl == null) continue;

                Console.WriteLine($"device: {device.Name}:");
                Console.WriteLine($"device path: {device.DevicePath}:");

                var settingsList = new List<object>();

                for (int i = 0; i <= 100; i++)
                {


                    int result = cameraControl.GetRange((CameraControlProperty)i,
                        out int min, out int max, out int steppingDelta,
                        out int defaultValue, out var flags);

                    if (result == 0)
                    {
                        Console.WriteLine($"Property: {i}, min: {min}, max: {max}, steppingDelta: {steppingDelta}");
                        Console.WriteLine($"defaultValue: {defaultValue}, flags: {flags}\n");

                        cameraControl.Get((CameraControlProperty)i, out int value, out var flags2);
                        Console.WriteLine($"currentValue: {value}, flags: {flags2}\n");

                        var props = new
                        {
                            property = i,
                            min = min,
                            max = max,
                            steppingDelta = steppingDelta,
                            defaultValue = defaultValue,
                            rangeFlags = flags,
                            currentValue = value,
                            currentFlags = flags2
                        };
                        settingsList.Add(props);
                    }
                }

                Console.WriteLine("-----------------------");


                cameraList.Add(new { device = device, settings = settingsList });


            }
            await PostDataAsync("http://127.0.0.1:1880/TOCCSetCameraConfig", cameraList);
        }

        private void dumpSettings()
        {
            object source = null;
            IBaseFilter idevice = null;
            //var cameraControl = null;
            Guid iid = typeof(IBaseFilter).GUID;
            foreach (DsDevice device in DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice))
            {
                try
                {
                    device.Mon.BindToObject(null, null, ref iid, out source);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"device failed: {device.Name}:");
                    //Console.WriteLine(ex.ToString());
                    continue;
                }

                idevice = (IBaseFilter)source;
                var cameraControl = idevice as IAMCameraControl;
                if (cameraControl == null) continue;

                Console.WriteLine($"device: {device.Name}:");
                foreach (CameraControlProperty i in Enum.GetValues(typeof(CameraControlProperty)))
                {
                    cameraControl.GetRange(i,
                        out int min, out int max, out int steppingDelta,
                        out int defaultValue, out var flags);

                    Console.WriteLine($"Property: {i}, min: {min}, max: {max}, steppingDelta: {steppingDelta}");
                    Console.WriteLine($"defaultValue: {defaultValue}, flags: {flags}\n");

                    cameraControl.Get(i, out int value, out var flags2);
                    Console.WriteLine($"currentValue: {value}, flags: {flags2}\n");

                }
                Console.WriteLine("-----------------------");


            }
        }

        private void buttonDump_Click(object sender, System.EventArgs e)
        {
            dumpAll();
        }

        private void buttonStartServer_Click(object sender, System.EventArgs e)
        {
            StartServer();
        }

        private void buttonStopServer_Click(object sender, System.EventArgs e)
        {
            StopServer();
        }

        private void buttonUp_Click(object sender, EventArgs e)
        {
            Up();
        }

        private void Up()
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get(CameraControlProperty.Tilt, out int value, out var flags);
            cameraControl.Set(CameraControlProperty.Tilt, value - 10, CameraControlFlags.Manual);
            Console.WriteLine($"Property: {CameraControlProperty.Tilt}, value: {value}");

        }

        private void buttonDown_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get(CameraControlProperty.Tilt, out int value, out var flags);
            cameraControl.Set(CameraControlProperty.Tilt, value + 10, CameraControlFlags.Manual);
            Console.WriteLine($"Property: {CameraControlProperty.Tilt}, value: {value}");
        }

        public void farLeft()
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Set((CameraControlProperty)10, 100, CameraControlFlags.Manual);
            //Console.WriteLine($"Property: {CameraControlProperty.Tilt}, value: {value}");
        }

        private void buttonLeft_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get(CameraControlProperty.Pan, out int value, out var flags);
            cameraControl.Set(CameraControlProperty.Pan, value - 10, CameraControlFlags.Manual);
            Console.WriteLine($"Property: {CameraControlProperty.Pan}, value: {value}");
        }

        private void buttonRight_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get(CameraControlProperty.Pan, out int value, out var flags);
            cameraControl.Set(CameraControlProperty.Pan, value + 10, CameraControlFlags.Manual);
            Console.WriteLine($"Property: {CameraControlProperty.Pan}, value: {value}");
        }



        private void buttonZoomIn_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get(CameraControlProperty.Zoom, out int value, out var flags);
            cameraControl.Set(CameraControlProperty.Zoom, value + 10, CameraControlFlags.Manual);
            Console.WriteLine($"Property: {CameraControlProperty.Zoom}, value: {value}");

        }

        private void buttonZoomOut_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get(CameraControlProperty.Zoom, out int value, out var flags);
            cameraControl.Set(CameraControlProperty.Zoom, value - 10, CameraControlFlags.Manual);
            Console.WriteLine($"Property: {CameraControlProperty.Zoom}, value: {value}");

        }

        private void buttonCenter_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.GetRange(CameraControlProperty.Pan,
                out int min, out int max, out int steppingDelta,
                out int defaultValue, out var flags);
            cameraControl.Set(CameraControlProperty.Pan, 0, CameraControlFlags.Manual);

            cameraControl.GetRange(CameraControlProperty.Tilt,
                out min, out max, out steppingDelta,
                out defaultValue, out flags);
            cameraControl.Set(CameraControlProperty.Tilt, defaultValue, CameraControlFlags.Manual);

            cameraControl.GetRange(CameraControlProperty.Zoom,
                out min, out max, out steppingDelta,
                out defaultValue, out flags);
            cameraControl.Set(CameraControlProperty.Zoom, defaultValue, CameraControlFlags.Manual);
        }

        private void buttonUpLimit_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.GetRange(CameraControlProperty.Tilt,
                out int min, out int max, out int steppingDelta,
                out int defaultValue, out var flags);
            cameraControl.Set(CameraControlProperty.Tilt, 60, CameraControlFlags.Manual);

        }

        private void buttonDownLimit_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.GetRange(CameraControlProperty.Tilt,
                out int min, out int max, out int steppingDelta,
                out int defaultValue, out var flags);
            cameraControl.Set(CameraControlProperty.Tilt, -60, CameraControlFlags.Manual);
        }

        private void buttonLeftLimit_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            //cameraControl.GetRange(CameraControlProperty.Pan,
            //    out int min, out int max, out int steppingDelta,
            //    out int defaultValue, out var flags);
            //cameraControl.Set((CameraControlProperty)10, 100, CameraControlFlags.Manual);
            cameraControl.Set(CameraControlProperty.Pan, -129, CameraControlFlags.Manual);

            //Console.WriteLine($"Property: {CameraControlProperty.Pan}, min: {min}, MAX: {max}, steppingDelta: {steppingDelta}");
        }
        private void buttonRightLimit_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            //cameraControl.GetRange(CameraControlProperty.Pan,
            //    out int min, out int max, out int steppingDelta,
            //    out int defaultValue, out var flags);
            cameraControl.Set(CameraControlProperty.Pan, 129, CameraControlFlags.Manual);
            //Console.WriteLine($"Property: {CameraControlProperty.Pan}, MIN: {min}, max: {max}, steppingDelta: {steppingDelta}");
        }

        private void buttonZoomInFull_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.GetRange(CameraControlProperty.Zoom,
                out int min, out int max, out int steppingDelta,
                out int defaultValue, out var flags);
            cameraControl.Set(CameraControlProperty.Zoom, max, CameraControlFlags.Manual);

        }

        private void buttonZoomOutFull_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.GetRange(CameraControlProperty.Zoom,
                out int min, out int max, out int steppingDelta,
                out int defaultValue, out var flags);
            cameraControl.Set(CameraControlProperty.Zoom, min, CameraControlFlags.Manual);

        }

        private void buttonFastZoomInFull_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Set((CameraControlProperty)13, 100, CameraControlFlags.Manual);

        }

        private void buttonFastZoomOutFull_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Set((CameraControlProperty)13, -100, CameraControlFlags.Manual);

        }
        private void button2_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get((CameraControlProperty)11, out int value, out var flags);
            cameraControl.Set((CameraControlProperty)10, 0, CameraControlFlags.Manual);
            cameraControl.Set((CameraControlProperty)11, -120, CameraControlFlags.Manual);
            //Console.WriteLine($"Property 11: value: {value}");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get((CameraControlProperty)11, out int value, out var flags);
            cameraControl.Set((CameraControlProperty)10, 0, CameraControlFlags.Manual);
            cameraControl.Set((CameraControlProperty)11, 120, CameraControlFlags.Manual);
            //Console.WriteLine($"Property 11: value: {value}");

        }

        private void buttonFastLeft_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get((CameraControlProperty)10, out int value, out var flags);
            cameraControl.Set((CameraControlProperty)11, 0, CameraControlFlags.Manual);
            cameraControl.Set((CameraControlProperty)10, 160, CameraControlFlags.Manual);
            Console.WriteLine($"Property 10: value: {value}");
        }

        private void buttonFastRight_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            cameraControl.Get((CameraControlProperty)10, out int value, out var flags);
            cameraControl.Set((CameraControlProperty)11, 0, CameraControlFlags.Manual);
            cameraControl.Set((CameraControlProperty)10, -160, CameraControlFlags.Manual);
            Console.WriteLine($"Property 10: value: {value}");

        }

        public void SelectDeviceByCameraNumber(int cameraNumber)
        {
            foreach (DataGridViewRow row in cameraGridView.Rows)
            {
                if (row.Cells["CameraNumber"].Value != null && int.Parse(row.Cells["CameraNumber"].Value.ToString()) == cameraNumber)
                {
                    row.Selected = true;
                    theDevicePath = row.Cells["DevicePath"].Value.ToString();
                    theDevice = CreateFilter(FilterCategory.VideoInputDevice, theDevicePath);
                    Console.WriteLine($"Selected device: {theDevicePath}");
                    return;
                }
            }
        }

        public void SelectDeviceBySystemCameraNumber(int cameraNumber)
        {
            foreach (DataGridViewRow row in cameraGridView.Rows)
            {
                if (row.Cells["SystemCameraNumber"].Value != null && (int)row.Cells["SystemCameraNumber"].Value == cameraNumber)
                {
                    row.Selected = true;
                    theDevicePath = row.Cells["DevicePath"].Value.ToString();
                    theDevice = CreateFilter(FilterCategory.VideoInputDevice, theDevicePath);
                    return;
                }
            }
        }

        public void SelectDeviceByDevicePath(string devicePath)
        {
            foreach (DataGridViewRow row in cameraGridView.Rows)
            {
                if (row.Cells["DevicePath"].Value != null && row.Cells["DevicePath"].Value.ToString() == devicePath)
                {
                    row.Selected = true;
                    theDevicePath = devicePath;
                    theDevice = CreateFilter(FilterCategory.VideoInputDevice, theDevicePath);
                    return;
                }
            }

        }

        public int GetSystemCameraNumberByDevicePath(string devicePath)
        {
            foreach (DataGridViewRow row in cameraGridView.Rows)
            {
                if (row.Cells["DevicePath"].Value != null && row.Cells["DevicePath"].Value.ToString() == devicePath)
                {
                    int systemCameraNumber;
                    if (int.TryParse(row.Cells["SystemCameraNumber"].Value.ToString(), out systemCameraNumber))
                    {
                        return systemCameraNumber;
                    }
                }
            }

            // If the device path is not found, return -1 or throw an exception
            return -1;
        }

        public int GetCameraNumberByDevicePath(string devicePath)
        {
            foreach (DataGridViewRow row in cameraGridView.Rows)
            {
                if (row.Cells["DevicePath"].Value != null && row.Cells["DevicePath"].Value.ToString() == devicePath)
                {
                    int cameraNumber;
                    if (int.TryParse(row.Cells["CameraNumber"].Value.ToString(), out cameraNumber))
                    {
                        return cameraNumber;
                    }
                }
            }

            // If the device path is not found, return -1 or throw an exception
            return -1;
        }

        private void CameraGridView_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {

            // Select the whole row when a cell is clicked
            if (e.RowIndex >= 0 && e.Button == MouseButtons.Left)
            {
                DataGridViewRow selectedRow = cameraGridView.Rows[e.RowIndex];
                cameraGridView.ClearSelection();
                selectedRow.Selected = true;
                cameraGridView.CurrentCell = cameraGridView.Rows[e.RowIndex].Cells[0];

                theDevicePath = selectedRow.Cells["DevicePath"].Value.ToString();
                Console.WriteLine($"Device Path: {theDevicePath}");
                theDevice = CreateFilter(FilterCategory.VideoInputDevice, theDevicePath);
            }
        }

        private int rowIndex = -1;
        private Rectangle dragBoxFromMouseDown;

        private void CameraGridView_MouseDown(object sender, MouseEventArgs e)
        {
            rowIndex = cameraGridView.HitTest(e.X, e.Y).RowIndex;
            if (rowIndex >= 0)
            {
                Size dragSize = SystemInformation.DragSize;
                dragBoxFromMouseDown = new Rectangle(new Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize);
            }
            else
            {
                dragBoxFromMouseDown = Rectangle.Empty;
            }
        }

        private void CameraGridView_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left && dragBoxFromMouseDown != Rectangle.Empty && !dragBoxFromMouseDown.Contains(e.X, e.Y))
            {
                cameraGridView.DoDragDrop(cameraGridView.Rows[rowIndex], DragDropEffects.Move);
            }
        }

        private void CameraGridView_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void CameraGridView_DragDrop(object sender, DragEventArgs e)
        {
            Point clientPoint = cameraGridView.PointToClient(new Point(e.X, e.Y));
            int destinationIndex = cameraGridView.HitTest(clientPoint.X, clientPoint.Y).RowIndex;
            if (destinationIndex >= 0 && destinationIndex != rowIndex)
            {
                DataGridViewRow rowToMove = e.Data.GetData(typeof(DataGridViewRow)) as DataGridViewRow;
                cameraGridView.Rows.RemoveAt(rowIndex);
                cameraGridView.Rows.Insert(destinationIndex, rowToMove);
                reorderCameras();
                SaveDataGridViewToJson(cameraGridView, "config.json");
            }
        }

        private void reorderCameras()
        {
            for (int i = 0; i < cameraGridView.Rows.Count; i++)
            {
                cameraGridView.Rows[i].Cells["CameraNumber"].Value = i + 1;
            }
            UpdateSystemCameraNumber(cameraGridView);
        }

        private void PurgeRedItemsButton_Click(object sender, EventArgs e)
        {
            // Loop through the rows in reverse order to avoid issues caused by row indices changing as rows are removed
            for (int rowIndex = cameraGridView.Rows.Count - 1; rowIndex >= 0; rowIndex--)
            {
                DataGridViewRow row = cameraGridView.Rows[rowIndex];

                // Check if the row has a red background
                if (row.DefaultCellStyle.BackColor == Color.Red)
                {
                    // Remove the row from the DataGridView
                    cameraGridView.Rows.RemoveAt(rowIndex);
                }
            }
            reorderCameras();
            SaveDataGridViewToJson(cameraGridView, "config.json");

        }
    }

    public class OSCReceiver
    {
        private Form1 _parent = null;
        private UDPListener _listener = null;
        private Dictionary<int, CameraControl> _threads;

        public OSCReceiver(int port, Form1 parent)
        {
            _parent = parent;

            HandleOscPacket callback = delegate (OscPacket packet)
            {
                var message = (OscMessage)packet;

                spawnCameraControl(message);
            };

            _listener = new UDPListener(port, callback);
            _threads = new Dictionary<int, CameraControl>();
        }

        public void stop()
        {
            Console.WriteLine("Cleanup OSCReceiver");
            _listener.Close();
        }

        public void spawnCameraControl(OscMessage message)
        {
            string devicePath = "none";
            CameraControl result = null;
            int systemCamera = 1;

            if (message.Address == "/SetCurrentDevice")
            {
                int cameraNumber = (int)message.Arguments[0];
                _parent.SelectDeviceByCameraNumber(cameraNumber);
                devicePath = _parent.theDevicePath;
                Console.WriteLine($"Set Device: {devicePath}");
            }
            else if (message.Address == "/SendCurrentValuesToEndpoint")
            {
                devicePath = _parent.theDevicePath;
            }
            else if (message.Address == "/FLYXY" && message.Arguments.Count == 2)
            {
                devicePath = _parent.theDevicePath;
                //Console.WriteLine($"/FLYXY device {devicePath}");
            }
            else if (message.Address == "/FLYZ")
            {
                devicePath = _parent.theDevicePath;
            }
            else if (message.Arguments[0] is string)
            {
                devicePath = (string)message.Arguments[0];
                _parent.SelectDeviceByDevicePath(devicePath);
                Console.WriteLine($"Device: {devicePath}");
            }
            else
            {
                int cameraNumber = (int)message.Arguments[0];
                _parent.SelectDeviceByCameraNumber(cameraNumber);
                devicePath = _parent.theDevicePath;
            }

            systemCamera = _parent.GetSystemCameraNumberByDevicePath(devicePath);

            //Console.WriteLine($"System Camera Number: {systemCamera} Device Path: {devicePath}");

            if (_threads.TryGetValue(systemCamera, out result))
            {
                result.receiveOSC(message);
            }
            else
            {
                CameraControl cc = new CameraControl(_parent, devicePath);
                cc.Start();
                cc.setDevice(devicePath);
                cc.receiveOSC(message);
                _threads.Add(systemCamera, cc);
            }

            //Console.WriteLine($"Total Threads: {_threads.Count}");
            // Remove completed threads

            //}
            //_threads.RemoveAll(t => !t.IsAlive);
        }
    }

    public class CameraControl
    {
        private Form1 _parent = null;

        private Thread _thread = null;

        private IBaseFilter theDevice = null;
        string theDevicePath = "unbound";
        private IAMCameraControl cameraControl = null;

        private int lastX = 0;
        private int lastY = 0;
        private int lastZ = 0;
        private int lastPan = 0;
        private int lastTilt = 0;
        private int lastZoom = 0;
        private int lastExposure = 0;
        private int lastFocus = 0;

        private int destPan = 0;
        private int destTilt = 0;
        private int destZoom = 0;

        private long startPanTicks = 0;
        private long startTiltTicks = 0;
        private long startZoomTicks = 0;

        private long elapsedTimeMS = 0;

        private long stopPanTicks = 0;
        private long stopTiltTicks = 0;
        private long stopZoomTicks = 0;

        private int panAmount = 1;
        private int tiltAmount = 1;
        private int zoomAmount = 1;

        private int movementSpeed = 30;

        private int panAdditionalSpeed = 0;
        private int tiltAdditionalSpeed = 0;
        private int zoomAdditionalSpeed = 0;
        private int maxAdditionalSpeed = 30;

        private int maxAdditionalPanSpeed = 30;
        private int maxAdditionalTiltSpeed = 30;
        private int maxAdditionalZoomSpeed = 30;

        int estPanTime = 0;
        int estTiltTime = 0;
        int estZoomTime = 0;


        private int panSpeed = 15;
        private int tiltSpeed = 15;
        private int zoomSpeed = 15;

        private int lastPanSpeed = 0;
        private int lastTiltSpeed = 0;
        private int lastZoomSpeed = 0;

        private int calculatedPanSpeed = 15;
        private int calculatedTiltSpeed = 15;
        private int calculatedZoomSpeed = 15;

        private int calculatedPanEndSpeed = 15;
        private int calculatedTiltEndSpeed = 15;
        private int calculatedZoomEndSpeed = 15;


        private bool stopPan = false;
        private bool stopTilt = false;
        private bool stopZoom = false;

        private bool zoomIn = false;

        private bool gotoActive = false;

        private bool zoomActive = false;

        double messageReceivedTimeTicks = 0;

        string remoteIPAddress;
        int remotePort;


        // For sending data out to Splunk in real-time.
        private TcpClient _tcpoutClient = null;
        private byte[] _tcpoutBytes = null;
        private NetworkStream _tcpoutStream = null;
        DateTimeOffset epoch = new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero);

        public CameraControl(Form1 parent, string devicePath)
        {
            _parent = parent;

            setDevice(devicePath);
            _thread = new Thread(new ThreadStart(cameraControlLoop));

            _tcpoutClient = new TcpClient("10.1.10.173", 30000);
            _tcpoutStream = _tcpoutClient.GetStream();
        }

        public void setDevice(string devicePath)
        {
            double startTimeTicks = DateTime.UtcNow.Ticks;
            object source = null;
            Guid iid = typeof(IBaseFilter).GUID;
            foreach (DsDevice device in DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice))
            {
                if (device.DevicePath.CompareTo(devicePath) == 0)
                {
                    try
                    {
                        device.Mon.BindToObject(null, null, ref iid, out source);
                        theDevicePath = device.DevicePath;
                        theDevice = (IBaseFilter)source;
                        cameraControl = theDevice as IAMCameraControl;
                        Console.WriteLine($"MS to find and set device: {(int)((DateTime.UtcNow.Ticks - startTimeTicks) / 10000)}ms {theDevicePath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"device failed: {ex}:");
                        //Console.WriteLine(ex.ToString());
                        continue;
                    }
                    break;
                }
            }
        }

        public void Stop()
        {
            Console.WriteLine("Cleanup on isle 7");
            _tcpoutClient.Close();
        }

        // Start the thread
        public void Start()
        {
            Console.WriteLine("Start()");
            _thread.Start();
        }

        // Wait for the thread to complete
        public void Join()
        {
            _thread.Join();
        }

        public bool IsAlive
        {
            get { return _thread.IsAlive; }
        }

        public void receiveOSC(OscMessage message)
        {
            messageReceivedTimeTicks = DateTime.UtcNow.Ticks;
            string address = message.Address;
            remoteIPAddress = message.remoteIPAddress;
            remotePort = message.remotePort;

            //Console.WriteLine($"Remote: {remoteIPAddress}:{remotePort}");

            if (cameraControl != null)
            {
                if (address == "/GetCurrentValues")
                {
                    // Must have at least one value passed.
                    int dummy = (int)message.Arguments[0];

                    cameraControl.Get(CameraControlProperty.Pan, out int currentPan, out var flags1);
                    cameraControl.Get(CameraControlProperty.Tilt, out int currentTilt, out var flags2);
                    cameraControl.Get(CameraControlProperty.Zoom, out int currentZoom, out var flags3);
                    cameraControl.Get(CameraControlProperty.Focus, out int currentFocus, out var flags4);

                    Console.WriteLine($"Sending Current PTZ: {currentPan} {currentTilt} {currentZoom} {currentFocus} \n");
                    Console.WriteLine($"To: {remoteIPAddress}:{remotePort}\n");

                    var responseMessage = new SharpOSC.OscMessage("/CurrentValues", currentPan, currentTilt, currentZoom, currentFocus, theDevicePath);
                    var sender = new SharpOSC.UDPSender(message.remoteIPAddress, message.remotePort);

                    sender.Send(responseMessage);
                }
                else if (address == "/SendCurrentValuesToEndpoint")
                {
                    Console.WriteLine($"/SendCurrentValuesToEndpoint");
                    // Must have at least one value passed.
                    remoteIPAddress = (string)message.Arguments[0];
                    remotePort = (int)message.Arguments[1];

                    cameraControl.Get(CameraControlProperty.Pan, out int currentPan, out var flags1);
                    cameraControl.Get(CameraControlProperty.Tilt, out int currentTilt, out var flags2);
                    cameraControl.Get(CameraControlProperty.Zoom, out int currentZoom, out var flags3);
                    cameraControl.Get(CameraControlProperty.Focus, out int currentFocus, out var flags4);

                    Console.WriteLine($"Sending Current PTZ: {currentPan} {currentTilt} {currentZoom} {currentFocus}");
                    Console.WriteLine($"To: {remoteIPAddress}:{remotePort}\n");

                    var cameraNumber = _parent.GetCameraNumberByDevicePath(theDevicePath);
                    var responseMessage = new SharpOSC.OscMessage("/CurrentValues", cameraNumber, currentPan, currentTilt, currentZoom, currentFocus);
                    var sender = new SharpOSC.UDPSender(remoteIPAddress, remotePort);
                    sender.Send(responseMessage);
                }
                else if (address == "/FLYXY")
                {
                    int valueX = (int)message.Arguments[0];
                    int valueY = ((int)message.Arguments[1]) * -1;

                    if (lastX != valueX)
                        cameraControl.Set((CameraControlProperty)10, valueX, CameraControlFlags.Manual);
                    if (lastY != valueY)
                        cameraControl.Set((CameraControlProperty)11, valueY, CameraControlFlags.Manual);

                    lastX = valueX;
                    lastY = valueY;
                    _parent.textLastMessage.Invoke((MethodInvoker)delegate
                    {
                        _parent.textLastMessage.Text = $"{address} {lastX} {lastY}";
                    });
                }
                else if (address == "/FLYZ")
                {
                    int valueZ = (int)message.Arguments[0];
                    remoteIPAddress = (string)message.Arguments[1];
                    remotePort = (int)message.Arguments[2];

                    if (valueZ != 0)
                    {
                        Console.WriteLine($"zoomActive=true");
                        zoomActive = true;
                    }
                    else
                    {
                        Console.WriteLine($"zoomActive=false");
                        zoomActive = false;
                    }

                    if (Math.Abs(lastZ - valueZ) > 5)
                    {
                        cameraControl.Set((CameraControlProperty)13, valueZ, CameraControlFlags.Manual);
                        lastZ = valueZ;
                    }

                    //_parent.textLastMessage.Invoke((MethodInvoker)delegate
                    //{
                    //    _parent.textLastMessage.Text = $"{address} {lastZ}";
                    //});
                }
                else if (address == "/PAN")
                {
                    int valuePan = (int)message.Arguments[0];

                    //if (lastPan != valuePan)
                    cameraControl.Set(CameraControlProperty.Pan, valuePan, CameraControlFlags.Manual);

                    lastPan = valuePan;
                    _parent.textLastMessage.Invoke((MethodInvoker)delegate
                    {
                        _parent.textLastMessage.Text = $"{address} {lastPan}";
                    });

                }
                else if (address == "/TILT")
                {
                    int valueTilt = (int)message.Arguments[0];

                    //if (lastTilt != valueTilt)
                    cameraControl.Set(CameraControlProperty.Tilt, valueTilt, CameraControlFlags.Manual);

                    lastTilt = valueTilt;
                    _parent.textLastMessage.Invoke((MethodInvoker)delegate
                    {
                        _parent.textLastMessage.Text = $"{address} {lastTilt}";
                    });

                }
                else if (address == "/ZOOM")
                {
                    int valueZoom = (int)message.Arguments[0];

                    //if (lastZoom != valueZoom)
                    cameraControl.Set(CameraControlProperty.Zoom, valueZoom, CameraControlFlags.Manual);

                    lastZoom = valueZoom;
                    _parent.textLastMessage.Invoke((MethodInvoker)delegate
                    {
                        _parent.textLastMessage.Text = $"{address} {lastZoom}";
                    });

                }
                else if (address == "/PTZS")
                {
                    Console.Write($"Arguments: {message.Arguments.Count}");

                    destPan = (int)message.Arguments[1];
                    destTilt = (int)message.Arguments[2];
                    destZoom = (int)message.Arguments[3];
                    movementSpeed = (int)message.Arguments[4];

                    PTZS(0);

                    _parent.textLastMessage.Invoke((MethodInvoker)delegate
                    {
                        _parent.textLastMessage.Text = $"PTZ:  {destPan} {destTilt} {destZoom}";
                    });
                }
                else if (address == "/FOCUS")
                {
                    int valueFocus = (int)message.Arguments[0];

                    if (lastFocus != valueFocus)
                        cameraControl.Set(CameraControlProperty.Focus, valueFocus, CameraControlFlags.Manual);

                    lastFocus = valueFocus;
                    _parent.textLastMessage.Invoke((MethodInvoker)delegate
                    {
                        _parent.textLastMessage.Text = $"{address} {lastFocus}";
                    });

                }
                else if (address == "/EXPOSURE")
                {
                    int valueExposure = (int)message.Arguments[0];

                    if (lastExposure != valueExposure)
                        cameraControl.Set(CameraControlProperty.Exposure, valueExposure, CameraControlFlags.Manual);

                    lastExposure = valueExposure;
                    _parent.textLastMessage.Invoke((MethodInvoker)delegate
                    {
                        _parent.textLastMessage.Text = $"{address} {lastExposure}";
                    });

                }
                else if (address == "/TRACKING")
                {
                    int valueTracking = (int)message.Arguments[0];

                    // Needs to send OSC to OBSBOT_WebCam software to enable or disable tracking.


                }
                else if (address == "/DumpValues")
                {
                    cameraControl.Get(CameraControlProperty.Pan, out int currentPan, out var flags1);
                    cameraControl.Get(CameraControlProperty.Tilt, out int currentTilt, out var flags2);
                    cameraControl.Get(CameraControlProperty.Zoom, out int currentZoom, out var flags3);
                    cameraControl.Get(CameraControlProperty.Focus, out int currentFocus, out var flags4);

                    //_listener.RemoteIpEndPoint.

                    Console.WriteLine($"Current PTZ: {currentPan} {currentTilt} {currentZoom} {currentFocus} \n");
                    //Console.WriteLine($"PORT: {_listener.RemoteIpEndPoint}");
                }
                else if (address == "/PanTest")
                {
                    int testSpeed = (int)message.Arguments[0];
                    int testPan = (int)message.Arguments[1];

                    panSpeedTest(testSpeed, testPan);

                }

            }
        }

        public double easeInOutQuad(double x)
        {
            return x < 0.5 ? 2 * x * x : 1 - Math.Pow(-2 * x + 2, 2) / 2;
        }

        public double easeOutQuad(double x)
        {
            return 1.0 - (1.0 - x) * (1.0 - x);
        }

        public double easeInQuad(double x)
        {
            return x * x;
        }

        public double easeOutQuint(double x)
        {
            return 1 - Math.Pow(1 - x, 5);
        }

        public double easeInQuint(double x)
        {
            return x * x * x * x;
        }

        public void PTZS(int speed)
        {
            cameraControl.Get(CameraControlProperty.Pan, out int currentPan, out var flags1);
            cameraControl.Get(CameraControlProperty.Tilt, out int currentTilt, out var flags2);
            cameraControl.Get(CameraControlProperty.Zoom, out int currentZoom, out var flags3);

            if (speed == 0)
            {
                calculatedPanSpeed = movementSpeed;
                calculatedTiltSpeed = movementSpeed;
                calculatedZoomSpeed = movementSpeed;

                calculatedPanEndSpeed = movementSpeed / 2;
                calculatedTiltEndSpeed = movementSpeed / 2;
                calculatedZoomEndSpeed = movementSpeed / 2;

                Console.WriteLine($"Current PTZ: {currentPan} {currentTilt} {currentZoom} \n");

                panAmount = Math.Abs(currentPan - destPan);
                tiltAmount = Math.Abs(currentTilt - destTilt);
                zoomAmount = Math.Abs(currentZoom - destZoom);

                Console.WriteLine($"PTZ Distance to move: {panAmount} {tiltAmount} {zoomAmount} \n");

                estPanTime = (int)(panAmount / (movementSpeed / 1000.0f));
                estTiltTime = (int)(tiltAmount / (movementSpeed / 1000.0f));
                estZoomTime = (int)(zoomAmount / (movementSpeed / 1000.0f));

                Console.WriteLine($"Time estimates: {estPanTime} {estTiltTime} {estZoomTime} \n");

                if (estPanTime >= estTiltTime && estPanTime >= estZoomTime)
                {
                    calculatedTiltSpeed = ((int)(((float)tiltAmount / (float)estPanTime) * 1000.0f));
                    calculatedZoomSpeed = ((int)(((float)zoomAmount / (float)estPanTime) * 1000.0f)) + 15;
                    Console.WriteLine($"Pan Time Longest: tiltSpeed: {tiltSpeed} zoomSpeed {zoomSpeed} \n");
                }
                else if (estTiltTime >= estPanTime && estTiltTime >= estZoomTime)
                {
                    calculatedPanSpeed = ((int)(((float)panAmount / (float)estTiltTime) * 1000.0f));
                    calculatedZoomSpeed = ((int)(((float)zoomAmount / (float)estTiltTime) * 1000.0f)) + 15;
                    Console.WriteLine($"Tilt Time Longest: panSpeed: {panSpeed} zoomSpeed {zoomSpeed} \n");
                }
                else if (estZoomTime >= estPanTime && estZoomTime >= estTiltTime)
                {
                    calculatedPanSpeed = ((int)(((float)panAmount / (float)estZoomTime) * 1000.0f));
                    calculatedTiltSpeed = ((int)(((float)tiltAmount / (float)estZoomTime) * 1000.0f));
                    Console.WriteLine($"Zoom Time Longest: panSpeed: {panSpeed} tiltSpeed {tiltSpeed} \n");
                }
                else
                {
                    Console.WriteLine($"This should never happen");
                }

                panSpeed = calculatedPanSpeed;
                tiltSpeed = calculatedTiltSpeed;
                zoomSpeed = calculatedZoomSpeed;

                calculatedPanEndSpeed = panSpeed / 2;
                calculatedTiltEndSpeed = tiltSpeed / 2;
                calculatedZoomEndSpeed = zoomSpeed / 2;

                maxAdditionalPanSpeed = (int)((float)panSpeed * 1.5);
                maxAdditionalTiltSpeed = (int)((float)tiltSpeed * 1.9);
                maxAdditionalZoomSpeed = zoomSpeed * 5;
            }
            else
            {
                panSpeed = speed;
                tiltSpeed = speed;
                zoomSpeed = speed;
            }

            Console.WriteLine($"Pre calc Speeds: {panSpeed} {tiltSpeed} {zoomSpeed}");

            if (currentPan < destPan)
            {
                panSpeed *= -1;
            }

            if (currentTilt < destTilt)
            {
                tiltSpeed *= -1;
            }

            if (currentZoom > destZoom)
            {
                zoomSpeed *= -1;
            }

            Console.WriteLine($"Calculated Speeds: {calculatedPanSpeed} {calculatedTiltSpeed} {calculatedZoomSpeed}");

            if (panAmount > 1)
            {
                gotoActive = true;
                startPanTicks = DateTime.UtcNow.Ticks;
                Console.WriteLine($"Set Pan Speed: {panSpeed}");
                cameraControl.Set((CameraControlProperty)10, panSpeed, CameraControlFlags.Manual);
            }

            if (tiltAmount > 1)
            {
                gotoActive = true;
                startTiltTicks = DateTime.UtcNow.Ticks;
                Console.WriteLine($"Set Tilt Speed: {tiltSpeed}");
                cameraControl.Set((CameraControlProperty)11, tiltSpeed, CameraControlFlags.Manual);
            }

            if (zoomAmount > 1)
            {
                gotoActive = true;
                startZoomTicks = DateTime.UtcNow.Ticks;
                Console.WriteLine($"Set Zoom Speed: {zoomSpeed}");
                cameraControl.Set((CameraControlProperty)14, zoomSpeed, CameraControlFlags.Manual);
            }
        }

        public void panSpeedTest(int speed, int dPan)
        {
            //cameraControl.Set(CameraControlProperty.Pan, -130, CameraControlFlags.Manual);

            destPan = dPan;

            movementSpeed = speed;

            panSpeed = speed;

            cameraControl.Get(CameraControlProperty.Pan, out int currentPan, out var flags1);

            Console.WriteLine($"Current PTZ: {currentPan} \n");

            panAmount = Math.Abs(currentPan - destPan);

            Console.WriteLine($"PTZ Distance to move: {panAmount}\n");

            estPanTime = (int)(panAmount / (movementSpeed / 1000.0f));

            Console.WriteLine($"Time estimates: {estPanTime}\n");


            if (currentPan < destPan)
            {
                panSpeed *= -1;
            }

            Console.WriteLine($"Speeds: {panSpeed}");

            gotoActive = true;
            stopTilt = true;
            stopZoom = true;

            startPanTicks = DateTime.UtcNow.Ticks;
            //Console.WriteLine($"Start pan time: {startPanTicks}");
            cameraControl.Set((CameraControlProperty)10, panSpeed, CameraControlFlags.Manual);

        }

        public void cameraControlLoop()
        {
            long lastTimeTicks = DateTime.UtcNow.Ticks;
           
            try
            {
                while (true)
                {
                    lastTimeTicks = DateTime.UtcNow.Ticks;

                    if (zoomActive)
                    {
                        cameraControl.Get(CameraControlProperty.Zoom, out int currentZoom, out var flags3);

                        if (currentZoom != lastZ)
                        {
                            int cameraNumber = _parent.GetCameraNumberByDevicePath(theDevicePath);
                            var responseMessage = new SharpOSC.OscMessage("/CurrentZoom", cameraNumber, currentZoom);
                            var sender = new SharpOSC.UDPSender("127.0.0.1", 4334);
                            sender.Send(responseMessage);
                        }
                    }
                    else if (gotoActive)
                    {
                        cameraControl.Get(CameraControlProperty.Pan, out int currentPan, out var flags1);
                        cameraControl.Get(CameraControlProperty.Tilt, out int currentTilt, out var flags2);
                        cameraControl.Get(CameraControlProperty.Zoom, out int currentZoom, out var flags3);

                        elapsedTimeMS = (DateTime.UtcNow.Ticks - startPanTicks) / 10000;

                        float panAmountRemaining = Math.Abs(currentPan - destPan);
                        float tiltAmountRemaining = Math.Abs(currentTilt - destTilt);
                        float zoomAmountRemaining = Math.Abs(currentZoom - destZoom);

                        double pctPanComplete = 0;
                        double pctTiltComplete = 0;
                        double pctZoomComplete = 0;

                        if (panAmount > 0)
                        {
                            pctPanComplete = (panAmount - panAmountRemaining) / (float)panAmount * 2;
                        }
                        else
                        {
                            pctPanComplete = 1.0;
                        }

                        if (tiltAmount > 0)
                        {
                            pctTiltComplete = (tiltAmount - tiltAmountRemaining) / (float)tiltAmount * 2;
                        }
                        else
                        {
                            pctTiltComplete = 1.0;
                        }

                        if (zoomAmount > 0)
                        {
                            pctZoomComplete = (zoomAmount - zoomAmountRemaining) / (float)zoomAmount * 2;
                        }
                        else
                        {
                            pctZoomComplete = 1.0;
                        }




                        if (pctPanComplete < 1.0)
                        {
                            panAdditionalSpeed = (int)((double)maxAdditionalPanSpeed * easeInQuad(pctPanComplete));
                            panSpeed = calculatedPanSpeed + panAdditionalSpeed;
                        }
                        else if (pctPanComplete >= 1.0)
                        {
                            pctPanComplete -= 1.0;

                            panAdditionalSpeed = (int)((double)maxAdditionalPanSpeed * easeOutQuad(1.0 - pctPanComplete));
                            panSpeed = calculatedPanEndSpeed + panAdditionalSpeed;
                        }

                        if (pctTiltComplete < 1.0)
                        {
                            tiltAdditionalSpeed = (int)((double)maxAdditionalTiltSpeed * easeInQuad(pctTiltComplete));
                            tiltSpeed = calculatedTiltSpeed + tiltAdditionalSpeed;
                        }
                        else if (pctTiltComplete >= 1.0)
                        {
                            pctTiltComplete -= 1.0;

                            tiltAdditionalSpeed = (int)((double)maxAdditionalTiltSpeed * easeOutQuad(1.0 - pctTiltComplete));
                            tiltSpeed = calculatedTiltEndSpeed + tiltAdditionalSpeed;
                        }

                        if (pctZoomComplete < 1.0)
                        {
                            zoomAdditionalSpeed = (int)((double)maxAdditionalZoomSpeed * easeInQuad(pctZoomComplete));
                            zoomSpeed = calculatedZoomSpeed + zoomAdditionalSpeed;
                        }
                        else if (pctZoomComplete >= 1.0)
                        {
                            pctZoomComplete -= 1.0;
                            zoomAdditionalSpeed = (int)((double)maxAdditionalZoomSpeed * easeOutQuad(1.0 - pctZoomComplete));
                            zoomSpeed = calculatedZoomEndSpeed + zoomAdditionalSpeed;
                        }

                        Console.WriteLine($"Additional Speeds: {panAdditionalSpeed} {tiltAdditionalSpeed} {zoomAdditionalSpeed}");

                        //panSpeed = calculatedPanSpeed + panAdditionalSpeed;
                        //tiltSpeed = calculatedTiltSpeed + tiltAdditionalSpeed;
                        //zoomSpeed = calculatedZoomSpeed + zoomAdditionalSpeed;

                        if (currentPan < destPan)
                        {
                            panSpeed *= -1;
                        }

                        if (currentTilt < destTilt)
                        {
                            tiltSpeed *= -1;
                        }

                        if (currentZoom > destZoom)
                        {
                            zoomSpeed *= -1;
                        }

                        Console.WriteLine($"Remaining amounts: {pctPanComplete:0.000} {pctTiltComplete:0.000} {pctZoomComplete:0.000}");
                        Console.WriteLine($"Tweened speeds: {panSpeed} {tiltSpeed} {zoomSpeed}");

                        //double nowEpochMS = ((double)(DateTime.UtcNow.Ticks) / 10000.0)/1000.0;

                        double nowEpochMS = ((double)(DateTimeOffset.UtcNow - epoch).TotalMilliseconds) / 1000.0;

                        Console.WriteLine($"{nowEpochMS:0000000000.000} panSpeed={panSpeed} tiltSpeed={tiltSpeed} zoomSpeed={zoomSpeed}");

                        _tcpoutBytes = Encoding.ASCII.GetBytes($"{nowEpochMS:0000000000.000} panSpeed={panSpeed} tiltSpeed={tiltSpeed} zoomSpeed={zoomSpeed}\n");
                        _tcpoutStream.Write(_tcpoutBytes, 0, _tcpoutBytes.Length);


                        if (panSpeed != lastPanSpeed)
                        {
                            lastPanSpeed = panSpeed;
                            cameraControl.Set((CameraControlProperty)10, panSpeed, CameraControlFlags.Manual);
                        }

                        if (tiltSpeed != lastTiltSpeed)
                        {
                            lastTiltSpeed = tiltSpeed;
                            cameraControl.Set((CameraControlProperty)11, tiltSpeed, CameraControlFlags.Manual);
                        }

                        if (zoomSpeed != lastZoomSpeed)
                        {
                            lastZoomSpeed = zoomSpeed;
                            cameraControl.Set((CameraControlProperty)13, zoomSpeed, CameraControlFlags.Manual);
                        }

                        double panSpeedFactor = (double)(1.0f - Math.Cos((double)((double)pctPanComplete * (double)Math.PI)));
                        double tiltSpeedFactor = (double)(1.0f - Math.Cos((double)((double)pctTiltComplete * (double)Math.PI)));
                        double zoomSpeedFactor = (double)(1.0f - Math.Cos((double)((double)pctZoomComplete * (double)Math.PI)));

                        //Console.WriteLine($"Speed factors: {panSpeedFactor:0.000} {tiltSpeedFactor:0.000} {zoomSpeedFactor:0.000}");

                        //Console.WriteLine($"Sin of elapsedTimeMS: {Math.Sin(elapsedTimeMS)}");

                        if (!stopPan && panSpeed < 0 && currentPan >= destPan)
                        {
                            Console.WriteLine($"Stop Pan currentPan >= destPan: {currentPan} {currentTilt} {currentZoom} \n");

                            stopPanTicks = DateTime.UtcNow.Ticks;
                            stopPan = true;
                        }
                        else if (!stopPan && panSpeed >= 0 && currentPan <= destPan)
                        {
                            Console.WriteLine($"Stop Pan currentPan >= destPan: {currentPan} {currentTilt} {currentZoom} \n");
                            stopPanTicks = DateTime.UtcNow.Ticks;
                            stopPan = true;
                        }


                        if (!stopTilt && tiltSpeed < 0 && currentTilt >= destTilt)
                        {
                            Console.WriteLine($"Stop Tilt currentTilt <= destTilt: {currentPan} {currentTilt} {currentZoom} \n");
                            stopTiltTicks = DateTime.UtcNow.Ticks;
                            stopTilt = true;
                        }
                        else if (!stopTilt && tiltSpeed >= 0 && currentTilt <= destTilt)
                        {
                            Console.WriteLine($"Stop Tilt currentTilt >= destTilt: {currentPan} {currentTilt} {currentZoom} \n");
                            stopTiltTicks = DateTime.UtcNow.Ticks;
                            stopTilt = true;
                        }


                        if (!stopZoom && zoomSpeed < 0 && (currentZoom <= destZoom || currentZoom <= 0))
                        {
                            Console.WriteLine($"Stop Zoom currentZoom <= destZoom: {destZoom} {currentZoom} {zoomSpeed}\n");
                            stopZoomTicks = DateTime.UtcNow.Ticks;
                            stopZoom = true;
                        }
                        else if (!stopZoom && zoomSpeed >= 0 && (currentZoom >= destZoom || currentZoom >= 100))
                        {
                            Console.WriteLine($"Stop Zoom currentZoom >= destZoom: {destZoom} {currentZoom} {zoomSpeed} \n");
                            stopZoomTicks = DateTime.UtcNow.Ticks;
                            stopZoom = true;
                        }

                        if (stopPan)
                            cameraControl.Set((CameraControlProperty)10, 0, CameraControlFlags.Manual);
                        if (stopTilt)
                            cameraControl.Set((CameraControlProperty)11, 0, CameraControlFlags.Manual);
                        if (stopZoom)
                            cameraControl.Set((CameraControlProperty)13, 0, CameraControlFlags.Manual);


                        /*if (stopPan || stopTilt || stopZoom)
                        {
                            cameraControl.Set((CameraControlProperty)10, 0, CameraControlFlags.Manual);
                            cameraControl.Set((CameraControlProperty)11, 0, CameraControlFlags.Manual);
                            cameraControl.Set((CameraControlProperty)13, 0, CameraControlFlags.Manual);
                        }
                        */
                        /*
                        if (stopPan && Math.Abs(currentPan - destPan) > 2)
                        {
                            stopPan = false;
                            PTZS(10);
                        }
                        if (stopTilt && Math.Abs(currentTilt - destTilt) > 2)
                        {
                            stopTilt = false;
                            PTZS(10);
                        }
                        if (stopZoom && Math.Abs(currentZoom - destZoom) > 2)
                        {
                            stopZoom = false;
                            PTZS(10);
                        }
                        */

                        if (stopPan && stopTilt && stopZoom)
                        {
                            gotoActive = false;
                            stopPan = false;
                            stopTilt = false;
                            stopZoom = false;

                            float panTimeMS = (stopPanTicks - startPanTicks) / 10000;
                            float tiltTimeMS = (stopTiltTicks - startTiltTicks) / 10000;
                            float zoomTimeMS = (stopZoomTicks - startZoomTicks) / 10000;


                            float panPerMS = (float)panAmount / panTimeMS;
                            float tiltPerMS = (float)tiltAmount / tiltTimeMS;
                            float zoomPerMS = (float)zoomAmount / zoomTimeMS;

                            //Console.WriteLine($"Done Moving Speeds: panTime: {panTimeMS} tiltTime: {tiltTimeMS} zoomTime: {zoomTimeMS}");
                            //Console.WriteLine($"Speed per unit: pan: {panPerMS:0.0000} tilt: {tiltPerMS:0.0000} zoom: {zoomPerMS:0.0000}");
                            //Console.WriteLine($"pan, {panSpeed}, {estPanTime}, {panTimeMS}, {panPerMS} ");
                            //Console.WriteLine($"tilt, {tiltSpeed}, {estTiltTime}, {tiltTimeMS}, {tiltPerMS} ");
                            //Console.WriteLine($"zoom, {zoomSpeed}, {estZoomTime}, {zoomTimeMS}, {zoomPerMS} ");


                        }

                        //Console.WriteLine($"PTZ: {currentPan} {currentTilt} {currentZoom} \n");
                    }

                    //gotoActive = true;
                    //stopPan = false;
                    //stopTilt = false;
                    //stopZoom = false;

                    Thread.Sleep(33);
                    //Console.WriteLine($"Ticks per loop: {DateTime.UtcNow.Ticks - lastTimeStamp}");
                }

            }
            catch (Exception ex)
            {
                // log errors
            }

        }
    }

    public class CameraDevice
    {
        public string Name { get; set; }
        public string DevicePath { get; set; }
    }
}
