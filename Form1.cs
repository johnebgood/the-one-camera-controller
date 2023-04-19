/****************************************************************************
While the underlying libraries are covered by LGPL, this sample is released 
as public domain.  It is distributed in the hope that it will be useful, but 
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY 
or FITNESS FOR A PARTICULAR PURPOSE.  

Written by oohansen@gmail.com
*****************************************************************************/

using DirectShowLib;
using SharpOSC;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Text;
using System.Collections.Generic;

namespace TheOneCameraControl
{
    public class Form1 : System.Windows.Forms.Form
    {
        private ComboBox comboDevice;
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
        CameraControl camControl = null;


        public Form1()
        {
            InitializeComponent();

            List<CameraDevice> cameraList = new List<CameraDevice>();



            //enumerate Video Input filters and add them to comboDevice
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

                cameraList.Add(new CameraDevice { Name = (string)device.Name, DevicePath = (string)device.DevicePath });
                //comboDevice.Items.Add(device.Name);
                theDevice = (IBaseFilter)source;
                theDevicePath = device.DevicePath;
                //break;
            }

            comboDevice.DataSource = cameraList;
            comboDevice.DisplayMember = "Name";
            comboDevice.ValueMember = "DevicePath";

            //Select first combobox item
            //if (comboDevice.Items.Count > 0)
            //{
            //    comboDevice.SelectedIndex = 0;
            //}

            //StartServer();

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
            this.comboDevice = new System.Windows.Forms.ComboBox();
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
            this.SuspendLayout();
            // 
            // comboDevice
            // 
            this.comboDevice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboDevice.Location = new System.Drawing.Point(8, 8);
            this.comboDevice.Name = "comboDevice";
            this.comboDevice.Size = new System.Drawing.Size(256, 21);
            this.comboDevice.TabIndex = 0;
            this.comboDevice.SelectedIndexChanged += new System.EventHandler(this.comboDevice_SelectedIndexChanged);
            // 
            // buttonDump
            // 
            this.buttonDump.Location = new System.Drawing.Point(272, 8);
            this.buttonDump.Name = "buttonDump";
            this.buttonDump.Size = new System.Drawing.Size(120, 24);
            this.buttonDump.TabIndex = 2;
            this.buttonDump.Text = "Dump Settings";
            this.buttonDump.Click += new System.EventHandler(this.buttonDump_Click);
            // 
            // buttonStopServer
            // 
            this.buttonStopServer.Location = new System.Drawing.Point(8, 415);
            this.buttonStopServer.Name = "buttonStopServer";
            this.buttonStopServer.Size = new System.Drawing.Size(110, 23);
            this.buttonStopServer.TabIndex = 4;
            this.buttonStopServer.Text = "Stop OSC Server";
            this.buttonStopServer.Click += new System.EventHandler(this.buttonStopServer_Click);
            // 
            // buttonStartServer
            // 
            this.buttonStartServer.Location = new System.Drawing.Point(8, 382);
            this.buttonStartServer.Name = "buttonStartServer";
            this.buttonStartServer.Size = new System.Drawing.Size(111, 23);
            this.buttonStartServer.TabIndex = 5;
            this.buttonStartServer.Text = "Start OSC Server";
            this.buttonStartServer.Click += new System.EventHandler(this.buttonStartServer_Click);
            // 
            // buttonUp
            // 
            this.buttonUp.Location = new System.Drawing.Point(155, 75);
            this.buttonUp.Name = "buttonUp";
            this.buttonUp.Size = new System.Drawing.Size(120, 24);
            this.buttonUp.TabIndex = 7;
            this.buttonUp.Text = "Up";
            this.buttonUp.Click += new System.EventHandler(this.buttonUp_Click);
            // 
            // buttonDown
            // 
            this.buttonDown.Location = new System.Drawing.Point(155, 190);
            this.buttonDown.Name = "buttonDown";
            this.buttonDown.Size = new System.Drawing.Size(120, 24);
            this.buttonDown.TabIndex = 8;
            this.buttonDown.Text = "Down";
            this.buttonDown.Click += new System.EventHandler(this.buttonDown_Click);
            // 
            // buttonLeft
            // 
            this.buttonLeft.Location = new System.Drawing.Point(108, 134);
            this.buttonLeft.Name = "buttonLeft";
            this.buttonLeft.Size = new System.Drawing.Size(44, 24);
            this.buttonLeft.TabIndex = 9;
            this.buttonLeft.Text = "Left";
            this.buttonLeft.Click += new System.EventHandler(this.buttonLeft_Click);
            // 
            // buttonRight
            // 
            this.buttonRight.Location = new System.Drawing.Point(272, 134);
            this.buttonRight.Name = "buttonRight";
            this.buttonRight.Size = new System.Drawing.Size(40, 24);
            this.buttonRight.TabIndex = 10;
            this.buttonRight.Text = "Right";
            this.buttonRight.Click += new System.EventHandler(this.buttonRight_Click);
            // 
            // buttonCenter
            // 
            this.buttonCenter.Location = new System.Drawing.Point(186, 134);
            this.buttonCenter.Name = "buttonCenter";
            this.buttonCenter.Size = new System.Drawing.Size(60, 24);
            this.buttonCenter.TabIndex = 13;
            this.buttonCenter.Text = "Center";
            this.buttonCenter.Click += new System.EventHandler(this.buttonCenter_Click);
            // 
            // buttonRightLimit
            // 
            this.buttonRightLimit.Location = new System.Drawing.Point(322, 134);
            this.buttonRightLimit.Name = "buttonRightLimit";
            this.buttonRightLimit.Size = new System.Drawing.Size(70, 24);
            this.buttonRightLimit.TabIndex = 14;
            this.buttonRightLimit.Text = "Right Limit";
            this.buttonRightLimit.Click += new System.EventHandler(this.buttonRightLimit_Click);
            // 
            // buttonLeftLimit
            // 
            this.buttonLeftLimit.Location = new System.Drawing.Point(28, 134);
            this.buttonLeftLimit.Name = "buttonLeftLimit";
            this.buttonLeftLimit.Size = new System.Drawing.Size(70, 24);
            this.buttonLeftLimit.TabIndex = 15;
            this.buttonLeftLimit.Text = "Left Limit";
            this.buttonLeftLimit.Click += new System.EventHandler(this.buttonLeftLimit_Click);
            // 
            // buttonUpLimit
            // 
            this.buttonUpLimit.Location = new System.Drawing.Point(155, 47);
            this.buttonUpLimit.Name = "buttonUpLimit";
            this.buttonUpLimit.Size = new System.Drawing.Size(120, 24);
            this.buttonUpLimit.TabIndex = 16;
            this.buttonUpLimit.Text = "Up Limit";
            this.buttonUpLimit.Click += new System.EventHandler(this.buttonUpLimit_Click);
            // 
            // buttonDownLimit
            // 
            this.buttonDownLimit.Location = new System.Drawing.Point(155, 218);
            this.buttonDownLimit.Name = "buttonDownLimit";
            this.buttonDownLimit.Size = new System.Drawing.Size(120, 24);
            this.buttonDownLimit.TabIndex = 17;
            this.buttonDownLimit.Text = "Down Limit";
            this.buttonDownLimit.Click += new System.EventHandler(this.buttonDownLimit_Click);
            // 
            // buttonZoomOutFull
            // 
            this.buttonZoomOutFull.Location = new System.Drawing.Point(8, 333);
            this.buttonZoomOutFull.Name = "buttonZoomOutFull";
            this.buttonZoomOutFull.Size = new System.Drawing.Size(120, 24);
            this.buttonZoomOutFull.TabIndex = 23;
            this.buttonZoomOutFull.Text = "Zoom Out Full";
            this.buttonZoomOutFull.Click += new System.EventHandler(this.buttonZoomOutFull_Click);
            // 
            // buttonZoomInFull
            // 
            this.buttonZoomInFull.Location = new System.Drawing.Point(8, 242);
            this.buttonZoomInFull.Name = "buttonZoomInFull";
            this.buttonZoomInFull.Size = new System.Drawing.Size(120, 23);
            this.buttonZoomInFull.TabIndex = 22;
            this.buttonZoomInFull.Text = "Zoom In Full";
            this.buttonZoomInFull.Click += new System.EventHandler(this.buttonZoomInFull_Click);
            // 
            // buttonZoomOut
            // 
            this.buttonZoomOut.Location = new System.Drawing.Point(8, 305);
            this.buttonZoomOut.Name = "buttonZoomOut";
            this.buttonZoomOut.Size = new System.Drawing.Size(120, 24);
            this.buttonZoomOut.TabIndex = 21;
            this.buttonZoomOut.Text = "Zoom Out";
            this.buttonZoomOut.Click += new System.EventHandler(this.buttonZoomOut_Click);
            // 
            // buttonZoomIn
            // 
            this.buttonZoomIn.Location = new System.Drawing.Point(8, 270);
            this.buttonZoomIn.Name = "buttonZoomIn";
            this.buttonZoomIn.Size = new System.Drawing.Size(120, 24);
            this.buttonZoomIn.TabIndex = 20;
            this.buttonZoomIn.Text = "Zoom In";
            this.buttonZoomIn.Click += new System.EventHandler(this.buttonZoomIn_Click);
            // 
            // buttonFastLeft
            // 
            this.buttonFastLeft.Location = new System.Drawing.Point(297, 333);
            this.buttonFastLeft.Name = "buttonFastLeft";
            this.buttonFastLeft.Size = new System.Drawing.Size(103, 24);
            this.buttonFastLeft.TabIndex = 24;
            this.buttonFastLeft.Text = "Fast 10 (-160) Left";
            this.buttonFastLeft.Click += new System.EventHandler(this.buttonFastLeft_Click);
            // 
            // buttonFastRight
            // 
            this.buttonFastRight.Location = new System.Drawing.Point(290, 305);
            this.buttonFastRight.Name = "buttonFastRight";
            this.buttonFastRight.Size = new System.Drawing.Size(110, 24);
            this.buttonFastRight.TabIndex = 25;
            this.buttonFastRight.Text = "Fast 10 (160) Right";
            this.buttonFastRight.Click += new System.EventHandler(this.buttonFastRight_Click);
            // 
            // buttonFastDown
            // 
            this.buttonFastDown.Location = new System.Drawing.Point(297, 233);
            this.buttonFastDown.Name = "buttonFastDown";
            this.buttonFastDown.Size = new System.Drawing.Size(110, 24);
            this.buttonFastDown.TabIndex = 27;
            this.buttonFastDown.Text = "Fast 11 (120) Up";
            this.buttonFastDown.Click += new System.EventHandler(this.button1_Click);
            // 
            // buttonFastUp
            // 
            this.buttonFastUp.Location = new System.Drawing.Point(297, 261);
            this.buttonFastUp.Name = "buttonFastUp";
            this.buttonFastUp.Size = new System.Drawing.Size(103, 24);
            this.buttonFastUp.TabIndex = 26;
            this.buttonFastUp.Text = "Fast 11 (-120) Down";
            this.buttonFastUp.Click += new System.EventHandler(this.button2_Click);
            // 
            // textPort
            // 
            this.textPort.Location = new System.Drawing.Point(160, 385);
            this.textPort.Name = "textPort";
            this.textPort.Size = new System.Drawing.Size(62, 20);
            this.textPort.TabIndex = 28;
            this.textPort.Text = "33333";
            this.textPort.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(130, 387);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 13);
            this.label1.TabIndex = 29;
            this.label1.Text = "Port:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(160, 422);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 30;
            this.label2.Text = "Last Message:";
            // 
            // textLastMessage
            // 
            this.textLastMessage.Location = new System.Drawing.Point(234, 420);
            this.textLastMessage.Name = "textLastMessage";
            this.textLastMessage.Size = new System.Drawing.Size(175, 20);
            this.textLastMessage.TabIndex = 31;
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(657, 671);
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
            this.Controls.Add(this.comboDevice);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "The One Camera Controller";
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
            if (camControl != null)
                camControl.stop();
            return;
        }

        public void StartServer()
        {
            buttonStartServer.Enabled = false;
            buttonStopServer.Enabled = true;
            textPort.Enabled = false;
            camControl = new CameraControl(int.Parse(this.textPort.Text), this);

            return;
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

        private void dumpAll()
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
                Console.WriteLine($"device path: {device.DevicePath}:");
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
                    }
                }
                Console.WriteLine("-----------------------");


            }
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

        private void comboDevice_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            string devicepath = "none";

            //Release COM objects
            if (theDevice != null)
            {
                Marshal.ReleaseComObject(theDevice);
                theDevice = null;
            }
            //Create the filter for the selected video input device
            try
            {
                if (comboDevice.Items.Count > 0)
                {
                    devicepath = (string)comboDevice.SelectedValue;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            theDevice = CreateFilter(FilterCategory.VideoInputDevice, devicepath);
            theDevicePath = devicepath;
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

            cameraControl.Set((CameraControlProperty)10, 160, CameraControlFlags.Manual);
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
    }

    public class CameraControl
    {
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

        private Form1 _parent = null;
        private UDPListener _listener = null;

        // For sending data out to Splunk in real-time.
        private TcpClient _tcpoutClient = null;
        private byte[] _tcpoutBytes = null;
        private NetworkStream _tcpoutStream = null;
        DateTimeOffset epoch = new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero);

        Boolean isSetup = false;
        public CameraControl(int port, Form1 parent)
        {
            _parent = parent;

            HandleOscPacket callback = delegate (OscPacket packet)
            {
                var messageReceived = (OscMessage)packet;

                receiveOSC(messageReceived);
            };

            _listener = new UDPListener(port, callback);
            _tcpoutClient = new TcpClient("10.1.10.173", 30000);
            _tcpoutStream = _tcpoutClient.GetStream();
        }

        public void stop()
        {
            Console.WriteLine("Cleanup on isle 7");
            _listener.Close();
            _tcpoutClient.Close();
        }
        private void receiveOSC(OscMessage message)
        {
            string address = message.Address;
            string remoteIPAddress = message.remoteIPAddress;
            int remotePort = message.remotePort;

            Console.WriteLine($"Remote: {remoteIPAddress}:{remotePort}");

            if (!isSetup)
            {
                object source = null;
                Guid iid = typeof(IBaseFilter).GUID;
                foreach (DsDevice device in DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice))
                {
                    if (device.DevicePath.CompareTo(_parent.theDevicePath) == 0)
                    {
                        try
                        {
                            device.Mon.BindToObject(null, null, ref iid, out source);
                            theDevicePath = device.DevicePath;
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
                theDevice = (IBaseFilter)source;
                cameraControl = theDevice as IAMCameraControl;
                isSetup = true;

                Thread workerThread = new Thread(new ThreadStart(cameraControlLoop));
                // Start secondary thread
                workerThread.Start();
            }
            else if (cameraControl != null)
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
                    destPan = (int)message.Arguments[0];
                    destTilt = (int)message.Arguments[1];
                    destZoom = (int)message.Arguments[2];
                    movementSpeed = (int)message.Arguments[3];

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

        public double easeOutQuint(double x) {
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

                calculatedPanEndSpeed = movementSpeed/2;
                calculatedTiltEndSpeed = movementSpeed/2;
                calculatedZoomEndSpeed = movementSpeed/2;

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
            long lastTimeStamp = DateTime.UtcNow.Ticks;
            try
            {
                while (true)
                {
                    lastTimeStamp = DateTime.UtcNow.Ticks;
                    if (gotoActive)
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

                        double nowEpochMS = ((double)(DateTimeOffset.UtcNow - epoch).TotalMilliseconds)/1000.0;

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

                    Thread.Sleep(50);
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
