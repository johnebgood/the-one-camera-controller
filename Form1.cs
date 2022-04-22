/****************************************************************************
While the underlying libraries are covered by LGPL, this sample is released 
as public domain.  It is distributed in the hope that it will be useful, but 
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY 
or FITNESS FOR A PARTICULAR PURPOSE.  

Written by oohansen@gmail.com
*****************************************************************************/

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.Threading;

using SharpOSC;
using DirectShowLib;

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
		[DllImport("oleaut32.dll", CharSet=CharSet.Unicode, ExactSpelling=true)]
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
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//

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
                comboDevice.Items.Add(device.DevicePath);
                theDevice = (IBaseFilter)source;
                theDevicePath = device.DevicePath;
                //break;
            }

			//Select first combobox item
			if (comboDevice.Items.Count > 0)
			{
				comboDevice.SelectedIndex = 0;
			}

            //StartServer();
            
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
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
            this.comboDevice.Location = new System.Drawing.Point(13, 12);
            this.comboDevice.Name = "comboDevice";
            this.comboDevice.Size = new System.Drawing.Size(409, 28);
            this.comboDevice.TabIndex = 0;
            this.comboDevice.SelectedIndexChanged += new System.EventHandler(this.comboDevice_SelectedIndexChanged);
            // 
            // buttonDump
            // 
            this.buttonDump.Location = new System.Drawing.Point(435, 12);
            this.buttonDump.Name = "buttonDump";
            this.buttonDump.Size = new System.Drawing.Size(192, 35);
            this.buttonDump.TabIndex = 2;
            this.buttonDump.Text = "Dump Settings";
            this.buttonDump.Click += new System.EventHandler(this.buttonDump_Click);
            // 
            // buttonStopServer
            // 
            this.buttonStopServer.Location = new System.Drawing.Point(13, 607);
            this.buttonStopServer.Name = "buttonStopServer";
            this.buttonStopServer.Size = new System.Drawing.Size(176, 33);
            this.buttonStopServer.TabIndex = 4;
            this.buttonStopServer.Text = "Stop OSC Server";
            this.buttonStopServer.Click += new System.EventHandler(this.buttonStopServer_Click);
            // 
            // buttonStartServer
            // 
            this.buttonStartServer.Location = new System.Drawing.Point(13, 559);
            this.buttonStartServer.Name = "buttonStartServer";
            this.buttonStartServer.Size = new System.Drawing.Size(178, 33);
            this.buttonStartServer.TabIndex = 5;
            this.buttonStartServer.Text = "Start OSC Server";
            this.buttonStartServer.Click += new System.EventHandler(this.buttonStartServer_Click);
            // 
            // buttonUp
            // 
            this.buttonUp.Location = new System.Drawing.Point(248, 110);
            this.buttonUp.Name = "buttonUp";
            this.buttonUp.Size = new System.Drawing.Size(192, 35);
            this.buttonUp.TabIndex = 7;
            this.buttonUp.Text = "Up";
            this.buttonUp.Click += new System.EventHandler(this.buttonUp_Click);
            // 
            // buttonDown
            // 
            this.buttonDown.Location = new System.Drawing.Point(248, 278);
            this.buttonDown.Name = "buttonDown";
            this.buttonDown.Size = new System.Drawing.Size(192, 35);
            this.buttonDown.TabIndex = 8;
            this.buttonDown.Text = "Down";
            this.buttonDown.Click += new System.EventHandler(this.buttonDown_Click);
            // 
            // buttonLeft
            // 
            this.buttonLeft.Location = new System.Drawing.Point(173, 196);
            this.buttonLeft.Name = "buttonLeft";
            this.buttonLeft.Size = new System.Drawing.Size(70, 35);
            this.buttonLeft.TabIndex = 9;
            this.buttonLeft.Text = "Left";
            this.buttonLeft.Click += new System.EventHandler(this.buttonLeft_Click);
            // 
            // buttonRight
            // 
            this.buttonRight.Location = new System.Drawing.Point(435, 196);
            this.buttonRight.Name = "buttonRight";
            this.buttonRight.Size = new System.Drawing.Size(64, 35);
            this.buttonRight.TabIndex = 10;
            this.buttonRight.Text = "Right";
            this.buttonRight.Click += new System.EventHandler(this.buttonRight_Click);
            // 
            // buttonCenter
            // 
            this.buttonCenter.Location = new System.Drawing.Point(298, 196);
            this.buttonCenter.Name = "buttonCenter";
            this.buttonCenter.Size = new System.Drawing.Size(95, 35);
            this.buttonCenter.TabIndex = 13;
            this.buttonCenter.Text = "Center";
            this.buttonCenter.Click += new System.EventHandler(this.buttonCenter_Click);
            // 
            // buttonRightLimit
            // 
            this.buttonRightLimit.Location = new System.Drawing.Point(515, 196);
            this.buttonRightLimit.Name = "buttonRightLimit";
            this.buttonRightLimit.Size = new System.Drawing.Size(112, 35);
            this.buttonRightLimit.TabIndex = 14;
            this.buttonRightLimit.Text = "Right Limit";
            this.buttonRightLimit.Click += new System.EventHandler(this.buttonRightLimit_Click);
            // 
            // buttonLeftLimit
            // 
            this.buttonLeftLimit.Location = new System.Drawing.Point(45, 196);
            this.buttonLeftLimit.Name = "buttonLeftLimit";
            this.buttonLeftLimit.Size = new System.Drawing.Size(112, 35);
            this.buttonLeftLimit.TabIndex = 15;
            this.buttonLeftLimit.Text = "Left Limit";
            this.buttonLeftLimit.Click += new System.EventHandler(this.buttonLeftLimit_Click);
            // 
            // buttonUpLimit
            // 
            this.buttonUpLimit.Location = new System.Drawing.Point(248, 69);
            this.buttonUpLimit.Name = "buttonUpLimit";
            this.buttonUpLimit.Size = new System.Drawing.Size(192, 35);
            this.buttonUpLimit.TabIndex = 16;
            this.buttonUpLimit.Text = "Up Limit";
            this.buttonUpLimit.Click += new System.EventHandler(this.buttonUpLimit_Click);
            // 
            // buttonDownLimit
            // 
            this.buttonDownLimit.Location = new System.Drawing.Point(248, 319);
            this.buttonDownLimit.Name = "buttonDownLimit";
            this.buttonDownLimit.Size = new System.Drawing.Size(192, 35);
            this.buttonDownLimit.TabIndex = 17;
            this.buttonDownLimit.Text = "Down Limit";
            this.buttonDownLimit.Click += new System.EventHandler(this.buttonDownLimit_Click);
            // 
            // buttonZoomOutFull
            // 
            this.buttonZoomOutFull.Location = new System.Drawing.Point(13, 487);
            this.buttonZoomOutFull.Name = "buttonZoomOutFull";
            this.buttonZoomOutFull.Size = new System.Drawing.Size(192, 35);
            this.buttonZoomOutFull.TabIndex = 23;
            this.buttonZoomOutFull.Text = "Zoom Out Full";
            this.buttonZoomOutFull.Click += new System.EventHandler(this.buttonZoomOutFull_Click);
            // 
            // buttonZoomInFull
            // 
            this.buttonZoomInFull.Location = new System.Drawing.Point(13, 353);
            this.buttonZoomInFull.Name = "buttonZoomInFull";
            this.buttonZoomInFull.Size = new System.Drawing.Size(192, 35);
            this.buttonZoomInFull.TabIndex = 22;
            this.buttonZoomInFull.Text = "Zoom In Full";
            this.buttonZoomInFull.Click += new System.EventHandler(this.buttonZoomInFull_Click);
            // 
            // buttonZoomOut
            // 
            this.buttonZoomOut.Location = new System.Drawing.Point(13, 446);
            this.buttonZoomOut.Name = "buttonZoomOut";
            this.buttonZoomOut.Size = new System.Drawing.Size(192, 35);
            this.buttonZoomOut.TabIndex = 21;
            this.buttonZoomOut.Text = "Zoom Out";
            this.buttonZoomOut.Click += new System.EventHandler(this.buttonZoomOut_Click);
            // 
            // buttonZoomIn
            // 
            this.buttonZoomIn.Location = new System.Drawing.Point(13, 394);
            this.buttonZoomIn.Name = "buttonZoomIn";
            this.buttonZoomIn.Size = new System.Drawing.Size(192, 35);
            this.buttonZoomIn.TabIndex = 20;
            this.buttonZoomIn.Text = "Zoom In";
            this.buttonZoomIn.Click += new System.EventHandler(this.buttonZoomIn_Click);
            // 
            // buttonFastLeft
            // 
            this.buttonFastLeft.Location = new System.Drawing.Point(475, 487);
            this.buttonFastLeft.Name = "buttonFastLeft";
            this.buttonFastLeft.Size = new System.Drawing.Size(165, 35);
            this.buttonFastLeft.TabIndex = 24;
            this.buttonFastLeft.Text = "Fast 10 (-160) Left";
            this.buttonFastLeft.Click += new System.EventHandler(this.buttonFastLeft_Click);
            // 
            // buttonFastRight
            // 
            this.buttonFastRight.Location = new System.Drawing.Point(464, 446);
            this.buttonFastRight.Name = "buttonFastRight";
            this.buttonFastRight.Size = new System.Drawing.Size(176, 35);
            this.buttonFastRight.TabIndex = 25;
            this.buttonFastRight.Text = "Fast 10 (160) Right";
            this.buttonFastRight.Click += new System.EventHandler(this.buttonFastRight_Click);
            // 
            // buttonFastDown
            // 
            this.buttonFastDown.Location = new System.Drawing.Point(475, 341);
            this.buttonFastDown.Name = "buttonFastDown";
            this.buttonFastDown.Size = new System.Drawing.Size(176, 35);
            this.buttonFastDown.TabIndex = 27;
            this.buttonFastDown.Text = "Fast 11 (120) Up";
            this.buttonFastDown.Click += new System.EventHandler(this.button1_Click);
            // 
            // buttonFastUp
            // 
            this.buttonFastUp.Location = new System.Drawing.Point(475, 382);
            this.buttonFastUp.Name = "buttonFastUp";
            this.buttonFastUp.Size = new System.Drawing.Size(165, 35);
            this.buttonFastUp.TabIndex = 26;
            this.buttonFastUp.Text = "Fast 11 (-120) Down";
            this.buttonFastUp.Click += new System.EventHandler(this.button2_Click);
            // 
            // textPort
            // 
            this.textPort.Location = new System.Drawing.Point(256, 562);
            this.textPort.Name = "textPort";
            this.textPort.Size = new System.Drawing.Size(100, 26);
            this.textPort.TabIndex = 28;
            this.textPort.Text = "33333";
            this.textPort.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(208, 565);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(42, 20);
            this.label1.TabIndex = 29;
            this.label1.Text = "Port:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(256, 617);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 20);
            this.label2.TabIndex = 30;
            this.label2.Text = "Last Message:";
            // 
            // textLastMessage
            // 
            this.textLastMessage.Location = new System.Drawing.Point(375, 614);
            this.textLastMessage.Name = "textLastMessage";
            this.textLastMessage.Size = new System.Drawing.Size(279, 26);
            this.textLastMessage.TabIndex = 31;
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(8, 19);
            this.ClientSize = new System.Drawing.Size(666, 652);
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
            if(camControl != null)
                camControl.stop();
            return;
		}

		public void StartServer()
		{
            buttonStartServer.Enabled = false;
            buttonStopServer.Enabled = true;
            textPort.Enabled = false;
            camControl = new CameraControl(int.Parse(this.textPort.Text),this);

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

                    if (result == 0) {
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
				catch(Exception ex)
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
			//Release COM objects
			if (theDevice != null)
			{
				Marshal.ReleaseComObject(theDevice);
				theDevice = null;
			}
			//Create the filter for the selected video input device
			string devicepath = comboDevice.SelectedItem.ToString();
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
			cameraControl.Set(CameraControlProperty.Tilt, value+10, CameraControlFlags.Manual);
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
			cameraControl.Set(CameraControlProperty.Pan, value-10, CameraControlFlags.Manual);
            Console.WriteLine($"Property: {CameraControlProperty.Pan}, value: {value}");
        }

		private void buttonRight_Click(object sender, EventArgs e)
        {
			var cameraControl = theDevice as IAMCameraControl;
			if (cameraControl == null) return;

			cameraControl.Get(CameraControlProperty.Pan, out int value, out var flags);
			cameraControl.Set(CameraControlProperty.Pan, value+10, CameraControlFlags.Manual);
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
            cameraControl.Set((CameraControlProperty)10, 160, CameraControlFlags.Manual);
            cameraControl.Set(CameraControlProperty.Pan, -100, CameraControlFlags.Manual);

            //Console.WriteLine($"Property: {CameraControlProperty.Pan}, min: {min}, MAX: {max}, steppingDelta: {steppingDelta}");
        }
        private void buttonRightLimit_Click(object sender, EventArgs e)
        {
            var cameraControl = theDevice as IAMCameraControl;
            if (cameraControl == null) return;

            //cameraControl.GetRange(CameraControlProperty.Pan,
            //    out int min, out int max, out int steppingDelta,
            //    out int defaultValue, out var flags);
            cameraControl.Set((CameraControlProperty)10, 160, CameraControlFlags.Manual);
            cameraControl.Set(CameraControlProperty.Pan, 100, CameraControlFlags.Manual);
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
        private IAMCameraControl cameraControl = null;
        private int lastX = 0;
        private int lastY = 0;
        private int lastZoom = 0;
        private Form1 _parent = null;
        private UDPListener _listener = null;

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
            
        }

        public void stop()
        {
            Console.WriteLine("Cleanup on isle 7");
            _listener.Close();
        }
        private void receiveOSC(OscMessage message)
        {
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
            }
            else if(cameraControl != null) {
                string address = message.Address;
                if (address == "/XY")
                {
                    int valueX = (int)message.Arguments[0];
                    int valueY = ((int)message.Arguments[1]) * -1;

                    if (lastX != valueX)
                        cameraControl.Set((CameraControlProperty)10, valueX, CameraControlFlags.Manual);
                    if (lastY != valueY)
                        cameraControl.Set((CameraControlProperty)11, valueY, CameraControlFlags.Manual);

                    lastX = valueX;
                    lastY = valueY;
                    _parent.textLastMessage.Invoke((MethodInvoker)delegate {
                        _parent.textLastMessage.Text = $"{address} {lastX} {lastY}";
                    });
                }
                else if (address == "/ZOOM")
                {
                    int valueZoom = (int)message.Arguments[0];

                    if (lastZoom != valueZoom)
                        cameraControl.Set((CameraControlProperty)13, valueZoom, CameraControlFlags.Manual);

                    lastZoom = valueZoom;
                    _parent.textLastMessage.Invoke((MethodInvoker)delegate {
                        _parent.textLastMessage.Text = $"{address} {lastZoom}";
                    });

                }

            }
        }

        public void OSCThread()
        {
            try
            {

            }
            catch (Exception ex)
            {
                // log errors
            }

        }
        public void cameraControlThread()
        {
            try
            {

            }
            catch (Exception ex)
            {
                // log errors
            }

        }
    }


}
