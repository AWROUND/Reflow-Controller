using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;
using System.Timers;
using System.Windows.Forms;

namespace ReflowController
{
	///<summary>
	/// Project: GenericHid
	/// 
	/// ***********************************************************************
	/// Software License Agreement
	///
	/// Licensor grants any person obtaining a copy of this software ("You") 
	/// a worldwide, royalty-free, non-exclusive license, for the duration of 
	/// the copyright, free of charge, to store and execute the Software in a 
	/// computer system and to incorporate the Software or any portion of it 
	/// in computer programs You write.   
	/// 
	/// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	/// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	/// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	/// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	/// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	/// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
	/// THE SOFTWARE.
	/// ***********************************************************************
	/// 
	/// Author             
	/// Jan Axelson        
	/// 
	/// This software was written using Visual Studio Express 2012 for Windows
	/// Desktop building for the .NET Framework v4.5.
	/// 
	/// Purpose: 
	/// Demonstrates USB communications with a generic HID-class device
	/// 
	/// Requirements:
	/// Windows Vista or later and an attached USB generic Human Interface Device (HID).
	/// (Does not run on Windows XP or earlier because .NET Framework 4.5 will not install on these OSes.) 
	/// 
	/// Description:
	/// Finds an attached device that matches the vendor and product IDs in the form's 
	/// text boxes.
	/// 
	/// Retrieves the device's capabilities.
	/// Sends and requests HID reports.
	/// 
	/// Uses the System.Management class and Windows Management Instrumentation (WMI) to detect 
	/// when a device is attached or removed.
	/// 
	/// A list box displays the data sent and received along with error and status messages.
	/// You can select data to send and 1-time or periodic transfers.
	/// 
	/// You can change the size of the host's Input report buffer and request to use control
	/// transfers only to exchange Input and Output reports.
	/// 
	/// To view additional debugging messages, in the Visual Studio development environment,
	/// from the main menu, select Build > Configuration Manager > Active Solution Configuration 
	/// and select Configuration > Debug and from the main menu, select View > Output.
	/// 
	/// The application uses asynchronous FileStreams to read Input reports and write Output 
	/// reports so the application's main thread doesn't have to wait for the device to retrieve a 
	/// report when the HID driver's buffer is empty or send a report when the device's endpoint is busy. 
	/// 
	/// For code that finds a device and opens handles to it, see the FindTheHid routine in frmMain.cs.
	/// For code that reads from the device, see GetInputReportViaInterruptTransfer, 
	/// GetInputReportViaControlTransfer, and GetFeatureReport in Hid.cs.
	/// For code that writes to the device, see SendInputReportViaInterruptTransfer, 
	/// SendInputReportViaControlTransfer, and SendFeatureReport in Hid.cs.
	/// 
	/// This project includes the following modules:
	/// 
	/// GenericHid.cs - runs the application.
	/// FrmMain.cs - routines specific to the form.
	/// Hid.cs - routines specific to HID communications.
	/// DeviceManagement.cs - routine for obtaining a handle to a device from its GUID.
	/// Debugging.cs - contains a routine for displaying API error messages.
	/// HidDeclarations.cs - Declarations for API functions used by Hid.cs.
	/// FileIODeclarations.cs - Declarations for file-related API functions.
	/// DeviceManagementDeclarations.cs - Declarations for API functions used by DeviceManagement.cs.
	/// DebuggingDeclarations.cs - Declarations for API functions used by Debugging.cs.
	/// 
	/// Companion device firmware for several device CPUs is available from www.Lvr.com/hidpage.htm
	/// You can use any generic HID (not a system mouse or keyboard) that sends and receives reports.
	/// This application will not detect or communicate with non-HID-class devices.
	/// 
	/// For more information about HIDs and USB, and additional example device firmware to use
	/// with this application, visit Lakeview Research at http://Lvr.com 
	/// Send comments, bug reports, etc. to jan@Lvr.com or post on my PORTS forum: http://www.lvr.com/forum 
	/// 
	/// V6.2
	/// 11/12/13
	/// Disabled form buttons when a transfer is in progress.
	/// Other minor edits for clarity and readability.
	/// Will NOT run on Windows XP or earlier, see below.
	/// 
	/// V6.1
	/// 10/28/13
	/// Uses the .NET System.Management class to detect device arrival and removal with WMI instead of Win32 RegisterDeviceNotification.
	/// Other minor edits.
	/// Will NOT run on Windows XP or earlier, see below.
	///  
	/// V6.0
	/// 2/8/13
	/// This version will NOT run on Windows XP or earlier because the code uses .NET Framework 4.5 to support asynchronous FileStreams.
	/// The .NET Framework 4.5 redistributable is compatible with Windows 8, Windows 7 SP1, Windows Server 2008 R2 SP1, 
	/// Windows Server 2008 SP2, Windows Vista SP2, and Windows Vista SP3.
	/// For compatibility, replaced ToInt32 with ToInt64 here:
	/// IntPtr pDevicePathName = new IntPtr(detailDataBuffer.ToInt64() + 4);
	/// and here:
	/// if ((deviceNotificationHandle.ToInt64() == IntPtr.Zero.ToInt64()))
	/// For compatibility if the charset isn't English, added System.Globalization.CultureInfo.InvariantCulture here:
	/// if ((String.Compare(DeviceNameString, mydevicePathName, true, System.Globalization.CultureInfo.InvariantCulture) == 0))
	/// Replaced all Microsoft.VisualBasic namespace code with other .NET equivalents.
	/// Revised user interface for more flexibility.
	/// Moved interrupt-transfer and other HID-specific code to Hid.cs.
	/// Used JetBrains ReSharper to clean up the code: http://www.jetbrains.com/resharper/
	/// 
	/// V5.0
	/// 3/30/11
	/// Replaced ReadFile and WriteFile with FileStreams. Thanks to Joe Dunne and John on my Ports forum for tips on this.
	/// Simplified Hid.cs.
	/// Replaced the form timer with a system timer.
	/// 
	/// V4.6
	/// 1/12/10
	/// Supports Vendor IDs and Product IDs up to FFFFh.
	///
	/// V4.52
	/// 11/10/09
	/// Changed HIDD_ATTRIBUTES to use UInt16
	/// 
	/// V4.51
	/// 2/11/09
	/// Moved Free_ and similar to Finally blocks to ensure they execute.
	/// 
	/// V4.5
	/// 2/9/09
	/// Changes to support 64-bit systems, memory management, and other corrections. 
	/// Big thanks to Peter Nielsen.
	///  
	/// </summary>

	internal class FrmMain : Form
	{
		#region '"Windows Form Designer generated code "'
		public FrmMain()
		//: base()
		{
			// This call is required by the Windows Form Designer.
			InitializeComponent();
		}
		// Form overrides dispose to clean up the component list.
		protected override void Dispose(bool Disposing1)
		{
			if (Disposing1)
			{
				if (components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose(Disposing1);
		}

		// Required by the Windows Form Designer
		private System.ComponentModel.IContainer components;
		public System.Windows.Forms.ToolTip ToolTip1;
		public System.Windows.Forms.TextBox TxtBytesReceived;
		public System.Windows.Forms.GroupBox FraBytesReceived;
		public System.Windows.Forms.CheckBox ChkAutoincrement;
		public System.Windows.Forms.ComboBox CboByte1;
		public System.Windows.Forms.ComboBox CboByte0;
		public System.Windows.Forms.GroupBox FraBytesToSend;
		public System.Windows.Forms.ListBox LstResults;
		// NOTE: The following procedure is required by the Windows Form Designer
		// It can be modified using the Windows Form Designer.
		// Do not modify it using the code editor.   
		internal System.Windows.Forms.GroupBox fraInputReportBufferSize;
		internal System.Windows.Forms.TextBox txtInputReportBufferSize;
		internal System.Windows.Forms.Button cmdInputReportBufferSize;
		internal System.Windows.Forms.GroupBox fraDeviceIdentifiers;
		internal System.Windows.Forms.Label lblVendorID;
		internal System.Windows.Forms.TextBox txtVendorID;
		internal System.Windows.Forms.Label lblProductID;
		internal System.Windows.Forms.TextBox txtProductID;
		private Button cmdGetInputReportInterrupt;
		public GroupBox fraInterruptTransfers;
		private Button cmdPeriodicTransfers;
		public GroupBox fraSendAndGetContinuous;
		private RadioButton radFeature;
		private RadioButton radInputOutputControl;
		private RadioButton radInputOutputInterrupt;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel1;
        private ToolStripStatusLabel toolStripStatusLabel2;
        private Button button1;
        private DataGridView dataGridView1;
        private DataGridViewTextBoxColumn Column1;
        private DataGridViewTextBoxColumn Column2;
        private DataGridViewTextBoxColumn Column3;
        private DataGridViewTextBoxColumn Column4;
        private GroupBox groupBox1;
        private TextBox OvenText;
        private TextBox FanText;
        private TextBox textBox14;
        private TextBox textBox13;
        private TextBox textBox12;
        private TextBox textBox11;
        private TextBox textBox10;
        private TextBox TemperatureText;
        private TextBox SetpointText;
        private TextBox textBox7;
        private TextBox textBox6;
        private TextBox textBox5;
        private TextBox textBox4;
        private TextBox textBox3;
        private TextBox textBox2;
        private TextBox StageText;
        private ToolStripStatusLabel toolStripStatusLabel3;
        private ToolStripStatusLabel toolStripStatusLabel4;
        private ToolStripStatusLabel toolStripStatusLabel5;
        private ToolStripStatusLabel toolStripStatusLabel6;
        private ToolStripStatusLabel toolStripStatusLabel7;
        private ToolStripStatusLabel toolStripStatusLabel8;
        private ToolStripStatusLabel toolStripStatusLabel9;
        private ToolStripStatusLabel toolStripStatusLabel10;
        private Button cmdSendOutputReportInterrupt;

		[System.Diagnostics.DebuggerStepThrough()]
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.FraBytesReceived = new System.Windows.Forms.GroupBox();
            this.TxtBytesReceived = new System.Windows.Forms.TextBox();
            this.FraBytesToSend = new System.Windows.Forms.GroupBox();
            this.ChkAutoincrement = new System.Windows.Forms.CheckBox();
            this.CboByte1 = new System.Windows.Forms.ComboBox();
            this.CboByte0 = new System.Windows.Forms.ComboBox();
            this.LstResults = new System.Windows.Forms.ListBox();
            this.fraInputReportBufferSize = new System.Windows.Forms.GroupBox();
            this.cmdInputReportBufferSize = new System.Windows.Forms.Button();
            this.txtInputReportBufferSize = new System.Windows.Forms.TextBox();
            this.fraDeviceIdentifiers = new System.Windows.Forms.GroupBox();
            this.txtProductID = new System.Windows.Forms.TextBox();
            this.lblProductID = new System.Windows.Forms.Label();
            this.txtVendorID = new System.Windows.Forms.TextBox();
            this.lblVendorID = new System.Windows.Forms.Label();
            this.cmdSendOutputReportInterrupt = new System.Windows.Forms.Button();
            this.cmdGetInputReportInterrupt = new System.Windows.Forms.Button();
            this.fraInterruptTransfers = new System.Windows.Forms.GroupBox();
            this.cmdPeriodicTransfers = new System.Windows.Forms.Button();
            this.fraSendAndGetContinuous = new System.Windows.Forms.GroupBox();
            this.radFeature = new System.Windows.Forms.RadioButton();
            this.radInputOutputControl = new System.Windows.Forms.RadioButton();
            this.radInputOutputInterrupt = new System.Windows.Forms.RadioButton();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.button1 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.OvenText = new System.Windows.Forms.TextBox();
            this.FanText = new System.Windows.Forms.TextBox();
            this.textBox14 = new System.Windows.Forms.TextBox();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.TemperatureText = new System.Windows.Forms.TextBox();
            this.SetpointText = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.StageText = new System.Windows.Forms.TextBox();
            this.toolStripStatusLabel3 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel4 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel5 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel6 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel7 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel8 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel9 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel10 = new System.Windows.Forms.ToolStripStatusLabel();
            this.FraBytesToSend.SuspendLayout();
            this.fraInputReportBufferSize.SuspendLayout();
            this.fraDeviceIdentifiers.SuspendLayout();
            this.fraInterruptTransfers.SuspendLayout();
            this.fraSendAndGetContinuous.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // FraBytesReceived
            // 
            this.FraBytesReceived.BackColor = System.Drawing.SystemColors.Control;
            this.FraBytesReceived.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FraBytesReceived.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FraBytesReceived.Location = new System.Drawing.Point(16, 272);
            this.FraBytesReceived.Name = "FraBytesReceived";
            this.FraBytesReceived.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.FraBytesReceived.Size = new System.Drawing.Size(112, 136);
            this.FraBytesReceived.TabIndex = 4;
            this.FraBytesReceived.TabStop = false;
            this.FraBytesReceived.Text = "Bytes Received";
            // 
            // TxtBytesReceived
            // 
            this.TxtBytesReceived.AcceptsReturn = true;
            this.TxtBytesReceived.BackColor = System.Drawing.SystemColors.Window;
            this.TxtBytesReceived.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.TxtBytesReceived.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtBytesReceived.ForeColor = System.Drawing.SystemColors.WindowText;
            this.TxtBytesReceived.Location = new System.Drawing.Point(16, 520);
            this.TxtBytesReceived.MaxLength = 0;
            this.TxtBytesReceived.Multiline = true;
            this.TxtBytesReceived.Name = "TxtBytesReceived";
            this.TxtBytesReceived.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.TxtBytesReceived.Size = new System.Drawing.Size(440, 96);
            this.TxtBytesReceived.TabIndex = 5;
            // 
            // FraBytesToSend
            // 
            this.FraBytesToSend.BackColor = System.Drawing.SystemColors.Control;
            this.FraBytesToSend.Controls.Add(this.ChkAutoincrement);
            this.FraBytesToSend.Controls.Add(this.CboByte1);
            this.FraBytesToSend.Controls.Add(this.CboByte0);
            this.FraBytesToSend.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FraBytesToSend.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FraBytesToSend.Location = new System.Drawing.Point(16, 128);
            this.FraBytesToSend.Name = "FraBytesToSend";
            this.FraBytesToSend.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.FraBytesToSend.Size = new System.Drawing.Size(160, 136);
            this.FraBytesToSend.TabIndex = 1;
            this.FraBytesToSend.TabStop = false;
            this.FraBytesToSend.Text = "Bytes to Send";
            // 
            // ChkAutoincrement
            // 
            this.ChkAutoincrement.BackColor = System.Drawing.SystemColors.Control;
            this.ChkAutoincrement.Cursor = System.Windows.Forms.Cursors.Default;
            this.ChkAutoincrement.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ChkAutoincrement.ForeColor = System.Drawing.SystemColors.ControlText;
            this.ChkAutoincrement.Location = new System.Drawing.Point(6, 95);
            this.ChkAutoincrement.Name = "ChkAutoincrement";
            this.ChkAutoincrement.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ChkAutoincrement.Size = new System.Drawing.Size(201, 35);
            this.ChkAutoincrement.TabIndex = 6;
            this.ChkAutoincrement.Text = "Autoincrement values";
            this.ChkAutoincrement.UseVisualStyleBackColor = false;
            // 
            // CboByte1
            // 
            this.CboByte1.BackColor = System.Drawing.SystemColors.Window;
            this.CboByte1.Cursor = System.Windows.Forms.Cursors.Default;
            this.CboByte1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CboByte1.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CboByte1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.CboByte1.Location = new System.Drawing.Point(8, 64);
            this.CboByte1.Name = "CboByte1";
            this.CboByte1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.CboByte1.Size = new System.Drawing.Size(101, 22);
            this.CboByte1.TabIndex = 3;
            // 
            // CboByte0
            // 
            this.CboByte0.BackColor = System.Drawing.SystemColors.Window;
            this.CboByte0.Cursor = System.Windows.Forms.Cursors.Default;
            this.CboByte0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CboByte0.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CboByte0.ForeColor = System.Drawing.SystemColors.WindowText;
            this.CboByte0.Location = new System.Drawing.Point(8, 24);
            this.CboByte0.Name = "CboByte0";
            this.CboByte0.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.CboByte0.Size = new System.Drawing.Size(101, 22);
            this.CboByte0.TabIndex = 2;
            // 
            // LstResults
            // 
            this.LstResults.BackColor = System.Drawing.SystemColors.Window;
            this.LstResults.Cursor = System.Windows.Forms.Cursors.Default;
            this.LstResults.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LstResults.ForeColor = System.Drawing.SystemColors.WindowText;
            this.LstResults.HorizontalScrollbar = true;
            this.LstResults.ItemHeight = 14;
            this.LstResults.Location = new System.Drawing.Point(0, 440);
            this.LstResults.Name = "LstResults";
            this.LstResults.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.LstResults.Size = new System.Drawing.Size(440, 74);
            this.LstResults.TabIndex = 0;
            // 
            // fraInputReportBufferSize
            // 
            this.fraInputReportBufferSize.Controls.Add(this.cmdInputReportBufferSize);
            this.fraInputReportBufferSize.Controls.Add(this.txtInputReportBufferSize);
            this.fraInputReportBufferSize.Location = new System.Drawing.Point(248, 16);
            this.fraInputReportBufferSize.Name = "fraInputReportBufferSize";
            this.fraInputReportBufferSize.Size = new System.Drawing.Size(208, 96);
            this.fraInputReportBufferSize.TabIndex = 9;
            this.fraInputReportBufferSize.TabStop = false;
            this.fraInputReportBufferSize.Text = "Input Report Buffer Size";
            // 
            // cmdInputReportBufferSize
            // 
            this.cmdInputReportBufferSize.Location = new System.Drawing.Point(96, 32);
            this.cmdInputReportBufferSize.Name = "cmdInputReportBufferSize";
            this.cmdInputReportBufferSize.Size = new System.Drawing.Size(96, 56);
            this.cmdInputReportBufferSize.TabIndex = 1;
            this.cmdInputReportBufferSize.Text = "Change Buffer Size";
            this.cmdInputReportBufferSize.Click += new System.EventHandler(this.cmdInputReportBufferSize_Click);
            // 
            // txtInputReportBufferSize
            // 
            this.txtInputReportBufferSize.Location = new System.Drawing.Point(16, 40);
            this.txtInputReportBufferSize.Name = "txtInputReportBufferSize";
            this.txtInputReportBufferSize.Size = new System.Drawing.Size(56, 20);
            this.txtInputReportBufferSize.TabIndex = 0;
            // 
            // fraDeviceIdentifiers
            // 
            this.fraDeviceIdentifiers.Controls.Add(this.txtProductID);
            this.fraDeviceIdentifiers.Controls.Add(this.lblProductID);
            this.fraDeviceIdentifiers.Controls.Add(this.txtVendorID);
            this.fraDeviceIdentifiers.Controls.Add(this.lblVendorID);
            this.fraDeviceIdentifiers.Location = new System.Drawing.Point(16, 16);
            this.fraDeviceIdentifiers.Name = "fraDeviceIdentifiers";
            this.fraDeviceIdentifiers.Size = new System.Drawing.Size(208, 96);
            this.fraDeviceIdentifiers.TabIndex = 10;
            this.fraDeviceIdentifiers.TabStop = false;
            this.fraDeviceIdentifiers.Text = "Device Identifiers";
            // 
            // txtProductID
            // 
            this.txtProductID.Location = new System.Drawing.Point(120, 56);
            this.txtProductID.Name = "txtProductID";
            this.txtProductID.Size = new System.Drawing.Size(72, 20);
            this.txtProductID.TabIndex = 3;
            this.txtProductID.Text = "1299";
            this.txtProductID.TextChanged += new System.EventHandler(this.txtProductID_TextChanged);
            // 
            // lblProductID
            // 
            this.lblProductID.Location = new System.Drawing.Point(16, 56);
            this.lblProductID.Name = "lblProductID";
            this.lblProductID.Size = new System.Drawing.Size(112, 23);
            this.lblProductID.TabIndex = 2;
            this.lblProductID.Text = "Product ID (hex):";
            // 
            // txtVendorID
            // 
            this.txtVendorID.Location = new System.Drawing.Point(120, 24);
            this.txtVendorID.Name = "txtVendorID";
            this.txtVendorID.Size = new System.Drawing.Size(72, 20);
            this.txtVendorID.TabIndex = 1;
            this.txtVendorID.Text = "0925";
            this.txtVendorID.TextChanged += new System.EventHandler(this.txtVendorID_TextChanged);
            // 
            // lblVendorID
            // 
            this.lblVendorID.Location = new System.Drawing.Point(16, 24);
            this.lblVendorID.Name = "lblVendorID";
            this.lblVendorID.Size = new System.Drawing.Size(112, 23);
            this.lblVendorID.TabIndex = 0;
            this.lblVendorID.Text = "Vendor ID (hex):";
            // 
            // cmdSendOutputReportInterrupt
            // 
            this.cmdSendOutputReportInterrupt.Location = new System.Drawing.Point(10, 27);
            this.cmdSendOutputReportInterrupt.Name = "cmdSendOutputReportInterrupt";
            this.cmdSendOutputReportInterrupt.Size = new System.Drawing.Size(118, 50);
            this.cmdSendOutputReportInterrupt.TabIndex = 12;
            this.cmdSendOutputReportInterrupt.Text = "Send Output Report";
            this.cmdSendOutputReportInterrupt.UseVisualStyleBackColor = true;
            this.cmdSendOutputReportInterrupt.Click += new System.EventHandler(this.cmdSendOutputReportInterrupt_Click);
            // 
            // cmdGetInputReportInterrupt
            // 
            this.cmdGetInputReportInterrupt.Location = new System.Drawing.Point(10, 83);
            this.cmdGetInputReportInterrupt.Name = "cmdGetInputReportInterrupt";
            this.cmdGetInputReportInterrupt.Size = new System.Drawing.Size(118, 50);
            this.cmdGetInputReportInterrupt.TabIndex = 13;
            this.cmdGetInputReportInterrupt.Text = "Get Input Report";
            this.cmdGetInputReportInterrupt.UseVisualStyleBackColor = true;
            this.cmdGetInputReportInterrupt.Click += new System.EventHandler(this.cmdGetInputReportInterrupt_Click);
            // 
            // fraInterruptTransfers
            // 
            this.fraInterruptTransfers.BackColor = System.Drawing.SystemColors.Control;
            this.fraInterruptTransfers.Controls.Add(this.cmdSendOutputReportInterrupt);
            this.fraInterruptTransfers.Controls.Add(this.cmdGetInputReportInterrupt);
            this.fraInterruptTransfers.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fraInterruptTransfers.ForeColor = System.Drawing.SystemColors.ControlText;
            this.fraInterruptTransfers.Location = new System.Drawing.Point(194, 128);
            this.fraInterruptTransfers.Name = "fraInterruptTransfers";
            this.fraInterruptTransfers.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.fraInterruptTransfers.Size = new System.Drawing.Size(145, 152);
            this.fraInterruptTransfers.TabIndex = 14;
            this.fraInterruptTransfers.TabStop = false;
            this.fraInterruptTransfers.Text = "Interrupt Transfers";
            // 
            // cmdPeriodicTransfers
            // 
            this.cmdPeriodicTransfers.Location = new System.Drawing.Point(153, 36);
            this.cmdPeriodicTransfers.Name = "cmdPeriodicTransfers";
            this.cmdPeriodicTransfers.Size = new System.Drawing.Size(118, 51);
            this.cmdPeriodicTransfers.TabIndex = 16;
            this.cmdPeriodicTransfers.Text = "Start";
            this.cmdPeriodicTransfers.UseVisualStyleBackColor = true;
            this.cmdPeriodicTransfers.Click += new System.EventHandler(this.cmdPeriodicTransfers_Click);
            // 
            // fraSendAndGetContinuous
            // 
            this.fraSendAndGetContinuous.BackColor = System.Drawing.SystemColors.Control;
            this.fraSendAndGetContinuous.Controls.Add(this.radFeature);
            this.fraSendAndGetContinuous.Controls.Add(this.radInputOutputControl);
            this.fraSendAndGetContinuous.Controls.Add(this.radInputOutputInterrupt);
            this.fraSendAndGetContinuous.Controls.Add(this.cmdPeriodicTransfers);
            this.fraSendAndGetContinuous.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fraSendAndGetContinuous.ForeColor = System.Drawing.SystemColors.ControlText;
            this.fraSendAndGetContinuous.Location = new System.Drawing.Point(161, 296);
            this.fraSendAndGetContinuous.Name = "fraSendAndGetContinuous";
            this.fraSendAndGetContinuous.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.fraSendAndGetContinuous.Size = new System.Drawing.Size(295, 112);
            this.fraSendAndGetContinuous.TabIndex = 17;
            this.fraSendAndGetContinuous.TabStop = false;
            this.fraSendAndGetContinuous.Text = "Send and Get Continuous";
            // 
            // radFeature
            // 
            this.radFeature.AutoSize = true;
            this.radFeature.Location = new System.Drawing.Point(17, 76);
            this.radFeature.Name = "radFeature";
            this.radFeature.Size = new System.Drawing.Size(62, 18);
            this.radFeature.TabIndex = 19;
            this.radFeature.TabStop = true;
            this.radFeature.Text = "Feature";
            this.radFeature.UseVisualStyleBackColor = true;
            this.radFeature.CheckedChanged += new System.EventHandler(this.radFeature_CheckedChanged);
            // 
            // radInputOutputControl
            // 
            this.radInputOutputControl.AutoSize = true;
            this.radInputOutputControl.Location = new System.Drawing.Point(17, 52);
            this.radInputOutputControl.Name = "radInputOutputControl";
            this.radInputOutputControl.Size = new System.Drawing.Size(120, 18);
            this.radInputOutputControl.TabIndex = 18;
            this.radInputOutputControl.TabStop = true;
            this.radInputOutputControl.Text = "Input Output Control";
            this.radInputOutputControl.UseVisualStyleBackColor = true;
            this.radInputOutputControl.CheckedChanged += new System.EventHandler(this.radInputOutputControl_CheckedChanged);
            // 
            // radInputOutputInterrupt
            // 
            this.radInputOutputInterrupt.AutoSize = true;
            this.radInputOutputInterrupt.Location = new System.Drawing.Point(17, 28);
            this.radInputOutputInterrupt.Name = "radInputOutputInterrupt";
            this.radInputOutputInterrupt.Size = new System.Drawing.Size(126, 18);
            this.radInputOutputInterrupt.TabIndex = 17;
            this.radInputOutputInterrupt.TabStop = true;
            this.radInputOutputInterrupt.Text = "Input Output Interrupt";
            this.radInputOutputInterrupt.UseVisualStyleBackColor = true;
            this.radInputOutputInterrupt.CheckedChanged += new System.EventHandler(this.radInputOutputInterrupt_CheckedChanged);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2,
            this.toolStripStatusLabel3,
            this.toolStripStatusLabel4,
            this.toolStripStatusLabel5,
            this.toolStripStatusLabel6,
            this.toolStripStatusLabel7,
            this.toolStripStatusLabel8,
            this.toolStripStatusLabel9,
            this.toolStripStatusLabel10});
            this.statusStrip1.Location = new System.Drawing.Point(0, 734);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(972, 22);
            this.statusStrip1.TabIndex = 18;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(129, 17);
            this.toolStripStatusLabel1.Text = "Communication Status";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel2.Text = "toolStripStatusLabel2";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 676);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 19;
            this.button1.Text = "Reflow";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Reflow_ClickAsync);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4});
            this.dataGridView1.Location = new System.Drawing.Point(524, 16);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(429, 593);
            this.dataGridView1.TabIndex = 20;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Time";
            this.Column1.Name = "Column1";
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Current Temp";
            this.Column2.Name = "Column2";
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Setpoint";
            this.Column3.Name = "Column3";
            // 
            // Column4
            // 
            this.Column4.HeaderText = "Heater";
            this.Column4.Name = "Column4";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.OvenText);
            this.groupBox1.Controls.Add(this.FanText);
            this.groupBox1.Controls.Add(this.textBox14);
            this.groupBox1.Controls.Add(this.textBox13);
            this.groupBox1.Controls.Add(this.textBox12);
            this.groupBox1.Controls.Add(this.textBox11);
            this.groupBox1.Controls.Add(this.textBox10);
            this.groupBox1.Controls.Add(this.TemperatureText);
            this.groupBox1.Controls.Add(this.SetpointText);
            this.groupBox1.Controls.Add(this.textBox7);
            this.groupBox1.Controls.Add(this.textBox6);
            this.groupBox1.Controls.Add(this.textBox5);
            this.groupBox1.Controls.Add(this.textBox4);
            this.groupBox1.Controls.Add(this.textBox3);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.StageText);
            this.groupBox1.Location = new System.Drawing.Point(525, 611);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(435, 120);
            this.groupBox1.TabIndex = 54;
            this.groupBox1.TabStop = false;
            // 
            // OvenText
            // 
            this.OvenText.Location = new System.Drawing.Point(328, 39);
            this.OvenText.Name = "OvenText";
            this.OvenText.Size = new System.Drawing.Size(100, 20);
            this.OvenText.TabIndex = 69;
            // 
            // FanText
            // 
            this.FanText.Location = new System.Drawing.Point(328, 65);
            this.FanText.Name = "FanText";
            this.FanText.Size = new System.Drawing.Size(100, 20);
            this.FanText.TabIndex = 68;
            // 
            // textBox14
            // 
            this.textBox14.Location = new System.Drawing.Point(328, 91);
            this.textBox14.Name = "textBox14";
            this.textBox14.Size = new System.Drawing.Size(100, 20);
            this.textBox14.TabIndex = 67;
            // 
            // textBox13
            // 
            this.textBox13.Location = new System.Drawing.Point(222, 15);
            this.textBox13.Name = "textBox13";
            this.textBox13.Size = new System.Drawing.Size(100, 20);
            this.textBox13.TabIndex = 66;
            this.textBox13.Text = "Process Stage";
            this.textBox13.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox12
            // 
            this.textBox12.Location = new System.Drawing.Point(222, 39);
            this.textBox12.Name = "textBox12";
            this.textBox12.Size = new System.Drawing.Size(100, 20);
            this.textBox12.TabIndex = 65;
            this.textBox12.Text = "Oven";
            this.textBox12.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox11
            // 
            this.textBox11.Location = new System.Drawing.Point(222, 65);
            this.textBox11.Name = "textBox11";
            this.textBox11.Size = new System.Drawing.Size(100, 20);
            this.textBox11.TabIndex = 64;
            this.textBox11.Text = "Fan";
            this.textBox11.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(222, 91);
            this.textBox10.Name = "textBox10";
            this.textBox10.Size = new System.Drawing.Size(100, 20);
            this.textBox10.TabIndex = 63;
            this.textBox10.Text = "Aux";
            this.textBox10.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // TemperatureText
            // 
            this.TemperatureText.Location = new System.Drawing.Point(116, 13);
            this.TemperatureText.Name = "TemperatureText";
            this.TemperatureText.Size = new System.Drawing.Size(100, 20);
            this.TemperatureText.TabIndex = 62;
            // 
            // SetpointText
            // 
            this.SetpointText.Location = new System.Drawing.Point(116, 39);
            this.SetpointText.Name = "SetpointText";
            this.SetpointText.Size = new System.Drawing.Size(100, 20);
            this.SetpointText.TabIndex = 61;
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(116, 65);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(100, 20);
            this.textBox7.TabIndex = 60;
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(116, 91);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(100, 20);
            this.textBox6.TabIndex = 59;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(10, 91);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(100, 20);
            this.textBox5.TabIndex = 58;
            this.textBox5.Text = "Elapsed Time";
            this.textBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(10, 65);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(100, 20);
            this.textBox4.TabIndex = 57;
            this.textBox4.Text = "Stage Time";
            this.textBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(10, 39);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(100, 20);
            this.textBox3.TabIndex = 56;
            this.textBox3.Text = "Setpoint �C";
            this.textBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(10, 13);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 20);
            this.textBox2.TabIndex = 55;
            this.textBox2.Text = "Temperature �C";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // StageText
            // 
            this.StageText.Location = new System.Drawing.Point(328, 15);
            this.StageText.Name = "StageText";
            this.StageText.Size = new System.Drawing.Size(100, 20);
            this.StageText.TabIndex = 54;
            // 
            // toolStripStatusLabel3
            // 
            this.toolStripStatusLabel3.AutoSize = false;
            this.toolStripStatusLabel3.MergeAction = System.Windows.Forms.MergeAction.Replace;
            this.toolStripStatusLabel3.Name = "toolStripStatusLabel3";
            this.toolStripStatusLabel3.Size = new System.Drawing.Size(30, 17);
            this.toolStripStatusLabel3.Text = "Kp=";
            // 
            // toolStripStatusLabel4
            // 
            this.toolStripStatusLabel4.AutoSize = false;
            this.toolStripStatusLabel4.Name = "toolStripStatusLabel4";
            this.toolStripStatusLabel4.Size = new System.Drawing.Size(30, 17);
            this.toolStripStatusLabel4.Text = "toolStripStatusLabel4";
            // 
            // toolStripStatusLabel5
            // 
            this.toolStripStatusLabel5.AutoSize = false;
            this.toolStripStatusLabel5.Name = "toolStripStatusLabel5";
            this.toolStripStatusLabel5.Size = new System.Drawing.Size(30, 17);
            this.toolStripStatusLabel5.Text = "Ki=";
            // 
            // toolStripStatusLabel6
            // 
            this.toolStripStatusLabel6.AutoSize = false;
            this.toolStripStatusLabel6.Name = "toolStripStatusLabel6";
            this.toolStripStatusLabel6.Size = new System.Drawing.Size(30, 17);
            this.toolStripStatusLabel6.Text = "toolStripStatusLabel6";
            // 
            // toolStripStatusLabel7
            // 
            this.toolStripStatusLabel7.AutoSize = false;
            this.toolStripStatusLabel7.Name = "toolStripStatusLabel7";
            this.toolStripStatusLabel7.Size = new System.Drawing.Size(30, 17);
            this.toolStripStatusLabel7.Text = "Kd=";
            // 
            // toolStripStatusLabel8
            // 
            this.toolStripStatusLabel8.AutoSize = false;
            this.toolStripStatusLabel8.Name = "toolStripStatusLabel8";
            this.toolStripStatusLabel8.Size = new System.Drawing.Size(30, 17);
            this.toolStripStatusLabel8.Text = "toolStripStatusLabel8";
            // 
            // toolStripStatusLabel9
            // 
            this.toolStripStatusLabel9.Name = "toolStripStatusLabel9";
            this.toolStripStatusLabel9.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel9.Text = "toolStripStatusLabel9";
            // 
            // toolStripStatusLabel10
            // 
            this.toolStripStatusLabel10.Name = "toolStripStatusLabel10";
            this.toolStripStatusLabel10.Size = new System.Drawing.Size(124, 17);
            this.toolStripStatusLabel10.Text = "toolStripStatusLabel10";
            // 
            // FrmMain
            // 
            this.ClientSize = new System.Drawing.Size(972, 756);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.LstResults);
            this.Controls.Add(this.fraSendAndGetContinuous);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.TxtBytesReceived);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.fraInterruptTransfers);
            this.Controls.Add(this.fraDeviceIdentifiers);
            this.Controls.Add(this.fraInputReportBufferSize);
            this.Controls.Add(this.FraBytesReceived);
            this.Controls.Add(this.FraBytesToSend);
            this.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Location = new System.Drawing.Point(21, 28);
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Reflow Controller";
            this.Closed += new System.EventHandler(this.frmMain_Closed);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.FraBytesToSend.ResumeLayout(false);
            this.fraInputReportBufferSize.ResumeLayout(false);
            this.fraInputReportBufferSize.PerformLayout();
            this.fraDeviceIdentifiers.ResumeLayout(false);
            this.fraDeviceIdentifiers.PerformLayout();
            this.fraInterruptTransfers.ResumeLayout(false);
            this.fraSendAndGetContinuous.ResumeLayout(false);
            this.fraSendAndGetContinuous.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		private Boolean _deviceDetected;
		private IntPtr _deviceNotificationHandle;
		private FileStream _deviceData;
		private FormActions _formActions;
		private SafeFileHandle _hidHandle;
		private String _hidUsage;
		private ManagementEventWatcher _deviceArrivedWatcher;
		private Boolean _deviceHandleObtained;
		private ManagementEventWatcher _deviceRemovedWatcher;
		private Int32 _myProductId;
		private Int32 _myVendorId;
		private Boolean _periodicTransfersRequested;
		private ReportReadOrWritten _readOrWritten;
		private ReportTypes _reportType;
		private SendOrGet _sendOrGet;
		private Boolean _transferInProgress;
		private TransferTypes _transferType;

        private string state;
        private Byte State;
        private string temperature;
        private Byte Temperature;
        private string setpoint;
        private Byte Setpoint;
        private string oven;
        private Byte Oven;
        private string fan;
        private Byte Fan;
        private string elapsed1;
        private Byte Elapsed1;
        private string elapsed2;
        private Byte Elapsed2;
        private string start;
        private Byte Start;
        private string kp;
        private Byte Kp;
        private string ki;
        private Byte Ki;
        private string kd;
        private Byte Kd;
        private string cycleTime;
        private Byte CycleTime;
        private string pTerm;
        private Byte PTerm;
        private string iTerm;
        private Byte ITerm;
        private string dTerm;
        private Byte DTerm;
        private string output;
        private Byte Output;

        private static System.Timers.Timer _periodicTransfers;

		private readonly Debugging _myDebugging = new Debugging(); //  For viewing results of API calls via Debug.Write.
		private readonly DeviceManagement _myDeviceManagement = new DeviceManagement();
		private Hid _myHid = new Hid();

		private enum FormActions
		{
			AddItemToListBox,
			DisableInputReportBufferSize,
			EnableGetInputReportInterruptTransfer,
			EnableInputReportBufferSize,
			EnableSendOutputReportInterrupt,
			ScrollToBottomOfListBox,
			SetInputReportBufferSize,
            ChangeBackgroundRed,
            ChangeBackgroundGreen
        }

		private enum ReportReadOrWritten
		{
			Read,
			Written
		}

		private enum ReportTypes
		{
			Input,
			Output,
			Feature
		}

		private enum SendOrGet
		{
			Send,
			Get
		}

		private enum TransferTypes
		{
			Control,
			Interrupt
		}

		private enum WmiDeviceProperties
		{
			Name,
			Caption,
			Description,
			Manufacturer,
			PNPDeviceID,
			DeviceID,
			ClassGUID
		}

		internal FrmMain FrmMy;

		//  This delegate has the same parameters as AccessForm.
		//  Used in accessing the application's form from a different thread.

		private delegate void MarshalDataToForm(FormActions action, String textToAdd);

		///  <summary>
		///  Performs various application-specific functions that
		///  involve accessing the application's form.
		///  </summary>
		///  
		///  <param name="action"> a FormActions member that names the action to perform on the form</param>
		///  <param name="formText"> text that the form displays or the code uses for 
		///  another purpose. Actions that don't use text ignore this parameter. </param>

		private void AccessForm(FormActions action, String formText)
		{
			try
			{
				//  Select an action to perform on the form:

				switch (action)
				{
					case FormActions.AddItemToListBox:

						LstResults.Items.Add(formText);
						break;

					case FormActions.DisableInputReportBufferSize:

						cmdInputReportBufferSize.Enabled = false;
						break;

					case FormActions.EnableGetInputReportInterruptTransfer:

						cmdGetInputReportInterrupt.Enabled = true;
						break;

					case FormActions.EnableInputReportBufferSize:

						cmdInputReportBufferSize.Enabled = true;
						break;

					case FormActions.EnableSendOutputReportInterrupt:

						cmdSendOutputReportInterrupt.Enabled = true;
						break;

					case FormActions.ScrollToBottomOfListBox:

						LstResults.SelectedIndex = LstResults.Items.Count - 1;
						break;

					case FormActions.SetInputReportBufferSize:

						txtInputReportBufferSize.Text = formText;
						break;

                    case FormActions.ChangeBackgroundRed:

                        toolStripStatusLabel2.BackColor = System.Drawing.Color.Red;
                        break;

                    case FormActions.ChangeBackgroundGreen:

                        toolStripStatusLabel2.BackColor = System.Drawing.Color.Green;
                        break;
                }
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Add a handler to detect arrival of devices using WMI.
		///  </summary>

		private void AddDeviceArrivedHandler()
		{
			const Int32 pollingIntervalSeconds = 3;
			var scope = new ManagementScope("root\\CIMV2");
			scope.Options.EnablePrivileges = true;

			try
			{
				var q = new WqlEventQuery();
				q.EventClassName = "__InstanceCreationEvent";
				q.WithinInterval = new TimeSpan(0, 0, pollingIntervalSeconds);
				q.Condition = @"TargetInstance ISA 'Win32_USBControllerdevice'";
				_deviceArrivedWatcher = new ManagementEventWatcher(scope, q);
				_deviceArrivedWatcher.EventArrived += DeviceAdded;

				_deviceArrivedWatcher.Start();
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.Message);
				if (_deviceArrivedWatcher != null)
					_deviceArrivedWatcher.Stop();
			}
		}

		///  <summary>
		///  Add a handler to detect removal of devices using WMI.
		///  </summary>

		private void AddDeviceRemovedHandler()
		{
			const Int32 pollingIntervalSeconds = 3;
			var scope = new ManagementScope("root\\CIMV2");
			scope.Options.EnablePrivileges = true;

			try
			{
				var q = new WqlEventQuery();
				q.EventClassName = "__InstanceDeletionEvent";
				q.WithinInterval = new TimeSpan(0, 0, pollingIntervalSeconds);
				q.Condition = @"TargetInstance ISA 'Win32_USBControllerdevice'";
				_deviceRemovedWatcher = new ManagementEventWatcher(scope, q);
				_deviceRemovedWatcher.EventArrived += DeviceRemoved;
				_deviceRemovedWatcher.Start();
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.Message);
				if (_deviceRemovedWatcher != null)
					_deviceRemovedWatcher.Stop();
			}
		}

		/// <summary>
		/// Close the handle and FileStreams for a device.
		/// </summary>
		/// 
		private void CloseCommunications()
		{
			if (_deviceData != null)
			{
				_deviceData.Close();
			}

			if ((_hidHandle != null) && (!(_hidHandle.IsInvalid)))
			{
				_hidHandle.Close();
			}

			// The next attempt to communicate will get a new handle and FileStreams.

			_deviceHandleObtained = false;
		}

		///  <summary>
		///  Search for a specific device.
		///  </summary>

		private void cmdFindDevice_Click(Object sender, EventArgs e)
		{
			try
			{
				if (_transferInProgress)
				{
					DisplayTransferInProgressMessage();
				}
				else
				{
					_deviceDetected = FindDeviceUsingWmi();
					if (_deviceDetected)
					{
						FindTheHid();
					}
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

        ///  <summary>
        ///  Search for a specific device.
        ///  </summary>

        private async void Reflow_ClickAsync(Object sender, EventArgs e)
        {
            const Int32 readTimeout = 5000;

            String byteValue = null;
            Byte[] inputReportBuffer = null;
            State = 1;

            while (State != 0)
            {
                try
                {
                    Boolean success = false;

                    //  If the device hasn't been detected, was removed, or timed out on a previous attempt
                    //  to access it, look for the device.

                    if (!_deviceHandleObtained)
                    {
                        _deviceHandleObtained = FindTheHid();
                    }

                    if (_deviceHandleObtained)
                    {
                        //  Don't attempt to exchange reports if valid handles aren't available
                        //  (as for a mouse or keyboard under Windows 2000 and later.)

                        if (!_hidHandle.IsInvalid)
                        {
                            //  Read an Input report.

                            //  Don't attempt to send an Input report if the HID has no Input report.
                            //  (The HID spec requires all HIDs to have an interrupt IN endpoint,
                            //  which suggests that all HIDs must support Input reports.)

                            if (_myHid.Capabilities.InputReportByteLength > 0)
                            {
                                //  Set the size of the Input report buffer. 

                                inputReportBuffer = new Byte[_myHid.Capabilities.InputReportByteLength];

                                if (_transferType.Equals(TransferTypes.Control))
                                {
                                    {
                                        _transferInProgress = true;

                                        //  Read a report using a control transfer.

                                        success = _myHid.GetInputReportViaControlTransfer(_hidHandle, ref inputReportBuffer);
                                        //cmdGetInputReportControl.Enabled = true;
                                        _transferInProgress = false;
                                    }
                                }
                                else
                                {
                                    {
                                        _transferInProgress = true;

                                        //  Read a report using interrupt transfers. 
                                        //  Timeout if no report available.
                                        //  To enable reading a report without blocking the calling thread, uses Filestream's ReadAsync method.                                               

                                        // Create a delegate to execute on a timeout.

                                        Action onReadTimeoutAction = OnReadTimeout;

                                        // The CancellationTokenSource specifies the timeout value and the action to take on a timeout.

                                        var cts = new CancellationTokenSource();

                                        // Cancel the read if it hasn't completed after a timeout.

                                        cts.CancelAfter(readTimeout);

                                        // Specify the function to call on a timeout.

                                        cts.Token.Register(onReadTimeoutAction);

                                        // Stops waiting when data is available or on timeout:

                                        Int32 bytesRead = await _myHid.GetInputReportViaInterruptTransfer(_deviceData, inputReportBuffer, cts);

                                        // Arrive here only if the operation completed.

                                        // Dispose to stop the timeout timer. 

                                        cts.Dispose();

                                        _transferInProgress = false;
                                        cmdGetInputReportInterrupt.Enabled = true;

                                        if (bytesRead > 0)
                                        {
                                            success = true;
                                            Debug.Print("bytes read (includes report ID) = " + Convert.ToString(bytesRead));
                                        }
                                    }
                                }
                            }
                            else
                            {
                                MyMarshalDataToForm(FormActions.AddItemToListBox, "No attempt to read an Input report was made.");
                                MyMarshalDataToForm(FormActions.AddItemToListBox, "The HID doesn't have an Input report.");
                            }
                        }
                        else
                        {
                            MyMarshalDataToForm(FormActions.AddItemToListBox, "Invalid handle.");
                            MyMarshalDataToForm(FormActions.AddItemToListBox,
                                                "No attempt to write an Output report or read an Input report was made.");
                        }

                        if (success)
                        {
                            DisplayReportData(inputReportBuffer, ReportTypes.Input, ReportReadOrWritten.Read);
                        }
                        else
                        {
                            CloseCommunications();
                            MyMarshalDataToForm(FormActions.AddItemToListBox, "The attempt to read an Input report has failed.");
                            ScrollToBottomOfListBox();
                        }
                    }

                }

                catch (Exception ex)
                {
                    DisplayException(Name, ex);
                    throw;
                }

                State = inputReportBuffer[1];
                Temperature = inputReportBuffer[2];
                Setpoint = inputReportBuffer[3];
                Oven = inputReportBuffer[4];
                Fan = inputReportBuffer[5];
                Elapsed1 = inputReportBuffer[6];
                Elapsed2 = inputReportBuffer[7];
                Start = inputReportBuffer[8];
                Kp = inputReportBuffer[9];
                Ki = inputReportBuffer[10];
                Kd = inputReportBuffer[11];
                CycleTime = inputReportBuffer[12];
                PTerm = inputReportBuffer[13];
                ITerm = inputReportBuffer[14];
                DTerm = inputReportBuffer[15];
                Output = inputReportBuffer[16];

                //Current stage
                switch (State)
                {
                    case 0: StageText.Text = "WAITING"; break;
                    case 1: StageText.Text = "PREHEAT"; break;
                    case 2: StageText.Text = "SOAK"; break;
                    case 3: StageText.Text = "HEATING"; break;
                    case 4: StageText.Text = "REFLOW"; break;
                    case 5: StageText.Text = "COOLING"; break;
                    case 6: StageText.Text = "BAKE"; break;
                }

                //Current temperature
                TemperatureText.Text = Convert.ToString(Temperature) + "\u00b0" + "C";

                //Setpoint temperature
                SetpointText.Text = Convert.ToString(Setpoint) + "\u00b0" + "C";

                //Heater state (On/Off)
                switch (Oven)
                {
                    case 0: OvenText.Text = "OFF"; break;
                    case 1: OvenText.Text = "ON"; break;
                }

                //Auxiliary state (On/Off)
                switch (Fan)
                {
                    case 0: FanText.Text = "OFF"; break;
                    case 1: FanText.Text = "ON"; break;
                }

                ////Current stage remaining time
                //if (elapsed == "0000") elapsed = "N/A";
                //StageTimeLabel.Text = elapsed;

            }

        }


        /// <summary>
        /// Request to get an Input report retrieved using interrupt transfers.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void cmdGetInputReportInterrupt_Click(object sender, EventArgs e)
		{
			try
			{
				if (_transferInProgress)
				{
					DisplayTransferInProgressMessage();
				}
				else
				{
					//  Don't allow another transfer request until this one completes.
					//  Move the focus away from the button to prevent the focus from 
					//  switching to the next control in the tab order on disabling the button.

					fraInterruptTransfers.Focus();
					cmdGetInputReportInterrupt.Enabled = false;
					_transferType = TransferTypes.Interrupt;
					RequestToGetInputReport();
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Set the number of Input reports the HID driver will store.
		///  </summary>

		private void cmdInputReportBufferSize_Click(Object sender, EventArgs e)
		{
			try
			{
				if (_transferInProgress)
				{
					DisplayTransferInProgressMessage();
				}
				else
				{
					SetInputReportBufferSize();
				}
			}
			catch
				(Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		/// <summary>
		/// Alternate sending and getting a report.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>

		private void cmdPeriodicTransfers_Click(object sender, EventArgs e)
		{
			try
			{
				if (cmdPeriodicTransfers.Text == "Start")
				{
					if (_transferInProgress)
					{
						DisplayTransferInProgressMessage();
					}
					else
					{
						_sendOrGet = SendOrGet.Send;
						PeriodicTransfersStart();
					}
				}
				else
				{
					PeriodicTransfersStop();
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		
		/// <summary>
		/// Request to send an Output report using an interrupt transfer.		
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>

		private void cmdSendOutputReportInterrupt_Click(object sender, EventArgs e)
		{
			try
			{
				if (_transferInProgress)
				{
					DisplayTransferInProgressMessage();
				}
				else
				{
					//  Don't allow another transfer request until this one completes.
					//  Move the focus away from the button to prevent the focus from 
					//  switching to the next control in the tab order on disabling the button.

					fraInterruptTransfers.Focus();
					cmdSendOutputReportInterrupt.Enabled = false;
					_transferType = TransferTypes.Interrupt;
					RequestToSendOutputReport();
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Called on arrival of any device.
		///  Calls a routine that searches to see if the desired device is present.
		///  </summary>

		private void DeviceAdded(object sender, EventArrivedEventArgs e)
		{
			try
			{
				Debug.WriteLine("A USB device has been inserted");

				_deviceDetected = FindDeviceUsingWmi();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Called if the user changes the Vendor ID or Product ID in the text box.
		///  </summary>

		private void DeviceHasChanged()
		{
			try
			{
				//  If a device was previously detected, stop receiving notifications about it.

				if (_deviceHandleObtained)
				{
					DeviceNotificationsStop();

					CloseCommunications();
				}
				// Look for a device that matches the Vendor ID and Product ID in the text boxes.

				FindTheHid();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Add handlers to detect device arrival and removal.
		///  </summary>

		private void DeviceNotificationsStart()
		{
			AddDeviceArrivedHandler();
			AddDeviceRemovedHandler();
		}

		///  <summary>
		///  Stop receiving notifications about device arrival and removal
		///  </summary>

		private void DeviceNotificationsStop()
		{
			try
			{
				if (_deviceArrivedWatcher != null)
					_deviceArrivedWatcher.Stop();
				if (_deviceRemovedWatcher != null)
					_deviceRemovedWatcher.Stop();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Called on removal of any device.
		///  Calls a routine that searches to see if the desired device is still present.
		///  </summary>
		/// 
		private void DeviceRemoved(object sender, EventArgs e)
		{
			try
			{
				Debug.WriteLine("A USB device has been removed");

				_deviceDetected = FindDeviceUsingWmi();

				if (!_deviceDetected)
				{
					_deviceHandleObtained = false;
					CloseCommunications();
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Displays received or written report data.
		///  </summary>
		///  
		///  <param name="buffer"> contains the report data. </param>			
		///  <param name="currentReportType" > "Input", "Output", or "Feature"</param>
		///  <param name="currentReadOrWritten" > "read" for Input and IN Feature reports, "written" for Output and OUT Feature reports.</param>

		private void DisplayReportData(Byte[] buffer, ReportTypes currentReportType, ReportReadOrWritten currentReadOrWritten)
		{
			try
			{
				Int32 count;

				LstResults.Items.Add(currentReportType.ToString() + " report has been " + currentReadOrWritten.ToString().ToLower() + ".");

				//  Display the report data received in the form's list box.

				LstResults.Items.Add(" Report ID: " + String.Format("{0:X2} ", buffer[0]));
				LstResults.Items.Add(" Report Data:");

				TxtBytesReceived.Text = "";

				for (count = 1; count <= buffer.Length - 1; count++)
				{
					//  Display bytes as 2-character Hex strings.

					String byteValue = String.Format("{0:X2} ", buffer[count]);

					LstResults.Items.Add(" " + byteValue);

					//  Display the received bytes in the text box.

					TxtBytesReceived.SelectionStart = TxtBytesReceived.Text.Length;
					//TxtBytesReceived.SelectedText = byteValue + Environment.NewLine;
                    TxtBytesReceived.SelectedText = byteValue;
                }
				ScrollToBottomOfListBox();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Display a message if the user clicks a button when a transfer is in progress.
		///  </summary>
		/// 
		private void DisplayTransferInProgressMessage()
		{
			AccessForm(FormActions.AddItemToListBox, "Command not executed because a transfer is in progress.");
			ScrollToBottomOfListBox();
		}

		///  <summary>
		///  Do periodic transfers.
		///  </summary>
		/// <param name="source"></param>
		/// <param name="e"></param>
		///  <remarks>
		///  The timer is enabled only if continuous (periodic) transfers have been requested.
		///  </remarks>		  

		private void DoPeriodicTransfers(object source, ElapsedEventArgs e)
		{
			try
			{
				PeriodicTransfers();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		/// <summary>
		/// Enable the command buttons on the form.
		/// Needed after attempting a transfer and device not found.
		/// </summary>
		/// 
		private void EnableFormControls()
		{
			cmdGetInputReportInterrupt.Enabled = true;
			cmdPeriodicTransfers.Enabled = true;
			cmdSendOutputReportInterrupt.Enabled = true;
		}

		///  <summary>
		///  Use the System.Management class to find a device by Vendor ID and Product ID using WMI. If found, display device properties.
		///  </summary>
		/// <remarks> 
		/// During debugging, if you stop the firmware but leave the device attached, the device may still be detected as present
		/// but will be unable to communicate. The device will show up in Windows Device Manager as well. 
		/// This situation is unlikely to occur with a final product.
		/// </remarks>

		private Boolean FindDeviceUsingWmi()
		{
			try
			{
				// Prepend "@" to string below to treat backslash as a normal character (not escape character):

				String deviceIdString = @"USB\VID_" + _myVendorId.ToString("X4") + "&PID_" + _myProductId.ToString("X4");

				_deviceDetected = false;
				var searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnPEntity");

				foreach (ManagementObject queryObj in searcher.Get())
				{
					if (queryObj["PNPDeviceID"].ToString().Contains(deviceIdString))
					{
						_deviceDetected = true;
						MyMarshalDataToForm(FormActions.AddItemToListBox, "--------");
						MyMarshalDataToForm(FormActions.AddItemToListBox, "My device found (WMI):");

						// Display device properties.

						foreach (WmiDeviceProperties wmiDeviceProperty in Enum.GetValues(typeof(WmiDeviceProperties)))
						{
							MyMarshalDataToForm(FormActions.AddItemToListBox, (wmiDeviceProperty.ToString() + ": " + queryObj[wmiDeviceProperty.ToString()]));
							Debug.WriteLine(wmiDeviceProperty.ToString() + ": {0}", queryObj[wmiDeviceProperty.ToString()]);
						}
						MyMarshalDataToForm(FormActions.AddItemToListBox, "--------");
						MyMarshalDataToForm(FormActions.ScrollToBottomOfListBox, "");
                        toolStripStatusLabel2.Text = "Controller connected";
                        MyMarshalDataToForm(FormActions.ChangeBackgroundGreen, "");
                    }
				}
				if (!_deviceDetected)
				{
					MyMarshalDataToForm(FormActions.AddItemToListBox, "My device not found (WMI)");
					MyMarshalDataToForm(FormActions.ScrollToBottomOfListBox, "");
                    toolStripStatusLabel2.Text = "Controller not found";
                    MyMarshalDataToForm(FormActions.ChangeBackgroundRed, "");
                }
				return _deviceDetected;
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Call HID functions that use Win32 API functions to locate a HID-class device
		///  by its Vendor ID and Product ID. Open a handle to the device.
		///  </summary>
		///          
		///  <returns>
		///   True if the device is detected, False if not detected.
		///  </returns>

		private Boolean FindTheHid()
		{
			var devicePathName = new String[128];
			String myDevicePathName = "";

			try
			{
				_deviceHandleObtained = false;
				CloseCommunications();

				//  Get the device's Vendor ID and Product ID from the form's text boxes.

				GetVendorAndProductIDsFromTextBoxes(ref _myVendorId, ref _myProductId);

				// Get the HID-class GUID.

				Guid hidGuid = _myHid.GetHidGuid();

				String functionName = "GetHidGuid";
				Debug.WriteLine(_myDebugging.ResultOfApiCall(functionName));
				Debug.WriteLine("  GUID for system HIDs: " + hidGuid.ToString());

				//  Fill an array with the device path names of all attached HIDs.

				Boolean availableHids = _myDeviceManagement.FindDeviceFromGuid(hidGuid, ref devicePathName);

				//  If there is at least one HID, attempt to read the Vendor ID and Product ID
				//  of each device until there is a match or all devices have been examined.

				if (availableHids)
				{
					Int32 memberIndex = 0;

					do
					{
						// Open the handle without read/write access to enable getting information about any HID, even system keyboards and mice.

						_hidHandle = _myHid.OpenHandle(devicePathName[memberIndex], false);

						functionName = "CreateFile";
						Debug.WriteLine(_myDebugging.ResultOfApiCall(functionName));
						Debug.WriteLine("  Returned handle: " + _hidHandle);

						if (!_hidHandle.IsInvalid)
						{
							// The returned handle is valid, 
							// so find out if this is the device we're looking for.

							_myHid.DeviceAttributes.Size = Marshal.SizeOf(_myHid.DeviceAttributes);

							Boolean success = _myHid.GetAttributes(_hidHandle, ref _myHid.DeviceAttributes);

							if (success)
							{
								Debug.WriteLine("  HIDD_ATTRIBUTES structure filled without error.");
								Debug.WriteLine("  Structure size: " + _myHid.DeviceAttributes.Size);
								Debug.WriteLine("  Vendor ID: " + Convert.ToString(_myHid.DeviceAttributes.VendorID, 16));
								Debug.WriteLine("  Product ID: " + Convert.ToString(_myHid.DeviceAttributes.ProductID, 16));
								Debug.WriteLine("  Version Number: " + Convert.ToString(_myHid.DeviceAttributes.VersionNumber, 16));

								if ((_myHid.DeviceAttributes.VendorID == _myVendorId) && (_myHid.DeviceAttributes.ProductID == _myProductId))
								{
									Debug.WriteLine("  Handle obtained to my device");

									//  Display the information in form's list box.

									MyMarshalDataToForm(FormActions.AddItemToListBox, "Handle obtained to my device:");
									MyMarshalDataToForm(FormActions.AddItemToListBox, "  Vendor ID= " + Convert.ToString(_myHid.DeviceAttributes.VendorID, 16));
									MyMarshalDataToForm(FormActions.AddItemToListBox, "  Product ID = " + Convert.ToString(_myHid.DeviceAttributes.ProductID, 16));
									MyMarshalDataToForm(FormActions.ScrollToBottomOfListBox, "");

									_deviceHandleObtained = true;

									myDevicePathName = devicePathName[memberIndex];
								}
								else
								{
									//  It's not a match, so close the handle.

									_deviceHandleObtained = false;
									_hidHandle.Close();
								}
							}
							else
							{
								//  There was a problem retrieving the information.

								Debug.WriteLine("  Error in filling HIDD_ATTRIBUTES structure.");
								_deviceHandleObtained = false;
								_hidHandle.Close();
							}
						}

						//  Keep looking until we find the device or there are no devices left to examine.

						memberIndex = memberIndex + 1;
					}
					while (!((_deviceHandleObtained || (memberIndex == devicePathName.Length))));
				}

				if (_deviceHandleObtained)
				{
					//  The device was detected.
					//  Learn the capabilities of the device.

					_myHid.Capabilities = _myHid.GetDeviceCapabilities(_hidHandle);

					//  Find out if the device is a system mouse or keyboard.

					_hidUsage = _myHid.GetHidUsage(_myHid.Capabilities);

					//  Get the Input report buffer size.

					GetInputReportBufferSize();
					MyMarshalDataToForm(FormActions.EnableInputReportBufferSize, "");

					//Close the handle and reopen it with read/write access.

					_hidHandle.Close();

					_hidHandle = _myHid.OpenHandle(myDevicePathName, true);

					if (_hidHandle.IsInvalid)
					{
						MyMarshalDataToForm(FormActions.AddItemToListBox, "The device is a system " + _hidUsage + ".");
						MyMarshalDataToForm(FormActions.AddItemToListBox, "Windows 2000 and later obtain exclusive access to Input and Output reports for this devices.");
						MyMarshalDataToForm(FormActions.AddItemToListBox, "Windows 8 also obtains exclusive access to Feature reports.");
						MyMarshalDataToForm(FormActions.ScrollToBottomOfListBox, "");
					}
					else
					{
						if (_myHid.Capabilities.InputReportByteLength > 0)
						{
							//  Set the size of the Input report buffer. 

							var inputReportBuffer = new Byte[_myHid.Capabilities.InputReportByteLength];

							_deviceData = new FileStream(_hidHandle, FileAccess.Read | FileAccess.Write, inputReportBuffer.Length, false);
						}

						if (_myHid.Capabilities.OutputReportByteLength > 0)
						{
							Byte[] outputReportBuffer = null;
						}
						//  Flush any waiting reports in the input buffer. (optional)

						_myHid.FlushQueue(_hidHandle);
					}
				}
				else
				{
					MyMarshalDataToForm(FormActions.AddItemToListBox, "Device not found.");
					MyMarshalDataToForm(FormActions.DisableInputReportBufferSize, "");
					EnableFormControls();
					MyMarshalDataToForm(FormActions.ScrollToBottomOfListBox, "");
				}
				return _deviceHandleObtained;
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Perform shutdown operations.
		///  </summary>

		private void frmMain_Closed(Object eventSender, EventArgs eventArgs)
		{
			try
			{
				Shutdown();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Perform startup operations.
		///  </summary>

		private void frmMain_Load(Object eventSender, EventArgs eventArgs)
		{
			try
			{
				FrmMy = this;
				Startup();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		private void GetBytesToSend()
		{
			try
			{
				//  Get the bytes to send in a report from the combo boxes.
				//  Increment the values if the autoincrement check box is selected.

				if (ChkAutoincrement.Checked)
				{
					if (CboByte0.SelectedIndex < 255)
					{
						CboByte0.SelectedIndex = CboByte0.SelectedIndex + 1;
					}
					else
					{
						CboByte0.SelectedIndex = 0;
					}
					if (CboByte1.SelectedIndex < 255)
					{
						CboByte1.SelectedIndex = CboByte1.SelectedIndex + 1;
					}
					else
					{
						CboByte1.SelectedIndex = 0;
					}
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Find and display the number of Input buffers
		///  (the number of Input reports the HID driver will store). 
		///  </summary>

		private void GetInputReportBufferSize()
		{
			Int32 numberOfInputBuffers = 0;
			Boolean success;

			try
			{
				//  Get the number of input buffers.

				_myHid.GetNumberOfInputBuffers(_hidHandle, ref numberOfInputBuffers);

				//  Display the result in the text box.

				MyMarshalDataToForm(FormActions.SetInputReportBufferSize, Convert.ToString(numberOfInputBuffers));
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Retrieve a Vendor ID and Product ID in hexadecimal 
		///  from the form's text boxes and convert the text to Int32s.
		///  </summary>
		///  
		///  <param name="myVendorId"> the Vendor ID</param>
		///  <param name="myProductId"> the Product ID</param>

		private void GetVendorAndProductIDsFromTextBoxes(ref Int32 myVendorId, ref Int32 myProductId)
		{
			try
			{
				myVendorId = Int32.Parse(txtVendorID.Text, NumberStyles.AllowHexSpecifier);
				myProductId = Int32.Parse(txtProductID.Text, NumberStyles.AllowHexSpecifier);
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Initialize the elements on the form.
		///  </summary>

		private void InitializeDisplay()
		{
			try
			{
				//  Create a dropdown list box for each byte to send in a report.
				//  Display the values as 2-character hex strings.

				Int16 count;
				for (count = 0; count <= 255; count++)
				{
					String byteValue = String.Format("{0:X2} ", count);
					FrmMy.CboByte0.Items.Insert(count, byteValue);
					FrmMy.CboByte1.Items.Insert(count, byteValue);
				}

				//  Select a default value for each box

				FrmMy.CboByte0.SelectedIndex = 0;
				FrmMy.CboByte1.SelectedIndex = 128;
				FrmMy.radInputOutputInterrupt.Checked = true;

				//  Check the autoincrement box to increment the values each time a report is sent.

				ChkAutoincrement.CheckState = CheckState.Checked;

				//  Don't allow the user to select an input report buffer size until there is
				//  a handle to a HID.

				cmdInputReportBufferSize.Focus();
				cmdInputReportBufferSize.Enabled = false;

				LstResults.Items.Add("For a more detailed event log, view debug statements in Visual Studio's Output window:");
				LstResults.Items.Add("Click Build > Configuration Manager > Active Solution Configuration > Debug > Close.");
				LstResults.Items.Add("Then click View > Output.");
				LstResults.Items.Add("");
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Enables accessing a form's controls from another thread 
		///  </summary>
		///  
		///  <param name="action"> a FormActions member that names the action to perform on the form </param>
		///  <param name="textToDisplay"> text that the form displays or the code uses for 
		///  another purpose. Actions that don't use text ignore this parameter.  </param>

		private void MyMarshalDataToForm(FormActions action, String textToDisplay)
		{
			try
			{
				object[] args = { action, textToDisplay };

				//  The AccessForm routine contains the code that accesses the form.

				MarshalDataToForm marshalDataToFormDelegate = AccessForm;

				//  Execute AccessForm, passing the parameters in args.

				Invoke(marshalDataToFormDelegate, args);
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		/// <summary>
		/// Timeout if read via interrupt transfer doesn't return.
		/// </summary>

		private void OnReadTimeout()
		{
			try
			{
				MyMarshalDataToForm(FormActions.AddItemToListBox, "The attempt to read a report timed out.");
				MyMarshalDataToForm(FormActions.ScrollToBottomOfListBox, "");
				CloseCommunications();
				MyMarshalDataToForm(FormActions.EnableGetInputReportInterruptTransfer, "");
				_transferInProgress = false;
				_sendOrGet = SendOrGet.Send;
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		/// <summary>
		/// Timeout if write via interrupt transfer doesn't return.
		/// </summary>

		private void OnWriteTimeout()
		{
			try
			{
				MyMarshalDataToForm(FormActions.AddItemToListBox, "The attempt to write a report timed out.");
				MyMarshalDataToForm(FormActions.ScrollToBottomOfListBox, "");
				CloseCommunications();
				MyMarshalDataToForm(FormActions.EnableSendOutputReportInterrupt, "");
				_transferInProgress = false;
				_sendOrGet = SendOrGet.Get;
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		/// <summary>
		/// Alternat sending and getting a report.
		/// </summary>

		private void PeriodicTransfers()
		{
			try
			{
				if (!_transferInProgress)
				{
					if (_reportType == ReportTypes.Feature)
					{
						//SendOrGetFeatureReport();
					}
					else
					{
						// Output and Input reports

						SendOutputReportOrGetInputReport();
					}
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		/// <summary>
		/// Start doing periodic transfers.
		/// </summary>

		private void PeriodicTransfersStart()
		{
			// Don't allow changing the transfer type while transfers are in progress.

			if (radFeature.Checked)
			{
				radInputOutputControl.Enabled = false;
				radInputOutputInterrupt.Enabled = false;
			}
			else if (radInputOutputControl.Checked)
			{
				radFeature.Enabled = false;
				radInputOutputInterrupt.Enabled = false;
			}
			else if (radInputOutputInterrupt.Checked)
			{
				radFeature.Enabled = false;
				radInputOutputControl.Enabled = false;
			}

			//  Change the command button's text.

			cmdPeriodicTransfers.Text = "Stop";

			//  Enable the timer event to trigger a set of transfers.

			_periodicTransfers.Start();

			cmdPeriodicTransfers.Enabled = true;

			if (radInputOutputInterrupt.Checked)
			{
				_transferType = TransferTypes.Interrupt;
				_reportType = ReportTypes.Output;
			}
			else if (radInputOutputControl.Checked)
			{
				_transferType = TransferTypes.Control;
				_reportType = ReportTypes.Output;
			}
			else if (radFeature.Checked)
			{
				_transferType = TransferTypes.Control;
				_reportType = ReportTypes.Feature;
			}
			_periodicTransfersRequested = true;
			PeriodicTransfers();
		}

		/// <summary>
		/// Stop doing periodic transfers.
		/// </summary>

		private void PeriodicTransfersStop()
		{
			//  Stop doing continuous transfers.

			_periodicTransfersRequested = false;

			// Disable the timer that triggers the transfers.	

			_periodicTransfers.Stop();
			cmdPeriodicTransfers.Enabled = true;

			//  Change the command button's text.

			cmdPeriodicTransfers.Text = "Start";

			// Re-allow changing the transfer type.

			radFeature.Enabled = true;
			radInputOutputControl.Enabled = true;
			radInputOutputInterrupt.Enabled = true;
		}

		private void radInputOutputControl_CheckedChanged(object sender, EventArgs e)
		{
		}

		private void radInputOutputInterrupt_CheckedChanged(object sender, EventArgs e)
		{
		}

		private void radFeature_CheckedChanged(object sender, EventArgs e)
		{
		}

	

		///  <summary>
		///  Request an Input report.
		///  Assumes report ID = 0.
		///  </summary>

		private async void RequestToGetInputReport()
		{
			const Int32 readTimeout = 5000;

			String byteValue = null;
			Byte[] inputReportBuffer = null;

			try
			{
				Boolean success = false;

				//  If the device hasn't been detected, was removed, or timed out on a previous attempt
				//  to access it, look for the device.

				if (!_deviceHandleObtained)
				{
					_deviceHandleObtained = FindTheHid();
				}

				if (_deviceHandleObtained)
				{
					//  Don't attempt to exchange reports if valid handles aren't available
					//  (as for a mouse or keyboard under Windows 2000 and later.)

					if (!_hidHandle.IsInvalid)
					{
						//  Read an Input report.

						//  Don't attempt to send an Input report if the HID has no Input report.
						//  (The HID spec requires all HIDs to have an interrupt IN endpoint,
						//  which suggests that all HIDs must support Input reports.)

						if (_myHid.Capabilities.InputReportByteLength > 0)
						{
							//  Set the size of the Input report buffer. 

							inputReportBuffer = new Byte[_myHid.Capabilities.InputReportByteLength];

							if (_transferType.Equals(TransferTypes.Control))
							{
								{
									_transferInProgress = true;

									//  Read a report using a control transfer.

									success = _myHid.GetInputReportViaControlTransfer(_hidHandle, ref inputReportBuffer);
									//cmdGetInputReportControl.Enabled = true;
									_transferInProgress = false;
								}
							}
							else
							{
								{
									_transferInProgress = true;

									//  Read a report using interrupt transfers. 
									//  Timeout if no report available.
									//  To enable reading a report without blocking the calling thread, uses Filestream's ReadAsync method.                                               

									// Create a delegate to execute on a timeout.

									Action onReadTimeoutAction = OnReadTimeout;

									// The CancellationTokenSource specifies the timeout value and the action to take on a timeout.

									var cts = new CancellationTokenSource();

									// Cancel the read if it hasn't completed after a timeout.

									cts.CancelAfter(readTimeout);

									// Specify the function to call on a timeout.

									cts.Token.Register(onReadTimeoutAction);

									// Stops waiting when data is available or on timeout:

									Int32 bytesRead = await _myHid.GetInputReportViaInterruptTransfer(_deviceData, inputReportBuffer, cts);

									// Arrive here only if the operation completed.

									// Dispose to stop the timeout timer. 

									cts.Dispose();

									_transferInProgress = false;
									cmdGetInputReportInterrupt.Enabled = true;

									if (bytesRead > 0)
									{
										success = true;
										Debug.Print("bytes read (includes report ID) = " + Convert.ToString(bytesRead));
									}
								}
							}
						}
						else
						{
							MyMarshalDataToForm(FormActions.AddItemToListBox, "No attempt to read an Input report was made.");
							MyMarshalDataToForm(FormActions.AddItemToListBox, "The HID doesn't have an Input report.");
						}
					}
					else
					{
						MyMarshalDataToForm(FormActions.AddItemToListBox, "Invalid handle.");
						MyMarshalDataToForm(FormActions.AddItemToListBox,
											"No attempt to write an Output report or read an Input report was made.");
					}

					if (success)
					{
						DisplayReportData(inputReportBuffer, ReportTypes.Input, ReportReadOrWritten.Read);
					}
					else
					{
						CloseCommunications();
						MyMarshalDataToForm(FormActions.AddItemToListBox, "The attempt to read an Input report has failed.");
						ScrollToBottomOfListBox();
					}
				}
			}

			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

	


		///  <summary>
		///  Sends an Output report.
		///  Assumes report ID = 0.
		///  </summary>

		private async void RequestToSendOutputReport()
		{
			const Int32 writeTimeout = 5000;
			String byteValue = null;

			try
			{
				//  If the device hasn't been detected, was removed, or timed out on a previous attempt
				//  to access it, look for the device.

				if (!_deviceHandleObtained)
				{
					_deviceHandleObtained = FindTheHid();
				}

				if (_deviceHandleObtained)
				{
					GetBytesToSend();
				}
				//  Don't attempt to exchange reports if valid handles aren't available
				//  (as for a mouse or keyboard.)

				if (!_hidHandle.IsInvalid)
				{
					//  Don't attempt to send an Output report if the HID has no Output report.

					if (_myHid.Capabilities.OutputReportByteLength > 0)
					{
						//  Set the size of the Output report buffer.   

						var outputReportBuffer = new Byte[_myHid.Capabilities.OutputReportByteLength];

						//  Store the report ID in the first byte of the buffer:

						outputReportBuffer[0] = 0;

						//  Store the report data following the report ID.
						//  Use the data in the combo boxes on the form.

						outputReportBuffer[1] = Convert.ToByte(CboByte0.SelectedIndex);

						if (outputReportBuffer.GetUpperBound(0) > 1)
						{
							outputReportBuffer[2] = Convert.ToByte(CboByte1.SelectedIndex);
						}

						//  Write a report.

						Boolean success;

						if (_transferType.Equals(TransferTypes.Control))
						{
							{
								_transferInProgress = true;

								//  Use a control transfer to send the report,
								//  even if the HID has an interrupt OUT endpoint.

								success = _myHid.SendOutputReportViaControlTransfer(_hidHandle, outputReportBuffer);

								_transferInProgress = false;
								//cmdSendOutputReportControl.Enabled = true;
							}
						}
						else
						{
							Debug.Print("interrupt");
							_transferInProgress = true;

							// The CancellationTokenSource specifies the timeout value and the action to take on a timeout.

							var cts = new CancellationTokenSource();

							// Create a delegate to execute on a timeout.

							Action onWriteTimeoutAction = OnWriteTimeout;

							// Cancel the read if it hasn't completed after a timeout.

							cts.CancelAfter(writeTimeout);

							// Specify the function to call on a timeout.

							cts.Token.Register(onWriteTimeoutAction);

							// Send an Output report and wait for completion or timeout.

							success = await _myHid.SendOutputReportViaInterruptTransfer(_deviceData, _hidHandle, outputReportBuffer, cts);

							// Get here only if the operation completes without a timeout.

							_transferInProgress = false;
							cmdSendOutputReportInterrupt.Enabled = true;

							// Dispose to stop the timeout timer.

							cts.Dispose();
						}
						if (success)
						{
							DisplayReportData(outputReportBuffer, ReportTypes.Output, ReportReadOrWritten.Written);
						}
						else
						{
							CloseCommunications();
							AccessForm(FormActions.AddItemToListBox, "The attempt to write an Output report failed.");
							ScrollToBottomOfListBox();
						}
					}
				}
				else
				{
					AccessForm(FormActions.AddItemToListBox, "The HID doesn't have an Output report.");
				}
			}

			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Scroll to the bottom of the list box and trim as needed.
		///  </summary>

		private void ScrollToBottomOfListBox()
		{
			try
			{
				LstResults.SelectedIndex = LstResults.Items.Count - 1;

				//  If the list box is getting too large, trim its contents by removing the earliest data.

				if (LstResults.Items.Count > 1000)
				{
					Int32 count;
					for (count = 1; count <= 500; count++)
					{
						LstResults.Items.RemoveAt(4);
					}
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}


		/// <summary>
		/// Request to send an Output report or get an Input report.
		/// </summary>

		private void SendOutputReportOrGetInputReport()
		{
			try
			{
				//  If the device hasn't been detected, was removed, or timed out on a previous attempt
				//  to access it, look for the device.

				if (!_deviceHandleObtained)
				{
					_deviceHandleObtained = FindTheHid();
				}

				if (_deviceHandleObtained)
				{
					if (_sendOrGet == SendOrGet.Send)
					{
						RequestToSendOutputReport();
						_sendOrGet = SendOrGet.Get;
					}
					else
					{
						RequestToGetInputReport();
						_sendOrGet = SendOrGet.Send;
					}
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Set the number of Input buffers (the number of Input reports 
		///  the host will store) from the value in the text box.
		///  </summary>

		private void SetInputReportBufferSize()
		{
			try
			{
				if (!_transferInProgress)
				{
					//  Get the number of buffers from the text box.

					Int32 numberOfInputBuffers = Convert.ToInt32(txtInputReportBufferSize.Text);

					//  Set the number of buffers.

					_myHid.SetNumberOfInputBuffers(_hidHandle, numberOfInputBuffers);

					//  Verify and display the result.

					GetInputReportBufferSize();
				}
				else
				{
					DisplayTransferInProgressMessage();
				}
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Perform actions that must execute when the program ends.
		///  </summary>

		private void Shutdown()
		{
			try
			{
				CloseCommunications();
				DeviceNotificationsStop();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Perform actions that must execute when the program starts.
		///  </summary>

		private void Startup()
		{
			const Int32 periodicTransferInterval = 1000;
			try
			{
				_myHid = new Hid();
				InitializeDisplay();

				_periodicTransfers = new System.Timers.Timer(periodicTransferInterval);
				_periodicTransfers.Elapsed += DoPeriodicTransfers;
				_periodicTransfers.Stop();
				_periodicTransfers.SynchronizingObject = this;

				//  Default USB Vendor ID and Product ID:

				txtVendorID.Text = "04DB";
				txtProductID.Text = "1234";

				GetVendorAndProductIDsFromTextBoxes(ref _myVendorId, ref _myProductId);

				DeviceNotificationsStart();
				FindDeviceUsingWmi();
				FindTheHid();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  The Product ID has changed in the text box. Call a routine to handle it.
		///  </summary>

		private void txtProductID_TextChanged(Object sender, EventArgs e)
		{
			try
			{
				DeviceHasChanged();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  The Vendor ID has changed in the text box. Call a routine to handle it.
		///  </summary>

		private void txtVendorID_TextChanged(Object sender, EventArgs e)
		{
			try
			{
				DeviceHasChanged();
			}
			catch (Exception ex)
			{
				DisplayException(Name, ex);
				throw;
			}
		}

		///  <summary>
		///  Provides a central mechanism for exception handling.
		///  Displays a message box that describes the exception.
		///  </summary>
		///  
		///  <param name="moduleName"> the module where the exception occurred. </param>
		///  <param name="e"> the exception </param>

		internal static void DisplayException(String moduleName, Exception e)
		{
			//  Create an error message.

			String message = "Exception: " + e.Message + Environment.NewLine + "Module: " + moduleName + Environment.NewLine + "Method: " + e.TargetSite.Name;

			const String caption = "Unexpected Exception";

			MessageBox.Show(message, caption, MessageBoxButtons.OK);
			Debug.Write(message);

			// Get the last error and display it. 

			Int32 error = Marshal.GetLastWin32Error();

			Debug.WriteLine("The last Win32 Error was: " + error);
		}

		[STAThread]
		internal static void Main() { Application.Run(new FrmMain()); }
		private static FrmMain _transDefaultFormFrmMain;
		internal static FrmMain TransDefaultFormFrmMain
		{
			get
			{
				if (_transDefaultFormFrmMain == null)
				{
					_transDefaultFormFrmMain = new FrmMain();
				}
				return _transDefaultFormFrmMain;
			}
		}
	}
}
