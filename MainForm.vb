Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports NationalInstruments.NI4882


Public Class MainForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        '
        ' Required for Windows Form Designer support
        '
        InitializeComponent()
        secondaryAddressComboBox.Items.Add("None")
        Dim i As Integer
        For i = 96 To 126
            secondaryAddressComboBox.Items.Add(i)
        Next
        secondaryAddressComboBox.SelectedIndex = 0


    End Sub 'New

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (GpibDevice Is Nothing) Then
                GpibDevice.Dispose()
            End If
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub 'Dispose

    Private WithEvents openButton As System.Windows.Forms.Button
    Private WithEvents closeButton As System.Windows.Forms.Button
    Private stringToWriteTextBox As System.Windows.Forms.TextBox
    Private WithEvents writeButton As System.Windows.Forms.Button
    Private stringToWriteLabel As System.Windows.Forms.Label
    Private WithEvents readButton As System.Windows.Forms.Button
    Private stringReadLabel As System.Windows.Forms.Label
    Private GpibDevice As Device
    Private boardIdNumericUpDown As System.Windows.Forms.NumericUpDown
    Private boardIdLabel As System.Windows.Forms.Label
    Private primaryAddressNumericUpDown As System.Windows.Forms.NumericUpDown
    Private primaryAddressLabel As System.Windows.Forms.Label
    Private secondaryAddressLabel As System.Windows.Forms.Label
    Private stringReadTextBox As System.Windows.Forms.TextBox
    Private components As System.ComponentModel.Container = Nothing
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TChart1 As Steema.TeeChart.TChart
    Friend WithEvents StepSize_TB As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents StartLambda_TB As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Filename_TB As System.Windows.Forms.TextBox
    Friend WithEvents FastLine1 As Steema.TeeChart.Styles.FastLine
    Friend WithEvents FSC_CB As System.Windows.Forms.CheckBox
    Friend WithEvents StopLambda_TB As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TestTime_TB As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Power_TB As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents WavemeterFreq_TB As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LaserFreq_TB As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents secondaryAddressComboBox As System.Windows.Forms.ComboBox

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.openButton = New System.Windows.Forms.Button()
        Me.closeButton = New System.Windows.Forms.Button()
        Me.stringToWriteTextBox = New System.Windows.Forms.TextBox()
        Me.stringReadTextBox = New System.Windows.Forms.TextBox()
        Me.writeButton = New System.Windows.Forms.Button()
        Me.stringToWriteLabel = New System.Windows.Forms.Label()
        Me.readButton = New System.Windows.Forms.Button()
        Me.stringReadLabel = New System.Windows.Forms.Label()
        Me.boardIdNumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.primaryAddressNumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.boardIdLabel = New System.Windows.Forms.Label()
        Me.primaryAddressLabel = New System.Windows.Forms.Label()
        Me.secondaryAddressLabel = New System.Windows.Forms.Label()
        Me.secondaryAddressComboBox = New System.Windows.Forms.ComboBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Filename_TB = New System.Windows.Forms.TextBox()
        Me.StepSize_TB = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.StartLambda_TB = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TChart1 = New Steema.TeeChart.TChart()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.FastLine1 = New Steema.TeeChart.Styles.FastLine()
        Me.FSC_CB = New System.Windows.Forms.CheckBox()
        Me.StopLambda_TB = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TestTime_TB = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Power_TB = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.WavemeterFreq_TB = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LaserFreq_TB = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        CType(Me.boardIdNumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.primaryAddressNumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'openButton
        '
        Me.openButton.Location = New System.Drawing.Point(19, 142)
        Me.openButton.Name = "openButton"
        Me.openButton.Size = New System.Drawing.Size(90, 26)
        Me.openButton.TabIndex = 2
        Me.openButton.Text = "&Open"
        '
        'closeButton
        '
        Me.closeButton.Enabled = False
        Me.closeButton.Location = New System.Drawing.Point(115, 142)
        Me.closeButton.Name = "closeButton"
        Me.closeButton.Size = New System.Drawing.Size(90, 26)
        Me.closeButton.TabIndex = 3
        Me.closeButton.Text = "&Close"
        '
        'stringToWriteTextBox
        '
        Me.stringToWriteTextBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stringToWriteTextBox.Enabled = False
        Me.stringToWriteTextBox.Location = New System.Drawing.Point(270, 45)
        Me.stringToWriteTextBox.Name = "stringToWriteTextBox"
        Me.stringToWriteTextBox.Size = New System.Drawing.Size(141, 22)
        Me.stringToWriteTextBox.TabIndex = 4
        Me.stringToWriteTextBox.Text = "*idn?\n"
        '
        'stringReadTextBox
        '
        Me.stringReadTextBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stringReadTextBox.Location = New System.Drawing.Point(270, 156)
        Me.stringReadTextBox.Multiline = True
        Me.stringReadTextBox.Name = "stringReadTextBox"
        Me.stringReadTextBox.ReadOnly = True
        Me.stringReadTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.stringReadTextBox.Size = New System.Drawing.Size(276, 100)
        Me.stringReadTextBox.TabIndex = 6
        '
        'writeButton
        '
        Me.writeButton.Enabled = False
        Me.writeButton.Location = New System.Drawing.Point(270, 82)
        Me.writeButton.Name = "writeButton"
        Me.writeButton.Size = New System.Drawing.Size(90, 27)
        Me.writeButton.TabIndex = 7
        Me.writeButton.Text = "&Write"
        '
        'stringToWriteLabel
        '
        Me.stringToWriteLabel.Location = New System.Drawing.Point(270, 18)
        Me.stringToWriteLabel.Name = "stringToWriteLabel"
        Me.stringToWriteLabel.Size = New System.Drawing.Size(120, 26)
        Me.stringToWriteLabel.TabIndex = 8
        Me.stringToWriteLabel.Text = "String to Write:"
        '
        'readButton
        '
        Me.readButton.Enabled = False
        Me.readButton.Location = New System.Drawing.Point(366, 82)
        Me.readButton.Name = "readButton"
        Me.readButton.Size = New System.Drawing.Size(90, 27)
        Me.readButton.TabIndex = 9
        Me.readButton.Text = "&Read"
        '
        'stringReadLabel
        '
        Me.stringReadLabel.Location = New System.Drawing.Point(270, 119)
        Me.stringReadLabel.Name = "stringReadLabel"
        Me.stringReadLabel.Size = New System.Drawing.Size(120, 27)
        Me.stringReadLabel.TabIndex = 10
        Me.stringReadLabel.Text = "String Read:"
        '
        'boardIdNumericUpDown
        '
        Me.boardIdNumericUpDown.Location = New System.Drawing.Point(163, 49)
        Me.boardIdNumericUpDown.Name = "boardIdNumericUpDown"
        Me.boardIdNumericUpDown.Size = New System.Drawing.Size(72, 22)
        Me.boardIdNumericUpDown.TabIndex = 11
        '
        'primaryAddressNumericUpDown
        '
        Me.primaryAddressNumericUpDown.Location = New System.Drawing.Point(163, 77)
        Me.primaryAddressNumericUpDown.Name = "primaryAddressNumericUpDown"
        Me.primaryAddressNumericUpDown.Size = New System.Drawing.Size(72, 22)
        Me.primaryAddressNumericUpDown.TabIndex = 12
        Me.primaryAddressNumericUpDown.Value = New Decimal(New Integer() {11, 0, 0, 0})
        '
        'boardIdLabel
        '
        Me.boardIdLabel.Location = New System.Drawing.Point(19, 49)
        Me.boardIdLabel.Name = "boardIdLabel"
        Me.boardIdLabel.Size = New System.Drawing.Size(86, 19)
        Me.boardIdLabel.TabIndex = 14
        Me.boardIdLabel.Text = "Board ID:"
        '
        'primaryAddressLabel
        '
        Me.primaryAddressLabel.Location = New System.Drawing.Point(19, 77)
        Me.primaryAddressLabel.Name = "primaryAddressLabel"
        Me.primaryAddressLabel.Size = New System.Drawing.Size(120, 19)
        Me.primaryAddressLabel.TabIndex = 15
        Me.primaryAddressLabel.Text = "Primary Address:"
        '
        'secondaryAddressLabel
        '
        Me.secondaryAddressLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.secondaryAddressLabel.Location = New System.Drawing.Point(19, 105)
        Me.secondaryAddressLabel.Name = "secondaryAddressLabel"
        Me.secondaryAddressLabel.Size = New System.Drawing.Size(134, 18)
        Me.secondaryAddressLabel.TabIndex = 16
        Me.secondaryAddressLabel.Text = "Secondary Address:"
        '
        'secondaryAddressComboBox
        '
        Me.secondaryAddressComboBox.Location = New System.Drawing.Point(163, 105)
        Me.secondaryAddressComboBox.Name = "secondaryAddressComboBox"
        Me.secondaryAddressComboBox.Size = New System.Drawing.Size(72, 24)
        Me.secondaryAddressComboBox.TabIndex = 17
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 13)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1046, 393)
        Me.TabControl1.TabIndex = 18
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Label11)
        Me.TabPage1.Controls.Add(Me.Label9)
        Me.TabPage1.Controls.Add(Me.Button4)
        Me.TabPage1.Controls.Add(Me.Button3)
        Me.TabPage1.Controls.Add(Me.Label10)
        Me.TabPage1.Controls.Add(Me.Power_TB)
        Me.TabPage1.Controls.Add(Me.TestTime_TB)
        Me.TabPage1.Controls.Add(Me.Label7)
        Me.TabPage1.Controls.Add(Me.StopLambda_TB)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.FSC_CB)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.Filename_TB)
        Me.TabPage1.Controls.Add(Me.StepSize_TB)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.StartLambda_TB)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.TChart1)
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1038, 364)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "TLS Control"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 336)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(65, 17)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Filename"
        '
        'Filename_TB
        '
        Me.Filename_TB.Location = New System.Drawing.Point(77, 333)
        Me.Filename_TB.Name = "Filename_TB"
        Me.Filename_TB.Size = New System.Drawing.Size(305, 22)
        Me.Filename_TB.TabIndex = 11
        Me.Filename_TB.Text = "Sweep Test"
        '
        'StepSize_TB
        '
        Me.StepSize_TB.Location = New System.Drawing.Point(8, 303)
        Me.StepSize_TB.Name = "StepSize_TB"
        Me.StepSize_TB.Size = New System.Drawing.Size(101, 22)
        Me.StepSize_TB.TabIndex = 10
        Me.StepSize_TB.Text = "0.001"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 282)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(153, 17)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Step Size (>=0.001nm)"
        '
        'StartLambda_TB
        '
        Me.StartLambda_TB.Location = New System.Drawing.Point(8, 252)
        Me.StartLambda_TB.Name = "StartLambda_TB"
        Me.StartLambda_TB.Size = New System.Drawing.Size(124, 22)
        Me.StartLambda_TB.TabIndex = 8
        Me.StartLambda_TB.Text = "1526.500"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 230)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(146, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Start Wavelength(nm)"
        '
        'TChart1
        '
        '
        '
        '
        Me.TChart1.Aspect.Chart3DPercent = 5
        '
        '
        '
        '
        '
        '
        '
        '
        '
        Me.TChart1.Axes.Bottom.Title.Caption = "Frequency(GHz)"
        Me.TChart1.Axes.Bottom.Title.Lines = New String() {"Frequency(GHz)"}
        '
        '
        '
        '
        '
        '
        Me.TChart1.Axes.Left.Title.Caption = "Power(dBm)"
        Me.TChart1.Axes.Left.Title.Lines = New String() {"Power(dBm)"}
        '
        '
        '
        '
        '
        '
        Me.TChart1.Axes.Top.Title.Caption = "Frequency Vs Power"
        Me.TChart1.Axes.Top.Title.Lines = New String() {"Frequency Vs Power"}
        '
        '
        '
        Me.TChart1.Header.Lines = New String() {"DUT Frequency Vs Power"}
        Me.TChart1.Location = New System.Drawing.Point(313, 8)
        Me.TChart1.Name = "TChart1"
        Me.TChart1.Series.Add(Me.FastLine1)
        Me.TChart1.Size = New System.Drawing.Size(577, 317)
        Me.TChart1.TabIndex = 5
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(163, 181)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(101, 26)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Start Sweep"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.WavemeterFreq_TB)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.LaserFreq_TB)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.Label5)
        Me.TabPage2.Controls.Add(Me.openButton)
        Me.TabPage2.Controls.Add(Me.secondaryAddressComboBox)
        Me.TabPage2.Controls.Add(Me.closeButton)
        Me.TabPage2.Controls.Add(Me.secondaryAddressLabel)
        Me.TabPage2.Controls.Add(Me.stringToWriteTextBox)
        Me.TabPage2.Controls.Add(Me.primaryAddressLabel)
        Me.TabPage2.Controls.Add(Me.stringReadTextBox)
        Me.TabPage2.Controls.Add(Me.boardIdLabel)
        Me.TabPage2.Controls.Add(Me.writeButton)
        Me.TabPage2.Controls.Add(Me.primaryAddressNumericUpDown)
        Me.TabPage2.Controls.Add(Me.stringToWriteLabel)
        Me.TabPage2.Controls.Add(Me.boardIdNumericUpDown)
        Me.TabPage2.Controls.Add(Me.readButton)
        Me.TabPage2.Controls.Add(Me.stringReadLabel)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1038, 364)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "GPIB Configuration"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label5.Location = New System.Drawing.Point(19, 15)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(216, 18)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "GPIB Test Panel"
        '
        'FastLine1
        '
        Me.FastLine1.Color = System.Drawing.Color.FromArgb(CType(CType(243, Byte), Integer), CType(CType(156, Byte), Integer), CType(CType(53, Byte), Integer))
        Me.FastLine1.ColorEach = False
        Me.FastLine1.DrawAllPoints = False
        '
        '
        '
        Me.FastLine1.LinePen.Color = System.Drawing.Color.FromArgb(CType(CType(243, Byte), Integer), CType(CType(156, Byte), Integer), CType(CType(53, Byte), Integer))
        Me.FastLine1.ShowInLegend = False
        Me.FastLine1.Title = "fastLine1"
        Me.FastLine1.TreatNulls = Steema.TeeChart.Styles.TreatNullsStyle.Ignore
        '
        '
        '
        Me.FastLine1.XValues.DataMember = "X"
        Me.FastLine1.XValues.Order = Steema.TeeChart.Styles.ValueListOrder.Ascending
        '
        '
        '
        Me.FastLine1.YValues.DataMember = "Y"
        '
        'FSC_CB
        '
        Me.FSC_CB.AutoSize = True
        Me.FSC_CB.Enabled = False
        Me.FSC_CB.Location = New System.Drawing.Point(400, 334)
        Me.FSC_CB.Name = "FSC_CB"
        Me.FSC_CB.Size = New System.Drawing.Size(236, 21)
        Me.FSC_CB.TabIndex = 13
        Me.FSC_CB.Text = "Enable FSC Mode (0.1pm Steps)"
        Me.FSC_CB.UseVisualStyleBackColor = True
        '
        'StopLambda_TB
        '
        Me.StopLambda_TB.Location = New System.Drawing.Point(160, 253)
        Me.StopLambda_TB.Name = "StopLambda_TB"
        Me.StopLambda_TB.Size = New System.Drawing.Size(124, 22)
        Me.StopLambda_TB.TabIndex = 17
        Me.StopLambda_TB.Text = "1526.550"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(160, 231)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(145, 17)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Stop Wavelength(nm)"
        '
        'TestTime_TB
        '
        Me.TestTime_TB.Location = New System.Drawing.Point(160, 304)
        Me.TestTime_TB.Name = "TestTime_TB"
        Me.TestTime_TB.Size = New System.Drawing.Size(101, 22)
        Me.TestTime_TB.TabIndex = 19
        Me.TestTime_TB.Text = "6s per step"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(160, 282)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(107, 17)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Test Time (Min)"
        '
        'Power_TB
        '
        Me.Power_TB.Location = New System.Drawing.Point(11, 77)
        Me.Power_TB.Name = "Power_TB"
        Me.Power_TB.Size = New System.Drawing.Size(121, 22)
        Me.Power_TB.TabIndex = 21
        Me.Power_TB.Text = "1.00"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(8, 57)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(125, 17)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Laser Power(dBm)"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(11, 181)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(101, 26)
        Me.Button3.TabIndex = 23
        Me.Button3.Text = "Enable Laser"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(267, 304)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(38, 26)
        Me.Button4.TabIndex = 24
        Me.Button4.Text = "?"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 30)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(194, 17)
        Me.Label9.TabIndex = 25
        Me.Label9.Text = "Tuning Range: 1500-1620nm"
        '
        'WavemeterFreq_TB
        '
        Me.WavemeterFreq_TB.Location = New System.Drawing.Point(104, 212)
        Me.WavemeterFreq_TB.Name = "WavemeterFreq_TB"
        Me.WavemeterFreq_TB.Size = New System.Drawing.Size(69, 22)
        Me.WavemeterFreq_TB.TabIndex = 23
        Me.WavemeterFreq_TB.Text = "GPIB 12"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(101, 192)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 17)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Wavelength Meter"
        '
        'LaserFreq_TB
        '
        Me.LaserFreq_TB.Location = New System.Drawing.Point(18, 212)
        Me.LaserFreq_TB.Name = "LaserFreq_TB"
        Me.LaserFreq_TB.Size = New System.Drawing.Size(69, 22)
        Me.LaserFreq_TB.TabIndex = 21
        Me.LaserFreq_TB.Text = "GPIB 11"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 192)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 17)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Tunics-PRI"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(8, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(181, 20)
        Me.Label11.TabIndex = 26
        Me.Label11.Text = "Photonetics TUNICS"
        '
        'MainForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(916, 418)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(221, 351)
        Me.Name = "MainForm"
        Me.Text = "NI-488.2 Simple Read/Write"
        CType(Me.boardIdNumericUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.primaryAddressNumericUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub 'InitializeComponent

#End Region

    <STAThread()> _
    Public Shared Sub Main()
        Application.Run(New MainForm())
    End Sub 'Main

    Private Sub SetupControlState(ByVal isSessionOpen As Boolean)
        boardIdNumericUpDown.Enabled = Not isSessionOpen
        primaryAddressNumericUpDown.Enabled = Not isSessionOpen
        secondaryAddressComboBox.Enabled = Not isSessionOpen
        openButton.Enabled = Not isSessionOpen
        closeButton.Enabled = isSessionOpen
        stringToWriteTextBox.Enabled = isSessionOpen
        writeButton.Enabled = isSessionOpen
        readButton.Enabled = isSessionOpen
        stringReadTextBox.Enabled = isSessionOpen
    End Sub 'SetupControlState

    Private Sub openButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles openButton.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim currentSecondaryAddress As Integer
            If secondaryAddressComboBox.SelectedIndex <> 0 Then
                currentSecondaryAddress = secondaryAddressComboBox.SelectedItem
            Else
                currentSecondaryAddress = 0
            End If

            GpibDevice = New Device(CInt(boardIdNumericUpDown.Value), CByte(primaryAddressNumericUpDown.Value), CByte(currentSecondaryAddress))
            SetupControlState(True)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub 'openButton_Click

    Private Sub closeButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles closeButton.Click
        Try
            GpibDevice.Dispose()
            SetupControlState(False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub 'closeButton_Click

    Private Sub writeButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles writeButton.Click
        Try
            GpibDevice.Write(ReplaceCommonEscapeSequences(stringToWriteTextBox.Text))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub 'writeButton_Click

    Private Sub readButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles readButton.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            stringReadTextBox.Text = InsertCommonEscapeSequences(GpibDevice.ReadString())
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub 'readButton_Click

    Private Function ReplaceCommonEscapeSequences(ByVal s As String) As String
        Return s.Replace("\n", ControlChars.Lf).Replace("\r", ControlChars.Cr)
    End Function 'ReplaceCommonEscapeSequences

    Private Function InsertCommonEscapeSequences(ByVal s As String) As String
        Return s.Replace(ControlChars.Lf, "\n").Replace(ControlChars.Cr, "\r")
    End Function 'InsertCommonEscapeSequences

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim LaserLambda As String = "", LaserPower As String = ""
        Dim WavemeterPower As String = "", WavemeterLambda As String = ""
        Dim SetPower As New ArrayList, MeasuredPower As New ArrayList, SetLambda As New ArrayList, MeasuredLambda As New ArrayList
        Dim Freq As Double = 0, Steps As Double = 1, Wavelength As Double = 0
        Dim Counter As Integer = 0

        Dim StartWavelength As Double = Convert.ToDouble(StartLambda_TB.Text)
        Dim StopWavelength As Double = Convert.ToDouble(StopLambda_TB.Text)
        Dim StepSize As Double = Convert.ToDouble(StepSize_TB.Text)

        Dim dataTable1 As New DataTable("DataSet")
        dataTable1.Columns.Add("X", GetType(Double))
        dataTable1.Columns.Add("Y", GetType(Double))

        FastLine1.AutoRepaint = False ' Redraw in live data mode
        FastLine1.DataSource = dataTable1 ' Set source for data
        'Set Ordering to none, to increment speed when adding points
        FastLine1.XValues.Order = Steema.TeeChart.Styles.ValueListOrder.None

        Steps = StepSize_TB.Text

        ' Reset Wavemeter
        ' Read Wavemeter and Power from the Agilent on GPIB 11
        Try
            GpibDevice = New Device(0, 12)
            ' GpibDevice.PrimaryAddress = 12
            SetupControlState(True)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
        ' Reset Agilent Wavemeter
        Try
            GpibDevice.Write("*RST") ' Reset command
            System.Threading.Thread.Sleep(500)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        GpibDevice.Dispose()
        SetupControlState(False)

        ' Set wavelength and read back from TLS
        ' *************************************
        If FSC_CB.Checked = True Then
            ' Loop the frequency setting until the frequency range has been fully tested
            Do Until Counter > 449
                'Select and Start Sweep of the tunable laser GPIB address 10
                Try
                    GpibDevice = New Device(0, 11)
                    'GpibDevice.PrimaryAddress = 11
                    SetupControlState(True)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    Windows.Forms.Cursor.Current = Cursors.Default
                End Try
                Try
                    If Counter = 0 Then
                        Wavelength = Convert.ToDouble(StartLambda_TB.Text) - 0.0224 ' Set -22.4pm or -0.0224nm for full range sweep
                        GpibDevice.Write("DBM") ' Set to dBm mode
                        System.Threading.Thread.Sleep(500)
                        GpibDevice.Write("L=" + StartLambda_TB.Text)
                        System.Threading.Thread.Sleep(500)
                    End If
                    GpibDevice.Write("FSCL=" + Convert.ToString(Wavelength)) ' Add 0.1pm to current wavelength set point
                    System.Threading.Thread.Sleep(500)
                    ' Read Wavelength
                    GpibDevice.Write("L?")
                    System.Threading.Thread.Sleep(500)
                    LaserLambda = GpibDevice.ReadString()
                    LaserLambda = LaserLambda.Replace("L", "")
                    LaserLambda = LaserLambda.Replace("=", "")
                    LaserLambda = LaserLambda.Replace("\n", "")
                    LaserLambda = LaserLambda.Replace(" ", "")
                    LaserLambda = LaserLambda.Remove(LaserLambda.Length - 1, 1) ' Remove last character

                    SetLambda.Add(Convert.ToDouble(LaserLambda))
                    ' Read Power Setting
                    GpibDevice.Write("P?")
                    System.Threading.Thread.Sleep(500)
                    LaserPower = GpibDevice.ReadString()
                    LaserPower = LaserPower.Replace("P", "")
                    LaserPower = LaserPower.Replace("=", "")
                    LaserPower = LaserPower.Replace("\n", "")
                    LaserPower = LaserPower.Replace(" ", "")
                    SetPower.Add(LaserPower)

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
                GpibDevice.Dispose()
                SetupControlState(False)
                ' Read Wavemeter and Power from the Agilent on GPIB 11
                Try
                    GpibDevice = New Device(0, 12)
                    ' GpibDevice.PrimaryAddress = 12
                    SetupControlState(True)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    Windows.Forms.Cursor.Current = Cursors.Default
                End Try
                ' Read frequency and power from Agilent Wavemeter
                Try
                    GpibDevice.Write(":MEAS:SCAL:POW:WAV?") '+1.52596000E-006\n
                    System.Threading.Thread.Sleep(500)
                    WavemeterLambda = GpibDevice.ReadString() ' +1.52596000E-006\n
                    System.Threading.Thread.Sleep(500)
                    WavemeterLambda = WavemeterLambda.Replace("\n", "") ' Trim "\n" characters
                    WavemeterLambda = WavemeterLambda.Replace(" ", "") ' Trim " " characters

                    GpibDevice.Write(":MEASure:ARRay:POWer?") ' 1,-1.37523400E+000\n
                    System.Threading.Thread.Sleep(500)
                    WavemeterPower = GpibDevice.ReadString()
                    System.Threading.Thread.Sleep(500)
                    WavemeterPower = WavemeterPower.Replace("\n", "") ' Trim "\n" characters
                    WavemeterPower = WavemeterPower.Replace(" ", "") ' Trim "\n" characters
                    WavemeterPower = WavemeterPower.Replace("+", "") ' Trim "\n" characters
                    WavemeterPower = WavemeterPower.Remove(0, 2) ' Remove 1st 2 characeters
                    WavemeterPower = WavemeterPower.Remove(WavemeterPower.Length - 1, 1) ' Remove last character
                    Dim Power As Double = WavemeterPower.Substring(WavemeterPower.Length - 3) ' Select Power digits
                    WavemeterPower = WavemeterPower.Remove(WavemeterPower.Length - 4, 4) ' Remove last 4 characters
                    Power = 10 ^ Power
                    WavemeterPower = WavemeterPower * Power

                    ' Convert to Decimal from Scientific notation
                    ' Add to arraylists
                    MeasuredPower.Add(Convert.ToDouble(WavemeterPower))
                    MeasuredLambda.Add(WavemeterLambda)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
                GpibDevice.Dispose()
                SetupControlState(False)

                ' Plot data live Measured Power(dBm) and Wavelength(nm)
                ' Create DataTable with variables to plot
                dataTable1.Rows.Add({SetLambda(Counter), MeasuredPower(Counter)})

                FastLine1.CheckDataSource()
                FastLine1.RefreshSeries() ' Update plot with new data
                FastLine1.Repaint() ' Redraw plot
                FastLine1.Draw() ' Redraw??

                ' Save Data to Log File
                'Create Test folder for UID if not already created
                Dim sPath As String = "C:\Test Folder" + "\" + Filename_TB.Text
                If (My.Computer.FileSystem.DirectoryExists(sPath) = False) Then
                    My.Computer.FileSystem.CreateDirectory(sPath + "\" + Filename_TB.Text)
                Else
                    'Nothing else happens, because the folder exists
                End If

                ' SAVE DATA TO LOG FILE
                ' Write test settings to file on initial loop of the test
                If Counter = 0 Then
                    My.Computer.FileSystem.WriteAllText(sPath + "\" + "Log File-" + Filename_TB.Text + ".txt", _
                    "Laser Power(dBm)," + "Laser Wavelength(nm)," + "Wavemeter Power(dBm)," + "Wavemeter Wavelenth(nm)" & vbCrLf, True) ' Headings for data columns
                End If
                If Counter <> 0 Then
                    ' Write to text file data in case the power is lost during test
                    My.Computer.FileSystem.WriteAllText(sPath + "\" + "Log File-" + Filename_TB.Text + ".txt", _
                        LaserPower + "," + LaserLambda + "," + WavemeterPower + "," + WavemeterLambda & vbCrLf, True)
                End If

                ' Counter and Wavelength Increment
                Wavelength += 0.1 ' Add 0.1pm (i.e. dlambda=+0.1pm) to wavelength setting
                Counter += 1
            Loop
        End If
        ' **************************
        If FSC_CB.Checked = False Then
            ' Loop the frequency setting until the frequency range has been fully tested
            Do Until Wavelength >= StopWavelength
                'Select and Start Sweep of the tunable laser GPIB address 10
                Try
                    GpibDevice = New Device(0, 11)
                    'GpibDevice.PrimaryAddress = 11
                    SetupControlState(True)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    Windows.Forms.Cursor.Current = Cursors.Default
                End Try
                Try
                    If Counter = 0 Then
                        Wavelength = Convert.ToDouble(StartLambda_TB.Text) ' Add 0.001nm or 1pm to current wavelength set point
                        GpibDevice.Write("DBM") ' Set to dBm mode
                        System.Threading.Thread.Sleep(500)
                        GpibDevice.Write("L=" + StartLambda_TB.Text)
                        System.Threading.Thread.Sleep(500)
                    End If
                    GpibDevice.Write("L=" + Convert.ToString(Wavelength)) ' Add 0.001nm or 1pm to current wavelength set point
                    System.Threading.Thread.Sleep(500)

                    ' Read Wavelength
                    GpibDevice.Write("L?")
                    System.Threading.Thread.Sleep(500)
                    LaserLambda = GpibDevice.ReadString()
                    LaserLambda = LaserLambda.Replace("L", "")
                    LaserLambda = LaserLambda.Replace("=", "")
                    LaserLambda = LaserLambda.Replace("\n", "")
                    LaserLambda = LaserLambda.Replace(" ", "")
                    LaserLambda = LaserLambda.Remove(LaserLambda.Length - 1, 1) ' Remove last character

                    SetLambda.Add(Convert.ToDouble(LaserLambda))
                    ' Read Power Setting
                    GpibDevice.Write("P?")
                    System.Threading.Thread.Sleep(500)
                    LaserPower = GpibDevice.ReadString()
                    LaserPower = LaserPower.Replace("P", "")
                    LaserPower = LaserPower.Replace("=", "")
                    LaserPower = LaserPower.Replace("\n", "")
                    LaserPower = LaserPower.Replace(" ", "")
                    SetPower.Add(LaserPower)

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
                GpibDevice.Dispose()
                SetupControlState(False)
                ' Read Wavemeter and Power from the Agilent on GPIB 11
                Try
                    GpibDevice = New Device(0, 12)
                    ' GpibDevice.PrimaryAddress = 12
                    SetupControlState(True)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    Windows.Forms.Cursor.Current = Cursors.Default
                End Try
                ' Read frequency and power from Agilent Wavemeter
                Try
                    GpibDevice.Write(":MEAS:SCAL:POW:WAV?") '+1.52596000E-006\n
                    System.Threading.Thread.Sleep(500)
                    WavemeterLambda = GpibDevice.ReadString() ' +1.52596000E-006\n
                    System.Threading.Thread.Sleep(500)
                    WavemeterLambda = WavemeterLambda.Replace("\n", "") ' Trim "\n" characters
                    WavemeterLambda = WavemeterLambda.Replace(" ", "") ' Trim " " characters

                    GpibDevice.Write(":MEASure:ARRay:POWer?") ' 1,-1.37523400E+000\n
                    System.Threading.Thread.Sleep(500)
                    WavemeterPower = GpibDevice.ReadString()
                    System.Threading.Thread.Sleep(500)
                    WavemeterPower = WavemeterPower.Replace("\n", "") ' Trim "\n" characters
                    WavemeterPower = WavemeterPower.Replace(" ", "") ' Trim "\n" characters
                    WavemeterPower = WavemeterPower.Replace("+", "") ' Trim "\n" characters
                    WavemeterPower = WavemeterPower.Remove(0, 2) ' Remove 1st 2 characeters
                    WavemeterPower = WavemeterPower.Remove(WavemeterPower.Length - 1, 1) ' Remove last character
                    Dim Power As Double = WavemeterPower.Substring(WavemeterPower.Length - 4) ' Select Power digits
                    WavemeterPower = WavemeterPower.Remove(WavemeterPower.Length - 5, 5) ' Remove last 5 characters
                    Power = 10 ^ Power
                    WavemeterPower = WavemeterPower * Power

                    ' Convert to Decimal from Scientific notation
                    ' Add to arraylists
                    MeasuredPower.Add(Convert.ToDouble(WavemeterPower))
                    MeasuredLambda.Add(WavemeterLambda)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
                GpibDevice.Dispose()
                SetupControlState(False)

                ' Plot data live Measured Power(dBm) and Wavelength(nm)
                ' Create DataTable with variables to plot
                dataTable1.Rows.Add({SetLambda(Counter), MeasuredPower(Counter)})

                FastLine1.CheckDataSource()
                FastLine1.RefreshSeries() ' Update plot with new data
                FastLine1.Repaint() ' Redraw plot
                'FastLine1.Draw() ' Redraw??

                ' Save Data to Log File
                'Create Test folder for UID if not already created
                Dim sPath As String = "C:\Test Folder" + "\" + Filename_TB.Text
                If (My.Computer.FileSystem.DirectoryExists(sPath) = False) Then
                    My.Computer.FileSystem.CreateDirectory(sPath + "\" + Filename_TB.Text)
                Else
                    'Nothing else happens, because the folder exists
                End If

                ' SAVE DATA TO LOG FILE
                ' Write test settings to file on initial loop of the test
                If Counter = 0 Then
                    My.Computer.FileSystem.WriteAllText(sPath + "\" + "Log File-" + Filename_TB.Text + ".txt", _
                    "Laser Power(dBm)," + "Laser Wavelength(nm)," + "Wavemeter Power(dBm)," + "Wavemeter Wavelenth(nm)" & vbCrLf, True) ' Headings for data columns
                End If
                If Counter <> 0 Then
                    ' Write to text file data in case the power is lost during test
                    My.Computer.FileSystem.WriteAllText(sPath + "\" + "Log File-" + Filename_TB.Text + ".txt", _
                        LaserPower + "," + LaserLambda + "," + WavemeterPower + "," + WavemeterLambda & vbCrLf, True)
                End If

                ' Counter and Wavelength Increment
                Wavelength += StepSize ' Add >=0.001nm to wavelength setting depending on GUI entry
                Counter += 1
            Loop
        End If ' Do if the FSC CHeckbox is False
        ' **************************************
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim XValue As Array = {1, 2, 3, 4, 5, 6, 7, 8}, YValue As Array = {0.6, 1, 1.5, 3.2, 5.3, 9.2, 45.3, 100}
        Dim Count As Integer = 0

        Dim dataTable1 As New DataTable("DataSet")
        dataTable1.Columns.Add("X", GetType(Double))
        dataTable1.Columns.Add("Y", GetType(Double))
        Do Until Count > XValue.Length - 1
            dataTable1.Rows.Add({XValue(Count), YValue(Count)})
            Count += 1
        Loop
        FastLine1.DataSource = dataTable1
        FastLine1.RefreshSeries()
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Turn on laser and set the power in dBm
        Try
            GpibDevice = New Device(0, 11)
            'GpibDevice.PrimaryAddress = 11
            SetupControlState(True)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
        Try
            GpibDevice.Write("P=" + Power_TB.Text) ' Set power level dBm
            GpibDevice.Write("ENABLE") ' Enable the laser output
            System.Threading.Thread.Sleep(500)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        GpibDevice.Dispose()
        SetupControlState(False)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TestTime_TB.Text = ((Convert.ToDouble(StopLambda_TB.Text) - Convert.ToDouble(StartLambda_TB.Text)) / Convert.ToDouble(StepSize_TB.Text) * 6) / 60
        TestTime_TB.Refresh()
    End Sub
End Class 'MainForm