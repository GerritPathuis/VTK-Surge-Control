Imports System.IO
Imports System.IO.Ports
Imports System.Math

Public Class Form1
    Dim time As Double

    Dim Cout(4) As Double      'Current Outputs

    Dim myPort As Array  'COM Ports detected on the system will be stored here
    Dim comOpen As Boolean
    Private Property ConnectionOK As Boolean

    Dim Flow_in, Flow_out As Double     '[m3/hr]
    Dim Temp_in, Temp_out As Double     '[Celsius]
    Dim Press_in, Press_out As Double   '[Pa]

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox16.Text =
        "Based on " & vbCrLf &
        "Anti Surge Control Test Procedure, Guus van Gemert 2017" & vbCrLf &
        "KIMA, Ammonia & Urea Fertilizer Project, " & vbCrLf &
        "Start-up blower HD2 407/1235/T16B, 3000 rpm, 700 kW" & vbCrLf &
        "VTK project P16.0078" & vbCrLf &
        "Bypass valve size is 6 inch, valve speed Is 3 seconds"

        TextBox17.Text =
        "Hardware " & vbCrLf &
        "4 channel Analog Input (4-20mAmp) USB Module, LucidControl AI4" & vbCrLf &
        "4 channel Analog Output (4-20mAmp) USB Module, LucidControl AO4" & vbCrLf &
        "Install a Windows drive, see Lucid-Control.com For download"

        TextBox18.Text =
        "Bypass valve sizing" & vbCrLf &
        "Size To handle 50% Or more Of the maximum flow" & vbCrLf &
        "Valve speed 3-5 seconds" & vbCrLf &
        " "

        TextBox19.Text =
        "Test set-up" & vbCrLf &
        "The Laptop is connect to a AO4 and AAI4 LucidControl unit." & vbCrLf &
        "The Laptop program generates 4 signals and receives 1 each are 4-20mAmp" & vbCrLf &
        "The send signals represent Flow inlet/outlet [Am3/hr], Temp fan inlet [c], " & vbCrLf &
        "Pressure [Pa] fan inlet and delta pressure over the fan [Pa]" & vbCrLf &
        "The receives signals represent the position of the bypass valve" & vbCrLf &
        " "

        TextBox20.Text =
        "Test procedure" & vbCrLf &
        "Start situation is stable, sitting on Fan Curve" & vbCrLf &
        "The system flow Coefficient Ksys is changed resulting  " & vbCrLf &
        "in moving to another spot on the fan Curve." & vbCrLf &
        "When we near the Surge-area the connected ASC must react by" & vbCrLf &
        "opening the bybass valve and returning to a save spot on the" & vbCrLf &
        "fan-Curve."
        Reset()
    End Sub

    Private Sub Reset()
        Init_Chart1()
        Timer1.Interval = 1000   'Berekeningsinterval 1 sec
        time = 0

        Timer1.Enabled = True
    End Sub

    Private Sub Init_Chart1()
        Try
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()

            Chart1.Series.Add("Valve stem")
            'Chart1.Series.Add("Valve stem")

            Chart1.ChartAreas.Add("ChartArea0")
            Chart1.Series(0).ChartArea = "ChartArea0"

            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line

            Chart1.Titles.Add("Surge Control Simulation")
            Chart1.Titles(0).Font = New Font("Arial", 16, System.Drawing.FontStyle.Bold)

            Chart1.Series(0).Name = ""  'Valve stem position"
            Chart1.Series(0).Color = Color.Black
            Chart1.Series(0).BorderWidth = 3

            Chart1.ChartAreas("ChartArea0").AxisX.Title = "[sec]"
            Chart1.ChartAreas("ChartArea0").AxisY.Title = "Valve opening"
            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart1.ChartAreas("ChartArea0").AxisY.Maximum = 100
            Chart1.ChartAreas("ChartArea0").AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Vertical
            'Chart1.Series(0).YAxisType = AxisType.Primary

        Catch ex As Exception
            MessageBox.Show("Init Chart1 failed")
        End Try
    End Sub
    Private Sub Draw_Chart1()
        Try
            'Chart1.Series(0).Points.AddXY(time, new_valve_pos)
        Catch ex As Exception
            MessageBox.Show("Draw chart1 failed")
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Reset() 'Reset button
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim setup_string As String

        setup_string = "LucidIoCtrl –dCOM1 –tV –c0,1,2,3 –w"
        setup_string &= Cout(1).ToString("0.000") & ","
        setup_string &= Cout(2).ToString("0.000") & ","
        setup_string &= Cout(3).ToString("0.000") & ","
        setup_string &= Cout(4).ToString("0.000") & vbCrLf

        'setup_string = "LucidIoCtrl –dCOM1 –tV –c0,1,2,3 –w5.000,2.500,1.250,0.625" & vbCrLf
        Send_data(setup_string)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        draw_Chart1()
        time += Timer1.Interval / 1000                  '[sec]

        ' calc_flow()
    End Sub

    Private Sub Serial_setup() 'Serial port setup
        If (Me.SerialPort1.IsOpen = True) Then ' Preventing exceptions
            Me.SerialPort1.DiscardInBuffer()
            Me.SerialPort1.Close()
        End If

        Try
            Me.myPort = SerialPort.GetPortNames() 'Get all com ports available
            For Each port In myPort
                Me.cmbPort.Items.Add(port)
            Next port
            Me.cmbPort.Text = cmbPort.Items.Item(0)    'Set cmbPort text to the first COM port detected
        Catch ex As Exception
            MsgBox("No com ports detected")
        End Try

        Me.cmbBaud.Items.Add(9600)     'Populate the cmbBaud Combo box to common baud rates used
        Me.cmbBaud.Items.Add(19200)
        Me.cmbBaud.Items.Add(38400)
        Me.cmbBaud.SelectedIndex = 0    'Set cmbBaud text to 9600 Baud 

        Me.SerialPort1.ReceivedBytesThreshold = 24    'wait EOF char or until there are x bytes in the buffer, include \n and \r !!!!
        Me.SerialPort1.ReadBufferSize = 4096
        Me.SerialPort1.DiscardNull = True              'important otherwise it will not work
        Me.SerialPort1.Parity = Parity.None
        Me.SerialPort1.StopBits = StopBits.One
        Me.SerialPort1.Handshake = Handshake.None
        Me.SerialPort1.ParityReplace = True
        btnDisconnect.Enabled = False                  'Initially Disconnect Button is Disabled
    End Sub

    Private Sub CmbPort_Click(sender As Object, e As EventArgs) Handles cmbPort.Click
        cmbPort.SelectedIndex = -1
        cmbPort.Items.Clear()
        Serial_setup()
    End Sub

    Private Sub BtnConnect_Click(sender As System.Object, e As System.EventArgs) Handles btnConnect.Click
        Me.SerialPort1.Close()                     'Close existing 
        If cmbPort.Text.Length = 0 Then
            MsgBox("Sorry, did not find any connected USB Balancers")
        Else
            Me.SerialPort1.PortName = cmbPort.Text         'Set SerialPort1 to the selected COM port at startup
            Me.SerialPort1.BaudRate = cmbBaud.Text         'Set Baud rate to the selected value on

            'Other Serial Port Property
            Me.SerialPort1.Parity = IO.Ports.Parity.None
            Me.SerialPort1.StopBits = IO.Ports.StopBits.One
            Me.SerialPort1.DataBits = 8                  'Open our serial port

            Try
                Me.SerialPort1.Open()
                btnConnect.Enabled = False              'Disable Connect button
                btnConnect.BackColor = Color.Yellow
                cmbPort.BackColor = Color.Yellow
                btnConnect.Text = "OK connected"
                btnDisconnect.Enabled = True            'and Enable Disconnect button
            Catch ex As Exception
                MsgBox("Error 654 Open: " & ex.Message)
            End Try

            Try
                Me.SerialPort1.DiscardInBuffer()        'empty inbuffer
                Me.SerialPort1.DiscardOutBuffer()       'empty outbuffer

                Me.SerialPort1.WriteLine("s")           'Real time samples to PC
            Catch ex As Exception
                MsgBox("Error 786 Open: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub BtnDisconnect_Click(sender As System.Object, e As System.EventArgs) Handles btnDisconnect.Click
        Try
            Me.SerialPort1.DiscardInBuffer()
            Me.SerialPort1.Close()             'Close our Serial Port
            Me.SerialPort1.Dispose()
            btnConnect.Enabled = True
            btnConnect.BackColor = Color.Red
            btnConnect.Text = "Connect"
            btnDisconnect.Enabled = False
        Catch ex As Exception
            MsgBox("Error 104 Open: " & ex.Message)
        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox9.Text = NumericUpDown26.Value.ToString 'dp fan
    End Sub

    Private Sub Send_data(rtbtransmit As String)
        Try
            If Me.SerialPort1.IsOpen = False Then
                Me.SerialPort1.WriteLine(rtbtransmit) 'The text contained in the txtText will be sent to the serial port as ascii
            End If
        Catch exc As IOException
            Console.WriteLine("Error nr 887 IO exception" & exc.Message)
        End Try
    End Sub

    Private Sub SerialPort1_DataReceived(sender As System.Object, e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Try
            ReceivedText(SerialPort1.ReadLine())    'Automatically called every time a data is received at the serialPortb
        Catch exc As IOException
            Console.WriteLine("Error 453 IO exception" & exc.Message)
        End Try
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown32.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown27.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown33.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown22.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown18.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged
        Dim Range(2) As String
        Dim K_sys, K_bypass, k_sum, K100, valve_open, dp, ro As Double
        Dim A, B, C, Qv_in, Qv_out As Double
        Dim Pin, Pout As Double
        Dim Tin, Tout As Double
        Dim γ As Double

        Range(0) = NumericUpDown28.Value - NumericUpDown27.Value    'Flow
        Range(1) = NumericUpDown29.Value - NumericUpDown30.Value    'Temp
        Range(2) = NumericUpDown31.Value - NumericUpDown32.Value    'Pressure

        ro = NumericUpDown19.Value          'Density [kg/Am3]
        A = NumericUpDown17.Value           'Fan Curve [-]
        B = NumericUpDown16.Value           'Fan Curve [-]
        C = NumericUpDown20.Value           'Fan Curve [-]
        K100 = NumericUpDown21.Value        'K-value at 100% open [-]
        valve_open = NumericUpDown33.Value  'Position bypass valve [%]
        Pin = NumericUpDown18.Value         'Pressure inlet fan [Pa]
        Tin = NumericUpDown23.Value         'Temp inlet fan [c]
        γ = NumericUpDown22.Value     'Poly tropic exponent γ

        If ro > 0 Then 'to prevent exceptions
            '----- step 1 determin the K values----
            K_bypass = K100 * valve_open
            K_sys = NumericUpDown25.Value
            k_sum = K_sys + K_bypass

            '----- step 2 determine qv---
            Double.TryParse(TextBox9.Text, dp)
            Qv_in = Sqrt(dp / ro * k_sum ^ 2)
            '----- step 3 determine new dp---
            dp = ro * (A * Qv_in ^ 2 + B * Qv_in + C)   'Fan curve

            '----- step 4 determine Temp outlet fan ---
            Tout = Tin * (1 + dp / Pin) ^ ((γ - 1) / γ)

            '----- step 5 determine Pressure outlet fan ---
            Pout = Pin + dp

            '----- step 6 determine Discharge flow fan ---
            Qv_out = Qv_in * (Pin / Pout) ^ γ


            TextBox7.Text = Round(K_bypass, 2).ToString   'Resistance Bypass valve 
            TextBox8.Text = Round(Qv_out, 0).ToString
            TextBox9.Text = Round(dp, 0)
            TextBox10.Text = Round(Tin, 1).ToString
            TextBox11.Text = Range(0).ToString
            TextBox12.Text = Range(1).ToString
            TextBox13.Text = Range(2).ToString
            TextBox14.Text = Round(K_sys, 2).ToString   'Resistance Total system
            TextBox15.Text = Round(Qv_in, 0).ToString
            TextBox21.Text = Round(Pin, 0).ToString     'Pressure inlet
            'TextBox21.Text = Round(valve_open, 0).ToString  'Bypass valve
            TextBox25.Text = Round(Tout, 1).ToString
        End If

        '---------- calc output currents
        Cout(1) = Calc_output("Flow", Qv_in)
        Cout(2) = Calc_output("Pressure", Pin)
        Cout(3) = Calc_output("Pressure", dp)
        Cout(4) = Calc_output("Temperature", Tin)

        TextBox1.Text = Round(Cout(1), 1).ToString  'Flow inlet/out Actual [Am3/hr]
        TextBox2.Text = Round(Cout(2), 1).ToString  'Pressure in [Pa]
        TextBox3.Text = Round(Cout(3), 1).ToString  'Delta P [Pa]
        TextBox23.Text = Round(Cout(4), 1).ToString 'Temp fan in [c]
    End Sub
    Private Function Calc_output(outType As String, value As Double) As Double
        Dim results, range, value_4ma As Double
        Select Case outType
            Case "Flow"
                value_4ma = NumericUpDown27.Value
                Double.TryParse(TextBox11.Text, range)
                results = (value - value_4ma) / range * 16.0 + 4.0
            Case "Temperature"
                value_4ma = NumericUpDown30.Value
                Double.TryParse(TextBox12.Text, range)
                results = (value - value_4ma) / range * 16.0 + 4.0
            Case "Pressure"
                value_4ma = NumericUpDown32.Value
                Double.TryParse(TextBox13.Text, range)
                results = (value - value_4ma) / range * 16.0 + 4.0
            Case Else
                MessageBox.Show("Oops error in Calc_output function")
        End Select
        Return (results)
    End Function

    Private Sub ReceivedText(ByVal intext As String)
        MessageBox.Show(intext)
    End Sub

End Class
