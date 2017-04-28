Imports System.Globalization
Imports System.IO
Imports System.IO.Ports
Imports System.Math
Imports System.Threading

Public Class Form1
    Dim time As Double

    Dim pv(4) As Double             'Process values 0,1,2,3,4
    Dim Cout(4) As Double           'Current Outputs
    Dim last_deviation As Double    'PID control
    Dim Pterm, Iterm, Dterm As Double

    Dim myPort As Array  'COM Ports detected on the system will be stored here
    Dim comOpen As Boolean
    Private Property ConnectionOK As Boolean

    Dim Flow_in, Flow_out As Double     '[m3/hr]
    Dim Temp_in, Temp_out As Double     '[Celsius]
    Dim Press_in, Press_out As Double   '[Pa]
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

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

        TextBox24.Text = "6.8" 'Test value [c]

        For i = 0 To 3
            pv(i) = 1       'Initial value
        Next

        Reset()
        Update_calc_screen()
    End Sub

    Private Sub Reset()
        Init_Chart1()
        Init_Chart2()
        Timer1.Interval = 1000   'Berekeningsinterval 1000 msec
        time = 0

        Timer1.Enabled = True
        TextBox31.Text = "50"   'PID controller output 50%
    End Sub
    Private Sub Init_Chart1()
        Dim i As Integer
        Try
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()
            Chart1.ChartAreas.Add("ChartArea0")

            For i = 0 To 4
                Chart1.Series.Add(i.ToString)
                Chart1.Series(i.ToString).ChartArea = "ChartArea0"
                Chart1.Series(i.ToString).ChartType = DataVisualization.Charting.SeriesChartType.Line
                Chart1.Series(i.ToString).BorderWidth = 1
            Next

            Chart1.Titles.Add("ASC testing")
            Chart1.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)

            Chart1.Series(0).Name = "Flow inlet"
            Chart1.Series(1).Name = "Pressure in"
            Chart1.Series(2).Name = "delta P"
            Chart1.Series(3).Name = "Temp in"
            Chart1.Series(4).Name = "Bypass valve"
            Chart1.Series(0).Color = Color.Black
            Chart1.Series(0).BorderWidth = 2
            Chart1.Series(2).BorderWidth = 2

            Chart1.ChartAreas("ChartArea0").AxisX.Title = "[sec]"
            Chart1.ChartAreas("ChartArea0").AxisY.Title = "mAmp"
            Chart1.ChartAreas("ChartArea0").AxisY.Minimum = 4
            Chart1.ChartAreas("ChartArea0").AxisY.Maximum = 20
            Chart1.ChartAreas("ChartArea0").AxisX.MajorTickMark.Size = 2
            Chart1.ChartAreas("ChartArea0").AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Vertical

        Catch ex As Exception
            MessageBox.Show("Init Chart1 failed")
        End Try
    End Sub
    Private Sub Init_Chart2()
        Dim i As Integer

        Try
            Chart2.Series.Clear()
            Chart2.ChartAreas.Clear()
            Chart2.Titles.Clear()
            Chart2.ChartAreas.Add("ChartArea1")

            For i = 0 To 4
                Chart2.Series.Add(i.ToString)
                Chart2.Series(i.ToString).ChartArea = "ChartArea1"
                Chart2.Series(i.ToString).ChartType = DataVisualization.Charting.SeriesChartType.Line
                Chart2.Series(i.ToString).BorderWidth = 1
            Next

            Chart2.Titles.Add("ASC testing")
            Chart2.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)

            Chart2.Series(0).Name = "Flow inlet"
            Chart2.Series(1).Name = "Pressure in"
            Chart2.Series(2).Name = "delta P"
            Chart2.Series(3).Name = "Temp in"
            Chart2.Series(4).Name = "Bypass valve"
            Chart2.Series(0).Color = Color.Black
            Chart2.Series(0).BorderWidth = 2
            Chart2.Series(2).BorderWidth = 2

            Chart2.ChartAreas("ChartArea1").AxisX.Title = "[sec]"
            Chart2.ChartAreas("ChartArea1").AxisY.Title = "mAmp"
            Chart2.ChartAreas("ChartArea1").AxisY.Minimum = 4
            Chart2.ChartAreas("ChartArea1").AxisY.Maximum = 20
            Chart2.ChartAreas("ChartArea1").AxisX.MajorTickMark.Size = 2
            Chart2.ChartAreas("ChartArea1").AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Vertical

        Catch ex As Exception
            MessageBox.Show("Init Chart2 failed")
        End Try
    End Sub

    Private Sub Draw_Chart1()
        Try
            Chart1.Series(0).Points.AddXY(time, Cout(1))    'Flow in
            Chart1.Series(1).Points.AddXY(time, Cout(2))    'Pressure in
            Chart1.Series(2).Points.AddXY(time, Cout(3))    'dP
            Chart1.Series(3).Points.AddXY(time, Cout(4))    'Temp in

            Chart2.Series(0).Points.AddXY(time, Cout(1))    'Flow in
            Chart2.Series(1).Points.AddXY(time, Cout(2))    'Pressure in
            Chart2.Series(2).Points.AddXY(time, Cout(3))    'dP
            Chart2.Series(3).Points.AddXY(time, Cout(4))    'Temp in
        Catch ex As Exception
            MessageBox.Show("AddXY failed")
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim setup_string As String

        '---- This is in the test Tab-------------
        setup_string = "LucidIoCtrl –dCOM5 –tV –c0,1,2,3 –w"
        setup_string &= NumericUpDown5.Value.ToString("0.000") & ","
        setup_string &= NumericUpDown10.Value.ToString("0.000") & ","
        setup_string &= NumericUpDown14.Value.ToString("0.000") & ","
        setup_string &= NumericUpDown15.Value.ToString("0.000") & vbCrLf

        'setup_string = "LucidIoCtrl –dCOM5 –tV –c0,1,2,3 –w5.000,2.500,1.250,0.625" 'example
        Send_data(setup_string)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Dim GetIoGroup(4) As Byte   'Voltage, see Page 22 en 24
        Dim GetId(4) As Byte        'Getid Page 30

        GetIoGroup(1) = &H48   'OPC= GetIoGroup
        GetIoGroup(2) = &HF    'Channel 1
        GetIoGroup(3) = &H1C   'Voltage
        GetIoGroup(4) = &H0    'LEN

        GetId(1) = &HC0        'OPC= GetId
        GetId(2) = &H0
        GetId(3) = &H0
        GetId(4) = &H0


        time += Timer1.Interval / 1000                  '[msec]--->[sec]
        Label1.Text = time.ToString("000.0")

        ' GetIoGroupa = BitConverter.GetBytes(Long.Parse(GetIoGroup, NumberStyles.AllowHexSpecifier))

        If SerialPort1.IsOpen Then
            '----- what is connected -------
            SerialPort1.Write(GetId, 1, 4)

            '----------LucidControl Input module -------------
            'SerialPort1.Write(GetIoGroup, 1, 4)

            'TextBox26.Text &= "."
        End If

        Update_calc_screen()
        Draw_Chart1()
        PID_controller()
    End Sub

    Private Sub Serial_setup() 'Serial port setup
        If (SerialPort1.IsOpen = True) Then ' Preventing exceptions
            SerialPort1.DiscardInBuffer()
            SerialPort1.Close()
        End If

        Try
            myPort = SerialPort.GetPortNames() 'Get all com ports available
            For Each port In myPort
                cmbPort.Items.Add(port)
            Next port
            cmbPort.Text = CType(cmbPort.Items.Item(0), String)    'Set cmbPort text to the first COM port detected
        Catch ex As Exception
            MsgBox("No COM ports detected")
        End Try

        cmbBaud.Items.Add(9600)     'Populate the cmbBaud Combo box to common baud rates used
        cmbBaud.Items.Add(19200)
        cmbBaud.Items.Add(38400)
        cmbBaud.SelectedIndex = 0    'Set cmbBaud text to 9600 Baud 

        SerialPort1.ReadBufferSize = 4096
        SerialPort1.DiscardNull = True              'important otherwise it will not work
        SerialPort1.Parity = Parity.None
        SerialPort1.StopBits = StopBits.One
        SerialPort1.Handshake = Handshake.None
        SerialPort1.ParityReplace = CByte(True)
        btnDisconnect.Enabled = False                  'Initially Disconnect Button is Disabled
    End Sub

    Private Sub CmbPort_Click(sender As Object, e As EventArgs) Handles cmbPort.Click
        cmbPort.SelectedIndex = -1
        cmbPort.Items.Clear()
        Serial_setup()
    End Sub

    Private Sub BtnConnect_Click(sender As System.Object, e As System.EventArgs) Handles btnConnect.Click
        SerialPort1.Close()                     'Close existing 
        If cmbPort.Text.Length = 0 Then
            MsgBox("Sorry, did not find any connected Lucid Controllers")
        Else
            SerialPort1.PortName = cmbPort.Text         'Set SerialPort1 to the selected COM port at startup
            SerialPort1.BaudRate = CInt(cmbBaud.Text)         'Set Baud rate to the selected value on

            'Other Serial Port Property
            SerialPort1.Parity = IO.Ports.Parity.None
            SerialPort1.StopBits = IO.Ports.StopBits.One
            SerialPort1.DataBits = 8                  'Open our serial port

            Try
                SerialPort1.Open()
                btnConnect.Enabled = False              'Disable Connect button
                btnConnect.BackColor = Color.Yellow
                cmbPort.BackColor = Color.Yellow
                btnConnect.Text = "OK connected"
                btnDisconnect.Enabled = True            'and Enable Disconnect button
            Catch ex As Exception
                MsgBox("Error 654 Open: " & ex.Message)
            End Try

            Try
                SerialPort1.DiscardInBuffer()        'empty inbuffer
                SerialPort1.DiscardOutBuffer()       'empty outbuffer
            Catch ex As Exception
                MsgBox("Error 786 Open: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub BtnDisconnect_Click(sender As System.Object, e As System.EventArgs) Handles btnDisconnect.Click
        Try
            SerialPort1.DiscardInBuffer()
            SerialPort1.Close()             'Close our Serial Port
            SerialPort1.Dispose()
            btnConnect.Enabled = True
            btnConnect.BackColor = Color.Red
            btnConnect.Text = "Connect"
            btnDisconnect.Enabled = False
        Catch ex As Exception
            MsgBox("Error 104 Open: " & ex.Message)
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Reset()
    End Sub
    Private Sub Send_data(rtbtransmit As String)
        Try
            If SerialPort1.IsOpen Then
                SerialPort1.WriteLine(rtbtransmit) 'The text contained in the txtText will be sent to the serial port as ascii
            End If
        Catch exc As IOException
            MsgBox("Error nr 887 IO exception" & exc.Message)
        End Try
    End Sub
    Private Sub SerialPort1_DataReceived(sender As System.Object, e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Dim intext_hex As String = String.Empty
        Dim intext As String = String.Empty

        'see https://msdn.microsoft.com/en-us/library/05cts4c3(v=vs.110).aspx
        Try
            intext_hex = SerialPort1.ReadLine()

            If Not String.IsNullOrEmpty(intext_hex) Then
                intext &= StrToHex(intext_hex) & vbCrLf
                Invoke(Sub() TextBox26.Text &= intext)
            End If
        Catch exc As IOException
            MsgBox("Error 453 IO exception" & exc.Message)
        End Try
    End Sub

    Public Function StrToHex(Data As String) As String
        Dim sVal As String
        Dim sHex As String = ""
        While Data.Length > 0
            sVal = Conversion.Hex(Strings.Asc(Data.Substring(0, 1).ToString()))
            Data = Data.Substring(1)
            If sVal.Length < 2 Then
                sHex = sHex & "0" & sVal
            Else
                sHex = sHex & sVal
            End If
        End While
        Return sHex
    End Function


    Private Sub Update_calc_screen()
        Dim Range(3) As String
        Dim K_sys, K_bypass, k_sum, K100, valve_open, dp, ro As Double
        Dim A, B, C, Qv_in, Qv_out, A1 As Double
        Dim Pin, Pout As Double
        Dim Tin, Tout As Double
        Dim γ As Double
        Dim p_time, period, amplitude As Double
        Dim Qv_a, Qv_b As Double
        Range(0) = CType(NumericUpDown28.Value - NumericUpDown27.Value, String)    'Flow
        Range(1) = CType(NumericUpDown29.Value - NumericUpDown30.Value, String)    'Temp
        Range(2) = CType(NumericUpDown31.Value - NumericUpDown32.Value, String)    'Pressure
        Range(3) = CType(NumericUpDown13.Value - NumericUpDown34.Value, String)    'Pressure

        ro = NumericUpDown19.Value                  'Density [kg/Am3]
        A = NumericUpDown17.Value                   'Fan Curve [-]
        B = NumericUpDown16.Value                   'Fan Curve [-]
        C = NumericUpDown20.Value                   'Fan Curve [-]
        K100 = NumericUpDown21.Value                'K-value at 100% open [-]
        valve_open = NumericUpDown33.Value / 100    'Position bypass valve [%]
        Pin = NumericUpDown18.Value                 'Pressure inlet fan [Pa]
        Tin = NumericUpDown23.Value                 'Temp inlet fan [c]
        γ = NumericUpDown22.Value                   'Poly tropic exponent γ

        If ro > 0 Then 'to prevent exceptions
            '----- step 1 determin the K values----
            K_bypass = K100 * valve_open
            period = NumericUpDown7.Value
            amplitude = NumericUpDown2.Value
            p_time = time Mod period
            Select Case True
                Case RadioButton1.Checked           'Flat line
                    K_sys = NumericUpDown25.Value
                Case RadioButton2.Checked           'Square wave
                    If (p_time > (period / 2)) Then
                        K_sys = NumericUpDown25.Value + amplitude / 2
                    Else
                        K_sys = NumericUpDown25.Value - amplitude / 2
                    End If
                Case RadioButton3.Checked           'Sine
                    K_sys = NumericUpDown25.Value + (amplitude / 2) * Sin(p_time / period * 2 * PI)
                Case RadioButton4.Checked           'Saw tooth
                    K_sys = NumericUpDown25.Value - (amplitude / 2) + amplitude * p_time / period
            End Select
            TextBox5.Text = Round(K_sys, 1).ToString
            TextBox6.Text = Round(K_bypass, 1).ToString

            '----- step 2 determine qv---
            '------ ABC formula ----------
            k_sum = K_sys + K_bypass
            A1 = A - 1 / k_sum ^ 2
            Qv_a = (-B + (Sqrt(B ^ 2 - 4 * A1 * C))) / (2 * A1)
            Qv_b = (-B - (Sqrt(B ^ 2 - 4 * A1 * C))) / (2 * A1)
            Qv_in = CDbl(IIf(Qv_a > 0, Qv_a, Qv_b))

            '----- step 3 determine new dp---
            dp = ro * (A * Qv_in ^ 2 + B * Qv_in + C)   'Fan curve

            '----- step 4 determine Temp outlet fan ---
            Tout = Tin * (1 + dp / Pin) ^ ((γ - 1) / γ)

            '----- step 5 determine Pressure outlet fan ---
            Pout = Pin + dp

            '----- step 6 determine Discharge flow fan ---
            Qv_out = Qv_in * (Pin / Pout) ^ γ

            '----- present the data ----
            TextBox7.Text = Round(K_bypass, 2).ToString   'Resistance Bypass valve 
            TextBox8.Text = Round(Qv_out, 0).ToString
            TextBox9.Text = Round(dp, 0).ToString
            TextBox10.Text = Round(Tin, 1).ToString
            TextBox11.Text = Range(0).ToString
            TextBox12.Text = Range(1).ToString
            TextBox13.Text = Range(2).ToString
            TextBox32.Text = Range(3).ToString
            TextBox14.Text = Round(K_sys, 2).ToString   'Resistance Total system
            TextBox15.Text = Round(Qv_in, 0).ToString
            TextBox21.Text = Round(Pin, 0).ToString     'Pressure inlet
            'TextBox21.Text = Round(valve_open, 0).ToString  'Bypass valve
            TextBox25.Text = Round(Tout, 1).ToString
        End If

        '-------- Surge Alarm-----------
        TextBox15.BackColor = CType(IIf(NumericUpDown1.Value < Qv_in, Color.White, Color.Red), Color)

        '---------- calc output currents
        Cout(1) = Convert_Units_to_mAmp("Flow", Qv_in)
        Cout(2) = Convert_Units_to_mAmp("Pressure", Pin)
        Cout(3) = Convert_Units_to_mAmp("Pressure", dp)
        Cout(4) = Convert_Units_to_mAmp("Temperature", Tin)

        TextBox1.Text = Round(Cout(1), 1).ToString  'Flow inlet/out Actual [Am3/hr]
        TextBox2.Text = Round(Cout(2), 1).ToString  'Pressure in [Pa]
        TextBox3.Text = Round(Cout(3), 1).ToString  'Delta P [Pa]
        TextBox23.Text = Round(Cout(4), 1).ToString 'Temp fan in [c]
    End Sub

    '------- Convert fysical units ----> mAmp's  
    Private Function Convert_Units_to_mAmp(outType As String, value As Double) As Double
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
            Case "Valve-positioner"
                value_4ma = NumericUpDown34.Value
                Double.TryParse(TextBox32.Text, range)
                ' MessageBox.Show(range.ToString)
                results = (value - value_4ma) / range * 16.0 + 4.0
            Case Else
                MessageBox.Show("Oops error in Convert_Units_to_mAmp function")
        End Select
        Return (results)
    End Function
    '------- Convert from mAmp's ----> fysical units
    Private Function Convert_mAmp_to_Units(outType As String, value As Double) As Double
        Dim results, range, value_4ma As Double
        Select Case outType
            Case "Flow"
                value_4ma = NumericUpDown27.Value
                Double.TryParse(TextBox11.Text, range)
                results = (value - 4) / 16 * range + value_4ma
            Case "Temperature"
                value_4ma = NumericUpDown30.Value
                Double.TryParse(TextBox12.Text, range)
                results = (value - 4) / 16 * range + value_4ma
            Case "Pressure"
                value_4ma = NumericUpDown32.Value
                Double.TryParse(TextBox13.Text, range)
                results = (value - 4) / 16 * range + value_4ma
            Case "Valve-positioner"
                value_4ma = NumericUpDown34.Value
                Double.TryParse(TextBox32.Text, range)
                results = (value - 4) / 16 * range + value_4ma
            Case Else
                MessageBox.Show("Oops error in Calc_in function")
        End Select
        Return (results)
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown6.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged
        Dim ro, dp, flow, Ks As Double

        'Calculate Ksys @ work point
        flow = NumericUpDown6.Value
        dp = NumericUpDown3.Value
        ro = NumericUpDown4.Value

        Ks = flow * Sqrt(ro / dp)
        TextBox4.Text = Round(Ks, 1).ToString
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Safe_to_file()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Reset()
    End Sub
    Private Sub PID_controller()
        Dim setpoint, deviation, PID_output, dt As Double
        Dim Kp, Ki, Kd As Double    'Setting PID controller 
        Dim pv, input_ma, output_ma, range, ddev As Double

        '----- input from Flow transmitter in [Mamp]----
        Double.TryParse(TextBox1.Text, input_ma)                '[mAmp]
        TextBox27.Text = Round(Convert_mAmp_to_Units("Flow", input_ma), 0).ToString  '[m3/hr]

        '------ Setting PID controller --------
        Kp = NumericUpDown9.Value
        Ki = NumericUpDown11.Value
        Kd = NumericUpDown12.Value

        '------ shift register with process values
        Double.TryParse(TextBox27.Text, pv)
        '------ time interval-----
        dt = Timer1.Interval / 1000     '[sec]
        ' If dt <= 0 Then dt = 1          'Preventing devision by zero errors 

        '------- convert Flow to procent----
        Double.TryParse(TextBox11.Text, range)
        setpoint = NumericUpDown8.Value
        deviation = (pv - setpoint) / range * 100       '[%]
        If deviation > 10000 Then deviation = 1          'for startup
        If deviation < -10000 Then deviation = 1         'for startup


        If CheckBox1.Checked And pv > 0 Then
            ddev = deviation - last_deviation               'change in deviation
            last_deviation = deviation

            '=========== Calculate PID controller==========
            Double.TryParse(TextBox31.Text, PID_output)
            Pterm = Kp * deviation                          'P action
            Iterm = Iterm + Ki * deviation * dt             'I action
            If Iterm > 100 Then Iterm = 100                 'anti-Windup
            If Iterm < 0 Then Iterm = 0                     'anti-Windup

            ' TextBox26.Text &= ddev.ToString & vbCrLf
            Dterm = Kd * ddev / dt  'D action
            PID_output = Pterm + Iterm + Dterm
            '=============================================
            output_ma = Convert_Units_to_mAmp("Valve-positioner", PID_output)

            '---------- present results ------------
            TextBox28.Text = Round(deviation, 2).ToString("0.00")
            TextBox29.Text = Round(input_ma, 2).ToString("0.00")    'input [mAmp]
            TextBox30.Text = Round(output_ma, 2).ToString("0.00")   'output [mAmp]
            TextBox31.Text = Round(PID_output, 2).ToString("0.00")  'output [m3/hr]

            TextBox33.Text = Round(Pterm, 2).ToString("0.00")
            TextBox34.Text = Round(Iterm, 2).ToString("0.00")
            TextBox35.Text = Round(Dterm, 2).ToString("0.00")
        End If
    End Sub
    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        Convert_bypass_valve_mAmp_to_position()
    End Sub
    Private Sub Safe_to_file()
        Dim file_name As String
        Dim dirpath_Home As String = "C:\Temp\"

        file_name = dirpath_Home & "ASC-Chart_" & DateTime.Now.ToString("yyyy-MM-dd_HH_mm_ss") & ".jpeg"

        If Directory.Exists(dirpath_Home) Then
            Chart1.SaveImage(file_name, Imaging.ImageFormat.Jpeg)
        Else
            MessageBox.Show("File is NOT saved" & vbCrLf & "Directory doen not exist" & vbCrLf & "Please create " & dirpath_Home)
        End If
    End Sub
    Private Sub Convert_bypass_valve_mAmp_to_position()
        Dim bypass_valve_position, tmp As Double
        Double.TryParse(TextBox24.Text, tmp)
        bypass_valve_position = (tmp - 4) / 16 * 100    '[%]
        TextBox22.Text = Round(bypass_valve_position, 0).ToString
    End Sub
End Class
