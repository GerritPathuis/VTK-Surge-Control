Imports System.Globalization
Imports System.IO
Imports System.IO.Ports
Imports System.Math
Imports System.Threading

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

        Reset()
        Update_calc_screen()
    End Sub

    Private Sub Reset()
        Init_Chart1()
        Init_Chart2()
        Timer1.Interval = 100   'Berekeningsinterval 500 msec
        time = 0

        Timer1.Enabled = True
    End Sub
    Private Sub Init_Chart1()
        Dim i As Integer
        Try
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()
            Chart1.ChartAreas.Add("ChartArea0")

            For i = 0 To 3
                Chart1.Series.Add(i.ToString)
                Chart1.Series(i.ToString).ChartArea = "ChartArea0"
                Chart1.Series(i.ToString).ChartType = DataVisualization.Charting.SeriesChartType.Line
                Chart1.Series(i.ToString).BorderWidth = 1
            Next

            Chart1.Titles.Add("ASC testing")
            Chart1.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)

            Chart1.Series(0).Name = "Flow in"
            Chart1.Series(1).Name = "Pressure in"
            Chart1.Series(2).Name = "delta P"
            Chart1.Series(3).Name = "Temp in"
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

            For i = 0 To 3
                Chart2.Series.Add(i.ToString)
                Chart2.Series(i.ToString).ChartArea = "ChartArea1"
                Chart2.Series(i.ToString).ChartType = DataVisualization.Charting.SeriesChartType.Line
                Chart2.Series(i.ToString).BorderWidth = 1
            Next

            Chart2.Titles.Add("ASC testing")
            Chart2.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)

            Chart2.Series(0).Name = "Flow in"
            Chart2.Series(1).Name = "Pressure in"
            Chart2.Series(2).Name = "delta P"
            Chart2.Series(3).Name = "Temp in"
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
        setup_string = "LucidIoCtrl –dCOM1 –tV –c0,1,2,3 –w"
        setup_string &= NumericUpDown5.Value.ToString("0.000") & ","
        setup_string &= NumericUpDown10.Value.ToString("0.000") & ","
        setup_string &= NumericUpDown14.Value.ToString("0.000") & ","
        setup_string &= NumericUpDown15.Value.ToString("0.000") & vbCrLf

        'setup_string = "LucidIoCtrl –dCOM1 –tV –c0,1,2,3 –w5.000,2.500,1.250,0.625" 'example
        Send_data(setup_string)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim setup_string_output, setup_string_input As String

        time += Timer1.Interval / 1000                  '[sec]
        Label1.Text = time.ToString("000.0")

        '----------LucidControl Output module -------------
        'Send new setting to the  current output
        setup_string_output = "LucidIoCtrl –dCOM1 –tV –c0,1,2,3 –w"
        setup_string_output &= Cout(1).ToString("0.000") & ","
        setup_string_output &= Cout(2).ToString("0.000") & ","
        setup_string_output &= Cout(3).ToString("0.000") & ","
        setup_string_output &= Cout(4).ToString("0.000") & vbCr
        'setup_string_output = "LucidIoCtrl –dCOM1 –tV –c0,1,2,3 –w5.000,2.500,1.250,0.625" 'Example
        Send_data(setup_string_output)

        '----------LucidControl Input module -------------
        setup_string_input = "LucidIoCtrl –dCOM1 –tV –c0,1,2,3 –r" & vbCr
        Send_data(setup_string_input)
        Update_calc_screen()
        Draw_Chart1()
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

        SerialPort1.ReceivedBytesThreshold = 24    'wait EOF char or until there are x bytes in the buffer, include \n and \r !!!!
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
        TextBox9.Text = NumericUpDown26.Value.ToString 'dp fan
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
        Try
            ReceivedText(SerialPort1.ReadLine())    'Automatically called every time a data is received at the serialPortb
        Catch exc As IOException
            MsgBox("Error 453 IO exception" & exc.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown25.ValueChanged, NumericUpDown33.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown22.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown18.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, RadioButton3.CheckedChanged, RadioButton2.CheckedChanged, RadioButton1.CheckedChanged, NumericUpDown7.ValueChanged, NumericUpDown2.ValueChanged
        Update_calc_screen()
    End Sub
    Private Sub Update_calc_screen()
        Dim Range(2) As String
        Dim K_sys, K_bypass, k_sum, K100, valve_open, dp, ro As Double
        Dim A, B, C, Qv_in, Qv_out As Double
        Dim Pin, Pout As Double
        Dim Tin, Tout As Double
        Dim γ As Double
        Dim p_time, period As Double
        Range(0) = CType(NumericUpDown28.Value - NumericUpDown27.Value, String)    'Flow
        Range(1) = CType(NumericUpDown29.Value - NumericUpDown30.Value, String)    'Temp
        Range(2) = CType(NumericUpDown31.Value - NumericUpDown32.Value, String)    'Pressure

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
            p_time = time Mod period
            Select Case True
                Case RadioButton1.Checked           'Do nothing
                    K_sys = NumericUpDown25.Value
                Case RadioButton2.Checked           'Square wave
                    If (p_time > (period / 2)) Then
                        K_sys = NumericUpDown25.Value + NumericUpDown2.Value
                    Else
                        K_sys = NumericUpDown25.Value - NumericUpDown2.Value
                    End If
                Case RadioButton3.Checked           'Sine
                    K_sys = NumericUpDown25.Value + NumericUpDown2.Value * Sin(p_time / period * 2 * PI)
            End Select
            TextBox5.Text = Round(K_sys, 1).ToString
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

            '----- present the data ----
            TextBox7.Text = Round(K_bypass, 2).ToString   'Resistance Bypass valve 
            TextBox8.Text = Round(Qv_out, 0).ToString
            TextBox9.Text = Round(dp, 0).ToString
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

        '-------- Surge Alarm-----------
        TextBox15.BackColor = CType(IIf(NumericUpDown1.Value < Qv_in, Color.White, Color.Red), Color)

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown6.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged
        Dim ro, dp, flow, Ks As Double

        'Calculate Ksys @ work point
        flow = NumericUpDown6.Value
        dp = NumericUpDown3.Value
        ro = NumericUpDown4.Value

        Ks = flow * Sqrt(ro / dp)
        TextBox4.Text = Round(Ks, 1).ToString
    End Sub

    Private Sub ReceivedText(ByVal intext As String)
        MessageBox.Show(intext)
    End Sub
    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        Calc_bypass_valve_position()
    End Sub

    Private Sub Calc_bypass_valve_position()
        Dim bypass_valve_position, tmp As Double
        Double.TryParse(TextBox24.Text, tmp)
        bypass_valve_position = (tmp - 4) / 16 * 100    '[%]
        TextBox22.Text = Round(bypass_valve_position, 0).ToString
    End Sub
End Class
