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
    Dim counter As Integer = 0
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

        TextBox24.Text = "6.8"  'Test value [c]
        Label108.Text = ""      'Communication Error codes

        For i = 0 To 3
            pv(i) = 1       'Initial value
        Next

        Reset()
        Update_calc_screen()
    End Sub

    Private Sub Reset()
        Init_Chart1()
        Init_Chart2()
        Timer1.Interval = 2000   'Berekeningsinterval 2000 msec
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
        SetOut()
    End Sub
    Private Sub SetOut()
        Dim SetIoGroup(4) As Byte   'Set-Voltage or Current, see Page 22 en 26
        Dim SetIoG As String = String.Empty
        Dim str_hex1 As String = String.Empty
        Dim str_hex2 As String = String.Empty
        Dim str_hex3 As String = String.Empty
        Dim str_hex4 As String = String.Empty
        Dim message_length As Integer = 0
        Dim bb() As Byte
        Dim ret As String

        SetIoGroup(1) = &H42   'OPC= SetIoGroup !!!!
        SetIoGroup(2) = &HF    'Channel 1...4 (0000.1111==0x0F)
        If RadioButton8.Checked Then
            SetIoGroup(3) = &H1D   'Volt (0 to 1,000,000 MicroVolt)
        Else
            SetIoGroup(3) = &H23   'Current  (0 to 1,000,000 MicroAmp)
        End If
        SetIoGroup(4) = &H10   'Len (4 x 4=16 bytes)

        '------ make Command string of the Command byte array---
        SetIoG = System.Text.Encoding.Default.GetString(SetIoGroup)

        '---------- now convert to hex-------
        SetIoG = String_ascii_to_Hex_ascii(SetIoG) '& "-"

        If RadioButton8.Checked Then
            'Voltage output, channel #1...4 
            str_hex1 = Hex(CDec(NumericUpDown5.Value * 10 ^ 6))
            str_hex2 = Hex(CDec(NumericUpDown10.Value * 10 ^ 6))
            str_hex3 = Hex(CDec(NumericUpDown14.Value * 10 ^ 6))
            str_hex4 = Hex(CDec(NumericUpDown15.Value * 10 ^ 6))
        Else
            '----------- current output, channel #1...4 -------------
            str_hex1 = Hex(CDec(NumericUpDown5.Value - 4) / 16 * 10 ^ 6)
            str_hex2 = Hex(CDec(NumericUpDown10.Value - 4) / 16 * 10 ^ 6)
            str_hex3 = Hex(CDec(NumericUpDown14.Value - 4) / 16 * 10 ^ 6)
            str_hex4 = Hex(CDec(NumericUpDown15.Value - 4) / 16 * 10 ^ 6)

        End If
        '------ convert to Big endian and ------
        '------ adding all string-sections to one string
        SetIoG &= To_big_endian(str_hex1)
        SetIoG &= To_big_endian(str_hex2)
        SetIoG &= To_big_endian(str_hex3)
        SetIoG &= To_big_endian(str_hex4)

        '------ convert to bytes and write to port-----
        TextBox26.Text &= "SetIoG= " & SetIoG & vbCrLf
        bb = HexStringToByteArray(SetIoG)

        If SerialPort2.IsOpen Then
            SerialPort2.Write(bb, 1, 20)
            ret = String.Join(",", Array.ConvertAll(bb, Function(byteValue) byteValue.ToString))
            TextBox26.Text &= "=" & ret & "=" & vbCrLf
        Else
            TextBox26.Text &= "SerialPort2 is closed" & vbCrLf
        End If
    End Sub


    Private Function To_big_endian(str_num As String) As String
        Dim return_val As String = String.Empty
        Dim bytes() As Byte = {&H0, &H0, &H0, &H0}
        Dim bytes_big() As Byte = {&H0, &H0, &H0, &H0}
        Dim b As Byte
        Dim byte_0x00() As Byte = {&H0, &H0, &H0, &H0}
        Dim byte_sum() As Byte = {&H0, &H0, &H0, &H0}
        Dim value, no_bytes As Integer

        value = Convert.ToInt32(str_num, 16)
        bytes = BitConverter.GetBytes(value)

        '------- determine number significant bytes
        If value > 2 ^ 32 Then MessageBox.Show("Problem in To_big_endian")
        Select Case value
            Case Is >= CInt(2 ^ 24)
                no_bytes = 4
            Case Is < CInt(2 ^ 24)
                no_bytes = 3
            Case Is < CInt(2 ^ 16)
                no_bytes = 2
            Case Is < CInt(2 ^ 8)
                no_bytes = 1
        End Select

        'MessageBox.Show("str_num=" & str_num & " No_bytes= " & no_bytes.ToString)
        'MessageBox.Show(" bytes(0)= " & Conversion.Hex(bytes(0)))
        'MessageBox.Show(" bytes(1)= " & Conversion.Hex(bytes(1)))
        'MessageBox.Show(" bytes(2)= " & Conversion.Hex(bytes(2)))
        'MessageBox.Show(" bytes(3)= " & Conversion.Hex(bytes(3)))

        Select Case no_bytes
            Case 1  '[1 bytes] move right and add zero
                bytes_big(0) = &H0
                bytes_big(1) = &H0
                bytes_big(2) = &H0
                bytes_big(3) = bytes(0)
            Case 2   '[2 bytes] move right and add zero
                bytes_big(0) = &H0
                bytes_big(1) = &H0
                bytes_big(2) = bytes(1)
                bytes_big(3) = bytes(0)
            Case 3          '[3 bytes] move right and add zero (This one works!!)
                bytes_big(0) = &H0
                bytes_big(1) = bytes(2)
                bytes_big(2) = bytes(1)
                bytes_big(3) = bytes(0)
            Case 4
                'Do nothing
        End Select
        Array.Reverse(bytes_big)    'Now reverse order

        '--------- make the string-------
        For Each b In bytes_big
            return_val += b.ToString("X2")
        Next

        ' MessageBox.Show("Little Endian=" & str_num.ToString & " To_big_endian= " & return_val.ToString)
        Return return_val
    End Function
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim t1, t2, t3, t4 As Double

        'Send result calculations to the outputs 
        If CheckBox3.Checked Then
            Double.TryParse(TextBox1.Text, t1)     'Flow
            Double.TryParse(TextBox2.Text, t2)     'Pinlet
            Double.TryParse(TextBox3.Text, t3)     'delta Pressure
            Double.TryParse(TextBox23.Text, t4)    'Tinlet

            'keeps things with the selected output range-----
            If t1 > NumericUpDown5.Maximum Then t1 = NumericUpDown5.Maximum
            If t2 > NumericUpDown10.Maximum Then t2 = NumericUpDown10.Maximum
            If t3 > NumericUpDown14.Maximum Then t3 = NumericUpDown14.Maximum
            If t4 > NumericUpDown15.Maximum Then t4 = NumericUpDown15.Maximum
            If t1 < NumericUpDown5.Minimum Then t1 = NumericUpDown5.Minimum
            If t2 < NumericUpDown10.Minimum Then t2 = NumericUpDown10.Minimum
            If t3 < NumericUpDown14.Minimum Then t3 = NumericUpDown14.Minimum
            If t4 < NumericUpDown15.Minimum Then t4 = NumericUpDown15.Minimum

            NumericUpDown5.Value = CDec(t1)     'Flow
            NumericUpDown10.Value = CDec(t2)    'Pinlet
            NumericUpDown14.Value = CDec(t3)    'delta Pressure
            NumericUpDown15.Value = CDec(t4)    'Tinlet
        End If

        GetIO()                                 'Get the feedback value
        If CheckBox3.Checked Then SetOut()      'Set the output values
        Update_calc_screen()
        Draw_Chart1()
        PID_controller()
    End Sub
    Private Sub GetIO()
        Dim GetIo(4) As Byte   'Get-Voltage, see Page 22 en 23

        '--------- update time on sceen-------
        time += Timer1.Interval * 0.001                  '[msec]--->[sec]
        Label1.Text = time.ToString("000.0")

        '--------- prepare request to Lucid Control------
        GetIo(1) = &H48   'OPC= GetIoGroup
        GetIo(2) = &H1    'Channel 1
        If RadioButton5.Checked Then
            GetIo(3) = &H1D   'Voltage range 0-100,000,000 mV (4Bytes)
            Label102.Text = "[Volt]"
        Else
            'GetIo(3) = &H23   'Amp range 0-1,000,000 mAmp (4Bytes)
            GetIo(3) = &H1D   'Amp range 0-1,000,000 mAmp (4Bytes)
            Label102.Text = "[mAmp]"
        End If
        GetIo(4) = &H0    'LEN

        If SerialPort1.IsOpen Then
            '-------LucidControl AI4, 10Volt Input module -------------
            SerialPort1.Write(GetIo, 1, 4)
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'Find ports----------
        combo_Port1.SelectedIndex = -1  'To instruments
        combo_Port1.Items.Clear()
        combo_Port2.SelectedIndex = -1  'From Bypass valve
        combo_Port2.Items.Clear()
        Serial_setup()
    End Sub

    Private Sub Serial_setup()                  'Serial ports setup
        combo_Baud.Items.Clear()

        If (SerialPort1.IsOpen = True) Then     'Write to Instruments
            SerialPort1.DiscardInBuffer()       'Preventing exceptions
            SerialPort1.Close()
        End If

        If (SerialPort2.IsOpen = True) Then     'Read from Bypass Valve 
            SerialPort2.DiscardInBuffer()       'Preventing exceptions
            SerialPort2.Close()
        End If
        Try
            myPort = SerialPort.GetPortNames() 'Get all com ports available
            For Each port In myPort
                combo_Port1.Items.Add(port)
                combo_Port2.Items.Add(port)
            Next port
            combo_Port1.Text = CType(combo_Port1.Items.Item(0), String)    'Set cmbPort text to the first COM port detected
            combo_Port2.Text = CType(combo_Port2.Items.Item(0), String)    'Set cmbPort text to the first COM port detected
        Catch ex As Exception
            MsgBox("No COM ports detected")
        End Try

        combo_Baud.Items.Add(9600)     'Populate the cmbBaud Combo box to common baud rates used
        combo_Baud.Items.Add(19200)
        combo_Baud.Items.Add(38400)
        combo_Baud.Items.Add(57600)
        combo_Baud.SelectedIndex = 0     'Set cmbBaud text to 9600 Baud 
    End Sub

    Private Sub BtnConnect_Click(sender As System.Object, e As System.EventArgs) Handles btnConnect.Click
        'Connect
        If combo_Port1.Text.Length = 0 Then
            MsgBox("Sorry, did not find any connected Lucid Controllers")
        Else
            SerialPort1.PortName = combo_Port1.Text         'Set SerialPort1 to the selected COM port at startup
            SerialPort1.BaudRate = CInt(combo_Baud.Text)    'Set Baud rate to the selected value on
            SerialPort1.Parity = Parity.None
            SerialPort1.StopBits = StopBits.One
            SerialPort1.Handshake = Handshake.None
            SerialPort1.DataBits = 8                        'Open our serial port
            SerialPort1.ReadBufferSize = 4096
            SerialPort1.ReceivedBytesThreshold = 4
            SerialPort1.DiscardNull = False                 'important otherwise it will not work

            SerialPort2.PortName = combo_Port2.Text         'Set SerialPort2 to the selected COM port at startup
            SerialPort2.BaudRate = CInt(combo_Baud.Text)    'Set Baud rate to the selected value on
            SerialPort2.Parity = Parity.None
            SerialPort2.StopBits = StopBits.One
            SerialPort2.Handshake = Handshake.None
            SerialPort2.DataBits = 8                        'Open our serial port
            SerialPort2.ReadBufferSize = 4096
            SerialPort2.ReceivedBytesThreshold = 4
            SerialPort2.DiscardNull = False                 'important otherwise it will not work

            Try
                SerialPort1.Open()
                SerialPort2.Open()
                btnConnect.Enabled = False              'Disable Connect button
                btnConnect.BackColor = Color.Yellow
                btnConnect.Text = "OK connected"
                btnDisconnect.Enabled = True            'and Enable Disconnect button
            Catch ex As Exception
                MsgBox("Error 654 Open: " & ex.Message)
            End Try

            combo_Port1.BackColor = CType(IIf(SerialPort1.IsOpen, Color.Yellow, Color.Red), Color)
            combo_Port2.BackColor = CType(IIf(SerialPort2.IsOpen, Color.Yellow, Color.Red), Color)
        End If
    End Sub

    Private Sub BtnDisconnect_Click(sender As System.Object, e As System.EventArgs) Handles btnDisconnect.Click
        'Disconnect ports
        Try
            SerialPort1.DiscardInBuffer()
            SerialPort1.Close()             'Close our Serial Port
            SerialPort1.Dispose()
            SerialPort2.DiscardInBuffer()
            SerialPort2.Close()             'Close our Serial Port
            SerialPort2.Dispose()

            btnConnect.Enabled = True
            btnConnect.BackColor = Color.Red
            combo_Port1.BackColor = Color.White
            combo_Port2.BackColor = Color.White
            btnConnect.Text = "Connect"
            btnDisconnect.Enabled = False
        Catch ex As Exception
            MsgBox("Error 104 Open: " & ex.Message)
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Reset()
    End Sub
    Private Sub SerialPort1_DataReceived(sender As System.Object, e As SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        Dim intext_hex As String = String.Empty
        Dim intext As String = String.Empty
        Dim status_code As String = String.Empty
        Dim status_OK As String = "00"
        Dim Value_channel_0_hex As String  'Lucid-Control AI4, 10V module
        Dim Value_channel_0_dec As Double  'Lucid-Control AI4, 10V module

        Dim Volt_channel_0 As Double    'Lucid-Control AI4, 10V module
        Dim bigE As String = String.Empty
        Dim bypass_pos As Integer = 0

        intext_hex = SerialPort1.ReadExisting       'Read the data
        intext = String_ascii_to_Hex_ascii(intext_hex)              'Convert data to hex
        '--------- Status Communication-------
        status_code = intext.Substring(0, 2)

        If String.Equals(status_code, status_OK) And (intext.Length = 12) Then

            '---------- Test value -----------
            'intext = "0004" & "D0121300"                   'Test value +1.2500 Volt
            'intext = "0004" & "A0252600"                   'Test value +2.5000 Volt
            'intext = "0004" & "404b4c00"                   'Test value +5.000 Volt
            'intext = "0004" & "C0B4B3FF"                   'Test value -5.000 Volt
            'intext = "0004" & "39FFFFFF"                   'Test value 0.000199 Volt

            Value_channel_0_hex = intext.Substring(4, 8)    'Hex

            '---- The received value is little-Endian (now reverse order)-----
            bigE = Value_channel_0_hex.Substring(6, 2)
            bigE &= Value_channel_0_hex.Substring(4, 2)
            bigE &= Value_channel_0_hex.Substring(2, 2)
            bigE &= Value_channel_0_hex.Substring(0, 2)

            '---------- calc the value---------
            Value_channel_0_dec = Convert.ToInt32(bigE, 16)          '[microVolt] Channel 0
            'If Value_channel_0_dec > Integer.MaxValue Then 'negative number
            '    Label111.Text = "neg value"
            '    Value_channel_0_dec = (2 ^ 32 - Value_channel_0_dec) * -1
            'End If
            Value_channel_0_dec /= 10 ^ 6     '[microV-->Volt] 

            '--------- Present data--------------
            Try
                'Invoke(Sub() TextBox38.Text = intext.Substring(4, 8))           'Hex 4 Bytes value
                Invoke(Sub() TextBox38.Text = bigE)     'Hex 4 Bytes value
                Invoke(Sub() TextBox39.Text = Value_channel_0_dec.ToString)     'Decimal
                Invoke(Sub() TextBox37.Text = Round(Value_channel_0_dec, 2).ToString("0.00"))    'Volt
                Invoke(Sub() TextBox26.Text &= intext & " ")

                '----------- bypass valve position-----------
                If RadioButton5.Checked Then
                    bypass_pos = CInt(Value_channel_0_dec / 10 * 100)  'Volt input
                Else
                    bypass_pos = CInt((Value_channel_0_dec - 4) / 16 * 100)  'Amp input
                End If
                If bypass_pos > 100 Then bypass_pos = 100   'max 100% open
                If bypass_pos < 0 Then bypass_pos = 0       'min 0% open
                Invoke(Sub() TextBox36.Text = bypass_pos.ToString)       'bypass % open
            Catch ex As Exception
            End Try
            '------- Feedback ON/OFF----------
            If CheckBox2.Checked Then 'Feedback from SCS controller ON/OFF
                Invoke(Sub() NumericUpDown33.Value = CDec(Round(Volt_channel_0 * 10, 0)))   'bypass % open
            End If
        Else
            counter += 1
            Invoke(Sub() Label108.Text = counter.ToString & " statuscode=" & String_ascii_to_Hex_ascii(status_code))
            'MessageBox.Show("Lucid Communication problem Status Code= " & status_code)
            SerialPort1.DiscardInBuffer()        'empty inbuffer
        End If
    End Sub

    Private Sub SerialPort2_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles SerialPort2.DataReceived
        '-------- Keep the buffer empty----------
        Dim intext_hex2 As String = String.Empty

        intext_hex2 = SerialPort2.ReadExisting       'Read the data
    End Sub

    Public Function String_ascii_to_Hex_ascii(Data As String) As String
        Dim sVal As String = String.Empty
        Dim sHex As String = String.Empty
        'see http://stackoverflow.com/questions/14017007/how-to-convert-a-hexadecimal-value-to-ascii
        'Example string   Ascii HEX= 0x48 0x45 0x58 = 72 69 88
        'Used for information received from Lucid-Control modules
        'Convert ascii-String to Hex-string
        While Data.Length > 0
            sVal = Hex(Strings.Asc(Data.Substring(0, 1).ToString()))
            Data = Data.Substring(1)
            If sVal.Length < 2 Then
                sHex = sHex & "0" & sVal
            Else
                sHex = sHex & sVal
            End If
            'sHex = sHex & " "  'for testing
        End While
        Return sHex
    End Function
    Public Function String_Hex_to_ascii(Data As String) As String
        Dim com As String = String.Empty
        'see http://stackoverflow.com/questions/14017007/how-to-convert-a-hexadecimal-value-to-ascii
        'Data = "484558"    'Example string   Ascii HEX= 0x48 0x45 0x58 = 72 69 88

        For x = 0 To Data.Length - 1 Step 2
            com &= ChrW(CInt("&H" & Data.Substring(x, 2)))
        Next
        Return com
    End Function
    Public Function HexStringToByteArray(hexString As String) As Byte()
        Dim com As String = String.Empty
        'see http://www.vbforums.com/showthread.php?643593-Hex-String-to-Byte-Array
        'hexString=  "01050001FFFF8FFB"  'Example string
        Dim length As Integer = hexString.Length
        Dim upperBound As Integer = length \ 2
        Dim bytes(upperBound) As Byte

        If length Mod 2 = 0 Then
            upperBound -= 1
        Else
            hexString = "0" & hexString
        End If

        For i As Integer = 0 To upperBound
            bytes(i) = Convert.ToByte(hexString.Substring(i * 2, 2), 16)
        Next
        Return bytes
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

        'Range is required for converting the signal to and from 4-20 mAmp
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
                Case RadioButton1.Checked           'Feedback from ASC controller
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

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim GetId(4) As Byte       'Getid 

        GetId(1) = &HC0        'OPC= GetId Page 30
        GetId(2) = &H0
        GetId(3) = &H0
        GetId(4) = &H0

        If SerialPort1.IsOpen Then
            SerialPort1.Write(GetId, 1, 4)
        Else
            MessageBox.Show("Cannot get ID Port is closed")
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox26.Clear()
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        GroupBox5.Text = "Outputs test values 0-5 Volt"
        '---- max and min
        NumericUpDown5.Minimum = 0
        NumericUpDown5.Maximum = 10
        NumericUpDown10.Minimum = 0
        NumericUpDown10.Maximum = 10
        NumericUpDown14.Minimum = 0
        NumericUpDown14.Maximum = 10
        NumericUpDown15.Minimum = 0
        NumericUpDown15.Maximum = 10
        '---- value
        NumericUpDown5.Value = 0
        NumericUpDown10.Value = 0
        NumericUpDown14.Value = 0
        NumericUpDown15.Value = 0

        Label19.Text = "[V]"
        Label23.Text = "[V]"
        Label109.Text = "[V]"
        Label110.Text = "[V]"
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        GroupBox5.Text = "Outputs test values 4-20 mAmp"
        '---- value ----
        NumericUpDown5.Value = 4
        NumericUpDown10.Value = 4
        NumericUpDown14.Value = 4
        NumericUpDown15.Value = 4
        '---- max and min
        NumericUpDown5.Minimum = 4
        NumericUpDown5.Maximum = 20
        NumericUpDown10.Minimum = 4
        NumericUpDown10.Maximum = 20
        NumericUpDown14.Minimum = 4
        NumericUpDown14.Maximum = 20
        NumericUpDown15.Minimum = 4
        NumericUpDown15.Maximum = 20

        Label19.Text = "[mA]"
        Label23.Text = "[mA]"
        Label109.Text = "[mA]"
        Label110.Text = "[mA]"
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim callib(4) As Byte       'callib 
        Dim result As DialogResult
        Dim ret As String

        callib(1) = &H52     'OPC= Calibrate 30
        callib(2) = &H1      'Channel 1
        callib(3) = &H0
        callib(4) = &H0
        Timer1.Stop()
        MessageBox.Show("For Calibration" & vbCrLf & "Bridge the channel inputs !!")
        result = MessageBox.Show("Are you sure", " ", MessageBoxButtons.OKCancel)
        If result = DialogResult.OK Then
            result = MessageBox.Show("Realy sure ???", " ", MessageBoxButtons.OKCancel)
            If result = DialogResult.OK Then
                Select Case NumericUpDown35.Value
                    Case 1
                        callib(2) = &H1
                    Case 2
                        callib(2) = &H2
                    Case 3
                        callib(2) = &H4
                    Case 4
                        callib(2) = &H8
                End Select

                If SerialPort1.IsOpen Then
                    SerialPort1.Write(callib, 1, 4)
                    ret = String.Join(",", Array.ConvertAll(callib, Function(byteValue) byteValue.ToString("X2")))
                    TextBox26.Text &= "Calibrate " & ret & vbCrLf
                Else
                    MessageBox.Show("Cannot get ID Port is closed")
                End If
            End If
        End If
        Timer1.Start()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Timer1.Stop()       'Freeze
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
        dt = Timer1.Interval * 0.001     '[sec]

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
