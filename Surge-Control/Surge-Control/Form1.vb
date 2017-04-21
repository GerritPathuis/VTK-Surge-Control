Imports System.IO
Imports System.IO.Ports
Imports System.Math

Public Class Form1
    Dim time As Double
    Dim new_valve_pos, old_valve_pos As Double
    Dim quotient As Double
    Dim Actual_fan_flow As Double
    Dim Cout(3) As Double      'Current Output

    Dim myPort As Array  'COM Ports detected on the system will be stored here
    Dim comOpen As Boolean
    Private Property ConnectionOK As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Reset()
    End Sub

    Private Sub Reset()
        Init_Chart1()
        Timer1.Interval = 1000   'Berekeningsinterval 1 sec
        time = 0
        new_valve_pos = 20      'bypass % open
        old_valve_pos = 20      'bypass % open
        Actual_fan_flow = 0.1   '[m3/s]

        Calc_quotient()
        calc_new_valve_position()
        calc_flow()
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
            Chart1.Series(0).Points.AddXY(time, new_valve_pos)
        Catch ex As Exception
            MessageBox.Show("Draw chart1 failed")
        End Try
    End Sub

    Private Sub Calc_new_valve_position()
        Dim setpoint As Double
        Dim increment_open As Double
        Dim increment_closed As Double

        increment_open = NumericUpDown8.Value
        increment_closed = NumericUpDown9.Value

        NumericUpDown6.Value = NumericUpDown11.Value / NumericUpDown12.Value ^ 2
        setpoint = NumericUpDown6.Value

        'Opening or Closing the valve 
        If CheckBox1.Checked Then
            If quotient < setpoint Then new_valve_pos = old_valve_pos + increment_open
            If quotient > setpoint Then new_valve_pos = old_valve_pos - increment_closed
        Else
            If quotient < setpoint Then new_valve_pos = old_valve_pos - increment_closed
            If quotient > setpoint Then new_valve_pos = old_valve_pos + increment_open
        End If

        'Valve position between 0 and 100 %
        If new_valve_pos > 100 Then new_valve_pos = 100     'Valve completely open
        If new_valve_pos < 0 Then new_valve_pos = 0         'Valve closed completely

        old_valve_pos = new_valve_pos
        TextBox1.Text = Math.Round(new_valve_pos, 1).ToString
    End Sub

    Private Sub calc_flow()
        Dim max_flow As Double
        Dim bypass_valve_pos, bypass_flow As Double
        Dim process_valve_open, Actual_process_flow As Double

        max_flow = NumericUpDown7.Value

        'Valves positions
        process_valve_open = NumericUpDown13.Value / 100    'Process valve position
        Double.TryParse(TextBox1.Text, bypass_valve_pos)    'Bypass valve

        Actual_process_flow = max_flow * process_valve_open
        Actual_fan_flow = max_flow * process_valve_open

        If RadioButton1.Checked Then bypass_flow = Actual_fan_flow * bypass_valve_pos / 100   'Lineair
        If RadioButton2.Checked Then bypass_flow = Actual_fan_flow * bypass_valve_pos / 100   'Quick opening
        If RadioButton3.Checked Then bypass_flow = Actual_fan_flow * bypass_valve_pos / 100   'Equal perceentage

        'Total flow is process + bypass flow
        Actual_fan_flow = Actual_process_flow + bypass_flow

        TextBox4.Text = Math.Round(Actual_process_flow, 2).ToString 'Process flow
        TextBox5.Text = Math.Round(bypass_flow, 2).ToString         'Bypass flow
        TextBox6.Text = Math.Round(Actual_fan_flow, 2).ToString     'Total fan flow
    End Sub
    Private Sub Calc_quotient()
        Dim dp, setpoint1, setpoint2 As Double

        dp = NumericUpDown2.Value           'mBar
        quotient = dp / Actual_fan_flow ^ 2
        TextBox2.Text = Math.Round(quotient, 0).ToString


        setpoint1 = NumericUpDown6.Value
        setpoint2 = setpoint1 * (1 + NumericUpDown4.Value / 100)

        'Warnings

        Select Case quotient
            Case quotient > setpoint1
                Label13.Text = "SURGING"
                Label13.BackColor = Color.Red
            Case quotient > setpoint1 And quotient < setpoint2
                Label13.Text = "NEAR SURGING"
                Label13.BackColor = Color.Red
            Case quotient > setpoint2
                Label13.Text = "ALL OK"
                Label13.BackColor = Color.LightGreen
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Reset() 'Reset button
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim setup_string As String
        Cout(0) = NumericUpDown5.Value
        Cout(1) = NumericUpDown10.Value
        Cout(2) = NumericUpDown14.Value
        Cout(3) = NumericUpDown15.Value

        setup_string = "LucidIoCtrl –dCOM1 –tV –c0,1,2,3 –w5.000,2.500,1.250,0.625" & vbCrLf
        Send_data(setup_string)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        draw_Chart1()
        time += Timer1.Interval / 1000                  '[sec]
        TextBox3.Text = Math.Round(time, 1).ToString
        Calc_quotient()
        calc_new_valve_position()
        calc_flow()
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub ReceivedText(ByVal intext As String)
        MessageBox.Show(intext)
        'Try
        '    If Me.TxtMbedMessage.InvokeRequired Then
        '        Dim x As New SetTextCallback(AddressOf ReceivedText)
        '        Me.Invoke(x, New Object() {(intext)})
        '    Else
        '        Me.TxtMbedMessage.AppendText(intext)
        '    End If
        'Catch ex As Exception
        '    MsgBox("Error 771 received text exception received" & ex.Message)
        'End Try
    End Sub


End Class
