Public Class Form1
    Dim time As Double
    Dim new_valve_pos, old_valve_pos As Double
    Dim quotient As Double
    Dim Actual_fan_flow As Double

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Reset()
    End Sub

    Private Sub Reset()
        init_Chart1()
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

    Private Sub init_Chart1()
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
    Private Sub draw_Chart1()
        Try
            Chart1.Series(0).Points.AddXY(time, new_valve_pos)
        Catch ex As Exception
            MessageBox.Show("Draw chart1 failed")
        End Try
    End Sub

    Private Sub calc_new_valve_position()
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

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        draw_Chart1()
        time += Timer1.Interval / 1000                  '[sec]
        TextBox3.Text = Math.Round(time, 1).ToString
        Calc_quotient()
        calc_new_valve_position()
        calc_flow()
    End Sub
End Class
