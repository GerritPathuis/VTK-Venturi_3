﻿Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word

Public Class Form1
    Dim _flow_kghr As Double                'Flow [kg/hr]
    Dim _flow_kgs As Double                 'Flow [kg/s]
    Dim _flow_m3sec As Double               'Flow [m3/s]
    Dim _dia_in As Double                   'Dia inlet
    Dim _dia_keel As Double                 'Dia keel
    Dim _β As Double                        'Diameter ratio 
    Dim _dyn_visco, _ρ As Double            'Medium info
    Dim _C_classic As Double                'Discharge coefficient
    Dim _Reynolds As Double                 'Reynolds
    Dim _area_in, _v_inlet As Double        'Venturi data
    Dim _p1_tap, _p2_tap, _dp_tap As Double 'Pressures [Pa]
    Dim _κa, _τ As Double                   'Kappa, tou
    Dim ξ_pr_loss As Double                 'Unrecovered pressure loss
    Dim zeta As Double                      'Resistance coeffi 
    Dim exp_factor, exp_factor1, exp_factor2, exp_factor3 As Double
    Dim A2a, A2b, a2c As Double

    '----------- directory's-----------
    Dim dirpath_Eng As String = "N:\Engineering\VBasic\Venturi_input\"
    Dim dirpath_Rap As String = "N:\Engineering\VBasic\Venturi_rapport_copy\"
    Dim dirpath_Home As String = "C:\Temp\"

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, TabPage7.Enter
        Calc_rectangle_venturi()
    End Sub
    'Shell flow metering Handbook chapter 6.2.11 page 99, 238 and 239

    Private Sub Calc_rectangle_venturi()
        Dim Inlet_W, Inlet_H, small_w, small_h As Double    'Dimensions
        Dim Inlet_sq As Double      'W/H ratio
        Dim Throught_sq As Double   'W/H ratio
        Dim m As Double
        Dim β As Double
        Dim DeIn As Double          'Diameter inlet
        Dim DeT As Double           'Diameter throught
        Dim α_venturi As Double     'Expansion coefficient
        Dim T_line As Double        'Line temperature (actual)
        Dim T_0 As Double           'Reference temperature
        Dim qm As Double            'Mass flow [kg/s]
        Dim X As Double             'Normal flow scale 0-10
        Dim ip As Double             'Inlet pressure
        Dim dp As Double            'Instrument dp
        Dim κ As Double = 1.4       'κ= Cp/Cv
        Dim Ka As Double            'ka factor
        Dim Cv As Double = 0.975     'factor rectangle venturi
        Dim ε As Double             'fluids expansivity
        Dim vis As Double           'Viscosity
        Dim Fs As Double
        Dim Reynolds As Double
        Dim W As Double = 0         ' Steam_water_ontent[%]

        Dim _area_inlet As Double    'Inlet
        Dim area_throut As Double   'Throut
        Dim v_inlet As Double       'Inlet speed

        vis = NumericUpDown6.Value / 1000       'Viscositie [m.Pa.s]
        Inlet_W = NumericUpDown12.Value         'Inlet width [mm]
        Inlet_H = NumericUpDown13.Value         'Inlet height [mm]
        Inlet_sq = Inlet_W / Inlet_H            'w/h ratio
        _area_inlet = Inlet_W * Inlet_H
        small_w = NumericUpDown14.Value         'Keel width [mm]
        small_h = NumericUpDown15.Value         'Keel height [mm]
        Throught_sq = small_w / small_h         'w/h ratio
        area_throut = small_w * small_h


        T_0 = NumericUpDown16.Value         'Reference temp
        T_line = NumericUpDown17.Value      'Line temp
        ip = _p1_tap / 10 ^ 5               'Operating pressure [Pa]->[bar]
        dp = _dp_tap / 10 ^ 5               'dp on instrument [Pa]->[bar]
        qm = NumericUpDown1.Value           'Flow [kg/hr]
        X = NumericUpDown19.Value           'Normal flow scale 0-10

        'Thermal expansion coefficient steel 
        α_venturi = 1.3 * 10 ^ -5           '[/K]

        m = (small_w * small_h) / (Inlet_W * Inlet_H)

        '============= Calc inlet diameter ============
        DeIn = 1.1284 * Sqrt(Inlet_W * Inlet_H) * (1 + α_venturi * (T_line - T_0))

        '============= Calc throat diameter ============
        DeT = 1.1284 * Sqrt(small_w * small_h) * (1 + α_venturi * (T_line - T_0))


        '=============   venturi dimensional β ratio  ============
        β = Sqrt(small_w * small_h / (Inlet_H * Inlet_W))

        '============ Fs Water content in steam in [% weight] =========
        Fs = 1 + 0.0074 * W

        '============= Calc fluids expansivity ============
        Ka = (((ip - dp) * (X / 10) ^ 2) / ip)

        ε = κ * Ka ^ (2 / κ) / (κ - 1)
        ε *= (1 - β ^ 4) / (1 - β ^ 4 * Ka ^ (2 / κ))
        ε *= (1 - Ka ^ ((κ - 1) / κ)) / (1 - Ka)
        ε = Sqrt(ε)

        '============ Mass flow [kg/s] ==============
        qm = 3.512407 * 10 ^ -5 * Cv * (1 / Sqrt(1 - β ^ 4))
        qm *= ε * X * DeT ^ 2 * Fs * Sqrt(dp * _ρ)

        '=========== _Reynolds ============
        Reynolds = 1.2732 * 10 ^ 6 * qm / (vis * DeIn)

        '========= inlet speed =============
        v_inlet = qm * _ρ / (Inlet_W * Inlet_H / 10 ^ 6)  '[m/s]


        '============ Check inlet dimensions ratio =============
        If Inlet_sq < 0.67 Or Inlet_sq > 1.5 Then
            TextBox44.BackColor = Color.Red
        Else
            TextBox44.BackColor = Color.LightGreen
        End If

        '============ Check throught dimensions ratio =============
        If Throught_sq < 0.67 Or Throught_sq > 1.5 Then
            TextBox45.BackColor = Color.Red
        Else
            TextBox45.BackColor = Color.LightGreen
        End If

        '============ Check Reynold =============
        If _Reynolds < 2.0 * 10 ^ 5 Or _Reynolds > 2.0 * 10 ^ 7 Then
            TextBox41.BackColor = Color.Red
        Else
            TextBox41.BackColor = Color.LightGreen
        End If

        TextBox31.BackColor = CType(IIf(DeIn <= 1200, Color.LightGreen, Color.Red), Color)

        TextBox14.Text = _area_inlet.ToString("0")
        TextBox28.Text = β.ToString("0.0000")
        TextBox30.Text = Ka.ToString("0.0000")
        TextBox40.Text = ε.ToString("0.0000")
        TextBox31.Text = DeIn.ToString("0")             'Inlet size temp comp.
        TextBox47.Text = DeT.ToString("0")              'Throut size temp comp.
        TextBox32.Text = _ρ.ToString("0.0000")          'Density

        TextBox33.Text = qm.ToString("0.00")            '[kg/s]
        TextBox42.Text = (qm * 3600).ToString("0")      '[kg/hr]
        TextBox50.Text = (qm * _ρ).ToString("0.00")     '[m3/s]
        TextBox49.Text = (qm * 3600 * _ρ).ToString("0") '[m3/hr]

        TextBox34.Text = dp.ToString("0.000")           '[bar]
        TextBox35.Text = ip.ToString("0.000")           '[bar]
        TextBox36.Text = vis.ToString("0.0000")         'Voscosity
        TextBox37.Text = area_throut.ToString("0")
        TextBox38.Text = α_venturi.ToString
        TextBox39.Text = κ.ToString
        TextBox41.Text = _Reynolds.ToString("0")

        TextBox43.Text = v_inlet.ToString("0.0")    '[m3/s]
        TextBox44.Text = Inlet_sq.ToString("0.0")   '[-]
        TextBox45.Text = Throught_sq.ToString("0.0") '[-]
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, NumericUpDown10.ValueChanged, TabPage5.Enter
        TextBox27.Text = Calc_dyn_vis.ToString("0.0")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox21.Text =
       "ISO 5167-1:2003" & vbCrLf &
       "ISO 5167-4:2003" & vbCrLf &
       "Classieke Venturi diameter 200-1200mm"

        '------------- Initial values----------------------
        _flow_kghr = 30000              '[kg/m3]
        _flow_kgs = _flow_kghr / 3600   '[kg/sec]
        _ρ = 1.2                        '[kg/m3]
        _κa = 1.4                       'Isentropic exponent
        _p1_tap = 101325                '[pa]

        _dia_in = 0.8                   '[m] classis venturi inlet diameter = outlet diameter
        _β = 0.5                        '[-]
        _C_classic = 0.985              'See ISO5167-4 section 5.5.4

        '========== dp range instrument [mbar] =============
        ComboBox1.Items.Add("    12.5")
        ComboBox1.Items.Add("    20")
        ComboBox1.Items.Add("    25")
        ComboBox1.Items.Add("    50")
        ComboBox1.Items.Add("   125")
        ComboBox1.Items.Add("   250")
        ComboBox1.SelectedIndex = 1

        Double.TryParse(CType(ComboBox1.SelectedItem, String), _dp_tap)
        _dp_tap *= 100                          '[mbar] --> [Pa]

        '--------- calc ---------------
        _flow_m3sec = _flow_kghr / (3600 * _ρ) '[m3/s]
        _area_in = Math.PI / 4 * _dia_in ^ 2   '[m2]
        _v_inlet = _flow_m3sec / _area_in      '[m/s] keel
        _p2_tap = _p1_tap - _dp_tap            '[Pa]
        _τ = _p2_tap / _p1_tap                 'Pressure ratio
        _dia_keel = _β * _dia_in               '[mm]

        '----------- terug zetten op het scherm-------------
        Present_results()
        Calc_venturi1()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown2.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown4.ValueChanged, ComboBox1.SelectedIndexChanged
        Calc_venturi1()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Save_tofile()
    End Sub

    Private Sub Calc_venturi1()
        Dim Ecc1, Ecc2, Ecc3 As Double
        Dim Dev1, Dev2, Dev3 As Double

        Get_data_from_screen()

        Ecc1 = 0        'Start lower limit of eccentricity [-]
        Ecc2 = 1.0      'Start upper limit of eccentricity [-]
        Ecc3 = 0.5      'In the middle of eccentricity [-]

        Dev1 = CDbl(Calc_A2(Ecc1))
        Dev2 = CDbl(Calc_A2(Ecc2))
        Dev3 = CDbl(Calc_A2(Ecc3))

        '-------------Iteratie 30x halveren moet voldoende zijn ---------------
        '---------- Exc= excentricity, looking for Deviation is zero ---------

        For jjr = 0 To 30
            If Dev1 * Dev3 < 0 Then
                Ecc2 = Ecc3
            Else
                Ecc1 = Ecc3
            End If
            Ecc3 = (Ecc1 + Ecc2) / 2
            Dev1 = CDbl(Calc_A2(Ecc1))
            Dev2 = CDbl(Calc_A2(Ecc2))
            Dev3 = CDbl(Calc_A2(Ecc3))
        Next jjr
        _β = Ecc3
        _dia_keel = _β * _dia_in

        '-------- Controle nulpunt zoek functie ----------------
        If Dev3 > 0.00001 Then
            GroupBox4.BackColor = Color.Red
        Else
            GroupBox4.BackColor = Color.Transparent
        End If

        '-------- Unrecovered pressure loss over the complete venturi assembly----
        'ξ_pr_loss = 0.15 * _dp_tap
        ξ_pr_loss = (-0.017 * _β + 0.191) * _dp_tap   'For 7 degree divergent section

        '--------- resistance coefficient venturi assembly 
        zeta = 2 * ξ_pr_loss / (_ρ * _v_inlet ^ 2)

        Draw_chart1()
        Present_results()
    End Sub

    Private Function Calc_A2(_βa As Double) As Double

        '----- calc -------------
        _p2_tap = _p1_tap - _dp_tap                '[Pa]
        _τ = _p2_tap / _p1_tap                     'Pressure ratio

        '---------- expansie factor ISI 5167-4 Equation 2---------
        exp_factor1 = _κa * _τ ^ (2 / _κa)
        exp_factor1 /= _κa - 1

        exp_factor2 = 1 - _βa ^ 4
        exp_factor2 /= 1 - _βa ^ 4 * _τ ^ (2 / _κa)

        exp_factor3 = 1 - _τ ^ ((_κa - 1) / _κa)
        exp_factor3 /= 1 - _τ

        exp_factor = Math.Sqrt(exp_factor1 * exp_factor2 * exp_factor3)

        '------------- iteratie-------------------
        _flow_kghr = NumericUpDown1.Value            '[kg/h] inlet
        _flow_kgs = _flow_kghr / 3600                '[kg/sec] inlet
        _flow_m3sec = _flow_kghr / (3600 * _ρ)       '[m3/s] inlet

        _area_in = Math.PI / 4 * _dia_in ^ 2         '[m2]
        _v_inlet = _flow_m3sec / _area_in            '[m/s] inlet

        _Reynolds = _v_inlet * _dia_in * _ρ / _dyn_visco

        '------- ISO5167-1:2003, SECTION 5.2 page 8-------------
        A2b = _C_classic * exp_factor * _βa ^ 2 / Math.Sqrt(1 - _βa ^ 4)
        A2a = 4 * _flow_kgs / (PI * _dia_in ^ 2 * Math.Sqrt(2 * _dp_tap * _ρ))

        a2c = A2a - A2b
        Return (a2c)
    End Function
    Private Sub Present_results()
        Try
            NumericUpDown1.Value = CDec(_flow_kghr)             '[kg/m3]
            NumericUpDown7.Value = CDec(_κa)                    'Isentropic exponent
            NumericUpDown2.Value = CDec(_ρ)                     '[kg/m3]
            NumericUpDown6.Value = CDec(_dyn_visco * 10 ^ 6)    '_dyn_visco
            NumericUpDown4.Value = CDec(_dia_in * 1000)         '[m] classis venturi inlet diameter = outlet diameter
            TextBox46.Text = _β.ToString("0.00")                '[-]

            TextBox1.Text = Math.Round(_dia_keel * 1000, 0).ToString     '[mm] keel diameter
            TextBox2.Text = _C_classic.ToString
            TextBox3.Text = Math.Round(_Reynolds, 0).ToString           '[-]
            TextBox4.Text = Math.Round(_v_inlet, 1).ToString            '[m/s]
            TextBox5.Text = Math.Round(exp_factor, 3).ToString          '[-]
            TextBox13.Text = Math.Round(_p2_tap / 100, 1).ToString      '[Pa]->[mBar]
            TextBox12.Text = Math.Round(_τ, 4).ToString
            TextBox15.Text = (_dia_in * 1000).ToString("0")             'Diameter inlet [m]->[mm]
            TextBox16.Text = _flow_m3sec.ToString("0.000")              'Flow [m3/s]
            TextBox51.Text = (_flow_m3sec * 3600).ToString("0")         'Flow [m3/h]
            TextBox17.Text = (_dia_keel * 1000).ToString("0")           'Diameter keel
            TextBox23.Text = (ξ_pr_loss / 100).ToString("0.00")         'Unrecovered pressure loss [Pa]->[mBar]
            TextBox26.Text = zeta.ToString("0.00")                      'Resistance coeffi venturi assembly
            TextBox48.Text = _flow_kgs.ToString("0.00")

            '------- _β check --------------
            If _β < 0.4 Or _β > 0.7 Then
                TextBox46.BackColor = Color.Red
            Else
                TextBox46.BackColor = Color.LightGreen
            End If

            '------- _τ check --------------
            If _τ < 0.75 Then
                TextBox12.BackColor = Color.Red
            Else
                TextBox12.BackColor = Color.LightGreen
            End If

            '------- _Reynolds check-----------
            If _Reynolds < 2.0 * 10 ^ 5 Or _Reynolds > 2.0 * 10 ^ 6 Then
                TextBox3.BackColor = Color.Red
                If _Reynolds < 2.0 * 10 ^ 5 Then Label10.Text = "_Reynolds, Te lage snelheid"
                If _Reynolds > 2.0 * 10 ^ 6 Then Label10.Text = "_Reynolds, Te Hoge snelheid"
            Else
                TextBox3.BackColor = Color.LightGreen
                Label10.Text = "_Reynolds OK"
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 845")  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Read_file()
    End Sub

    Private Sub Draw_chart1()
        Dim x, y As Double
        Try
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()
            Chart1.Series.Add("Series0")
            Chart1.ChartAreas.Add("ChartArea0")
            Chart1.Series(0).ChartArea = "ChartArea0"
            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart1.Titles.Add("Determine _β" & vbCrLf & "ISO 5167-1:2003, Section 5.2" & vbCrLf & "(A2 must be zero)")
            Chart1.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)
            Chart1.Series(0).Name = "Koppel[%]"
            Chart1.Series(0).Color = Color.Blue
            Chart1.Series(0).IsVisibleInLegend = False
            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart1.ChartAreas("ChartArea0").AxisX.Maximum = 1
            Chart1.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            Chart1.ChartAreas("ChartArea0").AxisY.Title = "Invariant A2"
            Chart1.ChartAreas("ChartArea0").AxisX.Title = "_β [-]"

            '------ data for the Chart -----------------------------
            For x = 0 To 1.0 Step 0.01
                y = CDbl(Calc_A2(x))
                Chart1.Series(0).Points.AddXY(x, y)
            Next x

            '------ data for the actual _β value -----------------
            Calc_A2(_β)
        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 206")  ' Show the exception's message.
        End Try
    End Sub
    '-------------------- Dimension of the Venturi ----------------
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabControl1.Enter, NumericUpDown9.ValueChanged, NumericUpDown3.ValueChanged, TabPage3.Enter
        Dim Length(10) As Double
        Dim deltad As Double
        Dim φ_divert As Double

        φ_divert = NumericUpDown9.Value                          'Diversion angle
        deltad = (_dia_in - _dia_keel) / 2
        TextBox15.Text = Round(_dia_in * 1000, 0).ToString       'Diameter in
        TextBox17.Text = Round(_dia_keel * 1000, 0).ToString     'Diameter keel

        Length(0) = 2 * _dia_in                                  'Bocht R=D bij fan intake
        Length(1) = 3 * _dia_in                                  'Recht in 
        Length(2) = deltad / Math.Tan(NumericUpDown3.Value * Math.PI / 180)       'Convergeren
        Length(3) = _dia_keel                                    'Meten
        Length(4) = deltad / Math.Tan(φ_divert * Math.PI / 180)       'Divergeren
        Length(5) = 4 * _dia_keel                                'Recht achter Venturi
        Length(6) = _dia_in / 4                                  'Lucht inlaat
        Length(7) = _dia_in                                      'Chinese hat
        Length(8) = Length(0) + Length(1) + Length(2) + Length(3) + Length(4) + Length(5) + Length(6) + Length(7)

        TextBox20.Text = Round(Length(0) * 1000, 0).ToString    'Bend
        TextBox6.Text = Round(Length(1) * 1000, 0).ToString
        TextBox7.Text = Round(Length(2) * 1000, 0).ToString
        TextBox8.Text = Round(Length(3) * 1000, 0).ToString
        TextBox9.Text = Round(Length(4) * 1000, 0).ToString
        TextBox10.Text = Round(Length(5) * 1000, 0).ToString    'Down stream 4xd
        TextBox18.Text = Round(Length(6) * 1000, 0).ToString    'Air intake
        TextBox19.Text = Round(Length(7) * 1000, 0).ToString    'Chinese hat
        TextBox11.Text = Round(Length(8) * 1000, 0).ToString    'Total

        TextBox22.Text = Round(_dia_keel * 1000, 0).ToString     'Length C
    End Sub

    Private Sub Get_data_from_screen()
        Try
            _flow_kghr = NumericUpDown1.Value               '[kg/hr]
            _κa = NumericUpDown7.Value                      'Isentropic exponent
            _ρ = NumericUpDown2.Value                       '[kg/m3]
            _dyn_visco = NumericUpDown6.Value / 10 ^ 6      'kin_visco
            _p1_tap = NumericUpDown11.Value * 100           '[mBar]->[pa]
            Double.TryParse(CType(ComboBox1.SelectedItem, String), _dp_tap)
            _dp_tap *= 100                                  '[mbar] --> [Pa]
            _dia_in = NumericUpDown4.Value / 1000           '[m] classis venturi inlet diameter = outlet diameter

        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 254")  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim oWord As Word.Application ' = Nothing

        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2, oPara4 As Word.Paragraph
        Dim row, font_sizze As Integer
        Dim ufilename As String

        Try
            oWord = New Word.Application()

            'Start Word and open the document template. 
            font_sizze = 9
            oWord = CType(CreateObject("Word.Application"), Word.Application)
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            'Insert a paragraph at the beginning of the document. 
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "VTK Engineering"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = font_sizze + 3
            oPara1.Range.Font.Bold = CInt(True)
            oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara2.Range.Font.Size = font_sizze + 1
            oPara2.Format.SpaceAfter = 1
            oPara2.Range.Font.Bold = CInt(False)
            oPara2.Range.Text = "Classical Venturi tube acc ISO5167-1-4:2003" & vbCrLf
            oPara2.Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Project Name"
            oTable.Cell(row, 2).Range.Text = TextBox24.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Project number "
            oTable.Cell(row, 2).Range.Text = TextBox25.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Author "
            oTable.Cell(row, 2).Range.Text = Environment.UserName
            row += 1
            oTable.Cell(row, 1).Range.Text = "Date "
            oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(2)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a 16 x 3 table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 23, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Input Data"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Air density _ρ"
            oTable.Cell(row, 2).Range.Text = NumericUpDown2.Value.ToString("0.00")
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Dynamic viscosity"
            oTable.Cell(row, 2).Range.Text = NumericUpDown6.Value.ToString("0.0")
            oTable.Cell(row, 3).Range.Text = "[Pa.s 10^-6]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Isentropic exponent"
            oTable.Cell(row, 2).Range.Text = NumericUpDown7.Value.ToString("0.0")
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inlet pressure"
            oTable.Cell(row, 2).Range.Text = NumericUpDown11.Value.ToString("0.0")
            oTable.Cell(row, 3).Range.Text = "[mBar abs]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "dp @ max flow"
            oTable.Cell(row, 2).Range.Text = (_dp_tap / 100).ToString("0.0")
            oTable.Cell(row, 3).Range.Text = "[mBar]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Mass flow"
            oTable.Cell(row, 2).Range.Text = _flow_kghr.ToString("0") & vbTab & _flow_kgs.ToString("0.00")
            oTable.Cell(row, 3).Range.Text = "[kg/h, kg/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Volume flow"
            oTable.Cell(row, 2).Range.Text = TextBox16.Text
            oTable.Cell(row, 3).Range.Text = "[m3/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inside Inlet diameter"
            oTable.Cell(row, 2).Range.Text = NumericUpDown4.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inside Throat diameter"
            oTable.Cell(row, 2).Range.Text = TextBox1.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "_β Diameter ration"
            oTable.Cell(row, 2).Range.Text = TextBox46.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inlet speed"
            oTable.Cell(row, 2).Range.Text = TextBox4.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Reynolds"
            oTable.Cell(row, 2).Range.Text = TextBox3.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Discharge Coefficient"
            oTable.Cell(row, 2).Range.Text = TextBox2.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "ε Expansibility factor"
            oTable.Cell(row, 2).Range.Text = TextBox5.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Uncovered pressure loss"
            oTable.Cell(row, 2).Range.Text = TextBox23.Text
            oTable.Cell(row, 3).Range.Text = "[mbar]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Upstream straight length"
            oTable.Cell(row, 2).Range.Text = TextBox10.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Converging angle, length"
            oTable.Cell(row, 2).Range.Text = NumericUpDown3.Value.ToString("0") & vbTab & TextBox7.Text
            oTable.Cell(row, 3).Range.Text = "[deg, mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Measuring section"
            oTable.Cell(row, 2).Range.Text = TextBox8.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Diverging angle, length"
            oTable.Cell(row, 2).Range.Text = NumericUpDown9.Value.ToString("0.0") & vbTab & TextBox9.Text
            oTable.Cell(row, 3).Range.Text = "[deg, mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Down stream straight length"
            oTable.Cell(row, 2).Range.Text = TextBox6.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Total Stack length"
            oTable.Cell(row, 2).Range.Text = TextBox11.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.4)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.2)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------save picture ---------------- 
            Draw_chart2()
            Chart2.SaveImage("c:\Temp\MainChart.gif", System.Drawing.Imaging.ImageFormat.Gif)
            oPara4 = oDoc.Content.Paragraphs.Add
            oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oPara4.Range.InlineShapes.AddPicture("c:\Temp\MainChart.gif")
            oPara4.Range.InlineShapes.Item(1).LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
            oPara4.Range.InlineShapes.Item(1).Width = 310
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '--------------Save file word file------------------
            'See https://msdn.microsoft.com/en-us/library/63w57f4b.aspx
            ufilename = "N:\Engineering\VBasic\Rapport_copy\Campbell_report_" & DateTime.Now.ToString("yyyy_MM_dd__HH_mm_ss") & ".docx"

            If Directory.Exists("N:\Engineering\VBasic\Rapport_copy") Then
                oWord.ActiveDocument.SaveAs(ufilename.ToString)
            End If
        Catch ex As Exception
            MessageBox.Show("Bestaat directory N:\Engineering\VBasic\Rapport_copy\ ? " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, TabPage4.Enter
        Draw_chart2()
    End Sub

    Private Sub Draw_chart2()
        Dim x, y As Double
        Try
            Chart2.Series.Clear()
            Chart2.ChartAreas.Clear()
            Chart2.Titles.Clear()
            Chart2.Series.Add("Series0")
            Chart2.ChartAreas.Add("ChartArea0")
            Chart2.Series(0).ChartArea = "ChartArea0"
            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart2.Titles.Add("Venturi flow computation acc. " & "ISO 5167-4:2003 Chapter 4")
            Chart2.Titles.Add("Discharge Coefficient= " & _C_classic.ToString & ", Dia.throat= " & Round(_dia_keel * 1000, 1).ToString & " [mm]" & ", _ρ= " & _ρ.ToString & " [kg/m3]" & ", K= " & _κa.ToString & " [-]")
            Chart2.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)
            Chart2.Series(0).Color = Color.Black
            Chart2.Series(0).IsVisibleInLegend = False
            Chart2.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart2.ChartAreas("ChartArea0").AxisX.MinorGrid.Enabled = True
            Chart2.ChartAreas("ChartArea0").AxisY.MinorGrid.Enabled = True
            Chart2.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            Chart2.ChartAreas("ChartArea0").AxisY.MinorTickMark.Enabled = True
            Chart2.ChartAreas("ChartArea0").AxisY.Title = "Flow [kg/hr]"
            Chart2.ChartAreas("ChartArea0").AxisX.Title = "dp_tap [Pa]"

            '----------------- data for the Chart -----------------
            '--------------- see ISO 5167-4 Equation 1-------------
            For x = 0 To _dp_tap Step 1
                y = _C_classic / Sqrt(1 - _β ^ 4)
                y *= exp_factor * PI / 4 * _dia_keel ^ 2 * Sqrt(2 * x * _ρ)
                y *= 3600                               '[kg/h]
                Chart2.Series(0).Points.AddXY(x, y)
            Next x
        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 476")  ' Show the exception's message.
        End Try
    End Sub

    Function Calc_dyn_vis() As Double
        Dim tk, _dyn_visco As Double

        'Dynamic Viscosity calculation (valid for -100K too 800K)
        'See http://onlinelibrary.wiley.com/doi/10.1002/9780470516430.app2/pdf

        tk = NumericUpDown10.Value + 273.15
        _dyn_visco = 1.458 * tk ^ 1.5 / (tk + 110.4)    '[Pas * 10^-6]

        Return (_dyn_visco)
    End Function
    'Save control settings and case_x_conditions to file
    Private Sub Save_tofile()
        Dim temp_string As String
        Dim filename As String = "PV_Calc_" & TextBox24.Text & "_" & TextBox25.Text & "_" & TextBox52.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".vtk"
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim i As Integer

        If String.IsNullOrEmpty(TextBox8.Text) Then
            TextBox8.Text = "-"
        End If

        temp_string = TextBox24.Text & ";" & TextBox25.Text & ";" & TextBox52.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all numeric, combobox, checkbox and radiobutton controls -----------------
        FindControlRecursive(all_num, Me, GetType(NumericUpDown))   'Find the control
        all_num = all_num.OrderBy(Function(x) x.Name).ToList()      'Alphabetical order
        For i = 0 To all_num.Count - 1
            Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
            temp_string &= grbx.Value.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all combobox controls and save
        FindControlRecursive(all_combo, Me, GetType(ComboBox))      'Find the control
        all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()   'Alphabetical order
        For i = 0 To all_combo.Count - 1
            Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
            temp_string &= grbx.SelectedItem.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all checkbox controls and save
        FindControlRecursive(all_check, Me, GetType(CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As CheckBox = CType(all_check(i), CheckBox)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all radio controls and save
        FindControlRecursive(all_radio, Me, GetType(RadioButton))   'Find the control
        all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_radio.Count - 1
            Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '--------- add notes -----
        temp_string &= TextBox53.Text & ";"

        '---- if path not exist then create one----------
        Try
            If (Not System.IO.Directory.Exists(dirpath_Home)) Then System.IO.Directory.CreateDirectory(dirpath_Home)
            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
        Catch ex As Exception
        End Try

        Try
            If CInt(temp_string.Length.ToString) > 100 Then      'String may be empty
                If Directory.Exists(dirpath_Eng) Then
                    File.WriteAllText(dirpath_Eng & filename, temp_string, Encoding.ASCII)      'used at VTK
                Else
                    File.WriteAllText(dirpath_Home & filename, temp_string, Encoding.ASCII)     'used at home
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Line 5062, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    'Retrieve control settings and case_x_conditions from file
    'Split the file string into 5 separate strings
    'Each string represents a control type (combobox, checkbox,..)
    'Then split up the secton string into part to read into the parameters
    Private Sub Read_file()
        Dim control_words(), words() As String
        Dim i As Integer
        Dim ttt As Double
        Dim k As Integer = 0
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        OpenFileDialog1.FileName = "Venturi*"
        If Directory.Exists(dirpath_Eng) Then
            OpenFileDialog1.InitialDirectory = dirpath_Eng  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_Home  'used at home
        End If

        OpenFileDialog1.Title = "Open a Text File"
        OpenFileDialog1.Filter = "VTK Files|*.vtk"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim readText As String = File.ReadAllText(OpenFileDialog1.FileName, Encoding.ASCII)
            control_words = readText.Split(separators1, StringSplitOptions.None) 'Split the read file content

            '----- retrieve case condition-----
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split first line the read file content
            TextBox24.Text = words(0)                  'Project number
            TextBox25.Text = words(1)                  'Project name
            TextBox52.Text = words(2)                  'Tag ID

            '---------- terugzetten numeric controls -----------------
            FindControlRecursive(all_num, Me, GetType(NumericUpDown))
            all_num = all_num.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(1).Split(separators, StringSplitOptions.None)     'Split the read file content
            For i = 0 To all_num.Count - 1
                Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal numeric controls--
                If (i < words.Length - 1) Then
                    If Not (Double.TryParse(words(i + 1), ttt)) Then MessageBox.Show("Numeric controls conversion problem occured")
                    If ttt <= grbx.Maximum And ttt >= grbx.Minimum Then
                        grbx.Value = CDec(ttt)          'OK
                    Else
                        grbx.Value = grbx.Minimum       'NOK
                        MessageBox.Show("Numeric controls value out of ousode min-max range, Minimum value is used")
                    End If
                Else
                    MessageBox.Show("Warning last Numeric controls not found in file")  'NOK
                End If
            Next

            '---------- terugzetten combobox controls -----------------
            FindControlRecursive(all_combo, Me, GetType(ComboBox))
            all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(2).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_combo.Count - 1
                Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    grbx.SelectedItem = words(i + 1)
                Else
                    MessageBox.Show("Warning last combobox not found in file")
                End If
            Next

            '---------- terugzetten checkbox controls -----------------
            FindControlRecursive(all_check, Me, GetType(CheckBox))
            all_check = all_check.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(3).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_check.Count - 1
                Dim grbx As CheckBox = CType(all_check(i), CheckBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last checkbox not found in file")
                End If
            Next

            '---------- terugzetten radiobuttons controls -----------------
            FindControlRecursive(all_radio, Me, GetType(RadioButton))
            all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(4).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_radio.Count - 1
                Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal radiobuttons--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last radiobutton not found in file")
                End If
            Next
            '---------- terugzetten Notes -- ---------------
            If control_words.Count > 5 Then
                words = control_words(5).Split(separators, StringSplitOptions.None) 'Split the read file content
                TextBox53.Clear()
                TextBox53.AppendText(words(1))
            Else
                MessageBox.Show("Warning Notes not found in file")
            End If
        End If
    End Sub

    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function
End Class
