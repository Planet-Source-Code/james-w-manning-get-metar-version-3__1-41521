VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optUnits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imperial"
      Height          =   255
      Index           =   1
      Left            =   4770
      TabIndex        =   22
      Top             =   4335
      Width           =   1215
   End
   Begin VB.OptionButton optUnits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Metric"
      Height          =   255
      Index           =   0
      Left            =   4770
      TabIndex        =   21
      Top             =   4095
      Value           =   -1  'True
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   5610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView vwCountry 
      Height          =   8895
      Left            =   15
      TabIndex        =   20
      Top             =   30
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   15690
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check every half hour"
      Height          =   240
      Left            =   4785
      TabIndex        =   19
      Top             =   3810
      Width           =   3870
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14520
      Top             =   255
   End
   Begin VB.ComboBox cboAirport 
      Height          =   315
      Left            =   5880
      TabIndex        =   18
      Top             =   330
      Width           =   6885
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Wind Direction"
      Height          =   2055
      Left            =   13365
      TabIndex        =   14
      Top             =   1185
      Width           =   1815
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   15
         Top             =   240
         Width           =   150
      End
      Begin VB.Shape Shape1 
         Height          =   1365
         Left            =   120
         Shape           =   3  'Circle
         Top             =   480
         Width           =   1665
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   960
         X2              =   1620
         Y1              =   1170
         Y2              =   1170
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4740
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Metar"
      Height          =   525
      Left            =   12870
      TabIndex        =   0
      Top             =   210
      Width           =   1425
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   10080
      X2              =   8565
      Y1              =   6480
      Y2              =   6495
   End
   Begin VB.Label lblPressure 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   6450
      TabIndex        =   17
      Top             =   3510
      Width           =   7185
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Atmospheric Pressure:"
      Height          =   195
      Left            =   4800
      TabIndex        =   16
      Top             =   3510
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   195
      Left            =   4800
      TabIndex        =   13
      Top             =   1110
      Width           =   390
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   5835
      TabIndex        =   12
      Top             =   1095
      Width           =   2685
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      Height          =   195
      Left            =   4800
      TabIndex        =   11
      Top             =   1500
      Width           =   390
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   5835
      TabIndex        =   10
      Top             =   1485
      Width           =   2685
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Wind:"
      Height          =   195
      Left            =   4800
      TabIndex        =   9
      Top             =   1950
      Width           =   420
   End
   Begin VB.Label lblWind 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   5835
      TabIndex        =   8
      Top             =   1935
      Width           =   6315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Visibility:"
      Height          =   195
      Left            =   4800
      TabIndex        =   7
      Top             =   2340
      Width           =   585
   End
   Begin VB.Label lblVis 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   5835
      TabIndex        =   6
      Top             =   2325
      Width           =   2685
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Temperature:"
      Height          =   195
      Left            =   4800
      TabIndex        =   5
      Top             =   2775
      Width           =   945
   End
   Begin VB.Label lblTemperature 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   5835
      TabIndex        =   4
      Top             =   2760
      Width           =   2685
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Conditions:"
      Height          =   195
      Left            =   4800
      TabIndex        =   3
      Top             =   3135
      Width           =   780
   End
   Begin VB.Label lblConditions 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   5835
      TabIndex        =   2
      Top             =   3015
      Width           =   6900
   End
   Begin VB.Label lblMetar 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   5130
      TabIndex        =   1
      Top             =   780
      Width           =   9045
   End
   Begin VB.Image Image1 
      Height          =   4545
      Left            =   7845
      Picture         =   "frmMain.frx":1138
      Top             =   4260
      Width           =   4530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' I apologize for no comments in this code. I will change that now...
' I've also included a pdf file for understanding METAR data.

Dim Metar As String 'This holds the Metar string minus the remarks (RMK Code)
Dim i As Integer, j As Integer 'Couple of counters
Dim Temp As String, temp2 As String 'Temp strings
Dim Data As String
Dim mData() As String 'This holds the current data token from the Metar string
'Dim Direct As Integer
Dim mDate As String 'Holds the date read from the METAR
Dim mTime As String 'Holds the time
Dim Tmr As Integer 'Holds the amount of passed time for storing in database
'Dim CurX As Integer
Dim mTemp As Integer 'Holds the current Temperature
Dim mConditions As String 'Holds the current conditions
Dim mAtmos As Double 'Holds atmospheric pressure in Kpa.
Dim WindSpd As Integer 'Holds the windspeed
Dim WindDir As Integer ' Holds the wind direction
Dim WindGust As Integer ' Holds the wind gust speed
Dim Visib As Integer ' Holds Visibility distance
Dim Imperial As Boolean ' if true, uses imperial measurements, otherwise metric.

Private Sub Check1_Click()
    'This is to check every 30 minutes for any weather change.
    If Check1.Value = 1 Then
        Timer1.Enabled = True
        Tmr = 0
        Check1.Caption = "Stop checking every 30 Minutes"
    Else
        Timer1.Enabled = False
        Check1.Caption = "Check every 30 minutes"
        Tmr = 0
    End If
End Sub

Private Sub Command1_Click()
    On Error GoTo Metar_err
    Dim TZone As String
    Dim Bias As Integer
    Metar = Inet1.OpenURL("http://weather.noaa.gov/cgi-bin/mgetmetar.pl?cccc=" & Mid$(cboAirport.Text, 1, 4), icString)
    i = InStr(1, Metar, "<P>The observation is:</P>")
    i = InStr(i, Metar, Mid$(cboAirport.Text, 1, 4))
    If InStr(1, Metar, "RMK") <> 0 Then
        Metar = Mid$(Metar, i, InStr(i, Metar, "RMK") - i)
    Else
        Metar = Mid$(Metar, i, InStr(i, Metar, Chr(10)) - i)
    End If
    Temp = ""
    lblConditions.Caption = Temp
    mData = Split(Metar, " ")
    For i = 0 To UBound(mData(), 1) - 1
        If mData(i) <> vbNullString Then Data = mData(i)
        If Data <> Mid$(cboAirport.Text, 1, 4) Then
            If InStr(1, Data, "Z") <> 0 And Val(Data) <> 0 Then
            'Date and time
                lblDate.Caption = MonthName(Month(Now)) & " " & Mid$(Data, 1, 2) & ", " & Year(Now)
                Bias = GetLocalTZ(TZone)
                lblTime.Caption = IIf(Val(Mid$(Data, 3, 2)) + Bias > 12, Val(Mid$(Data, 3, 2)) + Bias - 12, IIf(Val(Mid$(Data, 3, 2)) + Bias <= 0, Val(Mid$(Data, 3, 2)) + Bias + 12, Val(Mid$(Data, 3, 2)) + Bias)) & ":" & Mid$(Data, 5, 2) & " " & IIf(Val(Mid$(Data, 3, 2)) + Bias > 12, "PM", IIf(Val(Mid$(Data, 3, 2)) + Bias <= 0, "PM", "AM")) & " (" & TZone & ")"
            ElseIf InStr(1, Data, "KT") <> 0 Then
            'Wind Speed and direction as well as gusts if any
                temp2 = ""
                WindDir = Val(Mid$(Data, 1, 3))
                If InStr(1, Data, "G") <> 0 Then
                    WindSpd = Val(Mid$(Data, InStr(4, Data, "G") + 1, InStr(InStr(4, Data, "G") + 1, Data, "KT") - 2))
                    WindGust = Val(Mid$(Data, InStr(1, Data, "G") - 2, InStr(4, Data, "KT") - InStr(4, Data, "G") + 1))
                    temp2 = IIf(Not Imperial, Round(WindSpd * 1.852, 0) & " km/h", WindSpd & " knots (" & Round(WindSpd * 1.15, 0) & " mph)") & "with " & IIf(Not Imperial, Round(WindGust * 1.852, 0) & " km/h", WindGust & " knots (" & Round(WindGust * 1.15, 0) & " mph)") & " Gusts"
                Else
                    WindSpd = Val(Mid$(Data, 4, 2))
                    temp2 = IIf(Not Imperial, Round(WindSpd * 1.852, 0) & " km/h", WindSpd & " knots (" & Round(WindSpd * 1.15, 0) & " mph)")
                End If
                'IIf(InStr(1, Data, "G") <> 0, Round(Val(Mid$(Data, InStr(4, Data, "G") + 1, InStr(InStr(4, Data, "G") + 1, Data, "KT") - 2)) * 1.852, 0) & " km/h with " & Round(Val(Mid$(Data, InStr(1, Data, "G") - 2, InStr(4, Data, "KT") - InStr(4, Data, "G") + 1)) * 1.852, 0) & " km/h Gusts", Round(Val(Mid$(Data, 4, 2)) * 1.852, 0) & " km/h")
                Line1.X2 = Line1.X1 + Sin(3.14159 * WindDir / 180) * 660
                Line1.Y2 = Line1.Y1 - Cos(3.14159 * WindDir / 180) * 660
                lblWind.Caption = WindDir & " Degrees at " & temp2
            ElseIf InStr(1, Data, "SM") <> 0 Then
            'Visibility in Km
                If InStr(1, Data, "/") = 0 Then
                    Visib = Val(Mid$(Data, 1, InStr(1, Data, "SM") - 1))
                    lblVis.Caption = IIf(Not Imperial, Round(Visib * 1.609344) & " km", Visib & " Statute Miles")
                Else
                    Visib = Val(Mid$(Data, 1, InStr(1, Data, "/") - 1)) / Val(Mid$(Data, InStr(1, Data, "/") + 1, 1))
                    lblVis.Caption = IIf(Not Imperial, Round(Visib * 1.609344) & " km", Visib & " Statute Miles")
                End If
            ElseIf Mid$(Data, 1, 1) = "A" And Len(Data) = 5 Then
            'Atmospheric Pressure in Kpa
                mAtmos = Val(Mid$(Data, 2, 2) & "." & Mid$(Data, 4, 2))
                lblPressure.Caption = IIf(Not Imperial, Round(mAtmos * 3.386, 2) & " Kpa", mAtmos & " inHg (" & Round(mAtmos * 33.864, 2) & " mb)")
            ElseIf InStr(1, Data, "/") <> 0 And InStr(1, Data, "SM") = 0 Then
            'Temperature in °C
                mTemp = Val(IIf(Mid$(Data, 1, 1) = "M", "-", "") & Val(Mid$(Data, IIf(Mid$(Data, 1, 1) = "M", 2, 1), InStr(1, Data, "/") - 1)))
                Line2.X2 = Line2.X1 + Sin(3.14159 * (Round(mTemp / 140 * 360) - 44) / 180) * 1515
                Line2.Y2 = Line2.Y1 - Cos(3.14159 * (Round(mTemp / 140 * 360) - 44) / 180) * 1515
                Form1.Caption = IIf(Not Imperial, mTemp & "°C", Round((mTemp * 9 / 5) + 32, 0) & "°F")
                lblTemperature.Caption = IIf(Not Imperial, mTemp & "°C", Round((mTemp * 9 / 5) + 32, 0) & "°F")
            ElseIf Data = "BR" Then
                Temp = Temp & " Mist."
            Else
            'Anything else like cloud cover, and precipitation
                j = 1
                If InStr(1, Data, "-") <> 0 Then
                'Light precipitation
                    Data = Mid$(Data, 2, Len(Data) - 1)
                    Temp = Temp & " Light"
                ElseIf InStr(1, Data, "+") <> 0 Then
                'Heavy Precipitation
                    Data = Mid$(Data, 2, Len(Data) - 1)
                    Temp = Temp & " Heavy"
                End If
                Do Until j > Len(Data)
                'Snow, Showers, Rain, Blowing, Drifting, Drizzle, Haze.
                    Select Case Mid$(Data, j, 2)
                        Case "SN"
                            Temp = Temp & " Snow."
                        Case "SH"
                            Temp = Temp & " Showers."
                        Case "DR"
                            Temp = Temp & " Drifting."
                        Case "BL"
                            Temp = Temp & " Blowing."
                        Case "RA"
                            Temp = Temp & " Rain."
                        Case "DZ"
                            Temp = Temp & " Drizzle."
                        Case "HZ"
                            Temp = Temp & " Haze."
                    End Select
                    j = j + 2
                Loop
                'Cloud Conditions and Heights (ft)
                If Mid$(Data, 1, 3) = "OVC" Then
                    Temp = Temp & " Overcast clouds at " & Val(Mid$(Data, 4, 3)) * 100 & " feet."
                ElseIf Mid$(Data, 1, 3) = "BKN" Then
                    Temp = Temp & " Broken clouds at " & Val(Mid$(Data, 4, 3)) * 100 & " feet."
                ElseIf Mid$(Data, 1, 3) = "FEW" Then
                    Temp = Temp & " Few clouds at " & Val(Mid$(Data, 4, 3)) * 100 & " feet."
                ElseIf Mid$(Data, 1, 3) = "SCT" Then
                    Temp = Temp & " Scattered clouds at " & Val(Mid$(Data, 4, 3)) * 100 & " feet."
                ElseIf Mid$(Data, 1, 3) = "FG" Then
                    Temp = Temp & " Foggy."
                ElseIf Mid$(Data, 1, 3) = "CLR" Then
                    Temp = Temp & " Clear."
                End If
                If Temp <> "" Then
                    lblConditions.Caption = Temp
                    mConditions = Temp
                End If
            End If
        End If
        Data = ""
    Next i
    lblMetar.Caption = Metar
    Exit Sub
Metar_err:
    MsgBox "There is no data for " & Mid$(cboAirport.Text, 1, 4) & ", try again.", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
    Dim Tmp As String
    Dim Lst As Node
    Dim Country() As String, City() As String, Airport As String, oldTmp
    Dim Temp As String
    Dim cBreakout() As String
    Dim Cnt As Integer
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    
    Imperial = False 'It's a canadian program, so I'm biased! :-)
    
    Open App.Path & "\ICAO Name.cfg" For Input As #1
    cboAirport.Clear
    cboAirport.Text = "CYHM Hamilton Airport, Canada"
    j = 0
    k = 0
    ReDim Country(j) As String
    ReDim City(k) As String
    Dim strTZ As String
    Screen.MousePointer = 11
    frmWait.Show
    Do Until EOF(1)
        DoEvents
        Line Input #1, Tmp
        cboAirport.AddItem Tmp
        Temp = Right$(Tmp, Len(Tmp) - InStrRev(Tmp, ","))
        oldTemp = vbNullString
        For i = 0 To UBound(Country(), 1)
            If Country(i) = Temp Then
                oldTemp = Temp
            End If
        Next i
        If oldTemp = vbNullString Then
            ReDim Preserve Country(j) As String
            Country(j) = Temp
            ReDim Preserve City(k) As String
            City(k) = Tmp
            j = j + 1
        Else
            ReDim Preserve City(k) As String
            City(k) = Tmp
        End If
        k = k + 1
    Loop
    Close #1
    For i = 0 To j - 1
        DoEvents
        vwCountry.Nodes.Add , , Trim(Country(i)), Trim(Country(i)), 3
        For l = 0 To k - 1
            Erase cBreakout
            cBreakout = Split(City(l), ",")
            If Trim(cBreakout(UBound(cBreakout(), 1))) = Trim(Country(i)) Then
                vwCountry.Nodes.Add Trim(Country(i)), tvwChild, , Trim(cBreakout(0))
            End If
        Next l
    Next i
    Erase cBreakout
    Screen.MousePointer = 0
    frmWait.Hide
    Tmr = 0
    CurX = 0
    oldTemp = 0
    Timer1.Enabled = True
End Sub

Private Sub optUnits_Click(Index As Integer)
    If Index = 0 Then
        Imperial = False
    Else
        Imperial = True
    End If
    Call Command1_Click
End Sub

Private Sub Timer1_Timer()
'    Dim xData As ADODB.Connection
'    Dim xRs As ADODB.Recordset
    
    Tmr = Tmr + 1
    If Tmr = 1800 Then
'        Set xData = New ADODB.Connection
'        Set xRs = New ADODB.Recordset
'        With xData
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=ostrich;Persist Security Info=True;User ID=sa;Initial Catalog=Weather;Data Source=JMANNING"
'            .Open
'        End With
        Call Command1_Click
'        With xRs
'            .Open "SELECT * FROM [Weather Capture]", xData, adOpenKeyset, adLockOptimistic
'            .AddNew
'            !wDate = FormatDateTime(Now, vbShortDate)
'            !wTime = FormatDateTime(Now, vbShortTime)
'            !Temperature = mTemp
'            !AtmosPressure = mAtmos
'            !Conditions = mConditions
'            !Metar = Metar
'            .Update
'        End With
'        Set xRs = Nothing
'        Set xData = Nothing
        Tmr = 0
    End If
End Sub

Private Sub vwCountry_NodeClick(ByVal Node As MSComctlLib.Node)
    If Not Node.Parent Is Nothing Then
        cboAirport.Text = Node.Text
        Command1_Click
    End If
End Sub
