VERSION 5.00
Begin VB.UserControl tCal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   CanGetFocus     =   0   'False
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   ScaleHeight     =   5730
   ScaleWidth      =   5625
   Begin VB.Timer MonthTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   660
      Top             =   5235
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Left            =   360
      TabIndex        =   11
      Top             =   3240
      Width           =   1875
   End
   Begin VB.Shape UDshape 
      DrawMode        =   7  'Invert
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   2385
      Top             =   15
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape UDshape 
      DrawMode        =   7  'Invert
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   45
      Top             =   15
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Mon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   30
      TabIndex        =   10
      Top             =   360
      Width           =   270
   End
   Begin VB.Label labDayNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   345
      TabIndex        =   9
      Top             =   840
      Width           =   270
   End
   Begin VB.Label labMonth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "February 2001"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   285
      TabIndex        =   8
      Top             =   45
      Width           =   1620
   End
   Begin VB.Label labDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Tue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   1
      Left            =   315
      TabIndex        =   7
      Top             =   360
      Width           =   270
   End
   Begin VB.Label labDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Wed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   360
      Width           =   270
   End
   Begin VB.Label labDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Thu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   3
      Left            =   885
      TabIndex        =   5
      Top             =   360
      Width           =   270
   End
   Begin VB.Label labDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Fri"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   4
      Left            =   1170
      TabIndex        =   4
      Top             =   360
      Width           =   270
   End
   Begin VB.Label labDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Sat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   5
      Left            =   1455
      TabIndex        =   3
      Top             =   360
      Width           =   270
   End
   Begin VB.Label labDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Sun"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   6
      Left            =   1740
      TabIndex        =   2
      Top             =   360
      Width           =   270
   End
   Begin VB.Label labUD 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label labUD 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2085
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "tCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim MonthFontName As String
Dim MonthFontSize As Integer
Dim DayFontName As String
Dim DayFontSize As Integer
Dim DayNoFontName As String
Dim DayNoFontSize As Integer
Dim i7 As Integer
Dim CurDate As Date
Dim CurDateOff As Date
Dim CurIdxOFf As Integer
Dim Udtick As Long
Dim MAXtick As Integer
Dim UDdir As Integer
Dim SelDate As Date
Dim GetOut As Boolean
'Default Property Values:
Const m_def_Value = 0
'Property Variables:
Dim m_Value As Date
'Event Declarations:
Event DateChanged(dte As Date)







Private Sub labDayNo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'labDayNo(Index).ForeColor = &HFFFFFF
'labDayNo(Index).BackColor = &H0
CurIdxOFf = Index
SelDate = DateAdd("d", Index, CurDateOff)

Call FillDates
Call ShowSelDate

End Sub

Private Sub labDayNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'labDayNo(Index).ForeColor = &H0
'labDayNo(Index).BackColor = &HFFFFFF

End Sub



Private Sub labUD_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
UDshape(Index).Visible = False
MonthTimer.Enabled = False

End Sub



Private Sub UserControl_Initialize()
Call Arrange
CurDate = tINTDATE(Now)
SelDate = Now
m_def_SelDate = Now
Call FillDates
Call ShowSelDate

End Sub
Private Sub ShowSelDate()
Dim Doff As Integer
CurIdxOFf = DateDiff("d", CurDateOff, SelDate)
labDayNo(CurIdxOFf).ForeColor = &HFF&

RaiseEvent DateChanged(SelDate)

End Sub


Private Function tINTDATE(dte) As Date
Dim d As Integer
Dim Y As Integer
d = 1 - Day(dte)
tINTDATE = DateAdd("d", d, dte)



End Function


Private Sub Arrange()
Dim Vspc As Integer
Dim Hspc As Integer
Dim j As Integer
Dim X As Integer
Dim Y As Integer
Dim Xoff As Integer
Dim Yoff As Integer
UserControl.BackColor = &HFFFFFF
Vspc = 240
Hspc = 285
Xoff = 30
Yoff = 495
labUD(0).Top = Yoff - 480
labUD(1).Top = Yoff - 480
labUD(0).Left = Xoff + 15
labUD(1).Left = Xoff + 1755

UDshape(0).Top = Yoff - 450
UDshape(1).Top = Yoff - 450
UDshape(0).Left = Xoff + 15
UDshape(1).Left = Xoff + 1800
UDshape(0).Width = 165
UDshape(0).Height = 240
UDshape(1).Width = 165
UDshape(1).Height = 240

labMonth.Top = Yoff - 450
labMonth.Left = Xoff + 210

For Y = 1 To 5
   UserControl.Line (15, Y * Vspc + Yoff)-(2040, Y * Vspc + Yoff), &HC0C0C0
Next Y
For X = 1 To 6
   UserControl.Line (X * Hspc + Xoff - 15, Yoff)-(X * Hspc + Xoff - 15, Yoff + 1735 - 270), &HC0C0C0
Next X
UserControl.Line (0, 0)-(2100 - 60, 2100 - 135), &H0, B
For Y = 0 To 5
   For X = 0 To 6
      labDay(X).Top = Yoff - 180
      labDay(X).Left = X * Hspc + Xoff - 15
      labDay(X).Width = Hspc + 30
      labDay(X).Height = 180
      If j > 0 Then
         Load labDayNo(j)
         labDayNo(j).Visible = True
      End If
      labDayNo(j).Height = Vspc - 45
      labDayNo(j).Width = Hspc - 30
      labDayNo(j).Top = Y * Vspc + Yoff + 30
      labDayNo(j).Left = X * Hspc + Xoff + 15
      j = j + 1

   Next X
Next Y




End Sub






Private Sub FillDates()
Dim tDte As Date
Dim tMON As Integer
Dim Doff As Integer
Doff = 1 - Weekday(CurDate, 2)
tDte = DateAdd("d", Doff, CurDate)
CurDateOff = tDte
labMonth = Format(CurDate, "mmmm yyyy")
tMON = Month(CurDate)

CurIdxOFf = DateDiff("d", CurDateOff, SelDate)


For i7 = 0 To 41
   labDayNo(i7) = Format(tDte, "dd")
   If Month(tDte) = tMON Then
      labDayNo(i7).ForeColor = &H0
   Else
      labDayNo(i7).ForeColor = &HC0C0C0
   End If
   
   tDte = DateAdd("d", 1, tDte)
   
Next i7



End Sub

Private Sub MonthTimer_Timer()
Dim keepOFF As Integer
keepOFF = CurIdxOFf

Udtick = Udtick + 1
If Udtick > MAXtick Then
   CurDate = DateAdd("m", UDdir, CurDate)
   Call FillDates
   CurIdxOFf = keepOFF
   SelDate = DateAdd("d", CurIdxOFf, CurDateOff)
   Call ShowSelDate
End If

End Sub
Private Sub labUD_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim keepOFF As Integer
UDshape(Index).Visible = True
keepOFF = CurIdxOFf
If Button = 1 Then
   MAXtick = 1
   MonthTimer.Interval = 200
Else
   MAXtick = 5
   MonthTimer.Interval = 5
End If

Select Case Index
Case 0
   CurDate = DateAdd("m", -1, CurDate)
   Call FillDates
   CurIdxOFf = keepOFF
   SelDate = DateAdd("d", CurIdxOFf, CurDateOff)
   UDdir = -1
   Call ShowSelDate
Case 1
   CurDate = DateAdd("m", 1, CurDate)
   Call FillDates
   CurIdxOFf = keepOFF
   SelDate = DateAdd("d", CurIdxOFf, CurDateOff)
   UDdir = 1
   Call ShowSelDate
Case Else
End Select
Udtick = 0
MonthTimer.Enabled = True

End Sub

Private Sub UserControl_Resize()
If GetOut Then Exit Sub
GetOut = True
UserControl.Width = 2100 - 45
UserControl.Height = 2100 - 120
GetOut = False

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=3,0,0,0
Public Property Get Value() As Date
   Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Date)
   m_Value = New_Value
   PropertyChanged "Value"
CurDate = tINTDATE(m_Value)
SelDate = m_Value
Call FillDates
Call ShowSelDate
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_Value = m_def_Value
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_Value = PropBag.ReadProperty("Value", m_def_Value)

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

