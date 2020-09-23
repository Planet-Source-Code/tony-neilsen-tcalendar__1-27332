VERSION 5.00
Object = "{E5DD9816-AC55-11D5-86C8-00E0299370E5}#2.0#0"; "tpnCalendar.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   6900
   ClientLeft      =   2550
   ClientTop       =   2055
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   8370
   Begin tpnCalendar.tCal tCal1 
      Height          =   1980
      Left            =   885
      Top             =   1050
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3493
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   825
      TabIndex        =   1
      Text            =   "12/04/2001"
      Top             =   675
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Left            =   855
      TabIndex        =   0
      Top             =   255
      Width           =   2265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub tCal1_DateChanged(dte As Date)
Label1 = dte
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   tCal1.Value = CDate(Text1)
   Label1 = CDate(Text1)
End If

End Sub
