VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Let's ROLL some PICTURE "
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14205
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   14205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrFilm 
      Interval        =   500
      Left            =   -2880
      Top             =   25920
   End
   Begin VB.Timer tmrSlideGen 
      Interval        =   2000
      Left            =   -13440
      Top             =   8640
   End
   Begin VB.Label lblComment 
      Caption         =   "Enjoy the smile of beauty....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   11160
      Width           =   4335
   End
   Begin VB.Image imgTosi11 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   240
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Image imgTosi12 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Image imgTosi13 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Image imgTosi14 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Image imgTosi15 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   10440
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Image imgTosi16 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   12600
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Image imgTosi10 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   12480
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image imgTosi9 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image imgTosi8 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image imgTosi7 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image imgTosi6 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image imgTosi5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   240
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Press ""Esc"" to Exit..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   1
      Top             =   11160
      Width           =   4095
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Design: Fatos Halilaj - F5 for more.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   11160
      Width           =   4695
   End
   Begin VB.Image imgTosi4 
      Height          =   5295
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   7695
   End
   Begin VB.Image imgTosi3 
      Height          =   5295
      Left            =   0
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   7575
   End
   Begin VB.Image imgTosi2 
      Height          =   5775
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
   Begin VB.Image imgTosi1 
      Height          =   5775
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Demonstrating simple "Slide Show"
' Date: 2006/06/27
' Coded by: Fatos Halilaj
' if you have any comment or something like that, then mail me.
' mail: fatosi2005@yahoo.com
' =-=-=-=-=-=-=-=-=-=-=-=--code can be freely used in any place.
' -------------------------use it as you need...

' Hope to find some more time and I can give some more simple examples

Option Explicit
Rem SLIDE SHOW


Private Sub GetImage(imgTosi As Image)
Dim foto As Integer
foto = 1 + Int(16 * Rnd())     ' Generating random integers 1 to 16

' getting picture
imgTosi.Picture = LoadPicture(App.Path & "\images\pic" & foto & ".jpg")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyEscape Then
   MsgBox "Hope you loved this simple ''§LIDE §HOW'?'" & vbCrLf & _
    vbCrLf & "I think it's simple but looks great" & vbCrLf & Space$(38) & _
                        "¦" & ".....Author", vbInformation, "Peace out...Tosi"
   Unload Me
   End
ElseIf KeyCode = vbKeyF5 Then
    frmAuthor.Show

End If
End Sub


Private Sub Form_Load()
Dim message As String

message = "Welcome to our show"
message = message & vbCrLf & _
          "Hope you'll like this simple code and give your opinion - comments"
message = message & vbCrLf & vbCrLf & "When you click the OK button: " & vbCrLf & _
          "- To see more information about the author - press F5 function key" & vbCrLf & _
          "- To Exit the program - show, you may click the ESCAPE button on your keyboard"
message = message & vbCrLf & "- Do take care, code can be freely used but not with the pictures ok, I haven't ask" & _
                   vbCrLf & "those chicks to using them in our stuff, but demonstration is demonstration, ok"
message = message & vbCrLf & vbCrLf & Space$(26) & "Thanks, Tosi"

MsgBox message, vbInformation, "Welcome..."

  
  KeyPreview = True
End Sub



Private Sub tmrFilm_Timer()
   Call GetImage(imgTosi5)
   Call GetImage(imgTosi6)
   Call GetImage(imgTosi7)
   Call GetImage(imgTosi8)
   Call GetImage(imgTosi9)
   Call GetImage(imgTosi10)
   Call GetImage(imgTosi11)
   Call GetImage(imgTosi12)
   Call GetImage(imgTosi13)
   Call GetImage(imgTosi14)
   Call GetImage(imgTosi15)
   Call GetImage(imgTosi16)
End Sub

Private Sub tmrSlideGen_Timer()
   Call GetImage(imgTosi1)
   Call GetImage(imgTosi2)
   Call GetImage(imgTosi3)
   Call GetImage(imgTosi4)
End Sub

