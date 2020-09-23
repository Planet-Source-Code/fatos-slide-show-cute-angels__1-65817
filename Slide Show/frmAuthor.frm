VERSION 5.00
Begin VB.Form frmAuthor 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Author"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   -7200
      Top             =   480
   End
   Begin VB.Line lnMail 
      X1              =   2760
      X2              =   5280
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblMail 
      BackColor       =   &H80000008&
      Caption         =   "                                                                                        fatosi2005@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label lblCntct 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact me: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Line lnAuthor 
      BorderColor     =   &H0080C0FF&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   2520
      X2              =   5160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line lnCompany 
      BorderWidth     =   4
      X1              =   2160
      X2              =   6120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "BLI SOFT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   120
      Picture         =   "frmAuthor.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1785
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Fatos Halilaj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 KeyPreview = True
End Sub



Private Sub Timer1_Timer()
lblCompany.ForeColor = QBColor(Rnd * 13)
lblCntct.ForeColor = QBColor(Rnd * 10)
lnAuthor.BorderColor = QBColor(Rnd * 14)
lnAuthor.BorderStyle = Rnd * (5)
lnCompany.BorderColor = QBColor(Rnd * 14)
lnCompany.BorderStyle = Rnd * 6
lnMail.BorderColor = QBColor(Rnd * 11)
lnMail.BorderStyle = Rnd * 7
End Sub
