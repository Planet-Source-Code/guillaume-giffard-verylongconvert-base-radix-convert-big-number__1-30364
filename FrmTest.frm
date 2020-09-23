VERSION 5.00
Begin VB.Form FrmTest 
   Caption         =   "FrmTest"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9300
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7320
      TabIndex        =   10
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   2085
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3720
      Width           =   9015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Default         =   -1  'True
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   2085
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   9015
   End
   Begin VB.Label Label6 
      Caption         =   "Default = . (point)"
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Separator :"
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Output base :"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Input base :"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Output number :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Input number :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu MnAbout 
      Caption         =   "About..."
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text5.Text = "" Then Text5.Text = "."
    If Text3.Text <> "" And Text4.Text <> "" Then Text2.Text = VeryLongConvert(Text1.Text, CInt(Text3.Text), CInt(Text4.Text), Text5.Text)
End Sub

Private Sub MnAbout_Click()
    Call MsgBox("VeryLongConvert is a function that converts numbers up to 32,000 digits from any base between 2 and 36 to another." + (Chr(13) & Chr(10)) + (Chr(13) & Chr(10)) + "Programmed by Guillaume GIFFARD the 01/03/2002" + (Chr(13) & Chr(10)) + "Contact me at Guiland@mail.com" + (Chr(13) & Chr(10)) + (Chr(13) & Chr(10)) + "See and download my programmes at :" + (Chr(13) & Chr(10)) + "http://www.planet-source-code.com/vb", 64, "About VeryLongConvert")
End Sub
