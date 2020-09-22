VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9960
   ClientLeft      =   1950
   ClientTop       =   540
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   7455
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   315
      Left            =   5775
      TabIndex        =   40
      Top             =   9600
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   2625
      TabIndex        =   34
      Tag             =   "NotEmpty;Numeric;Min=100; Max=200;Display=You have entered wrong parameter;"
      Top             =   8325
      Width           =   4665
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   2625
      TabIndex        =   30
      Tag             =   "Max=10;Numeric;"
      Top             =   7350
      Width           =   4665
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   2625
      TabIndex        =   26
      Tag             =   "Max=10;Numeric;"
      Top             =   6600
      Width           =   4665
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   2625
      TabIndex        =   22
      Tag             =   "Numeric;"
      Top             =   5925
      Width           =   4665
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   2625
      TabIndex        =   17
      Tag             =   "Time;"
      Top             =   4800
      Width           =   4665
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2625
      TabIndex        =   12
      Tag             =   "Date;"
      Top             =   3150
      Width           =   4665
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2625
      TabIndex        =   8
      Tag             =   "Lcase;"
      Top             =   2475
      Width           =   4665
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2625
      TabIndex        =   4
      Tag             =   "Ucase;"
      Top             =   1800
      Width           =   4665
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2625
      TabIndex        =   0
      Tag             =   "NotEmpty;"
      Top             =   1125
      Width           =   4665
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   3
      Left            =   2625
      TabIndex        =   39
      Top             =   8700
      Width           =   4665
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"Form1.frx":009A
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   7440
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   7275
      X2              =   75
      Y1              =   9525
      Y2              =   9525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Combination of conditions:"
      Height          =   195
      Index           =   8
      Left            =   75
      TabIndex        =   37
      Top             =   7875
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "NotEmpty;Numeric;Min=100; Max=200;Display=You have entered wrong parameter;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   17
      Left            =   4125
      TabIndex        =   36
      Top             =   7875
      Width           =   3285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property TAG is set to:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   16
      Left            =   2625
      TabIndex        =   35
      Top             =   7875
      Width           =   1425
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   7275
      X2              =   75
      Y1              =   7725
      Y2              =   7725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Minimum = 10"
      Height          =   195
      Index           =   7
      Left            =   75
      TabIndex        =   33
      Top             =   7425
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Min=10;Numeric;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   15
      Left            =   4125
      TabIndex        =   32
      Top             =   7125
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property TAG is set to:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   14
      Left            =   2625
      TabIndex        =   31
      Top             =   7125
      Width           =   1425
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   7275
      X2              =   75
      Y1              =   6975
      Y2              =   6975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "maximum =10"
      Height          =   195
      Index           =   6
      Left            =   75
      TabIndex        =   29
      Top             =   6675
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Max=10;Numeric;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   13
      Left            =   4125
      TabIndex        =   28
      Top             =   6375
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property TAG is set to:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   12
      Left            =   2625
      TabIndex        =   27
      Top             =   6375
      Width           =   1425
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   7275
      X2              =   75
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   7275
      X2              =   75
      Y1              =   5625
      Y2              =   5625
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   7275
      X2              =   75
      Y1              =   4425
      Y2              =   4425
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   7275
      X2              =   75
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   7275
      X2              =   75
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   7275
      X2              =   75
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Numeric value expected:"
      Height          =   195
      Index           =   5
      Left            =   75
      TabIndex        =   25
      Top             =   6000
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numeric;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   11
      Left            =   4125
      TabIndex        =   24
      Top             =   5700
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property TAG is set to:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   10
      Left            =   2625
      TabIndex        =   23
      Top             =   5700
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"Form1.frx":01B5
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   2625
      TabIndex        =   21
      Top             =   5175
      Width           =   4665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Time;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   9
      Left            =   4125
      TabIndex        =   20
      Top             =   4575
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Time is expected:"
      Height          =   195
      Index           =   4
      Left            =   75
      TabIndex        =   19
      Top             =   4875
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property TAG is set to:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   8
      Left            =   2625
      TabIndex        =   18
      Top             =   4575
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"Form1.frx":0244
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   0
      Left            =   2625
      TabIndex        =   16
      Top             =   3525
      Width           =   4665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   7
      Left            =   4125
      TabIndex        =   15
      Top             =   2925
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Date is expected:"
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   14
      Top             =   3225
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property TAG is set to:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   6
      Left            =   2625
      TabIndex        =   13
      Top             =   2925
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lcase;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   5
      Left            =   4125
      TabIndex        =   11
      Top             =   2250
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "text is converted to lower case:"
      Height          =   195
      Index           =   2
      Left            =   75
      TabIndex        =   10
      Top             =   2550
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property TAG is set to:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   4
      Left            =   2625
      TabIndex        =   9
      Top             =   2250
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "text is converted to UPER CASE:"
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   7
      Top             =   1875
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ucase;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   3
      Left            =   4125
      TabIndex        =   6
      Top             =   1575
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property TAG is set to:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   2
      Left            =   2625
      TabIndex        =   5
      Top             =   1575
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "this textbox mustn't be empty!"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   3
      Top             =   1200
      Width           =   2070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NotEmpty;"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Index           =   1
      Left            =   4125
      TabIndex        =   2
      Top             =   900
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Property TAG is set to:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   2625
      TabIndex        =   1
      Top             =   900
      Width           =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim xVal As Boolean
    Dim ValError As Boolean
    Dim xRepeat As Single
    For xRepeat = 0 To Text1.Count - 1
        xVal = Validate(Text1(xRepeat))
        If Not xVal Then
            DisplayErrTbX Text1(xRepeat)
            ValError = True
            Exit For
        End If
    Next
    If ValError Then
        MsgBox "Please enter correct parameter!", vbCritical
        Text1(xRepeat).SetFocus
    Else
        MsgBox "Form Valiadtion is compleeted. Please vote for my code, or email me to: seby@email.si...", vbInformation
        Unload Me
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim xVal As Boolean
    xVal = Validate(Text1(Index))
    If Not xVal Then DisplayErrTbX Text1(Index)
End Sub
