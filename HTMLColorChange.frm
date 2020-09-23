VERSION 5.00
Begin VB.Form HTMLColorChange 
   BackColor       =   &H00FF8080&
   Caption         =   "HTML Color Fader by DrushDman"
   ClientHeight    =   5700
   ClientLeft      =   3105
   ClientTop       =   1290
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "HTMLColorChange.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   8070
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Caption         =   "Colors To Fade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   3135
      Begin VB.VScrollBar Blue5 
         Height          =   975
         LargeChange     =   3
         Left            =   2520
         SmallChange     =   3
         TabIndex        =   24
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar green5 
         Height          =   975
         LargeChange     =   3
         Left            =   2640
         SmallChange     =   3
         TabIndex        =   23
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar RED5 
         Height          =   975
         LargeChange     =   3
         Left            =   2760
         SmallChange     =   3
         TabIndex        =   22
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar bLuE1 
         Height          =   975
         LargeChange     =   3
         Left            =   120
         SmallChange     =   3
         TabIndex        =   17
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar gReEn1 
         Height          =   975
         LargeChange     =   3
         Left            =   240
         SmallChange     =   3
         TabIndex        =   16
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar rEd1 
         Height          =   975
         LargeChange     =   3
         Left            =   360
         SmallChange     =   3
         TabIndex        =   15
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar bLuE2 
         Height          =   975
         LargeChange     =   3
         Left            =   720
         SmallChange     =   3
         TabIndex        =   14
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar grEEn2 
         Height          =   975
         LargeChange     =   3
         Left            =   840
         SmallChange     =   3
         TabIndex        =   13
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar rED2 
         Height          =   975
         LargeChange     =   3
         Left            =   960
         SmallChange     =   3
         TabIndex        =   12
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar bLuE3 
         Height          =   975
         LargeChange     =   3
         Left            =   1320
         SmallChange     =   3
         TabIndex        =   11
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar gREEN3 
         Height          =   975
         LargeChange     =   3
         Left            =   1440
         SmallChange     =   3
         TabIndex        =   10
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar rED3 
         Height          =   975
         LargeChange     =   3
         Left            =   1560
         SmallChange     =   3
         TabIndex        =   9
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar Blue4 
         Height          =   975
         LargeChange     =   3
         Left            =   1920
         SmallChange     =   3
         TabIndex        =   8
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar Green4 
         Height          =   975
         LargeChange     =   3
         Left            =   2040
         SmallChange     =   3
         TabIndex        =   7
         Top             =   360
         Width           =   135
      End
      Begin VB.VScrollBar Red4 
         Height          =   975
         LargeChange     =   3
         Left            =   2160
         SmallChange     =   3
         TabIndex        =   6
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Color5 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Color3 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label color2 
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         TabIndex        =   20
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label color1 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Color4 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Caption         =   "Make Wavy"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      ScaleHeight     =   315
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4200
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2880
      Width           =   7815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Emberson Fx HTML Color Fader"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   20.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   465
      Index           =   1
      Left            =   1800
      TabIndex        =   29
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Text to Fade"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   120
      TabIndex        =   28
      Top             =   2520
      Width           =   1650
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   27
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "HTML Code"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   120
      TabIndex        =   26
      Top             =   3840
      Width           =   1590
   End
End
Attribute VB_Name = "HTMLColorChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BLUE1_Change()
'Whenever button is pulled it changes the color of the label on the bottom
color1.BackColor = RGB(rEd1.Value, gReEn1.Value, bLuE1.Value)
End Sub

Private Sub blue2_Change()
color2.BackColor = RGB(rED2.Value, grEEn2.Value, bLuE2.Value)

End Sub


Private Sub blue3_Change()
Color3.BackColor = RGB(rED3.Value, gREEN3.Value, bLuE3.Value)
End Sub

Private Sub Blue4_Change()
Color4.BackColor = RGB(Red4.Value, Green4.Value, Blue4.Value)
End Sub

Private Sub Blue5_Change()
Color5.BackColor = RGB(RED5.Value, green5.Value, Blue5.Value)
End Sub

Private Sub BLUE6_Change()
COLOR6.BackColor = RGB(RED6.Value, Green6.Value, BLUE6.Value)
End Sub

Private Sub Command1_Click()
FadedText$ = FadeByColor5(color1.BackColor, color2.BackColor, Color3.BackColor, Color4.BackColor, Color5.BackColor, Text1.Text, False)
Text2 = FadedText$
Call FadePreview(Picture1, Text2.Text)
End Sub


Private Sub Command2_Click()

End Sub


Private Sub Command3_Click()
MsgBox "This program was made by DrushDman,the creator of Emberson Fx Program Design.If you have any questions email me at DrushDman@aol.com or IM me on my aol name DrushDman,or my AIM name VolcomDman", , "About"
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
FadedText$ = FadeByColor3(color1.BackColor, color2.BackColor, Color3.BackColor, Text1.Text, True)
Text2 = FadedText$
Call FadePreview(Picture1, Text2)
End Sub


Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()

End Sub

Private Sub EFX2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Form_Load()
Call MonkEFade.FormFade(Me, vbBlue, vbGreen)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub GREEN1_Change()
color1.BackColor = RGB(rEd1.Value, gReEn1.Value, bLuE1.Value)
End Sub

Private Sub green2_Change()
color2.BackColor = RGB(rED2.Value, grEEn2.Value, bLuE2.Value)

End Sub

Private Sub green3_Change()
Color3.BackColor = RGB(rED3.Value, gREEN3.Value, bLuE3.Value)
End Sub

Private Sub Green4_Change()
Color4.BackColor = RGB(Red4.Value, Green4.Value, Blue4.Value)
End Sub

Private Sub green5_Change()
Color5.BackColor = RGB(RED5.Value, green5.Value, Blue5.Value)
End Sub

Private Sub Green6_Change()
COLOR6.BackColor = RGB(RED6.Value, Green6.Value, BLUE6.Value)
End Sub

Private Sub Min_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HIGH.Visible = True
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub RED1_Change()
color1.BackColor = RGB(rEd1.Value, gReEn1.Value, bLuE1.Value)
End Sub


Private Sub red2_Change()
color2.BackColor = RGB(rED2.Value, grEEn2.Value, bLuE2.Value)
End Sub


Private Sub red3_Change()
Color3.BackColor = RGB(rED3.Value, gREEN3.Value, bLuE3.Value)

End Sub


Private Sub VScroll1_Change()
Color4.BackColor = RGB(rED3.Value, gREEN3.Value, bLuE3.Value)
End Sub

Private Sub VScroll2_Change()
Color4.BackColor = RGB(rED3.Value, gREEN3.Value, bLuE3.Value)
End Sub

Private Sub VScroll3_Change()
Color4.BackColor = RGB(RED5.Value, green5.Value, Blue5.Value)
End Sub

Private Sub VScroll4_Change()
Color5.BackColor = RGB(rED3.Value, gREEN3.Value, bLuE3.Value)
End Sub

Private Sub VScroll5_Change()
Color5.BackColor = RGB(rED3.Value, gREEN3.Value, bLuE3.Value)
End Sub

Private Sub VScroll6_Change()
Color5.BackColor = RGB(rED3.Value, gREEN3.Value, bLuE3.Value)
End Sub

Private Sub Red4_Change()
Color4.BackColor = RGB(rED3.Value, gREEN3.Value, bLuE3.Value)
End Sub

Private Sub RED5_Change()
Color5.BackColor = RGB(RED5.Value, green5.Value, Blue5.Value)
End Sub

Private Sub RED6_Change()
COLOR6.BackColor = RGB(RED6.Value, Green6.Value, BLUE6.Value)
End Sub
