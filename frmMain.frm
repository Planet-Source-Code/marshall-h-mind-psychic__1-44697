VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mind Reader"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAni 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4965
      Top             =   3960
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C0C000&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label lblWow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Can't believe it? Press Go again!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   1500
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0442
      ForeColor       =   &H00FFFF00&
      Height          =   945
      Left            =   90
      TabIndex        =   2
      Top             =   3165
      Width           =   5280
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   72
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2025
      Left            =   120
      TabIndex        =   1
      Top             =   705
      Width           =   2430
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   2430
      Left            =   75
      Shape           =   3  'Circle
      Top             =   315
      Width           =   2520
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404000&
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   1890
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Number(99) As String

Dim MultipleOfNineChar As String
Dim Wondering As Boolean

Dim AniCount As Integer

Private Sub cmdGo_Click()
    If Not Wondering Then
        tmrAni.Enabled = True
        lblHelp.Visible = False
        Cls
    Else
        Form_Load
        lblResult.Left = 9
        Shape1.Left = 5
        Shape2.Left = 26
        lblResult.Caption = ""
        cmdGo.Caption = "Go"
        lblHelp.Visible = True
        lblWow.Visible = False
        Wondering = False
    End If
End Sub

Private Sub Form_Load()
    'go through all possible numbers, and if a
    'number is a multiple of 9, set it to a certain char
    Randomize
    b = Rnd * 40 + 1
    MultipleOfNineChar = LCase(Chr(b + 65))

    For i = 0 To 99
        If i = 9 Or i = 18 Or i = 27 Or i = 36 Or i = 45 Or i = 54 Or i = 63 Or i = 72 Or i = 81 Or i = 90 Then
            Number(i) = MultipleOfNineChar
        Else
            b = Rnd * 40 + 1
            Number(i) = LCase(Chr(b + 65))
        End If
    Next
    
    Dim Tn As Integer
    For y = 0 To 19
        CurrentY = y * 12
        For x = 6 To 11
            CurrentX = x * 30 + 15
            If Tn < 100 Then
                Font.Name = "Wingdings"
                Print Number(Tn)
                Font.Name = "Arial"
                CurrentY = y * 12
                CurrentX = x * 30 - 2
                Print Tn
            End If
            CurrentY = y * 12
            Tn = Tn + 1
        Next
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub tmrAni_Timer()
    AniCount = AniCount + 1
    If AniCount > 5 Then GoTo EndAni
    lblResult.Left = lblResult.Left + 19
    Shape1.Left = Shape1.Left + 19
    Shape2.Left = Shape2.Left + 19
    Exit Sub
EndAni:
    AniCount = 0
    tmrAni.Enabled = False
    lblResult.Caption = MultipleOfNineChar
    cmdGo.Caption = "Go Again"
    lblWow.Visible = True
    Wondering = True
End Sub
