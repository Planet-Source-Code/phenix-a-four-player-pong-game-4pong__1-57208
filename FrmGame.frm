VERSION 5.00
Begin VB.Form FrmGame 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "4Pong"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   13560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   9735
      Left            =   9960
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label LblInformation 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmGame.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   8280
         Width           =   3135
      End
      Begin VB.Label LblInformation 
         BackStyle       =   0  'Transparent
         Caption         =   "Please double click on a coloured name above to set who you will play as (up to 4 human players)."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   7320
         Width           =   3135
      End
      Begin VB.Label LblInformation 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmGame.frx":009B
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   5880
         Width           =   3135
      End
      Begin VB.Label LblScore 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   8
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label LblScore 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   7
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label LblScore 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label LblScore 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label LblPlayer 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 4"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label LblPlayer 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 3"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label LblPlayer 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label LblPlayer 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Timer TimGame 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3600
      Top             =   3000
   End
   Begin VB.Shape Ball 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      Height          =   375
      Index           =   2
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Ball 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      Height          =   375
      Index           =   1
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Ball 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      Height          =   375
      Index           =   0
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape CenterCircle 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   375
   End
   Begin VB.Line Border 
      BorderColor     =   &H0000FF00&
      Index           =   5
      X1              =   2880
      X2              =   4080
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Border 
      BorderColor     =   &H0000FF00&
      Index           =   4
      X1              =   2880
      X2              =   4080
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Border 
      BorderColor     =   &H0000FF00&
      Index           =   3
      X1              =   2880
      X2              =   4080
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Border 
      BorderColor     =   &H0000FF00&
      Index           =   2
      X1              =   2880
      X2              =   4080
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Border 
      BorderColor     =   &H0000FF00&
      Index           =   1
      X1              =   2880
      X2              =   4080
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Border 
      BorderColor     =   &H0000FF00&
      Index           =   0
      X1              =   2880
      X2              =   4080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Shape Pad 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   3960
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Pad 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   3600
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Pad 
      BorderColor     =   &H000080FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   3240
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Pad 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   2880
      Top             =   1560
      Width           =   255
   End
End
Attribute VB_Name = "FrmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''
'' 4Pong By Dominic 'Phenix' Black ''
''                                 ''
'' Feel free to use any code in    ''
'' your own programs, just give me ''
'' credits for the bits from this  ''
'' program.                        ''
''                                 ''
'' I hope my code is readable! :P  ''
''                                 ''
'' Any questions please email me;  ''
'' phenix@filesnetwork.com         ''
'''''''''''''''''''''''''''''''''''''


' Please comment on my code at: http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=57208&lngWId=1
' Thank you!


' Constant Settings
Const Pad_Width As Integer = 100
Const Pad_Height As Integer = 1250
Const Pad_Padding As Integer = 100
Const Pad_Speed As Integer = 75
Const Border_Differance As Integer = 50
Const Centre_Circle As Integer = 2000
Const Min_Speed As Integer = 5
Const Max_Speed As Integer = 30
Const Rebound_Effect As Single = 0.1
Const Form_Width As Integer = 9960 'Have to use consts because otherwise it is wrong...
Const Form_Height As Integer = 9960 ' ...using frmgame.width or frmgame.height
Const debugOn As Boolean = False

' Variables
Dim Computer(0 To 3) As Boolean
Dim Difficultly As Integer '1 = hard, 2 = median, 3 = easy
Dim hMom(0 To 2) As Single
Dim vMom(0 To 2) As Single
Dim MovePad(0 To 3, 1 To 2) As Integer
Dim Score(0 To 3) As Integer
Dim FunkyGraphics As Boolean
Dim Place As Integer
Dim Max_Score As Integer
Dim Game_Started As Boolean


Private Sub Reset_Game(Optional ByPass As Boolean = True)
    Dim DoReset As Boolean
    Dim i As Integer
        
    ' Check with the user if we're not intructed to bypass the msgbox
    If ByPass = True Then
        DoReset = True
    Else
        'If he really wants to reset let him, else continue
        If MsgBox("Are you sure you want to restart the game?", vbQuestion Or vbYesNo, "4Pong") = vbYes Then
            DoReset = True
            For i = 0 To 3
                Score(i) = 0
                LblScore(i).Caption = "0"
                Pad(i).Visible = True
            Next
            Place = 4
            
        Else
            DoReset = False
        End If
    End If
    
    'Reset window
    If DoReset = True Then
        'Pad 0 - Left
        With Pad(0)
            .Width = Pad_Width
            .Height = Pad_Height
            .Left = Pad_Padding
            .Top = (Form_Height / 2) - (Pad_Height / 2)
        End With
        
        'Pad 1 - Top
        With Pad(1)
            .Width = Pad_Height
            .Height = Pad_Width
            .Left = (Form_Width / 2) - (Pad_Height / 2)
            .Top = Pad_Padding
        End With
        
        'Pad 2 - Right
        With Pad(2)
            .Width = Pad_Width
            .Height = Pad_Height
            .Left = Form_Width - (Pad_Padding + Pad_Width)
            .Top = (Form_Height / 2) - (Pad_Height / 2)
        End With
        
        'Pad 3 - Bottom
        With Pad(3)
            .Width = Pad_Height
            .Height = Pad_Width
            .Left = (Form_Width / 2) - (Pad_Height / 2)
            .Top = Form_Height - (Pad_Padding + Pad_Width)
        End With
        
        'Border - Left
        With Border(0)
            .X1 = Pad_Padding - Border_Differance
            .X2 = Pad_Padding - Border_Differance
            .Y1 = Pad_Padding - Border_Differance
            .Y2 = Form_Height - (Pad_Padding - Border_Differance)
        End With
        
        'Border - Top
        With Border(1)
            .X1 = Pad_Padding - Border_Differance
            .X2 = Form_Width - (Pad_Padding - Border_Differance)
            .Y1 = Pad_Padding - Border_Differance
            .Y2 = Pad_Padding - Border_Differance
        End With
        
        'Border - Right
        With Border(2)
            .X1 = Form_Width - (Pad_Padding - Border_Differance)
            .X2 = Form_Width - (Pad_Padding - Border_Differance)
            .Y1 = Pad_Padding - Border_Differance
            .Y2 = Form_Height - (Pad_Padding - Border_Differance)
        End With
        
        'Border - Bottom
        With Border(3)
            .X1 = Form_Width - (Pad_Padding - Border_Differance)
            .X2 = Pad_Padding - Border_Differance
            .Y1 = Form_Height - (Pad_Padding - Border_Differance)
            .Y2 = Form_Height - (Pad_Padding - Border_Differance)
        End With
        
        'Border - Dia 1
        With Border(4)
            .X1 = Pad_Padding - Border_Differance
            .X2 = Form_Width - (Pad_Padding - Border_Differance)
            .Y1 = Pad_Padding - Border_Differance
            .Y2 = Form_Height - (Pad_Padding - Border_Differance)
        End With
        
        'Border - Dia 2
        With Border(5)
            .X1 = Pad_Padding - Border_Differance
            .X2 = Form_Width - (Pad_Padding - Border_Differance)
            .Y1 = Form_Height - (Pad_Padding - Border_Differance)
            .Y2 = Pad_Padding - Border_Differance
        End With
        
        'Centre Circle
        With CenterCircle
            .Left = (Form_Width / 2) - (Centre_Circle / 2)
            .Top = (Form_Height / 2) - (Centre_Circle / 2)
            .Width = Centre_Circle
            .Height = Centre_Circle
        End With
            
        'Set Ball 0 up
        With Ball(0)
            .Top = (Form_Height / 2) - (.Height / 2)
            .Left = (Form_Width / 2) - (.Width / 2)
            .Visible = True
        End With
        
        'Give speed to Ball 0
        RandomSpeeds 0
        
        'Reset pad movement speeds
        For i = 0 To 3
            MovePad(i, 1) = 0
            MovePad(i, 2) = 0
        Next
        
        'Set other balls up
        Ball(1).Visible = False
        Ball(2).Visible = False
        
        'Disable Timer
        If AllComputers = False Or Game_Started <> True Then
            TimGame.Enabled = False
            Game_Started = True
            If AllComputers = True Then TimGame.Interval = 1
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyNumpad8
            MovePad(3, 2) = -Pad_Speed
        Case vbKeyNumpad9
            MovePad(3, 2) = Pad_Speed
        Case vbKeyUp
            MovePad(2, 1) = -Pad_Speed
        Case vbKeyDown
            MovePad(2, 1) = Pad_Speed
        Case vbKeyK
            MovePad(1, 2) = -Pad_Speed
        Case vbKeyL
            MovePad(1, 2) = Pad_Speed
        Case vbKeyQ
            MovePad(0, 1) = -Pad_Speed
        Case vbKeyA
            MovePad(0, 1) = Pad_Speed
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyPause
            'Pause / Resume game
            If Place <> 1 Then TimGame.Enabled = Not TimGame.Enabled
        Case vbKeyF2
            'Player wants to restart
            Reset_Game False
            
        Case vbKeyNumpad8
            MovePad(3, 2) = 0
        Case vbKeyNumpad9
            MovePad(3, 2) = 0
        Case vbKeyUp
            MovePad(2, 1) = 0
        Case vbKeyDown
            MovePad(2, 1) = 0
        Case vbKeyK
            MovePad(1, 2) = 0
        Case vbKeyL
            MovePad(1, 2) = 0
        Case vbKeyQ
            MovePad(0, 1) = 0
        Case vbKeyA
            MovePad(0, 1) = 0
        
        'Debug Keys
        Case vbKeyF12
            'Check explodeballs function
            If debugOn = True Then ExplodeBalls
            
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo errorHnd
    
    'Load window positions and other settings
    Dim i As Integer
    
    For i = 0 To 3
        Computer(i) = True
        LblPlayer(i).Caption = "Computer"
    Next
    
    Place = 4
    Reset_Game
    
Retry:
    Difficultly = InputBox("Please enter a difficulty setting." & vbNewLine & vbNewLine & _
                            "1 for Hard" & vbNewLine & "2 for Average" & vbNewLine & "3 for Easy", _
                            "4Pong Settings", "1")
    Max_Score = InputBox("Please enter a max score before a player will be removed for letting the ball of his side that ammount of times", "4Pong Settings", "10")
    
    
    Game_Started = False
    
    Exit Sub
    
errorHnd:
    MsgBox "An error was dectected on your input please make sure that you enter numbers only.", vbCritical, "4Pong: Error"
    GoTo Retry
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are you sure you wish to quit?", vbQuestion Or vbYesNo, "4Pong") = vbNo Then
        Cancel = 1
    Else
        MsgBox "Thank you for downloading my four player pong game." & vbNewLine & vbNewLine & " - Dom", vbInformation Or vbOKOnly, "4Pong - Thank you"
        End
    End If
End Sub

Private Sub LblPlayer_DblClick(Index As Integer)
    If Game_Started <> True Then
        Dim newName As String
        
        newName = InputBox("Please enter a name for player " & (Index + 1) & "." & vbNewLine & vbNewLine & "(Enter computer for a computer player.)", "4Pong")
        
        If UCase(newName) = "COMPUTER" Then
            LblPlayer(Index).Caption = "Computer"
            Computer(Index) = True
        Else
            LblPlayer(Index).Caption = newName
            Computer(Index) = False
            
            Select Case Index
                Case 0
                    MsgBox newName & "'s controls are:" & vbNewLine & vbNewLine & _
                        "Q for up" & vbNewLine & "A for down", vbInformation Or _
                        vbOKOnly, "4Pong Player Controls"
                Case 1
                    MsgBox newName & "'s controls are:" & vbNewLine & vbNewLine & _
                        "< for left" & vbNewLine & "> for right", vbInformation Or _
                        vbOKOnly, "4Pong Player Controls"
                Case 2
                    MsgBox newName & "'s controls are:" & vbNewLine & vbNewLine & _
                        "Up Key for up" & vbNewLine & "Down Key for down", vbInformation Or _
                        vbOKOnly, "4Pong Player Controls"
                Case 3
                    MsgBox newName & "'s controls are:" & vbNewLine & vbNewLine & _
                        "8 (NumPad) for left" & vbNewLine & "9 (NumPad) for right", vbInformation Or _
                        vbOKOnly, "4Pong Player Controls"
            End Select
        End If
    End If
End Sub

Private Sub TimGame_Timer()
    Dim i, j As Integer
    Dim smallest, num As Integer
    
    If debugOn = True And FunkyGraphics <> False Then
        CenterCircle.Width = CenterCircle.Width - 50
        CenterCircle.Height = CenterCircle.Height - 50
        CenterCircle.Left = (Form_Width / 2) - (CenterCircle.Width / 2)
        CenterCircle.Top = (Form_Height / 2) - (CenterCircle.Height / 2)
        With Border(4)
            .X1 = Border(4).X1 - ((Border(3).X2 - Border(3).X1) / 30)
            .X2 = Border(4).X2 + ((Border(3).X2 - Border(3).X1) / 30)
        End With
        With Border(5)
            .Y1 = Border(5).Y1 - ((Border(2).Y2 - Border(2).Y1) / 30)
            .Y2 = Border(5).Y2 + ((Border(2).Y2 - Border(2).Y1) / 30)
        End With
        If CenterCircle.Height <= 500 Then FunkyGraphics = False
    ElseIf debugOn = True Then
        CenterCircle.Width = CenterCircle.Width + 50
        CenterCircle.Height = CenterCircle.Height + 50
        CenterCircle.Left = (Form_Width / 2) - (CenterCircle.Width / 2)
        CenterCircle.Top = (Form_Height / 2) - (CenterCircle.Height / 2)
        With Border(4)
            .X1 = Border(4).X1 + ((Border(3).X2 - Border(3).X1) / 30)
            .X2 = Border(4).X2 - ((Border(3).X2 - Border(3).X1) / 30)
        End With
        With Border(5)
            .Y1 = Border(5).Y1 + ((Border(2).Y2 - Border(2).Y1) / 30)
            .Y2 = Border(5).Y2 - ((Border(2).Y2 - Border(2).Y1) / 30)
        End With
        If CenterCircle.Height >= 2000 Then FunkyGraphics = True
    End If
    
    'Loop though players
    For i = 0 To 3
        If Computer(i) = False Then
            
        Else
            If i = 0 Then
                num = 32000
                For j = 0 To 2
                    If Ball(j).Visible = True Then
                        If num > Ball(j).Left - Pad(i).Left And (Ball(j).Left - Pad(i).Left) > 0 Then
                            smallest = j
                            num = Ball(j).Left - Pad(i).Left
                        End If
                    End If
                Next
                
                If Ball(smallest).Top < Pad(i).Top Then
                    MovePad(i, 1) = -(Pad_Speed / Difficultly)
                ElseIf Ball(smallest).Top > Pad(i).Top + Pad_Height Then
                    MovePad(i, 1) = (Pad_Speed / Difficultly)
                Else
                    MovePad(i, 1) = 0
                End If
            ElseIf i = 1 Then
                num = 32000
                For j = 0 To 2
                    If Ball(j).Visible = True Then
                        If num > Ball(j).Top - Pad(i).Top And (Ball(j).Top - Pad(i).Top) > 0 Then
                            smallest = j
                            num = Ball(j).Top - Pad(i).Top
                        End If
                    End If
                Next
                
                If Ball(smallest).Left < Pad(i).Left Then
                    MovePad(i, 2) = -(Pad_Speed / Difficultly)
                ElseIf Ball(smallest).Left > Pad(i).Left + Pad_Width Then
                    MovePad(i, 2) = (Pad_Speed / Difficultly)
                Else
                    MovePad(i, 2) = 0
                End If
            ElseIf i = 2 Then
                num = 32000
                For j = 0 To 2
                    If Ball(j).Visible = True Then
                        If num > Pad(i).Left - Ball(j).Left And (Pad(i).Left - Ball(j).Left) > 0 Then
                            smallest = j
                            num = Pad(i).Left - Ball(j).Left
                        End If
                    End If
                Next
                
                If Ball(smallest).Top < Pad(i).Top Then
                    MovePad(i, 1) = -(Pad_Speed / Difficultly)
                ElseIf Ball(smallest).Top > Pad(i).Top + Pad_Height Then
                    MovePad(i, 1) = (Pad_Speed / Difficultly)
                Else
                    MovePad(i, 1) = 0
                End If
            ElseIf i = 3 Then
                num = 32000
                For j = 0 To 2
                    If Ball(j).Visible = True Then
                        If num > Pad(i).Top - Ball(j).Top And (Pad(i).Top - Ball(j).Top) > 0 Then
                            smallest = j
                            num = Pad(i).Top - Ball(j).Top
                        End If
                    End If
                Next
                
                If Ball(smallest).Left < Pad(i).Left Then
                    MovePad(i, 2) = -(Pad_Speed / Difficultly)
                ElseIf Ball(smallest).Left > Pad(i).Left + Pad_Width Then
                    MovePad(i, 2) = (Pad_Speed / Difficultly)
                Else
                    MovePad(i, 2) = 0
                End If
            End If
        End If
    Next
    
    For i = 0 To 3
        Pad(i).Left = Pad(i).Left + MovePad(i, 2)
        Pad(i).Top = Pad(i).Top + MovePad(i, 1)
    Next
    
    If Pad(0).Top < Pad_Padding Then Pad(0).Top = Pad_Padding
    If Pad(2).Top < Pad_Padding Then Pad(2).Top = Pad_Padding
    If Pad(0).Top > Form_Height - (Pad_Padding + Pad_Width) Then Pad(0).Top = Form_Height - (Pad_Padding + Pad_Width)
    If Pad(2).Top > Form_Height - (Pad_Padding + Pad_Width) Then Pad(2).Top = Form_Height - (Pad_Padding + Pad_Width)
    If Pad(0).Left < Pad_Padding Then Pad(0).Left = Pad_Padding
    If Pad(2).Left > Form_Width - (Pad_Padding + Pad_Width) Then Pad(2).Left = Form_Width - (Pad_Padding + Pad_Width)
    
    If Pad(1).Top < Pad_Padding Then Pad(1).Top = Pad_Padding
    If Pad(3).Top < Form_Height - (Pad_Padding + Pad_Width) Then Pad(3).Top = Form_Height - (Pad_Padding + Pad_Width)
    If Pad(1).Left < Pad_Padding Then Pad(1).Left = Pad_Padding
    If Pad(1).Left > Form_Width - (Pad_Padding + Pad_Height) Then Pad(1).Left = Form_Width - (Pad_Padding + Pad_Height)
    If Pad(3).Left < Pad_Padding Then Pad(3).Left = Pad_Padding
    If Pad(3).Left > Form_Width - (Pad_Padding + Pad_Height) Then Pad(3).Left = Form_Width - (Pad_Padding + Pad_Height)
    
    'Loop though balls
    For i = 0 To 2
        ' If the ball is visible then
        If Ball(i).Visible = True Then
            'Move this ball
            Ball(i).Left = Ball(i).Left + hMom(i)
            Ball(i).Top = Ball(i).Top + vMom(i)
    
            If (Ball(i).Left <= Pad(0).Left + Pad(0).Width And _
            Ball(i).Top >= Pad(0).Top - Ball(i).Height And _
            Ball(i).Top <= Pad(0).Top + Pad(0).Height And _
            Ball(i).Left >= Pad(0).Left - Ball(i).Width) And Pad(0).Visible = True Then
                Ball(i).Left = Pad(0).Left + Pad(0).Width
                tmp = ((Pad(0).Top + (Pad(0).Height / 2)) - (Ball(i).Top + (Ball(i).Height / 2))) * Rebound_Effect
                hMom(i) = -hMom(i)
                vMom(i) = vMom(i) - tmp
            End If
            
            If Ball(i).Left >= Pad(2).Left - Ball(i).Width And _
            Ball(i).Top >= Pad(2).Top - Ball(i).Height And _
            Ball(i).Top <= Pad(2).Top + Pad(2).Height And _
            Ball(i).Left <= Pad(2).Left + Pad(2).Width + Ball(i).Width And Pad(2).Visible = True Then
                Ball(i).Left = Pad(2).Left - Ball(i).Width
                tmp = ((Pad(2).Top + (Pad(2).Height / 2)) - (Ball(i).Top + (Ball(i).Height / 2))) * Rebound_Effect
                hMom(i) = -hMom(i)
                vMom(i) = vMom(i) - tmp
            End If
            
            If Ball(i).Top <= Pad(1).Top + Pad(1).Height And _
            Ball(i).Left >= Pad(1).Left - Ball(i).Width And _
            Ball(i).Left <= Pad(1).Left + Pad(1).Width And _
            Ball(i).Top >= Pad(1).Top - Ball(i).Height And Pad(1).Visible = True Then
                Ball(i).Top = Pad(1).Top + Pad(1).Height
                tmp = ((Pad(1).Left + (Pad(1).Width / 2)) - (Ball(i).Left + (Ball(i).Width / 2))) * Rebound_Effect
                vMom(i) = -vMom(i)
                hMom(i) = hMom(i) - tmp
            End If
            
            If Ball(i).Top >= Pad(3).Top - Ball(i).Height And _
            Ball(i).Left >= Pad(3).Left - Ball(i).Width And _
            Ball(i).Left <= Pad(3).Left + Pad(3).Width And _
            Ball(i).Top <= Pad(3).Top + Pad(3).Height + Ball(i).Height And Pad(3).Visible = True Then
                Ball(i).Top = Pad(3).Top - Ball(i).Height
                tmp = ((Pad(3).Left + (Pad(3).Width / 2)) - (Ball(i).Left + (Ball(i).Width / 2))) * Rebound_Effect
                vMom(i) = -vMom(i)
                hMom(i) = hMom(i) - tmp
            End If
            
            If Pad(0).Visible = False And Ball(i).Left <= Border(0).X1 Then
                hMom(i) = -hMom(i)
            End If
            
            If Pad(2).Visible = False And Ball(i).Left >= Border(2).X1 - Ball(i).Width Then
                hMom(i) = -hMom(i)
            End If
            
            If Pad(1).Visible = False And Ball(i).Top <= Border(1).Y1 Then
                vMom(i) = -vMom(i)
            End If
            
            If Pad(3).Visible = False And Ball(i).Top >= Border(3).Y1 - Ball(i).Height Then
                vMom(i) = -vMom(i)
            End If
            
            'If it goes off an edge add a score to somebody
            If Ball(i).Left <= Border(0).X1 And Pad(0).Visible = True Then AddScore 0, i
            If Ball(i).Left >= Border(2).X1 And Pad(2).Visible = True Then AddScore 2, i
            If Ball(i).Top <= Border(1).Y1 And Pad(1).Visible = True Then AddScore 1, i
            If Ball(i).Top >= Border(3).Y1 And Pad(3).Visible = True Then AddScore 3, i
        End If
        
        If Rnd() > 0.95 Then
            If Rnd() > 0.95 Then
                If Rnd() > 0.95 Then
                    ExplodeBalls
                End If
            End If
        End If
    
    Next
    
End Sub

Private Sub AddScore(Player As Integer, ByVal ballNum As Integer)
    Dim i As Integer
    
    'Add a "bad" score to the player's who side it went off
    Score(Player) = Score(Player) + 1
    LblScore(Player).Caption = "-" & Score(Player)
    
    'Make the ball invisible (incase more are on the field)
    Ball(ballNum).Visible = False
    
    If Score(Player) >= Max_Score Then
        Pad(Player).Visible = False
        Select Case Place
            Case 2
                For i = 0 To 3
                    If Pad(i).Visible = True Then Exit For
                Next
                LblScore(i).Caption = "First Place"
                LblScore(Player).Caption = "Second Place"
                Place = 1
                TimGame.Enabled = False
            Case 3
                LblScore(Player).Caption = "Third Place"
                Place = 2
            Case 4
                LblScore(Player).Caption = "Last Place"
                Place = 3
        End Select
    End If
        
    'If no balls are left reset game window
    If CountBalls = 0 Then Reset_Game
End Sub

Private Sub ExplodeBalls()
    Dim i As Integer
    
    If CountBalls = 1 Then
        If Ball(0).Visible = False Then
            For i = 1 To 2
                If Ball(i).Visible = True Then
                    Ball(0).Left = Ball(i).Left
                    Ball(0).Top = Ball(i).Top
                    hMom(0) = hMom(i)
                    vMom(0) = vMom(i)
                End If
            Next
        End If
        
        'Set each ball into the same position and make it visible
        For i = 1 To 2
            Ball(i).Left = Ball(0).Left
            Ball(i).Top = Ball(0).Top
            Ball(i).Visible = True
        Next
        
        Ball(0).Visible = True
        
        'Give all balls new speeds
        For i = 1 To 2
            RandomSpeeds i
        Next
    End If
End Sub

Private Sub RandomSpeeds(ballNum As Integer)
    Randomize
    
    hMom(ballNum) = (Int(Rnd() * Max_Speed) - (Max_Speed / 2)) * 2
    vMom(ballNum) = (Int(Rnd() * Max_Speed) - (Max_Speed / 2)) * 2
    
    If hMom(ballNum) < 0 Then
        hMom(ballNum) = hMom(ballNum) - Min_Speed
    Else
        hMom(ballNum) = hMom(ballNum) + Min_Speed
    End If

    If vMom(ballNum) < 0 Then
        vMom(ballNum) = vMom(ballNum) - Min_Speed
    Else
        vMom(ballNum) = vMom(ballNum) + Min_Speed
    End If
End Sub

Private Function CountBalls() As Integer
    Dim i, count As Integer
    
    'Check how many balls are left
    count = 0
    For i = 0 To 2
        If Ball(i).Visible = True Then count = count + 1
    Next
    
    'Return Value
    CountBalls = count
End Function

Private Function AllComputers() As Boolean
    Dim i As Integer
    
    For i = 0 To 3
        If Computer(i) <> True And Pad(i).Visible = True Then
            AllComputers = False
            Exit Function
        End If
    Next
    
    AllComputers = True
End Function
