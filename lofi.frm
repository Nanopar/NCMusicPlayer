VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nanop's Chill Music Player"
   ClientHeight    =   4635
   ClientLeft      =   3450
   ClientTop       =   2655
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "lofi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5160
   Begin VB.ComboBox ambienceSelector 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   14
      Text            =   "Select Ambience"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Timer wmptimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   3720
   End
   Begin VB.HScrollBar playtime 
      Height          =   135
      Left            =   240
      MousePointer    =   10  'Up Arrow
      SmallChange     =   5
      TabIndex        =   12
      Top             =   4200
      Width           =   4575
   End
   Begin VB.ComboBox albums 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "lofi.frx":29376
      Left            =   240
      List            =   "lofi.frx":2937D
      TabIndex        =   10
      Text            =   "Select Playlist"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2625
      Left            =   4560
      Max             =   100
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      Top             =   1200
      Value           =   1
      Width           =   255
   End
   Begin VB.CheckBox shuffle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shuffle"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   480
      Top             =   4440
   End
   Begin VB.HScrollBar volume_slider 
      Height          =   255
      Left            =   960
      Max             =   100
      MousePointer    =   9  'Size W E
      SmallChange     =   25
      TabIndex        =   4
      Top             =   840
      Value           =   100
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   4440
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   3975
      TabIndex        =   2
      Top             =   240
      Width           =   3975
      Begin VB.Label lblScroll 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nothing is playing. Play a song!"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   3120
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2100
      Left            =   240
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label time_finish 
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   615
   End
   Begin WMPLibCtl.WindowsMediaPlayer RainPlayer 
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   661
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP1 
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   661
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Volume:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape2 
      Height          =   4335
      Left            =   120
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label text_pause_play 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   4320
      Shape           =   5  'Rounded Square
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2025 Nanoparticles
' This is a fucking mess.
' Since I wanted this to be a one-off app, I am NOT cleaning my code up.
' (I wrote this while simultaneously learning the VB language. Please excuse yanderedev type code..)
'
' TODO: Fix inconsistent variable names (already started here)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const LB_SETCURSEL = &H186

Dim fixed_path As String
Dim is_playing As Boolean
Dim title_playing As String
Dim actual_path As String

Dim last_song As String
Dim last_song_selected As String
Dim last_playlist As String
Dim is_simulated_click As Boolean


Dim hovering_playtime As Boolean

Private Sub load_music()
    List1.Clear
    fixed_path = actual_path
    last_playlist = albums.List(albums.ListIndex)
    
    If albums.List(albums.ListIndex) <> "Root" Then
        fixed_path = fixed_path & albums.List(albums.ListIndex) & "\"
    End If
    
    Debug.Print "changed playlist! new: "; fixed_path
    
    Dim song As String
    song = Dir(fixed_path & "*.mp3")
    
    Do While song <> ""
        List1.AddItem song
        song = Dir()
    Loop
    
    If List1.ListCount = 0 Then
        List1.AddItem ("Directory Is Empty!")
    End If
End Sub

Private Sub load_ambience()
    ambienceSelector.Clear
    ambienceSelector.AddItem "No Ambience"
    ambienceSelector.ListIndex = 0
    fixed_path = actual_path
    Dim ambence_path As String
    
    ambence_path = fixed_path & "\ambience\"
    
    Dim song As String
    song = Dir(ambence_path & "*.mp3")
    
    Do While song <> ""
        ambienceSelector.AddItem song
        song = Dir()
    Loop
End Sub

Private Sub albums_Change()
    load_music
End Sub

Private Sub albums_Click()
    If last_playlist <> albums.List(albums.ListIndex) Then
        load_music
    End If
End Sub



Private Sub ambienceSelector_Click()
    If ambienceSelector.List(ambienceSelector.ListIndex) = "No Ambience" Then
        RainPlayer.Controls.stop
        Exit Sub
    End If
    RainPlayer.URL = actual_path & "\ambience\" & ambienceSelector.List(ambienceSelector.ListIndex)
    RainPlayer.Controls.play
    RainPlayer.settings.volume = 100 - VScroll1.Value
    RainPlayer.settings.setMode "loop", True
End Sub

Private Sub Form_Load()
    Dim arg As String
    is_simulated_click = False
    hovering_playtime = False
    arg = Trim$(Command$)
    
    fixed_path = App.Path
    If Right(fixed_path, 1) <> "\" Then
        fixed_path = fixed_path & "\"
    End If

    title_playing = "Nothing is playing.."
    is_playing = False
    fixed_path = fixed_path & "music\"
    actual_path = fixed_path

    Dim folder As String
    folder = Dir(fixed_path, vbDirectory)

    Do While folder <> ""
        If (GetAttr(fixed_path & folder) And vbDirectory) = vbDirectory Then
            If folder <> "." And folder <> ".." And folder <> "ambience" Then
                albums.AddItem folder
            End If
        End If
        folder = Dir()
    Loop

    albums.ListIndex = 0
    load_music

    If List1.ListCount = 0 Then
        List1.AddItem ("Directory Is Empty!")
    End If
    WMP1.settings.volume = 100
    
    load_ambience
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WMP1.Controls.stop
End Sub

Private Sub List1_Click()
    If List1.List(List1.ListIndex) = "Directory Is Empty!" Then
        Exit Sub
    End If
    
    
    
    Dim selected_file As String
    selected_file = fixed_path & List1.List(List1.ListIndex)
    title_playing = "Now Playing: " & List1.List(List1.ListIndex)
    Debug.Print is_simulated_click

    If is_simulated_click = True Or last_song_selected <> List1.List(List1.ListIndex) Then
        WMP1.URL = selected_file
        DoEvents
        is_playing = True
        Call update_play_state
        last_song_selected = List1.List(List1.ListIndex)
        is_simulated_click = False
    End If
End Sub

Private Sub playtime_Change()
    If Me.ActiveControl = playtime Then
        WMP1.Controls.currentPosition = playtime.Value
        List1.SetFocus
    End If
End Sub

Private Sub playtime_Scroll()
    WMP1.Controls.currentPosition = playtime.Value
End Sub

Private Sub shuffle_Click()
    If shuffle.Value = 1 Then
        ShuffleList List1
    End If
End Sub

Private Sub text_pause_play_Click()
    is_playing = Not is_playing
    Call update_play_state
End Sub

Private Sub update_play_state()
    If is_playing = True Then
        text_pause_play.Caption = "Pause"
        WMP1.Controls.play
    Else
        text_pause_play.Caption = "Play"
        WMP1.Controls.pause
    End If
End Sub

Private Sub Timer1_Timer()
    playtime.Value = Int(WMP1.Controls.currentPosition)
    time.Caption = WMP1.Controls.currentPositionString

    lblScroll.Left = lblScroll.Left - 50
    lblScroll.Caption = title_playing

    If lblScroll.Left + lblScroll.Width < 0 Then
        lblScroll.Left = picContainer.Width
    End If
End Sub

Private Sub SetVolume(percent As Integer)
    If percent < 0 Then percent = 0
    If percent > 100 Then percent = 100
    WMP1.settings.volume = percent
End Sub

Private Sub volume_slider_Change()
    SetVolume volume_slider.Value
End Sub

Private Sub volume_slider_Scroll()
    SetVolume volume_slider.Value
End Sub

Private Sub WMP1_PlayStateChange(ByVal NewState As Long)
    If NewState = 8 Then
        last_song = List1.List(List1.ListIndex)
        Call play_next_track
    End If

    If NewState = 10 Then
        Timer2.Enabled = True
    End If

    If NewState = 3 Then ' Playing
        time_finish.Caption = WMP1.currentMedia.durationString
        playtime.Max = WMP1.currentMedia.duration
        wmptimer.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    Debug.Print last_song
    WMP1.Controls.play
    Timer2.Enabled = False
End Sub

Private Sub play_next_track()
    If List1.ListIndex < List1.ListCount - 1 Then
        Dim new_index As Long
        new_index = List1.ListIndex + 1

        is_simulated_click = True
        SendMessage List1.hWnd, LB_SETCURSEL, new_index, 0
        List1_Click
    Else
        is_simulated_click = True
        is_playing = False
        Call text_pause_play_Click
        loop_back_to_start
    End If
End Sub

Private Sub ShuffleList(lst As ListBox)
    Dim i As Integer, j As Integer
    Dim temp As String
    Dim count As Integer

    count = lst.ListCount
    Randomize

    For i = count - 1 To 1 Step -1
        j = Int(Rnd * (i + 1))
        temp = lst.List(i)
        lst.List(i) = lst.List(j)
        lst.List(j) = temp
    Next i
End Sub

Private Sub SelectListItemByName(lst As ListBox, name_to_select As String)
    Dim i As Integer
    For i = 0 To lst.ListCount - 1
        If LCase(lst.List(i)) = LCase(name_to_select) Then
            lst.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub loop_back_to_start()
    If shuffle.Value = 1 Then
        ShuffleList List1
    End If
    SendMessage List1.hWnd, LB_SETCURSEL, 0, 0
    List1_Click
End Sub

Private Sub VScroll1_Change()
    RainPlayer.settings.volume = 100 - VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    RainPlayer.settings.volume = 100 - VScroll1.Value
End Sub


