VERSION 5.00
Begin VB.Form frmEditor_MapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Weather"
      Height          =   1695
      Left            =   120
      TabIndex        =   50
      Top             =   5760
      Width           =   2055
      Begin VB.ComboBox CmbWeather 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   120
         List            =   "frmMapProperties.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   480
         Width           =   1815
      End
      Begin VB.HScrollBar scrlWeatherIntensity 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   51
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   1920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Weather Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblWeatherIntensity 
         Caption         =   "Intensity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Effects"
      Height          =   2775
      Left            =   2280
      TabIndex        =   35
      Top             =   5040
      Width           =   4215
      Begin VB.HScrollBar scrlRed 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   42
         Top             =   1800
         Value           =   255
         Width           =   1095
      End
      Begin VB.HScrollBar scrlGreen 
         Height          =   255
         Left            =   1560
         Max             =   255
         TabIndex        =   41
         Top             =   1800
         Value           =   255
         Width           =   1095
      End
      Begin VB.HScrollBar scrlBlue 
         Height          =   255
         Left            =   3000
         Max             =   255
         TabIndex        =   40
         Top             =   1800
         Value           =   255
         Width           =   1095
      End
      Begin VB.HScrollBar scrlAlpha 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   39
         Top             =   2400
         Width           =   1095
      End
      Begin VB.HScrollBar scrlFogOpacity 
         Height          =   255
         Left            =   2160
         Max             =   255
         TabIndex        =   38
         Top             =   480
         Width           =   1815
      End
      Begin VB.HScrollBar ScrlFog 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   37
         Top             =   480
         Width           =   1815
      End
      Begin VB.HScrollBar ScrlFogSpeed 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   36
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   4080
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblR 
         Caption         =   "Red: 255"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblG 
         Caption         =   "Green: 255"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   1560
         TabIndex        =   48
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblB 
         Caption         =   "Blue: 255"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3000
         TabIndex        =   47
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblA 
         Caption         =   "Alpha: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblFogOpacity 
         Caption         =   "Fog Opacity: 0"
         Height          =   255
         Left            =   2160
         TabIndex        =   45
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFog 
         Caption         =   "Fog: None"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFogSpeed 
         Caption         =   "Fog Speed: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   810
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Boss"
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   4800
      Width           =   2055
      Begin VB.HScrollBar scrlBoss 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblBoss 
         Caption         =   "Boss: None"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Music"
      Height          =   3495
      Left            =   4440
      TabIndex        =   27
      Top             =   1440
      Width           =   2055
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ListBox lstMusic 
         Height          =   2205
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   2055
      Begin VB.TextBox txtMaxX 
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   20
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   19
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   18
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Current map: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   855
      Left            =   2280
      TabIndex        =   13
      Top             =   480
      Width           =   4215
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0045
         Left            =   960
         List            =   "frmMapProperties.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moral:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Boot Map:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Boot X:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Boot Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   630
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      Height          =   3495
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   255
         Left            =   1320
         TabIndex        =   57
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Paste"
         Height          =   255
         Left            =   720
         TabIndex        =   56
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Copy"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   3120
         Width           =   615
      End
      Begin VB.ListBox lstNpcs 
         Height          =   2400
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEditor_MapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NpcIndexCopy As Long

Private Sub cmdPlay_Click()
    Stop_Music
    Play_Music lstMusic.list(lstMusic.ListIndex)
End Sub

Private Sub cmdStop_Click()
    Stop_Music
End Sub

Private Sub cmdOk_Click()
    Dim X As Long, X2 As Long
    Dim Y As Long, Y2 As Long
    Dim tempArr() As TileRec

    If Not IsNumeric(txtMaxX.text) Then txtMaxX.text = Map.MapData.maxX
    If Val(txtMaxX.text) < 1 Then txtMaxX.text = 1
    If Val(txtMaxX.text) > MAX_BYTE Then txtMaxX.text = MAX_BYTE
    If Not IsNumeric(txtMaxY.text) Then txtMaxY.text = Map.MapData.maxY
    If Val(txtMaxY.text) < 1 Then txtMaxY.text = 1
    If Val(txtMaxY.text) > MAX_BYTE Then txtMaxY.text = MAX_BYTE

    With Map.MapData
        .Name = Trim$(txtName.text)

        If lstMusic.ListIndex >= 0 Then
            .Music = lstMusic.list(lstMusic.ListIndex)
        Else
            .Music = vbNullString
        End If

        .Up = Val(txtUp.text)
        .Down = Val(txtDown.text)
        .Left = Val(txtLeft.text)
        .Right = Val(txtRight.text)
        .Moral = cmbMoral.ListIndex
        .BootMap = Val(txtBootMap.text)
        .BootX = Val(txtBootX.text)
        .BootY = Val(txtBootY.text)
        
        .Weather = CmbWeather.ListIndex
        .WeatherIntensity = scrlWeatherIntensity.Value
        
        .Fog = ScrlFog.Value
        .FogSpeed = ScrlFogSpeed.Value
        .FogOpacity = scrlFogOpacity.Value
        
        .Red = scrlRed.Value
        .Green = scrlGreen.Value
        .Blue = scrlBlue.Value
        .alpha = scrlAlpha.Value
        
        .BossNpc = scrlBoss.Value
        ' set the data before changing it
        tempArr = Map.TileData.Tile
        X2 = Map.MapData.maxX
        Y2 = Map.MapData.maxY
        ' change the data
        .maxX = Val(txtMaxX.text)
        .maxY = Val(txtMaxY.text)

        If X2 > .maxX Then X2 = .maxX
        If Y2 > .maxY Then Y2 = .maxY
        ' redim the map size
        ReDim Map.TileData.Tile(0 To .maxX, 0 To .maxY)

        For X = 0 To X2
            For Y = 0 To Y2
                Map.TileData.Tile(X, Y) = tempArr(X, Y)
            Next
        Next

    End With

    ' cache the shit
    initAutotiles
    Unload frmEditor_MapProperties
    ClearTempTile
End Sub

Private Sub cmdCancel_Click()
    Unload frmEditor_MapProperties
End Sub

Private Sub Command1_Click()
    If Not lstNpcs.ListCount > 0 Then Exit Sub

    NpcIndexCopy = Map.MapData.Npc(lstNpcs.ListIndex + 1)
End Sub

Private Sub Command2_Click()
    Dim tmpIndex As Long, X As Long
    
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    Map.MapData.Npc(lstNpcs.ListIndex + 1) = NpcIndexCopy
    
    ' re-load the list
    tmpIndex = lstNpcs.ListIndex
    If ((lstNpcs.ListIndex + 1)) < MAX_MAP_NPCS Then tmpIndex = tmpIndex + 1
    
    lstNpcs.Clear

    For X = 1 To MAX_MAP_NPCS

        If Map.MapData.Npc(X) > 0 Then
            lstNpcs.AddItem X & ": " & Trim$(Npc(Map.MapData.Npc(X)).Name)
        Else
            lstNpcs.AddItem X & ": No NPC"
        End If

    Next

    lstNpcs.ListIndex = tmpIndex
End Sub

Private Sub Command3_Click()
    Dim X As Long, tmpIndex As Long

    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    Map.MapData.Npc(lstNpcs.ListIndex + 1) = 0
    
    ' re-load the list
    tmpIndex = lstNpcs.ListIndex
    
    If ((lstNpcs.ListIndex + 1)) < MAX_MAP_NPCS Then tmpIndex = tmpIndex + 1
    
    lstNpcs.Clear

    For X = 1 To MAX_MAP_NPCS

        If Map.MapData.Npc(X) > 0 Then
            lstNpcs.AddItem X & ": " & Trim$(Npc(Map.MapData.Npc(X)).Name)
        Else
            lstNpcs.AddItem X & ": No NPC"
        End If

    Next

    lstNpcs.ListIndex = tmpIndex
End Sub

Private Sub Form_Load()
    scrlRed.min = 0
    scrlGreen.min = 0
    scrlBlue.min = 0
    scrlAlpha.min = 0
    scrlRed.max = 255
    scrlGreen.max = 255
    scrlBlue.max = 255
    scrlAlpha.max = 255
End Sub

Private Sub lstNpcs_Click()
    Dim tmpString() As String
    Dim NpcNum As Long

    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    ' set the combo box properly
    tmpString = Split(lstNpcs.list(lstNpcs.ListIndex))
    NpcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
    cmbNpc.ListIndex = Map.MapData.Npc(NpcNum)
End Sub

Private Sub cmbNpc_Click()
    Dim tmpString() As String
    Dim NpcNum As Long
    Dim X As Long, tmpIndex As Long

    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    ' set the combo box properly
    tmpString = Split(cmbNpc.list(cmbNpc.ListIndex))

    ' make sure it's not a clear
    If Not cmbNpc.list(cmbNpc.ListIndex) = "No NPC" Then
        NpcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        Map.MapData.Npc(lstNpcs.ListIndex + 1) = NpcNum
    Else
        Map.MapData.Npc(lstNpcs.ListIndex + 1) = 0
    End If

    ' re-load the list
    tmpIndex = lstNpcs.ListIndex
    lstNpcs.Clear

    For X = 1 To MAX_MAP_NPCS

        If Map.MapData.Npc(X) > 0 Then
            lstNpcs.AddItem X & ": " & Trim$(Npc(Map.MapData.Npc(X)).Name)
        Else
            lstNpcs.AddItem X & ": No NPC"
        End If

    Next

    lstNpcs.ListIndex = tmpIndex
End Sub

Private Sub scrlAlpha_Change()
    lblA.caption = "Alpha: " & scrlAlpha.Value
End Sub

Private Sub scrlBlue_Change()
    lblB.caption = "Blue: " & scrlBlue.Value
End Sub

Private Sub scrlBoss_Change()

    If scrlBoss.Value > 0 Then
        lblBoss.caption = "Boss Npc: " & Trim$(Npc(Map.MapData.Npc(scrlBoss.Value)).Name)
    Else
        lblBoss.caption = "Boss Npc: None"
    End If

End Sub

Private Sub ScrlFog_Change()
    If ScrlFog.Value = 0 Then
        lblFog.caption = "Fog: None."
    Else
        lblFog.caption = "Fog: " & ScrlFog.Value
    End If
End Sub

Private Sub scrlFogOpacity_Change()
    lblFogOpacity.caption = "Fog Opacity: " & scrlFogOpacity.Value
End Sub

Private Sub ScrlFogSpeed_Change()
    lblFogSpeed.caption = "Fog Speed: " & ScrlFogSpeed.Value
End Sub

Private Sub scrlGreen_Change()
    lblG.caption = "Green: " & scrlGreen.Value
End Sub

Private Sub scrlRed_Change()
    lblR.caption = "Red: " & scrlRed.Value
End Sub

Private Sub scrlWeatherIntensity_Change()
    lblWeatherIntensity.caption = "Intensity: " & scrlWeatherIntensity.Value
End Sub
