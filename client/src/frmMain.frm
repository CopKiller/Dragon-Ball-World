VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crystalshire"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19200
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1280
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdAttWindow 
      Caption         =   "Atualizar Janela"
      Height          =   195
      Left            =   6360
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picIntro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAttWindow_Click()
    Dim tempWindow_Count As Long
    Dim tempzOrder_Win As Long
    Dim WindowName As String
    Dim callProcedure As Long
    Dim bytes() As Byte
    
    Dim windowIndex As Long
    
    'Nome da janela para obter o �ndice da janela no processamento
    WindowName = "winNpcChat"
    'Nome da sub de cria��o da janela
    callProcedure = GetAddress(AddressOf CreateWindow_NpcChat)
    
    'Obtem o indice da janela
    windowIndex = GetWindowIndex(WindowName)
    
    ' Notifica que est� atualizando uma janela em um m�todo de processamento, para n�o acrescentar mais redimensionamento
    windowUpdated = True
    controlUpdated = True
    
    With Windows(windowIndex).Window
        'Dados tempor�rios
        tempWindow_Count = windowCount
        tempzOrder_Win = .zOrder
        
        windowCount = windowIndex
        '//M�todo elaborado para alterar uma variavel privada no modulo de processamento
        Call SetzOrder_Win(.zOrder)
        
        CallWindowProc callProcedure, 1, bytes, 0, 0
        '//M�todo elaborado para alterar uma variavel privada no modulo de processamento
        Call SetzOrder_Win(tempzOrder_Win)
        
        windowCount = tempWindow_Count
    End With
    
    ' Desativa a abordagem de evitar uma sobrecarga de redimensionamento
    windowUpdated = False
End Sub

' Form
Private Sub Form_Unload(Cancel As Integer)
    DestroyGame
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeyPresses(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not InGame Then Exit Sub

    Select Case KeyCode
        ' handles screenshot mode
    Case vbKeyF11
        If MyIndex <= 0 Then Exit Sub

        If GetPlayerAccess(MyIndex) > 0 Then
            screenshotMode = Not screenshotMode
        End If
        Exit Sub

        ' handles form
    Case vbKeyInsert
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            frmAdmin.Show
        Else
            If frmMain.BorderStyle = 0 Then
                frmMain.BorderStyle = 1
            Else
                frmMain.BorderStyle = 0
            End If
            frmMain.caption = frmMain.caption
        End If

        Exit Sub
    End Select
End Sub

Private Sub Form_DblClick()
    HandleGuiMouse EntityStates.DoubleClick
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

End Sub
