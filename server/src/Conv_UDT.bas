Attribute VB_Name = "Conv_UDT"
Option Explicit

Public Conversation(1 To MAX_CONVS) As ConvWrapperRec
Public EmptyConv As ConvWrapperRec

Private Const CONV_LENGTH As Integer = 500
Private Const REPONSE_LENGTH As Integer = 100

Private Type ConvRec
    Talk As String * CONV_LENGTH
    rText(1 To 4) As String * REPONSE_LENGTH
    rTarget(1 To 4) As Long
    EventType As Long
    EventNum As Long
End Type

Private Type ConvWrapperRec
    Name As String * NAME_LENGTH
    chatCount As Long
    Conv() As ConvRec
End Type
