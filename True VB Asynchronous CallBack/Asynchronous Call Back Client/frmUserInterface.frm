VERSION 5.00
Begin VB.Form frmUserInterface 
   Caption         =   "Client Interface"
   ClientHeight    =   1008
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   1008
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLongProcess 
      Caption         =   "Do Something that Takes A Long Time"
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3492
   End
End
Attribute VB_Name = "frmUserInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This project is currently being used as a template for me.
'I meant this to be very basic (pardon the almost pun).
'If you don't like it sorry (rewrite it and teach me a better way).
'If you want to give me Praise and offer me a position in the iLLuMNaTi
'call me at firstmorningstar@yahoo.com.

'This is the Magic this is how the client and server know about each other
 Implements IClientCallBack
 
Private MyAsyncWorker As CallBackServer.InitializeTheAsync

Private Sub cmdLongProcess_Click()
    Set MyAsyncWorker = New CallBackServer.InitializeTheAsync
   
   'Send an instance of yourself to the server so it knows who you are
    MyAsyncWorker.InitializeAsync Me
    
   'You also do not need to do any cleanup (ex. Set MyAsyncWorker = Nothing)
   'it is taken care of by the server
End Sub

Private Sub IClientCallBack_HeardSomething(strMessage As String)
'This is where you recieve the message from the server
'One thing I have done in the past is format the messages in
'XML and done a "select" structure in order to determine what
'Functions or Subs to call
    MsgBox strMessage
End Sub
