VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InitializeTheAsync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Sub InitializeAsync(ClinetInterface As IClientCallBack) 'Pass in all the parameters to this SUB
'Set universal property telling me who I need to call back
  Set MyClientInterface = ClinetInterface
  InitializeTimerAndProcess
End Sub
