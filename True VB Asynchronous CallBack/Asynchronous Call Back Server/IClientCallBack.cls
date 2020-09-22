VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IClientCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Prototype the interface that the clients will use
'This is were you change the code to pass an object of some sort
'ex(CHANGE "strMessage As String" TO "rs As ADODB.RecordSet")
 Public Sub HeardSomething(strMessage As String)
 
 End Sub
