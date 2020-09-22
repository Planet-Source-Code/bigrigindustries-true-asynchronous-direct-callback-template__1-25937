Attribute VB_Name = "TimerInterface"
'   Why this is in a .bas module
'"AddressOf APITimerFireEvent"
'I hope that explains it.

'Used for pausing for Five seconds (The Long Process)
 Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
'This is the name of the client application we need to call back
 Public MyClientInterface As IClientCallBack
 
'This is used to kill the Async process
 Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
 Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'This is the timer APIs that will start a new thread
 Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
 Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
 
 Private TimerIdentifier As Long

Public Sub InitializeTimerAndProcess()
  
 'Start setup for the timer
  TimerIdentifier = SetTimer(0, 0, 1, AddressOf APITimerFireEvent)
  
End Sub
Private Sub APITimerFireEvent() 'The thread has called to be created
  
  'Kill the timer so that the thread is only created once
   KillTimer 0, TimerIdentifier
  
  'Call the sub that will do the processing on the sepreate thread
   DoTheLongProcess
   
End Sub

Private Sub DoTheLongProcess()
      Sleep 5000
      MyClientInterface.HeardSomething "I just did something that took five seconds to do and then called you back."
      
     'CleanUp
      TerminateProcess GetCurrentProcess, 0
End Sub


