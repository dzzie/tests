Attribute VB_Name = "Module1"
Option Explicit

'this has to be in a module because you need to get the
'address of the function with addressof operator

Function MyCallBackFunction(ByVal myIntArg As Long, ByVal myIntArg2 As Long) As Long
    
    MsgBox "In my CallBack arg1 = " & myIntArg & " arg2=" & myIntArg2, vbInformation
    
    MyCallBackFunction = myIntArg + myIntArg2
    
End Function
