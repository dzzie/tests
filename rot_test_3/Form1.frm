VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   135
      TabIndex        =   1
      Top             =   1755
      Width           =   4245
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public myClass As Class1

Private rot As New CRemotlyScriptable

Private Sub Form_Load()
    
    Dim cnt As Long
    Dim b As Boolean
    
    If App.PrevInstance Then
        MsgBox "prev instance"
        End
    End If
    
    'any new classes created will after this will use new flags and be accessible..
    'note if you run this in the IDE, the patch routines will only run once and stay in effect even after new debug session..
    cnt = rot.makeAllPublic(Me)
    List1.AddItem cnt & " classes patched"
    Set myClass = New Class1
    
    'if not using above then
    'Set myClass = New Class1
    'rot.makePublic myClass
    'Set myClass = New Class1 'now we need a new instance created for flags to take effect
    'note any child classes inside this one will not be accessible have to set all public
    
    b = rot.RegisterObj(Me, "MySpecialProject.Form1")
    List1.AddItem "MySpecialProject.Form1 = " & b
    
    b = rot.RegisterObj(myClass, "MySpecialProject.myClass")
    List1.AddItem "MySpecialProject.myClass = " & b
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rot = Nothing 'cleanup - required to release refs or rot will keep project alive..
End Sub
