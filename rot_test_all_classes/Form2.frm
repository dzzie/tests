VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "SubObj of myClass"
      Height          =   510
      Left            =   1350
      TabIndex        =   4
      Top             =   2655
      Width           =   1995
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Direct Access to class"
      Height          =   510
      Left            =   1350
      TabIndex        =   3
      Top             =   1980
      Width           =   1995
   End
   Begin VB.CommandButton Command3 
      Caption         =   "access class through form pub var"
      Height          =   510
      Left            =   1350
      TabIndex        =   2
      Top             =   1350
      Width           =   2040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   510
      Left            =   2745
      TabIndex        =   1
      Top             =   450
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   630
      TabIndex        =   0
      Top             =   495
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
 
    Private m_oForm1 As Object
    Private m_PatchedClass As Object
    
Private Sub Command3_Click()
    On Error Resume Next
    
    'if not made public error is:
        'A property or method call cannot include a reference to a private object, either as an argument or as a return value
    
    m_oForm1.myClass.alert "test"
    
    If Err.Number <> 0 Then
        MsgBox "calling m_oForm1.myClass.alert failed: " & Err.Description, vbExclamation
    End If
    
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    
    If m_PatchedClass Is Nothing Then
        Set m_PatchedClass = GetObject("MySpecialProject.myClass")
        If m_PatchedClass Is Nothing Then
            MsgBox "Still can not GetObject(MySpecialProject.myClass)"
            Exit Sub
        End If
    End If
    
    m_PatchedClass.alert "test"
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
    
End Sub

Private Sub Command5_Click()
    On Error GoTo hell
    MsgBox m_PatchedClass.obj2.getVal("final test")
    Exit Sub
hell:    MsgBox Err.Description
End Sub

    Private Sub Form_Load()
        On Error Resume Next
        Dim msg As String
        
        Set m_oForm1 = GetObject("MySpecialProject.Form1")
        Set m_PatchedClass = GetObject("MySpecialProject.myClass")
        
        If Err.Number <> 0 Then
        
            msg = "m_oForm1 is nothing? = " & (m_oForm1 Is Nothing) & vbCrLf & _
                  "m_PatchedClass is nothing = " & (m_PatchedClass Is Nothing) & vbCrLf & _
                   "Error: " & Err.Description
                   
            MsgBox msg
            
        End If
        
    End Sub
 
    Private Sub Command1_Click()
        m_oForm1.List1.AddItem "Test " & Now
    End Sub
 
    Private Sub Command2_Click()
        m_oForm1.List1.RemoveItem 0
    End Sub
