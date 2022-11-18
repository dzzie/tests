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

'rot code from wqweto
'https://www.vbforums.com/showthread.php?879529-project-one-workng-with-project2

Private Declare Function GetRunningObjectTable Lib "ole32" (ByVal dwReserved As Long, pResult As IUnknown) As Long
Private Declare Function CreateFileMoniker Lib "ole32" (ByVal lpszPathName As Long, pResult As IUnknown) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal newVal As Long)
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Const RWE = &H40

Private m_lCookie As Long
Private m_lCookie2 As Long

Public myClass As Class1

Function dbg(x)
    Debug.Print x
    List2.AddItem x
    'MsgBox x
End Function

Function makeAllClassesPublic(Optional obj As Object = Nothing, Optional doPatch As Boolean = False) As Long

    Dim pVTbl As Long, pObjInfo As Long, pObjTable As Long, newVal As Long, oldMemProt As Long, r As Long
    Dim pObjType As Long, patched As Long
    Dim objTable As New CObjectTable, co As CObject
    
    Const flag = &H800
    
    If obj Is Nothing Then Set obj = Me 'any object will lead us back to the full table of all objects
    
    dbg "Start Obj = " & TypeName(obj)
    
    GetMem4 ByVal ObjPtr(obj), pVTbl                ' Pointer to any internal vb objects vTable.
    GetMem4 ByVal pVTbl - 4&, pObjInfo              ' Pointer to tObjectInfo structure.
    GetMem4 ByVal pObjInfo + 4, pObjTable           ' Pointer to tObjectTable structure.

    objTable.LoadFromMem pObjTable
    
    dbg objTable.Objects.Count & " objects loaded from table"
    
    For Each co In objTable.Objects
        
        dbg "Obj " & co.name & ".Type = " & Hex(co.ObjectType) & " isClass = " & co.isClass()
        
        If co.isClass() And doPatch Then
            
            'Debug.Print co.dump & vbCrLf

            If (co.ObjectType And flag) = 0 Then
                
                pObjType = co.offsetOf_ObjType()
                newVal = co.ObjectType Xor flag
                
                If VirtualProtect(ByVal pObjType, 4, RWE, oldMemProt) <> 0 Then
                    
                    dbg vbTab & "Patching to: " & Hex(newVal)
                    PutMem4 ByVal pObjType, newVal
    
                    'GetMem4 ByVal pObjType, newVal
                    'dbg "Sanity Check: " & Hex(newVal)
                    
                    patched = patched + 1
                    VirtualProtect ByVal pObjType, 4, oldMemProt, r
                Else
                    dbg "virt prot failed"
                End If
                
            Else
                dbg "Cant patch flag already set?"
            End If
            
         End If
         
    Next

    makeAllClassesPublic = patched
    
End Function


''mostly from elroy
'Function makePublic(obj As Object, Optional ByRef orgVal As Long, Optional andPatch As Boolean = True) As Boolean
'
'    Dim pVTbl As Long, pObjInfo As Long, pObj As Long, newVal As Long, oldMemProt As Long, r As Long
'    Dim pObjTypeField As Long
'
'    Const flag = &H800 '1000 0000 0000
'
'    dbg "Typename(obj) = " & TypeName(obj)
'
'    GetMem4 ByVal ObjPtr(obj), pVTbl                 ' Pointer to vTable.
'    'dbg "pVTbl=" & Hex(pVTbl)
'
'    GetMem4 ByVal pVTbl - 4&, pObjInfo              ' Pointer to tObjectInfo structure.
'    'dbg "pObjInfo =" & Hex(pObjInfo)
'
'    GetMem4 ByVal pObjInfo + &H18&, pObj            ' Pointer to tObject     structure.
'    'dbg "pObj=" & Hex(pObj)
'
'    pObjTypeField = pObj + &H28&
'    GetMem4 ByVal pObjTypeField, orgVal             ' objType value
'
'    dbg "Current ObjType: " & Hex(orgVal)
'
'    If andPatch Then
'        If (orgVal And flag) = 0 Then  'in IDE = 0x18883 flag is set?
'
'            newVal = orgVal Xor flag
'
'            If VirtualProtect(ByVal pObjTypeField, 4, RWE, oldMemProt) <> 0 Then
'                dbg "Patching to: " & Hex(newVal)
'                PutMem4 ByVal pObjTypeField, newVal
'
'                GetMem4 ByVal pObjTypeField, newVal
'                dbg "Sanity Check: " & Hex(newVal)
'
'                makePublic = True
'                VirtualProtect ByVal pObjTypeField, 4, oldMemProt, r
'            Else
'                dbg "virt prot failed"
'            End If
'
'        Else
'            dbg "Cant patch flag already set?"
'        End If
'    End If
'
'End Function
 

Private Sub Form_Load()
    
    Dim cnt As Long
    
    If App.PrevInstance Then
        MsgBox "prev instance"
        End
    End If

    cnt = makeAllClassesPublic(Me, True)
    List1.AddItem cnt & " classes patched"
    
    Set myClass = New Class1 'any new classes created will now take on the new flags and be accessible..
    
    m_lCookie = PutObject(Me, "MySpecialProject.Form1")
    List1.AddItem "MySpecialProject.Form1 cookie: " & Hex(m_lCookie)
    
    m_lCookie2 = PutObject(myClass, "MySpecialProject.myClass")
    List1.AddItem "MySpecialProject.myClass cookie: " & Hex(m_lCookie2)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RevokeObject m_lCookie
    RevokeObject m_lCookie2
End Sub

Public Function PutObject(oObj As Object, sPathName As String, Optional ByVal Flags As Long) As Long
    Const ROTFLAGS_REGISTRATIONKEEPSALIVE As Long = 1
    Const IDX_REGISTER  As Long = 3
    Dim hResult         As Long
    Dim pROT            As IUnknown
    Dim pMoniker        As IUnknown
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        Err.Raise hResult, "GetRunningObjectTable"
    End If
    hResult = CreateFileMoniker(StrPtr(sPathName), pMoniker)
    If hResult < 0 Then
        Err.Raise hResult, "CreateFileMoniker"
    End If
    DispCallByVtbl pROT, IDX_REGISTER, ROTFLAGS_REGISTRATIONKEEPSALIVE Or Flags, ObjPtr(oObj), ObjPtr(pMoniker), VarPtr(PutObject)
End Function

Public Sub RevokeObject(ByVal lCookie As Long)
    Const IDX_REVOKE    As Long = 4
    Dim hResult         As Long
    Dim pROT            As IUnknown
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        Err.Raise hResult, "GetRunningObjectTable"
    End If
    DispCallByVtbl pROT, IDX_REVOKE, lCookie
End Sub

Private Function DispCallByVtbl(pUnk As IUnknown, ByVal lIndex As Long, ParamArray a() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = a
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function
