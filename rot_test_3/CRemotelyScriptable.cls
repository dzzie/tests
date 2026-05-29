VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRemotelyScriptable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'dzzie@yahoo.com
'http://sandsprite.com

'rot registration code from wqweto
'https://www.vbforums.com/showthread.php?879529-project-one-workng-with-project2

Private Declare Function vbaCheckType Lib "msvbvm60" Alias "__vbaCheckType" (ByVal pObj As Any, ByRef pIID As Any) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetRunningObjectTable Lib "ole32" (ByVal dwReserved As Long, pResult As IUnknown) As Long
Private Declare Function CreateFileMoniker Lib "ole32" (ByVal lpszPathName As Long, pResult As IUnknown) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal newVal As Long)
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Const RWE = &H40

Private Type tObjectTable
    lNull1 As Long              ' 0x00 (00d)
    aExecProj As Long           ' 0x04 (04d) Pointer to a memory structure
    aProjectInfo2 As Long       ' 0x08 (08d) Pointer to Project Info 2
    Const1 As Long              ' 0x0C
    Null2 As Long               ' 0x10
    lpProjectObject As Long     ' 0x14
    uuidObj(15) As Byte         'converted from 4 flags dzzie (from vb.idc)
    fCompileType As Integer     ' 0x28 (40d) Internal flag used during compilation
    ObjectCount1 As Integer     ' 0x2A
    iCompiledObjects As Integer ' 0x2C (44d) Number of objects compiled.
    iObjectsInUse As Integer    ' 0x2E (46d) Updated in the IDE to correspond the total number ' but will go up or down when initializing/unloading modules.
    lpObjectArray As Long       ' 0x30
    fIdeFlag As Long            ' 0x34
    lpIdeData As Long           ' 0x38
    lpIdeData2 As Long          ' 0x3C
    aProjectName As Long        ' 0x40      NTS
    LangID1  As Long            ' 0x44
    LangID2  As Long            ' 0x48
    lpIdeData3  As Long         ' 0x4C
    dwIdentifier  As Long       ' 0x50
End Type

Private Type tObject
    aObjectInfo As Long         ' 0x00  Pointer to the Object Info for this Object.
    Const1 As Long              ' 0x04  Always set to -1 after compiling.
    aPublicBytes As Long        ' 0x08  Pointer to Public Variable Size integers
    aStaticBytes As Long        ' 0x0C  Pointer to Static Variables Struct
    aModulePublic As Long       ' 0x10  Pointer to Public Variables in DATA section
    aModuleStatic As Long       ' 0x14  Pointer to Static Variables in DATA section
    aObjectName As Long         ' 0x18  Name of the Object.
    ProcCount As Long           ' 0x1C  Number of Methods in Object
    aProcNamesArray As Long     ' 0x20  If present, pointer to Method names array.
    oStaticVars As Long         ' 0x24  Offset to Static Vars from aModuleStatic
    objectType As Long          ' 0x28  Flags defining the Object Type.
    Null3 As Long               ' 0x2C  Not valid after compilation.
End Type

Private handles As New Collection

Private Sub Class_Terminate()
    Dim h
    Debug.Print Hex(ObjPtr(Me)) & " CRemotlyScriptable Terminating closing " & handles.Count & " handles..."
    For Each h In handles
        RevokeObject h
    Next
End Sub

'.text:66008130 _IID_AreYouABasicInstance dd 0B6C9465 11CFD082 0A0004F8B 4270FC9
Private Function isVB6Obj(o As IUnknown) As Boolean
    Dim g(3) As Long: g(0) = &HB6C9465: g(1) = &H11CFD082: g(2) = &HA0004F8B: g(3) = &H4270FC9
    isVB6Obj = vbaCheckType(o, g(0))
End Function

Function dbg(x)
    Debug.Print x
    'List2.AddItem x
    'MsgBox x
End Function

Private Function ExtractName(va As Long) As String
    Dim i As Long, b As Byte
    For i = 0 To 255
        GetMem1 va + i, b
        If b = 0 Then Exit For
        ExtractName = ExtractName & Chr(b)
    Next
End Function

Private Function isClass(objectType As Long) As Boolean
     
    Dim v
    For Each v In Array(1146883, 1277955, 98339, 100355, 1148931, &H118043, &H18003, &H138803)
        If v = objectType Then
            isClass = True
            Exit Function
        End If
    Next
    
End Function

Function makeAllPublic(internalObj As Object) As Long

    Dim pVTbl As Long, pObjInfo As Long, pObjTable As Long, newVal As Long, oldMemProt As Long, r As Long
    Dim pObjType As Long, patched As Long
    Dim objTable As tObjectTable
    Dim tObj As tObject, i As Long, va As Long
    
    Const flag = &H800
    
    If internalObj Is Nothing Then Exit Function
        'Set internalObj = VB.Forms(1)
        
    If Not isVB6Obj(internalObj) Then
        MsgBox "Not a vb6 object " & TypeName(internalObj), vbExclamation
        Exit Function
    End If
    
    dbg "Start Obj = " & TypeName(internalObj)
    
    GetMem4 ByVal ObjPtr(internalObj), pVTbl        ' Pointer to any internal vb objects vTable.
    GetMem4 ByVal pVTbl - 4&, pObjInfo              ' Pointer to tObjectInfo structure.
    GetMem4 ByVal pObjInfo + 4, pObjTable           ' Pointer to tObjectTable structure.

    CopyMemory ByVal VarPtr(objTable), ByVal pObjTable, &H54
    
    dbg objTable.ObjectCount1 & " objects in table"
    
    'bad data? vb does have some limitations here ?
    If objTable.ObjectCount1 < 1 Or objTable.ObjectCount1 > 255 Then Exit Function
    
    Const objSize As Long = &H30
    
    For i = 0 To objTable.ObjectCount1 - 1
        
        va = objTable.lpObjectArray + (i * objSize)
        CopyMemory ByVal VarPtr(tObj), ByVal va, objSize
        Debug.Print "Loaded Obj: " & ExtractName(tObj.aObjectName)
        
        If isClass(tObj.objectType) Then
        
            If (tObj.objectType And flag) = 0 Then
                
                pObjType = va + &H28
                newVal = tObj.objectType Xor flag
                
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

    makeAllPublic = patched
    
End Function

Function makePublic(obj As Object) As Boolean

    Dim pVTbl As Long, pObjInfo As Long, pObj As Long, newVal As Long, oldMemProt As Long, r As Long
    Dim pObjTypeField As Long, orgVal  As Long

    Const flag = &H800

    If obj Is Nothing Then Exit Function
    
    If Not isVB6Obj(obj) Then
        MsgBox "Not a vb6 object " & TypeName(obj), vbExclamation
        Exit Function
    End If
    
    dbg "Typename(obj) = " & TypeName(obj)
   
    GetMem4 ByVal ObjPtr(obj), pVTbl                ' Pointer to vTable.
    GetMem4 ByVal pVTbl - 4&, pObjInfo              ' Pointer to tObjectInfo structure.
    GetMem4 ByVal pObjInfo + &H18&, pObj            ' Pointer to tObject structure.

    pObjTypeField = pObj + &H28&
    GetMem4 ByVal pObjTypeField, orgVal             ' objType value
    'dbg "Current ObjType: " & Hex(orgVal)

    If Not isClass(orgVal) Then
        MsgBox "Only Classes can be made public"
        Exit Function
    End If
 
    If (orgVal And flag) = 0 Then

        newVal = orgVal Xor flag

        If VirtualProtect(ByVal pObjTypeField, 4, RWE, oldMemProt) <> 0 Then
            dbg "Patching to: " & Hex(newVal)
            PutMem4 ByVal pObjTypeField, newVal

            'GetMem4 ByVal pObjTypeField, newVal
            'dbg "Sanity Check: " & Hex(newVal)

            makePublic = True
            VirtualProtect ByVal pObjTypeField, 4, oldMemProt, r
        Else
            dbg "virt prot failed"
        End If

    Else
        dbg "Cant patch flag already set?"
    End If
     

End Function
 


Public Function RegisterObj(oObj As Object, sPathName As String, Optional ByVal Flags As Long, Optional ByRef errCode As Long) As Boolean
    
    Const ROTFLAGS_REGISTRATIONKEEPSALIVE As Long = 1
    Const IDX_REGISTER  As Long = 3
    Dim hResult         As Long
    Dim pROT            As IUnknown
    Dim pMoniker        As IUnknown
    Dim handle As Long
    
    RemoveObj sPathName 'in case already used ok if not
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        errCode = 1
        Exit Function
    End If

    hResult = CreateFileMoniker(StrPtr(sPathName), pMoniker)
    If hResult < 0 Then
        errCode = 2
        Exit Function
    End If
    
    DispCallByVtbl pROT, IDX_REGISTER, ROTFLAGS_REGISTRATIONKEEPSALIVE Or Flags, ObjPtr(oObj), ObjPtr(pMoniker), VarPtr(handle)
    If handle = 0 Then
        errCode = 3
        Exit Function
    End If
    
    handles.Add handle, sPathName
    RegisterObj = True
    
End Function

Public Function RemoveObj(sPathName As String) As Boolean
    Dim h As Long
    
    If KeyExistsInCollection(handles, sPathName) Then
        h = handles(sPathName)
        RevokeObject h
        handles.Remove sPathName
        RemoveObj = True
    End If
  
End Function

Private Sub RevokeObject(ByVal lCookie As Long)
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

Private Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

