VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObjectTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com - sandsprite.com

Public baseOffset As Long
Public Objects As New Collection   'of CObject

Public lNull1 As Long
Public aExecProj As Long           ' Pointer to VB Project Exec COM Object.
Public aProjectInfo2 As Long       ' Secondary Project Information
Public Const1 As Long              ' Always set to -1 after compiling. Unused
Public Null2 As Long
Public lpProjectObject As Long     ' Pointer to in-memory Project Data.
Public uuidObj As Variant          ' GUID of the Object Table. 16 byte array
Public fCompileType As Integer     ' Internal flag used during compilation.
Public ObjectCount1 As Integer     ' Total objects present in Project.
Public iCompiledObjects As Integer ' Equal to above after compiling.
Public iObjectsInUse As Integer
Public lpObjectArray As Long       ' Pointer to Object Descriptors
Public fIdeFlag As Long            ' Flag/Pointer used in IDE only.
Public lpIdeData As Long           ' Flag/Pointer used in IDE only.
Public lpIdeData2 As Long          ' Flag/Pointer used in IDE only.
Public aProjectName As Long        ' Pointer to Project Name
Public LangID1 As Long
Public LangID2 As Long
Public lpIdeData3 As Long          ' Flag/Pointer used in IDE only.
Public dwIdentifier As Long        ' Template Version of Structure.

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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Property Get structSize() As Long
    structSize = &H54
End Property

Friend Sub LoadFromMem(va As Long)
    
    Dim t As tObjectTable, co As CObject, i As Long
    
    baseOffset = va
    
    CopyMemory ByVal VarPtr(t), ByVal va, structSize
    
    With t
        lNull1 = .lNull1
        aExecProj = .aExecProj
        aProjectInfo2 = .aProjectInfo2
        Const1 = .Const1
        Null2 = .Null2
        lpProjectObject = .lpProjectObject
        uuidObj = .uuidObj
        fCompileType = .fCompileType
        ObjectCount1 = .ObjectCount1
        iCompiledObjects = .iCompiledObjects
        iObjectsInUse = .iObjectsInUse
        lpObjectArray = .lpObjectArray
        fIdeFlag = .fIdeFlag
        lpIdeData = .lpIdeData
        lpIdeData2 = .lpIdeData2
        aProjectName = .aProjectName
        LangID1 = .LangID1
        LangID2 = .LangID2
        lpIdeData3 = .lpIdeData3
        dwIdentifier = .dwIdentifier
    End With
        
    If ObjectCount1 < 1 Then Exit Sub
    If ObjectCount1 > 255 Then Exit Sub 'bad data? vb does have some limitations here ?
    
    Const objSize As Long = &H30
    
    For i = 0 To ObjectCount1 - 1
        va = lpObjectArray + (i * objSize)
        Set co = New CObject
        co.LoadFromMem va, i
        Objects.Add co
        Debug.Print "Loaded Obj: " & co.name
    Next
    
    Debug.Print "Loaded " & Objects.Count & " objects"

End Sub


Function dump(Optional prefix = vbTab) As String
        
        Dim fields() As String
        Dim ret() As String

        Dim i As Long, v, offset As Long, sz As Long

        fields = Split("lNull1,aExecProj,aProjectInfo2,Const1,Null2,lpProjectObject,uuidObj,fCompileType,ObjectCount1,iCompiledObjects,iObjectsInUse,lpObjectArray,fIdeFlag,lpIdeData,lpIdeData2,aProjectName,LangID1,LangID2,lpIdeData3,dwIdentifier", ",")
                
        push ret, "VA       Off   Name    = Value (epiProject)"
        For i = 0 To UBound(fields)
            v = CallByName(Me, fields(i), VbGet)
            If fields(i) = "uuidObj" Then
                sz = 16
            Else
                sz = sizeForType(v)
            End If
            push ret, Hex(baseOffset + offset) & "   " & Right("00" & Hex(offset), 3) & "   " & fields(i) & " = " & Hex(v)
            offset = offset + sz
        Next
        
        dump = prefix & Join(ret, vbCrLf & prefix)
        
End Function

Private Function sizeForType(v) As Long
    
    If TypeName(v) = "Integer" Then
        sizeForType = 2
    ElseIf TypeName(v) = "Long" Then
        sizeForType = 4
    ElseIf TypeName(v) = "Byte" Then
        sizeForType = 1
    Else
        sizeForType = LenB(v)
    End If
    
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
