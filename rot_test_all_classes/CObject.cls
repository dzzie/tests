VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com - sandsprite.com

Public index As Long
Public name As String
Public baseOffset As Long

Public aObjectInfo As Long            ' 0x00
Public Const1 As Long                 ' 0x04
Public aPublicBytes As Long           ' 0x08 (08d) Pointer to Public Variable Size integers
Public aStaticBytes As Long           ' 0x0C (12d) Pointer to Static Variables Struct (describes size of each var in alloc)
Public aModulePublic As Long          ' 0x10 (16d) Memory Pointer to Public Variables
Public aModuleStatic As Long          ' 0x14 (20d) Pointer to Static Variables
Public aObjectName As Long            ' 0x18  NTS
Public ProcCount As Long              ' 0x1C event sinks, funcs, subs, (api declares
Public aProcNamesArray As Long        ' 0x20 when non-zero
Public oStaticVars As Long            ' 0x24 (36d) objPtr(me) + oStaticVars = alloc for static vars
Public ObjectType As Long             ' 0x28
Public Null3 As Long                  ' 0x2C

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
    ObjectType As Long          ' 0x28  Flags defining the Object Type.
    Null3 As Long               ' 0x2C  Not valid after compilation.
End Type

Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Property Get structSize() As Long
    structSize = &H30
End Property

Property Get offsetOf_ObjType() As Long
    offsetOf_ObjType = baseOffset + &H28
End Property

Property Get hasOptInfo() As Boolean
    If (ObjectType And 2) = 2 Then hasOptInfo = True
End Property

Property Get isClass() As Boolean
     
    Dim v
    For Each v In Array(1146883, 1277955, 98339, 100355, 1148931, &H118043, &H18003, &H138803)
        If v = ObjectType Then
            isClass = True
            Exit Property
        End If
    Next
    
End Property

'mostly from moog
'https://www.rapidtables.com/convert/number/hex-to-binary.html?x=50430083
'tObject.ObjectTyper Properties...24 bits used
'0x80 =                                 1000 0000
'hexcodes: 18083, 1118803,180A3,180C3
'#########################################################
'form               0000 0001 1000 0000 1000 0011 --> 18083
'                   0000 0001 1000 0000 1010 0011 --> 180A3
'                   0000 0001 1000 0000 1100 0011 --> 180C3

'module             0000 0001 1000 0000 0000 0001 --> 18001
'                   0000 0001 1000 0000 0100 0001 --> 18041
'                   0000 0001 1000 0000 0010 0001 --> 18021

'class              0001 0001 1000 0000 0000 0011 --> 118003
'                   0001 0011 1000 1000 0000 0011 --> 138803  global class
'                   0001 0001 1000 0000 0100 0011 --> 118043
'                   0001 0011 1000 0000 0000 0011 --> 138003
'                   0000 0001 1000 0000 0010 0011 --> 18023
'                   0000 0001 1000 1000 0000 0011 --> 18803
'                   0000 0001 1000 0000 0000 0011 --> 18003  - vb5
'                   0001 0001 1000 1000 0000 0011 --> 118803

'usercontrol        0001 1101 1010 0000 0000 0011 --> 1DA003
'                   0001 1101 1010 0000 0010 0011 --> 1DA023
'                   0001 1101 1010 1000 0000 0011 --> 1DA803
'                   0001 1101 1010 0000 0100 0011 --> 1DA043

'propertypage       0001 0101 1000 0000 0000 0011 --> 158003
'                   0001 0101 1000 0000 0100 0011 --> 158043

'datareport    0001 0000 0001 1000 0000 1000 0011 -->1018083
'              0001 0000 0001 1000 0000 1100 0011 -->10180C3

'DataEnv  0101 0000 0100 0011 0000 0000 1000 0011 -->50430083 https://flylib.com/books/en/3.405.1.73/1/
'                 |    | |||    |  |    |||    |
'data report   ---+    | |||    |  |    |||    |      0x800000 1000 0000 0000 0000 0000 0000
'                      | |||    |  |    |||    |
'HasPublicInterface ---+ |||    |  |    |||    |      0x100000    1 0000 0000 0000 0000 0000
'HasPublicEvents --------+||    |  |    |||    |       0x80000      1000 0000 0000 0000 0000
'IsCreatable/Visible? ----+|    |  |    |||    |       0x40000       100 0000 0000 0000 0000
'global class -------------+    |  |    |||    |       0x20000        10 0000 0000 0000 0000
'Same as "HasPublicEvents" -----+  |    |||    |
'                               |  |    |||    |
'usercontrol     ---------------+  |    |||    |        0x2000             10 0000 0000 0000
'ocx/dll     ----------------------+    |||    |         0x800                1000 0000 0000
'form      -----------------------------+||    |          0x80                     1000 0000
'opt cmp text ---------------------------+|    |          0x40                      100 0000
'vb5     ---------------------------------+    |          0x20                       10 0000
'HasOptInfo     -------------------------------+           0x2                            10
'                                              |
'module    ------------------------------------+           0x2                            10
'
' https://www.rapidtables.com/convert/number/hex-to-binary.html?x=50430083

'11 unknown flags: &h400000,&h200000,&h10000,&h8000,&h4000,&h400,&h200,&h100,&h10,&h8,&h4

'data report          0x800000 1000 0000 0000 0000 0000 0000
'unknown 1                      100 0000 0000 0000 0000 0000  400000
'unknown 2                       10 0000 0000 0000 0000 0000  200000
'HasPublicInterface   0x100000    1 0000 0000 0000 0000 0000
'HasPublicEvents       0x80000      1000 0000 0000 0000 0000
'IsCreatable/Visible   0x40000       100 0000 0000 0000 0000
'global class          0x20000        10 0000 0000 0000 0000
'unknown 3                             1 0000 0000 0000 0000  10000  'everything has this but unnamed...
'unknown 4                               1000 0000 0000 0000   8000
'unknown 5                                100 0000 0000 0000   4000
'usercontrol            0x2000             10 0000 0000 0000
'ocx/dll                 0x800                1000 0000 0000
'unknown 6                                     100 0000 0000   400
'unknown 7                                      10 0000 0000   200
'unknown 8                                       1 0000 0000   100
'form                     0x80                     1000 0000
'opt cmp text             0x40                      100 0000
'vb5                      0x20                       10 0000
'unknown 9                                            1 0000   10
'unknown 10                                             1000    8
'unknown 11                                              100    4
'HasOptInfo                0x2                            10
'module                    0x2                            10
'unknown 12                                                1


Friend Sub LoadFromMem(va As Long, index As Long)

    Dim t As tObject
    Dim b As Byte, i As Long

    baseOffset = va
    Me.index = i
    
    CopyMemory ByVal VarPtr(t), ByVal va, structSize
     
    With t
        aObjectInfo = .aObjectInfo
        Const1 = .Const1
        aPublicBytes = .aPublicBytes
        aStaticBytes = .aStaticBytes
        aModulePublic = .aModulePublic
        aModuleStatic = .aModuleStatic
        aObjectName = .aObjectName
        ProcCount = .ProcCount
        aProcNamesArray = .aProcNamesArray
        oStaticVars = .oStaticVars
        ObjectType = .ObjectType
        Null3 = .Null3
    End With
    
    For i = 0 To 255
        GetMem1 aObjectName + i, b
        If b = 0 Then Exit For
        name = name & Chr(b)
    Next
    
    'objInfo.LoadFromMem aObjectInfo, Me

End Sub


Function dump(Optional indent = 1) As String

    Dim x(), prefix
 
    If indent > 0 Then prefix = String(indent, vbTab)
     
    'push x, prefix & "index: " & index
    push x, prefix & "name: " & name
    push x, prefix & "baseOffset: " & Hex(baseOffset)
    push x, prefix & "ObjectType: " & Hex(ObjectType)
    push x, prefix & "lpObjInfo: " & Hex(aObjectInfo)
    'push x, prefix & "ProcCount: " & ProcCount
    push x, prefix & "isClass: " & isClass()
    
    dump = Join(x, vbCrLf)
    
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

