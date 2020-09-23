VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum TYPEKIND
  TKIND_ENUM = 0
  TKIND_RECORD
  TKIND_MODULE
  TKIND_INTERFACE
  TKIND_DISPATCH
  TKIND_COCLASS
  TKIND_ALIAS
  TKIND_UNION
  TKIND_MAX
End Enum

Private Enum CALLCONV
  CC_CDECL = 1
  CC_MSCPASCAL
  CC_PASCAL = 2
  CC_MACPASCAL
  CC_STDCALL
  CC_RESERVED
  CC_SYSCALL
  CC_MPWCDECL
  CC_MPWPASCAL
  CC_MAX
End Enum
    
Private Enum FUNCKIND
  FUNC_VIRTUAL
  FUNC_PUREVIRTUAL
  FUNC_NONVIRTUAL
  FUNC_STATIC
  FUNC_DISPATCH
End Enum
    
Private Enum INVOKEKIND
  INVOKE_FUNC = &H1
  INVOKE_PROPERTYGET = &H2
  INVOKE_PROPERTYPUT = &H4
  INVOKE_PROPERTYPUTREF = &H8
End Enum

Private lastElems As Object
Private lastParams As Parameters

Private mFile As String
Private mTlbName As String
Private mHelpFile As String
Private mDocumentation As String
Private mVersion As String
Private mInfo As TypeLibInfo
Private mTypeInfo As TypeInfos
Private mDescription As String

Public Event Error(ByVal strError As String)

'--> Constructor
Private Sub Class_Initialize()
  mFile = vbNullString
  mTlbName = vbNullString
  mHelpFile = vbNullString
  mDocumentation = vbNullString
  mVersion = vbNullString
  Set mInfo = New TypeLibInfo
End Sub

'--> Destructor
Private Sub Class_Terminate()

End Sub

'--> Property File
Public Property Get Filename() As String
  Filename = mFile
End Property

Public Property Get Description() As String
  Description = mDescription
End Property

Public Property Let Filename(ByVal File As String)
  If Exist(File) Then
    mFile = File
  Else
   RaiseEvent Error("File does not exist")
  End If
End Property

'--> Property TlbName
Public Property Get TypeLibName() As String
  TypeLibName = mTlbName
End Property

'--> Help File
Public Property Get Helpfile() As String
  Helpfile = mHelpFile
End Property

'--> Documentation
Public Property Get Documentation() As String
  Documentation = mDocumentation
End Property

'--> Version
Public Property Get Version() As String
  Version = mVersion
End Property

'--> File Exists Function
Private Function Exist(strFile As String) As Long

On Error GoTo HandleErr

If Len(strFile) > 0 Then
  If Len(Dir(strFile)) <> 0 Then
    Exist = -1
  End If
End If

Exit Function

HandleErr:
  Err = 0
  Exist = 0
  Exit Function
  
End Function

'--> Generate DLL StartReport
Public Sub GenerateReport()

Dim o As Object

  On Error GoTo errHandler
    '--> instance tlbInf32.dll
    Set mInfo = TLIApplication.TypeLibInfoFromFile(mFile)
    mInfo.SetMemberFilters
    '--> default information
    mTlbName = mInfo.Name
    mHelpFile = mInfo.Helpfile
    mDocumentation = mInfo.HelpString
    mVersion = mInfo.MajorVersion & "." & mInfo.MinorVersion
        
  Exit Sub
  
errHandler:
  RaiseEvent Error(Err.Description)
  Err.Clear

End Sub

Public Function GetModules() As String()

Dim o As Object
Dim strModules() As String
Dim counter As Long

On Error GoTo errHandler

  '--> Modules
  Set mTypeInfo = mInfo.TypeInfos
  counter = 0
  '--> prevent ubound error
  ReDim strModules(0)
  For Each o In mTypeInfo
    ReDim Preserve strModules(counter)
    strModules(counter) = CStr(o.Name)
    counter = counter + 1
  Next
  
  GetModules = strModules

Exit Function

errHandler:
  RaiseEvent Error(Err.Description)
  Err.Clear
  
End Function

Public Function ChangeTypeInfosSelection(Index As Long) As String()

Dim o As Object, strTmp As String
Dim counter As Long
Dim strFunction() As String

  '--> prevent ubound error
  ReDim strFunction(0)
  counter = 1

On Error GoTo errHandler

  Set o = mTypeInfo(Index + 1)
    
  mHelpFile = o.HelpContext
  mDocumentation = o.HelpString
  Select Case o.TYPEKIND
    Case TKIND_ENUM
      strTmp = "Enum"
      Set lastElems = o.Members
    Case TKIND_RECORD
      strTmp = "Record"
      Set lastElems = o.Members
    Case TKIND_MODULE
      strTmp = "Module"
      Set lastElems = o.Members
    Case TKIND_INTERFACE
      strTmp = "Interface"
      Set lastElems = o.Members
    Case TKIND_DISPATCH
      Set lastElems = o.Members
      ChangeTypeInfosSelection = strFunction
      Exit Function
    Case TKIND_COCLASS
      strTmp = "CoClass"
      Set lastElems = o.Interfaces
    Case TKIND_ALIAS
      strTmp = "Alias (**)"
    Case TKIND_UNION
      strTmp = "Union"
      Set lastElems = o.Members
  End Select

  
  For Each o In lastElems
    ReDim Preserve strFunction(counter)
    strFunction(counter) = CStr(o.Name)
    counter = counter + 1
  Next

  ChangeTypeInfosSelection = strFunction

  Exit Function
  
errHandler:
  RaiseEvent Error(Err.Description)
  Err.Clear

End Function

Public Function ChangeElementsSelection(Index As Long) As String()

Dim o As Object, strTmp As String
Dim oType As Object
Dim strType As String, strDescr As String
Dim strElement() As String
Dim counter As Long

  ReDim strElement(0)
  counter = 1

  Set o = lastElems.Item(Index + 1)

  mHelpFile = o.HelpContext
  mDocumentation = o.HelpString

On Error Resume Next
'  If Err <> 0 Then
'    Select Case o.TypeInfoKind
'      Case TKIND_INTERFACE
'        frm.lEMemberID = "Interface"
'      Case TKIND_DISPATCH
'        frm.lEMemberID = "Dispatch"
'      Case Else
'        frm.lEMemberID = "Unknown element"
'    End Select
'    Exit Function
'  End If

  Select Case o.DescKind
    Case DESCKIND_FUNCDESC 'TYPE_FUNCTION
      Set lastParams = o.Parameters
      For Each o In lastParams
        ReDim Preserve strElement(counter)
        strElement(counter) = o.Name
        counter = counter + 1
      Next
      'LoadList frm.lstParameters, lastParams, CheckAttr:=False
      Set oType = o.ReturnType
      strType = TypeToString(oType)
      strType = TypeToString(oType, True)
      If strType = "" And o.INVOKEKIND = INVOKE_FUNC Then
        strDescr = "Sub " + o.Name + "(" + ListParameters(lastParams) + ")"
      Else
        strDescr = InvokeKindToString(o.INVOKEKIND) + " " + o.Name + _
                "(" + ListParameters(lastParams) + ")" + " As " + strType
      End If
    Case DESCKIND_VARDESC 'TYPE_CONSTANT
      Set oType = o.ReturnType
      strDescr = "Const " + o.Name + " = " & o.Value
  End Select

  mDescription = strDescr
  ChangeElementsSelection = strElement

On Error GoTo 0

End Function

Private Function TypeToString(rType As Object, Optional Normalize As Boolean = False) As String
Dim st As String, rt, udrt As Object, IsArr As Boolean
  rt = rType.VarType
  
  If (rt And VT_ARRAY) = VT_ARRAY Then
    st = "Array("
    IsArr = True
    rt = rt - VT_ARRAY
  End If
  Select Case rt
    Case vbEmpty      ' Empty (uninitialized)
      'st = st + "Empty"
      st = st + rType.TypeInfo.Name
    Case vbNull       ' Null (no valid data)
      st = st + "Null"
    Case vbInteger    ' Integer
      st = st + "Integer"
    Case vbLong       ' Long integer
      st = st + "Long"
    Case vbSingle     ' Single-precision floating-point number
      st = st + "Single"
    Case vbDouble     ' Double-precision floating-point number
      st = st + "Double"
    Case vbCurrency   ' Currency value
      st = st + "Currency"
    Case vbDate       ' Date value
      st = st + "Date"
    Case vbString     ' String
      st = st + "String"
    Case vbObject     ' Object
      st = st + "Object"
    Case vbError      ' Error value
      st = st + "Error"
    Case vbBoolean    ' Boolean value
      st = st + "Boolean"
    Case vbVariant    ' Variant (used only with arrays of variants)
      st = st + "Variant"
    Case vbDataObject ' A data access object
      st = st + "DataObject"
    Case vbDecimal    ' Decimal value
      st = st + "Decimal"
    Case vbByte       ' Byte value
      st = st + "Byte"
    Case vbArray      ' Array
      st = st + "Array"
    Case VT_VOID
      st = st
    Case VT_PTR
      st = st + TypeToString(rType.PointerDesc)
    Case VT_USERDEFINED
      Set udrt = rType.UserDefinedDesc
      st = st + udrt.Name
    Case VT_I1                    ' signed char
      st = st + "*signed char*"
    Case VT_UI1                   ' unsigned char
      st = st + "*unsigned char*"
    Case VT_UI2                   ' unsigned short
      st = st + "*signed short*"
    Case VT_UI4                   ' unsigned short
      st = st + "*unsigned short*"
    Case VT_I8                    ' signed 64-bit int
      st = st + "*signed 64-bit int*"
    Case VT_UI8                   ' unsigned 64-bit int
      st = st + "*unsigned 64-bit int*"
    Case VT_INT                   ' signed machine int
      st = st + "*signed machine int*"
    Case VT_UINT                  ' unsigned machine int
      st = st + "*unsigned machine int*"
    Case VT_HRESULT               '
      st = st + "*HRESULT*"
    Case VT_SAFEARRAY             ' (use VT_ARRAY in VARIANT)
      st = st + "*SAFEARRAY*"
    Case VT_CARRAY                ' C style array
      st = st + "*C style array*"
    Case VT_LPSTR                 ' null terminated string
      st = st + "*LPSTR*"
    Case VT_LPWSTR                ' wide null terminated string
      st = st + "*LPWSTR*"
    Case Else
      st = st + "*Unknown* (&H" & Hex(rt) & ")"
  End Select
  If IsArr Then st = st + ")"
  If Normalize Then st = TypeStrNormalized(st)
  TypeToString = st
End Function

Private Function TypeStrNormalized(sType As String) As String
  If Left$(sType, 1) = "*" And Right$(sType, 1) = "*" Then
    TypeStrNormalized = "<?>"
  Else
    TypeStrNormalized = sType
  End If
End Function

Private Function ListParameters(coll As Object) As String
Dim o As Object, sParams As String, sType As String
  sParams = ""
  For Each o In coll
    If sParams <> "" Then sParams = sParams + ", "
    sParams = sParams + o.Name
    sType = TypeToString(o.VarTypeInfo, True)
    If sType <> "" Then sParams = sParams + " As " + sType
  Next
  ListParameters = sParams
End Function

Private Function InvokeKindToString(ik) As String
Dim sTmp As String
  Select Case ik
    Case INVOKE_FUNC
      sTmp = "Function"
    Case INVOKE_PROPERTYGET
      sTmp = "Property Get"
    Case INVOKE_PROPERTYPUT
      sTmp = "Property Put"
    Case INVOKE_PROPERTYPUTREF
      sTmp = "Property Put Ref"
    Case Else
      sTmp = "*Unknown*(" & ik & ")"
  End Select
  InvokeKindToString = sTmp
End Function

