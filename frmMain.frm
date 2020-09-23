VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "DLLBrowser"
   ClientHeight    =   5445
   ClientLeft      =   660
   ClientTop       =   3075
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7830
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5700
      TabIndex        =   7
      Top             =   4980
      Width           =   1995
   End
   Begin VB.TextBox txtFunction 
      Enabled         =   0   'False
      Height          =   465
      Left            =   30
      TabIndex        =   6
      Top             =   4380
      Width           =   7665
   End
   Begin VB.ListBox listArguments 
      Height          =   2400
      Left            =   5190
      TabIndex        =   5
      Top             =   1920
      Width           =   2505
   End
   Begin VB.ListBox listFuntions 
      Height          =   2400
      Left            =   2610
      TabIndex        =   4
      Top             =   1920
      Width           =   2505
   End
   Begin VB.ListBox listModules 
      Height          =   2400
      Left            =   30
      TabIndex        =   3
      Top             =   1920
      Width           =   2505
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "..."
      Height          =   255
      Left            =   4410
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   4275
   End
   Begin MSComDlg.CommonDialog comDiag 
      Left            =   5070
      Top             =   4890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblHelpstring 
      Caption         =   "Helpstring: "
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   3450
      TabIndex        =   9
      Top             =   1020
      Width           =   3285
   End
   Begin VB.Label lblHelpfile 
      Caption         =   "Helpfile:"
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   90
      TabIndex        =   8
      Top             =   1020
      Width           =   3285
   End
   Begin VB.Label lblTop 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8385
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-->Reference to tlbinf32.dll
'-->Design by Mähr Stefan @2004
'-->Use this code at your own risk

Private strCurFile As String
Private WithEvents oInfo As cInfo
Attribute oInfo.VB_VarHelpID = -1

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdSelect_Click()

Me.listArguments.Clear
Me.listFuntions.Clear
Me.listModules.Clear

Dim strModules() As String
Dim i As Integer

'--> Select file
  With Me.comDiag
    On Error Resume Next
    .Filter = "Type Libraries (*.tlb,*.olb,*.dll)|*.tlb;*.olb;*.dll" & _
                              "|All Files (*.*)|*.*"
    .DialogTitle = "Select Type Library"
    .InitDir = strLastPath
    .ShowOpen
    strCurFile = .Filename
    If Err <> 0 Then Err.Clear
  End With
  
  If Len(strCurFile) = 0 Then Exit Sub
  Me.txtFile.Text = strCurFile
  strLastPath = strCurFile
  
  Set oInfo = New cInfo
  Call RefreshInfo
  
  strModules = oInfo.GetModules
  For i = 0 To UBound(strModules)
    Me.listModules.AddItem strModules(i)
  Next
  
End Sub

Private Sub Form_Load()
  Call LoadSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call SaveSettings
End Sub

Private Sub listFuntions_Click()
  Me.listArguments.Clear
  If oInfo Is Nothing Then Exit Sub
  
  Dim strArguments() As String
  strArguments = oInfo.ChangeElementsSelection(Me.listFuntions.ListIndex)
  Dim i As Long
  For i = 1 To UBound(strArguments)
    Me.listArguments.AddItem strArguments(i)
  Next
  Me.txtFunction.Text = vbNullString
  Me.txtFunction.Text = oInfo.Description
  Call RefreshInfo
  
End Sub

Private Sub listModules_Click()

  Me.listFuntions.Clear
  Me.listArguments.Clear

  If oInfo Is Nothing Then Exit Sub
  Dim strFunction() As String
  strFunction = oInfo.ChangeTypeInfosSelection(Me.listModules.ListIndex)
  Dim i As Long
  For i = 1 To UBound(strFunction) - 1
    Me.listFuntions.AddItem strFunction(i)
  Next
  Call RefreshInfo
End Sub

Private Sub oInfo_Error(ByVal strError As String)
  '--> Error event handling
  MsgBox strError
End Sub

Private Sub RefreshInfo()

  Me.lblHelpfile.Caption = vbNullString
  Me.lblHelpstring.Caption = vbNullString
  Me.lblTop.Caption = vbNullString
  
  With oInfo
    .Filename = strCurFile
    .GenerateReport
    Me.lblHelpfile.Caption = Me.lblHelpfile.Caption & " " & .Helpfile
    Me.lblHelpstring.Caption = Me.lblHelpstring.Caption & " " & .Documentation
    Me.lblTop.Caption = .TypeLibName & " Version " & .Version
  End With

End Sub
