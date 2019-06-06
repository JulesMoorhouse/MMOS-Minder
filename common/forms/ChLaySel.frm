VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildLayoutSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Layout Selection"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1305
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Layout Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmChildLayoutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim llngCol1Width As Long
Dim llngCol2Width As Long
Dim mstrLayoutsPath As String
Dim mstrLayoutsFound As String
Dim mlngLayoutType As Long
Dim lbooAcceptableLayoutFound As Boolean

Public Property Get LayoutType() As String

    LayoutType = mlngLayoutType
    
End Property
Public Property Let LayoutType(pstrLayoutType As String)

    mlngLayoutType = pstrLayoutType

End Property
Public Property Get LayoutsFound() As String

    LayoutsFound = mstrLayoutsFound
    
End Property
Public Property Let LayoutsFound(pstrLayoutsFound As String)

    mstrLayoutsFound = pstrLayoutsFound

End Property
Public Property Get LayoutsPath() As String

    LayoutsPath = mstrLayoutsPath
    
End Property
Public Property Let LayoutsPath(pstrLayoutsPath As String)

    mstrLayoutsPath = pstrLayoutsPath

End Property

Private Sub cmdSelect_Click()
Dim lstrTempFile As String

    If ListView1.SelectedItem.Index = -1 Then
        MsgBox "You must select a Layout!", , gconstrTitlPrefix & "Layouts"
        Exit Sub
    End If
    
    'MsgBox ListView1.SelectedItem.Text, , ListView1.SelectedItem.SubItems(1)
    
    lstrTempFile = GetTempDir & "L" & Format(Now(), "MMDDSSN")
    Decrypt mstrLayoutsPath & ListView1.SelectedItem.SubItems(1), gconEncryptDataFile, lstrTempFile
    DoEvents

    If ReadTempReportLayoutFile(lstrTempFile) = False Then
        Unload Me
        Exit Sub
    End If
    DoEvents
    Kill lstrTempFile
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    llngCol1Width = ListView1.ColumnHeaders.Item(1).Width
    llngCol2Width = ListView1.ColumnHeaders.Item(2).Width
    
    LoadLayouts
    If ListView1.ListItems.Count > 0 Then
        Set ListView1.SelectedItem = ListView1.ListItems(1)
    End If

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    ListView1.ColumnHeaders.Item(1).Width = llngCol1Width
    ListView1.ColumnHeaders.Item(2).Width = llngCol2Width
    
End Sub

Sub LoadLayouts()
Dim litmAdd As ListItem
Dim lstrFoundReportFiles As String
Dim lstrTempFile As String

Dim llngLayoutType As Long
Dim lstrReportDescription As String
Dim lstrFileLen As String
Dim lstrFileDate As String

Dim lstrRegLayoutType As String
Dim lstrRegReportDescription As String
Dim lstrRegFileLen As String
Dim lstrRegFileDate As String
   
Dim lstrTempPath As String

    lstrTempPath = GetTempDir
    
    lbooAcceptableLayoutFound = False
    lstrFoundReportFiles = Dir(mstrLayoutsPath & "*.rpt", vbNormal)
    
    If lstrFoundReportFiles = "" Then
        mstrLayoutsFound = "No Report Layouts Files Found!"
        Exit Sub
    Else
        mstrLayoutsFound = ""
    End If
    
    Do Until lstrFoundReportFiles = ""
        If Right$(UCase$(lstrFoundReportFiles), 4) = ".RPT" Then
        
            lstrFileDate = Format(FileDateTime(mstrLayoutsPath & lstrFoundReportFiles), "YYMMDDSSN")
            lstrFileLen = FileLen(mstrLayoutsPath & lstrFoundReportFiles)
            
            lstrRegFileLen = GetSetting(gstrIniAppName, lstrFoundReportFiles, "FileLen")
            lstrRegFileDate = GetSetting(gstrIniAppName, lstrFoundReportFiles, "DateTime")
            lstrRegReportDescription = GetSetting(gstrIniAppName, lstrFoundReportFiles, "Description")
            lstrRegLayoutType = Val(GetSetting(gstrIniAppName, lstrFoundReportFiles, "LayoutType"))
            
            If lstrFileLen = lstrRegFileLen And lstrFileDate = lstrRegFileDate Then
                lstrReportDescription = lstrRegReportDescription
                llngLayoutType = Val(lstrRegLayoutType)
            End If
            
            If Trim$(lstrReportDescription) = "" Then
                lstrTempFile = lstrTempPath & "L" & Format(Now(), "MMDDSSN")
                Decrypt mstrLayoutsPath & lstrFoundReportFiles, gconEncryptDataFile, lstrTempFile
                
                If ReadTempReportLayoutFile(lstrTempFile) = False Then
                    Exit Sub
                End If
                
                lstrReportDescription = gstrReportLayout.strLayoutName
                llngLayoutType = gstrReportLayout.lngLayoutType
                
                SaveSetting gstrIniAppName, lstrFoundReportFiles, "Description", lstrReportDescription
                SaveSetting gstrIniAppName, lstrFoundReportFiles, "LayoutType", llngLayoutType
                SaveSetting gstrIniAppName, lstrFoundReportFiles, "DateTime", lstrFileDate
                SaveSetting gstrIniAppName, lstrFoundReportFiles, "FileLen", lstrFileLen
                ClearReportingLayoutType
                Kill lstrTempFile
            End If
            
            If llngLayoutType = LayoutType Then
                Set litmAdd = ListView1.ListItems.Add(Text:=lstrReportDescription)
                litmAdd.SubItems(1) = lstrFoundReportFiles
                lbooAcceptableLayoutFound = True
            End If
            
            llngLayoutType = -1
            lstrReportDescription = ""
            lstrFileDate = ""
            lstrFileLen = ""
        End If
        lstrFoundReportFiles = Dir
    Loop
    
    If lbooAcceptableLayoutFound = True Then
        Set ListView1.SelectedItem = ListView1.ListItems(1)
    Else
        mstrLayoutsFound = "No Layouts for this report have been found!"
    End If
End Sub
