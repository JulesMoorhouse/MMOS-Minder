VERSION 5.00
Begin VB.Form frmAccount 
   Caption         =   "Account"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10515
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   10515
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPAD 
      Caption         =   "&PO Collect"
      Height          =   360
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7235
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelpWhat 
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Click on me then use me to find info on things on the screen"
      Top             =   7235
      Width           =   375
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   375
      Left            =   5362
      TabIndex        =   51
      Top             =   6000
      Width           =   4335
      Begin VB.TextBox txtEmailAddress 
         Height          =   288
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   24
         Top             =   0
         Width           =   3012
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Delivery Address Details"
      Height          =   4695
      Left            =   5362
      TabIndex        =   44
      Top             =   1200
      Width           =   5062
      Begin VB.CommandButton cmdCopyName 
         Caption         =   "C&opy =>"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Width           =   825
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "C&opy All=>"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2280
         Width           =   1305
      End
      Begin VB.TextBox txtDeliverySurname 
         Height          =   288
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   15
         Top             =   960
         Width           =   2532
      End
      Begin VB.TextBox txtDeliveryInitials 
         Height          =   288
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   14
         Top             =   600
         Width           =   1812
      End
      Begin VB.TextBox txtDeliverySalutation 
         Height          =   288
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDeliverAddress1 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1560
         Width           =   3012
      End
      Begin VB.TextBox txtDeliverAddress3 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   18
         Top             =   2280
         Width           =   3012
      End
      Begin VB.TextBox txtDeliverAddress2 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   17
         Top             =   1920
         Width           =   3012
      End
      Begin VB.TextBox txtDeliverAddress5 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   20
         Top             =   3000
         Width           =   3012
      End
      Begin VB.TextBox txtDeliverAddress4 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   19
         Top             =   2640
         Width           =   3012
      End
      Begin VB.TextBox txtDeliverPostcode 
         Height          =   288
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   21
         Top             =   3360
         Width           =   1812
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Surname"
         Height          =   255
         Left            =   600
         TabIndex        =   49
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "First"
         Height          =   255
         Left            =   960
         TabIndex        =   48
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Salutation"
         Height          =   255
         Left            =   600
         TabIndex        =   47
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         Height          =   495
         Left            =   720
         TabIndex        =   46
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Post Code"
         Height          =   255
         Left            =   600
         TabIndex        =   45
         Top             =   3360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Caller Address Details"
      Height          =   4695
      Left            =   120
      TabIndex        =   37
      Top             =   1200
      Width           =   5062
      Begin VB.TextBox txtEveTelephone 
         Height          =   288
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   10
         Top             =   4200
         Width           =   1812
      End
      Begin VB.TextBox txtAddress1 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1560
         Width           =   3012
      End
      Begin VB.TextBox txtAddress3 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   5
         Top             =   2280
         Width           =   3012
      End
      Begin VB.TextBox txtAddress2 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1920
         Width           =   3012
      End
      Begin VB.TextBox txtAddress5 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   7
         Top             =   3000
         Width           =   3012
      End
      Begin VB.TextBox txtAddress4 
         Height          =   288
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2640
         Width           =   3012
      End
      Begin VB.TextBox txtPostcode 
         Height          =   288
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   8
         Top             =   3360
         Width           =   1812
      End
      Begin VB.TextBox txtSurname 
         Height          =   288
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtInitials 
         Height          =   288
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   1
         Top             =   600
         Width           =   1812
      End
      Begin VB.TextBox txtTelephone 
         Height          =   288
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   9
         Top             =   3840
         Width           =   1812
      End
      Begin VB.TextBox txtSalutation 
         Height          =   288
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblEveningTelephone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Evening Telephone"
         Height          =   375
         Left            =   0
         TabIndex        =   50
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Surname"
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "First"
         Height          =   255
         Left            =   1080
         TabIndex        =   42
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         Height          =   255
         Left            =   720
         TabIndex        =   41
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Post Code"
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Telephone"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Salutation"
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdNotePad 
      Caption         =   "&Notepad"
      Height          =   360
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort Order"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7235
      Width           =   1305
   End
   Begin VB.ComboBox cboAccountStatus 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Timer timActivity 
      Interval        =   20000
      Left            =   4920
      Top             =   6480
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   360
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7235
      Width           =   1305
   End
   Begin VB.CheckBox chkReceiveMailings 
      Alignment       =   1  'Right Justify
      Caption         =   "Receive Mailings"
      Height          =   315
      Left            =   5400
      TabIndex        =   25
      Top             =   6480
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.ComboBox cboAccountType 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   6000
      Width           =   3012
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   360
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7235
      Width           =   1305
   End
   Begin VB.TextBox lblCustomerNumber 
      BackColor       =   &H8000000F&
      Height          =   288
      Left            =   9090
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1332
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   53
      Top             =   7080
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1244
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1852
   End
   Begin VB.Label lblAccountStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Status"
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblInfoType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Advice Note Address Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label lblAcctType 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Type"
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblCustNum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Number"
      Height          =   255
      Left            =   7635
      TabIndex        =   33
      Top             =   6600
      Width           =   1335
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lstrAccountType() As String
Dim lstrAccountStatus() As String
Dim mstrRoute As String
Dim mstrCriteria As String
Dim mstrCriteriaType As String
Dim lstrScreenHelpFile As String
Public Property Let Route(pstrRoute As String)

    mstrRoute = pstrRoute

End Property
Public Property Get Route() As String

    Route = mstrRoute
    
End Property
Public Property Let Criteria(pstrCriteria As String)

    mstrCriteria = pstrCriteria

End Property
Public Property Get Criteria() As String

    Criteria = mstrCriteria
    
End Property
Public Property Let CriteriaType(pstrCriteriaType As String)

    mstrCriteriaType = pstrCriteriaType

End Property
Public Property Get CriteriaType() As String

    CriteriaType = mstrCriteriaType
    
End Property

Private Sub cboAccountStatus_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtEmailAddress.SetFocus
    End If
    
End Sub

Private Sub cboAccountType_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cboAccountStatus.SetFocus
    End If
    
End Sub

Private Sub chkReceiveMailings_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        cmdNext.SetFocus
    End If
    
End Sub

Private Sub cmdCopyName_Click()
    
    With gstrCustomerAccount
        .strDeliverySalutation = .strSalutation
        .strDeliveryInitials = .strInitials
        .strDeliverySurname = .strSurname
    End With
    txtDeliverySalutation = txtSalutation
    txtDeliveryInitials = txtInitials
    txtDeliverySurname = txtSurname
    
End Sub

Private Sub cmdHelpWhat_Click()

    Me.WhatsThisMode
    
End Sub

Private Sub cmdNotePad_Click()

    frmChildCuNotes.Show vbModal
    
End Sub
Private Sub cmdAbort_Click()
Dim lintRetVal As Variant
Dim lstrRetAbort As String

    frmChildAbortOptions.Style = 1
    frmChildAbortOptions.Show vbModal
    lstrRetAbort = frmChildAbortOptions.AbortOption
    Unload frmChildAbortOptions
    
    Select Case lstrRetAbort
    Case "BACK"
        Exit Sub
    Case "ABORT"
       
    End Select
    
    lintRetVal = MsgBox("WARNING: By aborting this order no information will be saved!", vbYesNo + vbExclamation, gconstrTitlPrefix & "Abort Order")

    If lintRetVal = vbYes Then
        ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
    
        ClearCustomerAcount
        ClearAdviceNote
        ClearGen
        
        Unload Me
        frmAbout.Show
    
    End If
    
End Sub
Private Sub cmdConvert_Click()

    ConvertContact txtSurname, txtSalutation, txtInitials, txtSurname

End Sub
Private Sub cmdCopy_Click()

    With gstrCustomerAccount
        .strDeliverySalutation = .strSalutation
        .strDeliveryInitials = .strInitials
        .strDeliverySurname = .strSurname
        .strDeliveryAdd1 = .strAdd1
        .strDeliveryAdd2 = .strAdd2
        .strDeliveryAdd3 = .strAdd3
        .strDeliveryAdd4 = .strAdd4
        .strDeliveryAdd5 = .strAdd5
        .strDeliveryPostcode = .strPostcode
    End With
    txtDeliverySalutation = txtSalutation
    txtDeliveryInitials = txtInitials
    txtDeliverySurname = txtSurname
    txtDeliverAddress1 = txtAddress1
    txtDeliverAddress2 = txtAddress2
    txtDeliverAddress3 = txtAddress3
    txtDeliverAddress4 = txtAddress4
    txtDeliverAddress5 = txtAddress5
    txtDeliverPostcode = txtPostcode
    
    If Me.Route <> gconstrOrderModify Then
        cboAccountType.SetFocus
    End If
    
End Sub
Private Sub cmdHelp_Click()

    glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, lstrScreenHelpFile, HH_DISPLAY_TOPIC, 0)
    
End Sub
Public Sub cmdNext_Click()
Dim lstrEmailAddressError As String

    If Trim$(txtDeliverySalutation & txtDeliveryInitials & txtDeliverySurname & txtDeliverAddress1 & txtDeliverAddress2 & _
        txtDeliverAddress3 & txtDeliverAddress4 & _
        txtDeliverAddress5 & txtDeliverPostcode) = "" Then
        cmdCopy_Click
    End If
    
    If Trim$(txtDeliverySalutation & txtDeliveryInitials & txtDeliverySurname) = "" Then
        cmdCopyName_Click
    End If
    
    If cboAccountType.Visible = True And Trim$(cboAccountType) = "" Then
        MsgBox "You must enter the Account Type!", , gconstrTitlPrefix & "Mandatory Field"
        cboAccountType.SetFocus
        Exit Sub
    End If
    
    lstrEmailAddressError = CheckEmailAddress(txtEmailAddress)
    If lstrEmailAddressError <> "" And Trim(txtEmailAddress) <> "" Then
        MsgBox lstrEmailAddressError, , gconstrTitlPrefix & "Mandatory Field"
        txtEmailAddress.SetFocus
        Exit Sub
    End If
    
    SaveLocalFields
        
    Select Case Me.Route
    Case gconstrEntry
        UpdateAccount gstrCustomerAccount.lngCustNum
        Set gstrCurrentLoadedForm = frmOrdDetails
        frmOrdDetails.Route = Me.Route
        Unload Me
        frmOrdDetails.Show

    Case gconstrAccount
        UpdateAccount gstrCustomerAccount.lngCustNum
        ToggleAcctInUseBy gstrCustomerAccount.lngCustNum, False
        Unload Me
        ClearAdviceNote
        ClearCustomerAcount
        ClearGen
        Set gstrCurrentLoadedForm = frmAbout
        frmAbout.Show

    Case gconstrOrderModify
        Set gstrCurrentLoadedForm = frmOrdDetails
        frmOrdDetails.Route = Me.Route
        Unload Me
        frmOrdDetails.Show

    End Select
    
End Sub

Private Sub cmdPAD_Click()

    ShowPOCollectForm
    
End Sub

Private Sub Form_Activate()
Dim lintRetVal As Integer

    If gbooQAOK = True Then
        If mstrCriteriaType = "Postcode" And mstrCriteria <> "" And txtPostcode = mstrCriteria Then
            txtAddress1.SetFocus
            AddressFill txtAddress1
            Wait 1
            SendKeys "{ENTER}"
            SendKeys "" & mstrCriteria & "{ENTER}"
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case (KeyCode)
    Case vbKeyF1
        'Call cmdHelp_Click
    End Select
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
    FillList "Account Type", cboAccountType, lstrAccountType()
    FillList "Account Status", cboAccountStatus, lstrAccountStatus()
    
    GetLocalFields
    
    If UCase$(App.ProductName) <> "CLIENT" Then
        cmdPAD.Visible = False
    Else
        cmdPAD.Visible = True
    End If
    
    Select Case mstrCriteriaType
    Case "Name"
        If Trim$(txtSurname) = "" Then
            txtSurname = mstrCriteria
        End If

    Case "Postcode"
        If Trim$(txtPostcode) = "" Then
            txtPostcode = mstrCriteria
        End If

    Case "Phone" 'Not implemented yet
        If Trim$(txtTelephone) = "" Then
            txtTelephone = mstrCriteria
        End If

    End Select
        
    Select Case Me.Route
    Case gconstrAccount
        lblInfoType.Caption = "Account Address Information"
        cmdNext.Caption = "&Close"
        cboAccountType.Visible = True
        lblAcctType.Visible = True

        chkReceiveMailings.Visible = False
        cboAccountStatus.Visible = True
        lblAccountStatus.Visible = True
        
        lblEveningTelephone.Visible = True
        txtEveTelephone.Visible = True
        lblEmail.Visible = True
        txtEmailAddress.Visible = True
        
        cmdAbort.Visible = False

    Case gconstrEntry
        lblInfoType.Caption = "Advice Note && Account Address Information"
        cmdNext.Caption = "&Next"
        cboAccountType.Visible = True
        lblAcctType.Visible = True

        chkReceiveMailings.Visible = False
        cboAccountStatus.Visible = True
        lblAccountStatus.Visible = True
        cmdAbort.Visible = True
        
        lblEveningTelephone.Visible = True
        txtEveTelephone.Visible = True
        lblEmail.Visible = True
        txtEmailAddress.Visible = True

    Case gconstrOrderModify, gconstrEnquiry
        lblInfoType.Caption = "Advice Note Address Information"
        cmdNext.Caption = "&Next"
        cboAccountType.Visible = False
        lblAcctType.Visible = False
        chkReceiveMailings.Visible = False
        cboAccountStatus.Visible = False
        lblAccountStatus.Visible = False
        cmdAbort.Visible = False
        
        lblEveningTelephone.Visible = False
        txtEveTelephone.Visible = False
        lblEmail.Visible = False
        txtEmailAddress.Visible = False

    End Select
    
    ShowBanner Me, Me.Route
    
    SetupHelpFileReqs
    
End Sub
Sub GetLocalFields()

    Select Case Me.Route
    Case gconstrAccount, gconstrEntry
        With gstrCustomerAccount
            lblCustomerNumber = "M" & Trim$(.lngCustNum)
            txtSalutation = ProperCase$(.strSalutation)
            txtSurname = ProperCase(.strSurname)
            txtInitials = ProperCase(.strInitials)
            txtAddress1 = ProperCase(.strAdd1)
            txtAddress2 = ProperCase(.strAdd2)
            txtAddress3 = ProperCase(.strAdd3)
            txtAddress4 = ProperCase(.strAdd4)
            txtAddress5 = ProperCase(.strAdd5)
            txtPostcode = Trim$(UCase$(.strPostcode))
            txtTelephone = Trim$(FixTel(.strTelephoneNum))
            txtEveTelephone = Trim$(FixTel(.strEveTelephoneNum))
            
            txtDeliverySalutation = ProperCase$(.strDeliverySalutation)
            txtDeliverySurname = ProperCase(.strDeliverySurname)
            txtDeliveryInitials = ProperCase(.strDeliveryInitials)
            
            txtDeliverAddress1 = ProperCase(.strDeliveryAdd1)
            txtDeliverAddress2 = ProperCase(.strDeliveryAdd2)
            txtDeliverAddress3 = ProperCase(.strDeliveryAdd3)
            txtDeliverAddress4 = ProperCase(.strDeliveryAdd4)
            txtDeliverAddress5 = ProperCase(.strDeliveryAdd5)
            txtDeliverPostcode = Trim$(UCase$(.strDeliveryPostcode))
            
            SelectListItem Trim$(.strAccountType), cboAccountType, lstrAccountType()
            
            Select Case UCase$(Trim$(.strReceiveMailings))
            Case "Y"
                chkReceiveMailings.Value = 1 'Checked
            Case "N"
                chkReceiveMailings.Value = 0 'UnChecked
            End Select
                
            SelectListItem Trim$(.strAccountStatus), cboAccountStatus, lstrAccountStatus()
            If Trim$(cboAccountStatus) = "" Then
                SelectListItem "ALL", cboAccountStatus, lstrAccountStatus()
            End If
        
            txtEmailAddress = Trim$(.strEmail)
        End With

    Case gconstrOrderModify, gconstrEnquiry ' you should never get here under Equiry
        With gstrAdviceNoteOrder
            lblCustomerNumber = "M" & Trim$(.lngCustNum)
            txtSalutation = ProperCase$(.strSalutation)
            txtSurname = ProperCase(.strSurname)
            txtInitials = UCase$(.strInitials)
            txtAddress1 = ProperCase(.strAdd1)
            txtAddress2 = ProperCase(.strAdd2)
            txtAddress3 = ProperCase(.strAdd3)
            txtAddress4 = ProperCase(.strAdd4)
            txtAddress5 = ProperCase(.strAdd5)
            txtPostcode = Trim$(UCase$(.strPostcode))
            txtTelephone = Trim$(FixTel(.strTelephoneNum))
            
            txtDeliverySalutation = ProperCase$(.strDeliverySalutation)
            txtDeliverySurname = ProperCase(.strDeliverySurname)
            txtDeliveryInitials = UCase$(.strDeliveryInitials)
            
            txtDeliverAddress1 = ProperCase(.strDeliveryAdd1)
            txtDeliverAddress2 = ProperCase(.strDeliveryAdd2)
            txtDeliverAddress3 = ProperCase(.strDeliveryAdd3)
            txtDeliverAddress4 = ProperCase(.strDeliveryAdd4)
            txtDeliverAddress5 = ProperCase(.strDeliveryAdd5)
            txtDeliverPostcode = Trim$(UCase$(.strDeliveryPostcode))
            
        End With
    End Select
    
End Sub
Sub SaveLocalFields()

    Select Case Me.Route
    Case gconstrAccount, gconstrEntry
        With gstrCustomerAccount
            .strSalutation = ProperCase(txtSalutation)
            .strSurname = ProperCase(txtSurname)
            .strInitials = ProperCase(txtInitials)
            .strAdd1 = ProperCase(txtAddress1)
            .strAdd2 = ProperCase(txtAddress2)
            .strAdd3 = ProperCase(txtAddress3)
            .strAdd4 = ProperCase(txtAddress4)
            .strAdd5 = ProperCase(txtAddress5)
            
            .strPostcode = UCase$(txtPostcode)
            .strTelephoneNum = Trim$(FixTel(txtTelephone))
            .strEveTelephoneNum = Trim$(FixTel(txtEveTelephone))
            
            .strDeliverySalutation = ProperCase$(txtDeliverySalutation)
            .strDeliverySurname = ProperCase(txtDeliverySurname)
            .strDeliveryInitials = ProperCase(txtDeliveryInitials)
            
            .strDeliveryAdd1 = ProperCase(txtDeliverAddress1)
            .strDeliveryAdd2 = ProperCase(txtDeliverAddress2)
            .strDeliveryAdd3 = ProperCase(txtDeliverAddress3)
            .strDeliveryAdd4 = ProperCase(txtDeliverAddress4)
            .strDeliveryAdd5 = ProperCase(txtDeliverAddress5)
            .strDeliveryPostcode = UCase$(txtDeliverPostcode)
            
            .strAccountType = Trim$(NotNull(cboAccountType, lstrAccountType))
            
            Select Case chkReceiveMailings.Value
            Case 0
                .strReceiveMailings = "N"
            Case 1
                .strReceiveMailings = "Y"
            End Select
            
        End With
    Case gconstrOrderModify, gconstrEnquiry ' you should never get here under Equiry
        With gstrAdviceNoteOrder
            .strSalutation = ProperCase(txtSalutation)
            .strSurname = ProperCase(txtSurname)
            .strInitials = UCase$(txtInitials)
            .strAdd1 = ProperCase(txtAddress1)
            .strAdd2 = ProperCase(txtAddress2)
            .strAdd3 = ProperCase(txtAddress3)
            .strAdd4 = ProperCase(txtAddress4)
            .strAdd5 = ProperCase(txtAddress5)
            
            .strPostcode = UCase$(txtPostcode)
            .strTelephoneNum = Trim$(FixTel(txtTelephone))
            '.strEveTelephoneNum = Trim$(FixTel(txtEveTelephone)) 
            
            .strDeliverySalutation = ProperCase$(txtDeliverySalutation)
            .strDeliverySurname = ProperCase(txtDeliverySurname)
            .strDeliveryInitials = UCase$(txtDeliveryInitials)
            
            .strDeliveryAdd1 = ProperCase(txtDeliverAddress1)
            .strDeliveryAdd2 = ProperCase(txtDeliverAddress2)
            .strDeliveryAdd3 = ProperCase(txtDeliverAddress3)
            .strDeliveryAdd4 = ProperCase(txtDeliverAddress4)
            .strDeliveryAdd5 = ProperCase(txtDeliverAddress5)
            .strDeliveryPostcode = UCase$(txtDeliverPostcode)
            
        End With
    End Select
    
End Sub
Private Sub Form_Paint()

    RefreshMenu Me
    
End Sub
Private Sub Form_Resize()
Dim llngFormHalfWidth As Long
    
    On Error Resume Next
    
    llngFormHalfWidth = Me.Width / 2
    
    With cmdNext
        .Top = Me.Height - gconlongButtonTop
        '.Left = llngFormHalfWidth - 240 'Me.Width - 1545
        .Left = (Me.Width - .Width) - 180
    End With
    
    With cmdHelpWhat
        .Top = Me.Height - gconlongButtonTop
        .Left = 120
    End With
    
    With cmdHelp
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelpWhat.Left + cmdHelpWhat.Width + 105
    End With
    
    With cmdAbort
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdHelp.Width + cmdHelp.Left + 120
    End With
    
    With cmdNotePad
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdAbort.Width + cmdAbort.Left + 120
    End With
    
    With cmdPAD
        .Top = Me.Height - gconlongButtonTop
        .Left = cmdNotePad.Width + cmdNotePad.Left + 120
    End With
    
    With lblInfoType
        .Left = 0
        .Width = Me.Width
    End With
    
    With Frame1
        .Width = llngFormHalfWidth - 240
    End With
        
    With Frame2
        .Left = llngFormHalfWidth + 60 '120
        .Width = llngFormHalfWidth - 240
    End With

    With Frame3
        .Left = Frame2.Left
    End With
    
    With chkReceiveMailings
        .Left = Frame2.Left + 38
    End With
    
    With lblCustomerNumber
        .Left = Me.Width - (.Width + 180)
    End With
    
    With lblCustNum
        .Left = ((lblCustomerNumber.Left - .Width) - 120)
    End With

End Sub
Private Sub timActivity_Timer()

    CheckActivity
    
End Sub
Private Sub txtAddress1_GotFocus()
    
    SetSelected Me
    If gbooQAOK = True Then
        ShowStatus 39
    End If
    
End Sub
Private Sub txtAddress1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        KeyCode = 0
        AddressFill txtAddress1
    End If

    If KeyCode = 13 Then 'Carriage return
        txtAddress2.SetFocus
    End If
    
End Sub
Private Sub txtAddress1_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtAddress1_LostFocus()
    
    txtAddress1 = ProperCase(txtAddress1)
    
End Sub
Private Sub txtAddress2_GotFocus()
    
    SetSelected Me

End Sub

Private Sub txtAddress2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtAddress3.SetFocus
    End If
    
End Sub

Private Sub txtAddress2_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtAddress2_LostFocus()

        txtAddress2 = ProperCase(txtAddress2)

End Sub
Private Sub txtAddress3_GotFocus()
    
    SetSelected Me

End Sub

Private Sub txtAddress3_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtAddress4.SetFocus
    End If
    
End Sub

Private Sub txtAddress3_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtAddress3_LostFocus()

    txtAddress3 = ProperCase(txtAddress3)

End Sub
Private Sub txtAddress4_GotFocus()
    
    SetSelected Me

End Sub

Private Sub txtAddress4_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtAddress5.SetFocus
    End If
    
End Sub

Private Sub txtAddress4_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtAddress4_LostFocus()

    txtAddress4 = ProperCase(txtAddress4)

End Sub
Private Sub txtAddress5_GotFocus()
    
    SetSelected Me

End Sub

Private Sub txtAddress5_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtPostcode.SetFocus
    End If
    
End Sub

Private Sub txtAddress5_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtAddress5_LostFocus()

    txtAddress5 = ProperCase(txtAddress5)

End Sub
Private Sub txtDeliverAddress1_GotFocus()

    SetSelected Me
    If gbooQAOK = True Then
        ShowStatus 39
    End If
End Sub

Private Sub txtDeliverAddress1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtDeliverAddress2.SetFocus
    End If
    
End Sub

Private Sub txtDeliverAddress1_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtDeliverAddress1_LostFocus()

    txtDeliverAddress1 = ProperCase(txtDeliverAddress1)

End Sub
Private Sub txtDeliverAddress2_GotFocus()

    SetSelected Me

End Sub

Private Sub txtDeliverAddress2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtDeliverAddress3.SetFocus
    End If
    
End Sub

Private Sub txtDeliverAddress2_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtDeliverAddress2_LostFocus()

    txtDeliverAddress2 = ProperCase(txtDeliverAddress2)

End Sub
Private Sub txtDeliverAddress3_GotFocus()

    SetSelected Me

End Sub

Private Sub txtDeliverAddress3_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtDeliverAddress4.SetFocus
    End If
    
End Sub

Private Sub txtDeliverAddress3_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtDeliverAddress3_LostFocus()

    txtDeliverAddress3 = ProperCase(txtDeliverAddress3)

End Sub
Private Sub txtDeliverAddress4_GotFocus()

    SetSelected Me

End Sub

Private Sub txtDeliverAddress4_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtDeliverAddress5.SetFocus
    End If
    
End Sub

Private Sub txtDeliverAddress4_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtDeliverAddress4_LostFocus()

    txtDeliverAddress4 = ProperCase(txtDeliverAddress4)

End Sub
Private Sub txtDeliverAddress5_GotFocus()

    SetSelected Me

End Sub

Private Sub txtDeliverAddress5_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtDeliverPostcode.SetFocus
    End If
    
End Sub

Private Sub txtDeliverAddress5_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtDeliverAddress5_LostFocus()
    
    txtDeliverAddress5 = ProperCase(txtDeliverAddress5)

End Sub
Private Sub txtDeliverPostcode_GotFocus()
    
    SetSelected Me

End Sub

Private Sub txtDeliverPostcode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        If cboAccountType.Visible = True Then
            cboAccountType.SetFocus
        Else
            cmdNext.SetFocus
        End If
    End If
    
End Sub

Private Sub txtDeliverPostcode_KeyPress(KeyAscii As Integer)
    
    KeyAscii = CheckKeyAsciiValidPostCode(KeyAscii)
    
End Sub
Private Sub txtDeliverPostcode_LostFocus()

    txtDeliverPostcode = ProperCase(txtDeliverPostcode)

End Sub
Private Sub txtDeliveryInitials_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtDeliveryInitials_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtDeliverySurname.SetFocus
    End If
    
End Sub

Private Sub txtDeliveryInitials_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtDeliveryInitials_LostFocus()

    txtDeliveryInitials = ProperCase(txtDeliveryInitials)
    
End Sub
Private Sub txtDeliverySalutation_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtDeliverySalutation_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtDeliveryInitials.SetFocus
    End If
    
End Sub

Private Sub txtDeliverySalutation_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtDeliverySalutation_LostFocus()

    txtDeliverySalutation = ProperCase(txtDeliverySalutation)
    
End Sub
Private Sub txtDeliverySurname_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtDeliverySurname_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtDeliverAddress1.SetFocus
    End If
    
End Sub

Private Sub txtDeliverySurname_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtDeliverySurname_LostFocus()

    txtDeliverySurname = ProperCase(txtDeliverySurname)
    
End Sub
Private Sub txtEmailAddress_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtEmailAddress_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        If chkReceiveMailings.Visible = True Then
            chkReceiveMailings.SetFocus
        Else
            cmdNext.SetFocus
        End If
    End If
    
End Sub

Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidEmail(KeyAscii)
    
End Sub
Private Sub txtEveTelephone_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtEveTelephone_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        'cmdCopyName.SetFocus
        txtDeliverySalutation.SetFocus
    End If
    
End Sub

Private Sub txtEveTelephone_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtEveTelephone_LostFocus()

    txtEveTelephone = Trim$(FixTel(txtEveTelephone))
    
End Sub
Private Sub txtInitials_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtInitials_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtSurname.SetFocus
    End If
    
End Sub

Private Sub txtInitials_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtInitials_LostFocus()

    txtInitials = ProperCase(txtInitials)
    
End Sub
Private Sub txtPostCode_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtPostcode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtTelephone.SetFocus
    End If
    
End Sub

Private Sub txtPostcode_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValidPostCode(KeyAscii)
    
End Sub
Private Sub txtPostcode_LostFocus()
    
    txtPostcode = UCase$(txtPostcode)
    
End Sub
Private Sub txtSalutation_GotFocus()

    SetSelected Me
        
End Sub

Private Sub txtSalutation_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtInitials.SetFocus
    End If
    
End Sub

Private Sub txtSalutation_KeyPress(KeyAscii As Integer)
    
    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtSalutation_LostFocus()

    txtSalutation = ProperCase(txtSalutation)
    
End Sub
Private Sub txtSurname_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtSurname_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        txtAddress1.SetFocus
    End If
    
End Sub

Private Sub txtSurname_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtSurname_LostFocus()

    txtSurname = ProperCase(txtSurname)
    
End Sub
Private Sub txtTelephone_GotFocus()

    SetSelected Me
    
End Sub

Private Sub txtTelephone_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then 'Carriage return
        If txtEveTelephone.Visible = True Then
            txtEveTelephone.SetFocus
        Else
            txtDeliverySalutation.SetFocus
        End If
    End If
    
End Sub

Private Sub txtTelephone_KeyPress(KeyAscii As Integer)

    KeyAscii = CheckKeyAsciiValid(KeyAscii)
    
End Sub
Private Sub txtTelephone_LostFocus()

    txtTelephone = Trim$(FixTel(txtTelephone))
    
End Sub
Public Function FindFieldError() As Boolean
Dim lstrEmailAddressError As String

    FindFieldError = False
    
    If Trim$(txtDeliverySalutation & txtDeliveryInitials & txtDeliverySurname & txtDeliverAddress1 & txtDeliverAddress2 & _
        txtDeliverAddress3 & txtDeliverAddress4 & _
        txtDeliverAddress5 & txtDeliverPostcode) = "" Then
        
        cmdCopy_Click
    End If
    
    If Trim$(txtDeliverySalutation & txtDeliveryInitials & txtDeliverySurname) = "" Then
        cmdCopyName_Click
    End If
    
    If cboAccountType.Visible = True And Trim$(cboAccountType) = "" Then
        MsgBox "You must enter the Account Type!", , gconstrTitlPrefix & "Mandatory Field"
        FindFieldError = True
        Exit Function
    End If
    
    lstrEmailAddressError = CheckEmailAddress(txtEmailAddress)
    If lstrEmailAddressError <> "" And Trim(txtEmailAddress) <> "" Then
        MsgBox lstrEmailAddressError, , gconstrTitlPrefix & "Mandatory Field"
        FindFieldError = True
        Exit Function
    End If
    
End Function
Sub SetupHelpFileReqs()

    App.HelpFile = gstrHelpFileBase & gconstrHelpPopupFileParam
    cmdHelpWhat.Picture = frmButtons.imgHelpWhatsThis.Picture
    lstrScreenHelpFile = gstrHelpFileBase & "::/Account.xml>WhatsScreen"
    
    ctlBanner1.WhatsThisHelpID = IDH_ACCT_MAIN
    ctlBanner1.WhatIsID = IDH_ACCT_MAIN
    
    ctlBottomLine1.WhatsThisHelpID = IDH_ACCT_MAIN
    
    cmdCopyName.WhatsThisHelpID = IDH_ACCT_COPY
    cmdCopy.WhatsThisHelpID = IDH_ACCT_COPYALL
    
    cmdHelp.WhatsThisHelpID = IDH_STANDARD_HELP
    cmdNext.WhatsThisHelpID = IDH_STANDARD_NEXT
    cmdAbort.WhatsThisHelpID = IDH_STANDARD_ABORT
    cmdNotePad.WhatsThisHelpID = IDH_STANDARD_CUNOTES
    
    txtSalutation.WhatsThisHelpID = IDH_ACCT_SALU
    txtInitials.WhatsThisHelpID = IDH_ACCT_FIRST
    txtSurname.WhatsThisHelpID = IDH_ACCT_SURN
    txtAddress1.WhatsThisHelpID = IDH_ACCT_ADDR
    txtAddress2.WhatsThisHelpID = IDH_ACCT_ADDR
    txtAddress3.WhatsThisHelpID = IDH_ACCT_ADDR
    txtAddress4.WhatsThisHelpID = IDH_ACCT_ADDR
    txtAddress5.WhatsThisHelpID = IDH_ACCT_ADDR
    txtPostcode.WhatsThisHelpID = IDH_ACCT_POSTC
    txtDeliverySalutation.WhatsThisHelpID = IDH_ACCT_DEL_SALU
    txtDeliveryInitials.WhatsThisHelpID = IDH_ACCT_DEL_FIRST
    txtDeliverySurname.WhatsThisHelpID = IDH_ACCT_DEL_SURN
    txtDeliverAddress1.WhatsThisHelpID = IDH_ACCT_DEL_ADDR
    txtDeliverAddress2.WhatsThisHelpID = IDH_ACCT_DEL_ADDR
    txtDeliverAddress3.WhatsThisHelpID = IDH_ACCT_DEL_ADDR
    txtDeliverAddress4.WhatsThisHelpID = IDH_ACCT_DEL_ADDR
    txtDeliverAddress5.WhatsThisHelpID = IDH_ACCT_DEL_ADDR
    txtDeliverPostcode.WhatsThisHelpID = IDH_ACCT_DEL_POSTC
    
    txtTelephone.WhatsThisHelpID = IDH_ACCT_TEL
    txtEveTelephone.WhatsThisHelpID = IDH_ACCT_EVETEL
    
    cboAccountType.WhatsThisHelpID = IDH_ACCT_ACCTTYPE
    cboAccountStatus.WhatsThisHelpID = IDH_ACCT_ACTSTAT
    txtEmailAddress.WhatsThisHelpID = IDH_ACCT_EMAIL
    
End Sub
