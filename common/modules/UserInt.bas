Attribute VB_Name = "modUserInterface"
Option Explicit

Global Const IDH_STANDARD_HELP = 10000
Global Const IDH_STANDARD_BACK = 10010
Global Const IDH_STANDARD_PREVIOUS = 10020
Global Const IDH_STANDARD_ABORT = 10030
Global Const IDH_STANDARD_NEXT = 10040
Global Const IDH_STANDARD_CUNOTES = 10050
Global Const IDH_STANDARD_CONSNOTE = 10060
Global Const IDH_STANDARD_FIND = 10070
Global Const IDH_STANDARD_VWADNOT = 10080
Global Const IDH_STANDARD_LBLGRITOTFOUND = 10090
Global Const IDH_STANDARD_SEARCHBY = 10100
Global Const IDH_STANDARD_CHICASHBOOK = 10110

Global Const IDH_CAS_MAIN = 11000
Global Const IDH_CAS_ADDNEW = 11010
Global Const IDH_CAS_SELECT = 11020
Global Const IDH_CAS_TXTCNUM = 11030
Global Const IDH_CAS_TXTSURN = 11040
Global Const IDH_CAS_TXTPOST = 11050
Global Const IDH_CAS_LSTACCT = 11060
Global Const IDH_CAS_FIND = 11070

Global Const IDH_ACCT_MAIN = 12000
Global Const IDH_ACCT_COPY = 12010
Global Const IDH_ACCT_COPYALL = 12020
Global Const IDH_ACCT_SALU = 12030
Global Const IDH_ACCT_DEL_SALU = 12040
Global Const IDH_ACCT_FIRST = 12050
Global Const IDH_ACCT_DEL_FIRST = 12060
Global Const IDH_ACCT_SURN = 12070
Global Const IDH_ACCT_DEL_SURN = 12080
Global Const IDH_ACCT_ADDR = 12090
Global Const IDH_ACCT_DEL_ADDR = 12100
Global Const IDH_ACCT_POSTC = 12110
Global Const IDH_ACCT_DEL_POSTC = 12120
Global Const IDH_ACCT_TEL = 12130
Global Const IDH_ACCT_EVETEL = 12140
Global Const IDH_ACCT_ACCTTYPE = 12150
Global Const IDH_ACCT_ACTSTAT = 12160
Global Const IDH_ACCT_EMAIL = 12170

Global Const IDH_ORDDETS_MAIN = 14000
Global Const IDH_ORDDETS_ORDSTY = 14010
Global Const IDH_ORDDETS_MEDIA = 14020
Global Const IDH_ORDDETS_DELDATE = 14030
Global Const IDH_ORDDETS_COURIER = 14040
Global Const IDH_ORDDETS_PAYTY1 = 14050
Global Const IDH_ORDDETS_PAYTY2 = 14060
Global Const IDH_ORDDETS_CCNUM = 14070
Global Const IDH_ORDDETS_EXPMONTH = 14080
Global Const IDH_ORDDETS_EXPYEAR = 14090
Global Const IDH_ORDDETS_CARDNAME = 14100
Global Const IDH_ORDDETS_CARDTYPE = 14110
Global Const IDH_ORDDETS_CARDISS = 14120
Global Const IDH_ORDDETS_VALMONTH = 14130
Global Const IDH_ORDDETS_VALYEAR = 14140
Global Const IDH_ORDDETS_ORDCODE = 14150
Global Const IDH_ORDDETS_OVERSEAS = 14160

Global Const IDH_ORDER_MAIN = 15000
Global Const IDH_ORDER_ADDPRODS = 15010
Global Const IDH_ORDER_REFRESH = 15020
Global Const IDH_ORDER_POSTAGE = 15030
Global Const IDH_ORDER_UNDERPAY = 15040
Global Const IDH_ORDER_GRIDOL = 15060
Global Const IDH_ORDER_PAY1 = 15070
Global Const IDH_ORDER_PAY2 = 15080
Global Const IDH_ORDER_LBTOTPAY = 15090
Global Const IDH_ORDER_VAT = 15100
Global Const IDH_ORDER_LBLORDTOT = 15110
Global Const IDH_ORDER_TXTPOST = 15120
Global Const IDH_ORDER_TOTINCVAT = 15130
Global Const IDH_ORDER_DONAT = 15140
Global Const IDH_ORDER_RECON = 15150
Global Const IDH_ORDER_TXTUNDERP = 15160
Global Const IDH_ORDER_LBLCASH = 15170

Global Const IDH_CHIGENOPS_LIST = 16000

Global Const IDH_CHIPRODS_MAIN = 17000
Global Const IDH_CHIPRODS_GRIDPRODS = 17010
Global Const IDH_CHIPRODS_SEARCHCRIT = 17020
    
Global Const IDH_ORDHIST_MAIN = 18000
Global Const IDH_ORDHIST_MODIFY = 18010
Global Const IDH_ORDHIST_CONSIGNDETSVIEW = 18020
Global Const IDH_ORDHIST_GRIDADV = 18030
Global Const IDH_ORDHIST_GRIDOLM = 18040

Global Const IDH_FINANCE_MAIN = 19000
Global Const IDH_FINANCE_SEARCHCRIT = 19010

Global Const IDH_PACKING_MAIN = 20000
Global Const IDH_PACKING_CONFIRM = 20010
Global Const IDH_PACKING_GRIDOLM = 20020
Global Const IDH_PACKING_SEACRIT = 20030

Global Const IDH_ORDMNT_MAIN = 21000
Global Const IDH_ORDMNT_REFUND = 21010
Global Const IDH_ORDMNT_SORTBY = 21020
Global Const IDH_ORDMNT_GRIDADV = 21030
Global Const IDH_ORDMNT_SEARCHCRIT = 21040
    
'Admin
Global Const IDH_LABLAY_MAIN = 30000
Global Const IDH_LABLAY_LABTYPE = 30010
Global Const IDH_LABLAY_LABSACROSS = 30020
Global Const IDH_LABLAY_LADSDOWN = 30030
Global Const IDH_LABLAY_LINESBETW = 30040
Global Const IDH_LABLAY_CHARSLEFTORI = 30050
Global Const IDH_LABLAY_TOPMARG = 30060
Global Const IDH_LABLAY_LEFTMARG = 30070

Global Const IDH_REFDATA_MAIN = 31000
Global Const IDH_STOCKVIEW_MAIN = 32000
Global Const IDH_STOCKVIEW_GRIDPRODS = 32010
Global Const IDH_STOCKVIEW_UPDATESTOCK = 32020
Global Const IDH_STOCKVIEW_EXPORTSTOCK = 32030
Global Const IDH_STOCKVIEW_UPDATESUBSTS = 32040

Global Const IDH_SYSOPS_MAIN = 33000
Global Const IDH_SYSOPS_UNLOCK = 33010

Global Const IDH_UPGRADE_MAIN = 34000
Global Const IDH_UPGRADE_ZIPFILE = 34010
Global Const IDH_UPGRADE_BROWSE = 34020
Global Const IDH_UPGRADE_TESTPROGS = 34030
Global Const IDH_UPGRADE_UPGRADE = 34040
Global Const IDH_UPGRADE_REDEPLOY = 34050
Global Const IDH_UPGRADE_OLDPROGS = 34060
Global Const IDH_UPGRADE_REVERT = 34070
Global Const IDH_UPGRADE_USERS = 34080
Global Const IDH_UPGRADE_REFRESH = 34090
Global Const IDH_UPGRADE_COMPACT = 34100
Global Const IDH_UPGRADE_DBCHECK = 34110
Global Const IDH_UPGRADE_DEPLSUPPS = 34120

Global Const IDH_USERS_MAIN = 35000
Global Const IDH_USERS_GRIDUSERS = 35010
Global Const IDH_USERS_NEWPASS = 35020
Global Const IDH_USERS_CONFPASS = 35030
Global Const IDH_USERS_CHANGEPASS = 35040

'Reps
Global Const IDH_CHILABOPS_LABTYPE = 50010
Global Const IDH_CHILABOPS_LABSACROSS = 50020
Global Const IDH_CHILABOPS_LADSDOWN = 50030
Global Const IDH_CHILABOPS_LINESBETW = 50040
Global Const IDH_CHILABOPS_CHARSLEFTORI = 50050
Global Const IDH_CHILABOPS_TOPMARG = 50060
Global Const IDH_CHILABOPS_LEFTMARG = 50070

Global Const IDH_DEDUPE_MAIN = 51000
Global Const IDH_DEDUPE_SEACRIT = 51010
Global Const IDH_DEDUPE_SEATYPE = 51020
Global Const IDH_DEDUPE_GRIDDUPS = 51030
Global Const IDH_DEDUPE_MERGE = 51040
Global Const IDH_DEDUPE_CLEAR = 51050

Global Const IDH_PFORCE_MAIN = 52000
Global Const IDH_PFORCE_SEATCRIT = 52010
Global Const IDH_PFORCE_GRIDCONS = 52020
Global Const IDH_PFORCE_REPRINT = 52030
Global Const IDH_PFORCE_CREATEEFILE = 52040
Global Const IDH_PFORCE_PRINTMANI = 52050
Global Const IDH_PFORCE_DUMP = 52060
Global Const IDH_PFORCE_PRINTTHERM = 52070

Global Const IDH_GENREPS_MAIN = 53000
Global Const IDH_GENREPS_STARTDATE = 53010
Global Const IDH_GENREPS_ENDDATE = 53020
Global Const IDH_GENREPS_RANGEORDSTART = 53030
Global Const IDH_GENREPS_RANGEORDCMDSTART = 53040
Global Const IDH_GENREPS_RANDEORDEND = 53050
Global Const IDH_GENREPS_RANGEORDCMDEND = 53060
Global Const IDH_GENREPS_REPSTABS = 53070
Global Const IDH_GENREPS_REPLIST = 53080
Global Const IDH_GENREPS_CMDSELECT = 53090
Global Const IDH_GENREPS_OUTPUT = 53100
Global Const IDH_GENREPS_PRINTSET = 53110

Global Const IDH_SUMRY_MAIN = 54000

'Training card pass through params
Const mTraLiteGetStad = 10000
Const mTraLiteGetStadValues = 10030
Const mTraLiteGetStadProds = 10040
Const mTraLiteGetStadEssen = 10050

Global glngCurrentHelpHandle As Long

Global gstrHelpFileBase As String

Global Const gconstrHelpPopupFileParam = "::/textpopups.txt"
'
Global Const mnuClientGoOrderEntry = "&Order Entry" & vbTab & "Ctrl+Shift+O"
Global Const mnuClientGoEnquiry = "Order &Enquiry" & vbTab & "Ctrl+Shift+E"
Global Const mnuClientGoAcctMaint = "&Account Maintenance" & vbTab & "Ctrl+Shift+A"
Global Const mnuClientGoFinance = "&Finance" & vbTab & "Ctrl+Shift+F"
Global Const mnuClientGoPacking = "&Packing" & vbTab & "Ctrl+Shift+P"
Global Const mnuClientGoOrderMaint = "Order &Maintenance" & vbTab & "Ctrl+Shift+M"

Global Const mnuMgrGoGenReps = "General Reporting"
Global Const mnuMgrGoDistribution = "Distribution"
Global Const mnuMgrGoMarkSets = "Marketing Settings"
Global Const mnuMgrGoAgentReps = "Agent Reporting"
Global Const mnuMgrGoSumInfo = "Summary Info"
Global Const mnuMgrGoDupHand = "Duplicate Handling"
        
Global Const mnuMaintGoRefData = "Reference Data"
Global Const mnuMaintGoStockMan = "Stock Management"
Global Const mnuMaintGoSysOps = "System Options"
Global Const mnuMaintGoSysMain = "System Maintenance"
                
'Global gbooForceAppClose As Boolean
Enum ForceClose
    fcDontClose = 0
    fcCompleteClose = 1
    fcCloseKeepDB = 2
End Enum

Global gintForceAppClose As ForceClose
Global Const mnuFileHistoryOptionModify = "&Modify This Order"
Global Const mnuFileHistoryOptionModifyItem = 0
Global Const mnuFileHistoryOptionHistory = "Orders &History"
Global Const mnuFileHistoryOptionHistoryItem = 1
Global Const mnuFileHistoryOptionPack = "&Pack This Order"
Global Const mnuFileHistoryOptionPackItem = 2

Dim mstrControls() As String 'Object
Global mlngUImnuFilePageSetup   As Long

Type HistroryValue
    lngCustNum  As Long
    lngOrderNum As Long
    strCustName As String
    strType As String
End Type

Type FileHistoryItems
    strItemValue As String
    lngCustNum As Long
    lngOrderNum As Long
End Type
Global mudtUImnuFileHistory(5)     As FileHistoryItems

Global gstrTempKeyFail As String
Global gbooJustPreLoading       As Boolean

Global gstrUserMode As String
Global Const gconstrTestingMode = "TESTING"
Global Const gconstrLiveMode = "LIVE"

'
Global glngNumOfColours As Single

Type TableAndFields
    strSourceTable As String
    strAllowZeroLength As String
    strType As String
    strName As String
    strSize As String
    strDataUpdatable As String
    strDefaultValue As String
    strRequired As String
End Type

'General term used throughout progs
Global Const gconstrMainMenu = "MAINMENU"

'Client use
Global Const gconstrEnquiry = "ENQUIRY"
Global Const gconstrEntry = "ENTRY"
Global Const gconstrOrderModify = "OrderModify"
Global Const gconstrAccount = "ACCOUNTMAINT"
Global Const gconstrPacking = "PACKING"
Global Const gconstrOrdMaint = "ORDMAINT"
Global Const gconstrFinance = "FINANCE"

'Lite Only
Global Const gconstrLiteValues = "LITEVALUES"
Global Const gconstrLiteProducts = "LITEPRODUCTS"
Global Const gconstrLiteEssential = "LITEESSENTIAL"

'Admin Use
Global Const gconstrReferenceData = "REFDATA"
Global Const gconstrStockManagement = "STOCKMAN"
Global Const gconstrSystemOptions = "SYSOPTS"
Global Const gconstrSystemManagement = "SYSMAN"
Global Const gconstrAdminRoute = "ADMINROUTE"
Global Const gconstrConfigRoute = "CONFIGROUTE"
Global gstrRefDataSubTitle1 As String
Global gstrRefDataSubTitle2 As String
Global gintRefDataSubButton As Integer

'Reporting
Global Const gconstrGenralReporting = "GENREPS"
Global Const gconstrMarketingData = "MARKETDAT"
Global Const gconstrAgentReporting = "AGENTREP"
Global Const gconstrSummaryInfo = "SUMINFO"
Global Const gconstrDuplicateHandling = "DUPHAND"

Global Const gconstrThermalPrintRun = "THERMALRUN"
Global Const gconstrConsignmentNorm = "CONSIGNNORM"

'Configure
Global Const gconstrConfigNetInstall = "CONFNETINSTALL"
Global Const gconstrConfigFilesPaths = "CONFFILESPATHS"
Global Const gconstrConfigTables = "CONFTABLES"

'----MDI Form constants
Global gconUITopPos As Long
Global Const gconUILeftPos = 250
Global Const gconUIGap = 500
Global Const gconUIHeight = 700
Global Const gconUIWidth = 700

Global gintUILastButtonHighlighted As Integer

'Scroll Buttons
Global Const gconUISpace = 15
Global Const gconUIPosLeftFromEdgeScroll = 150
Global Const gconUIPosBottomFromEdgeScroll = 150
Global Const gconUIButtonHeight = 250
Global Const gconUIButtonWidth = 250
Global gbooUIScrollButtonClicked As Boolean

Global Const gconUIButtonTopPosDefault = 250
Global Const gconUIBottomPosDefault = 250

Global glngUIBottomPlottedPos As Long

Global gintUINumberofButtonsDraw As Integer
Global gstrUILastButtonRoute As String

'----MDI Form constants
Sub UnloadLastForm()

    Unload gstrCurrentLoadedForm
    
End Sub
Sub ShowBox(pobjObject As Form, pintBoxIndex As Integer, booShadow As Boolean, pbooVisible As Boolean)
Dim llngCurrentTop As Long
Dim llngExtraHeightAdj As Long
Dim llngNEColour As Long
Dim llngSWColour As Long

    If pbooVisible = False Then
        llngNEColour = pobjObject.picListBar.BackColor
        llngSWColour = pobjObject.picListBar.BackColor
    Else
        If booShadow = True Then
            llngNEColour = vbBlack
            llngSWColour = vbWhite
        Else
            llngNEColour = vbWhite
            llngSWColour = vbBlack
        
        End If
    End If
    
    llngExtraHeightAdj = (gconUIHeight + gconUIGap) * pintBoxIndex
    
    llngCurrentTop = gconUITopPos + llngExtraHeightAdj
    
    'Top
    pobjObject.picListBar.Line (gconUILeftPos, llngCurrentTop)-(gconUILeftPos + gconUIWidth, llngCurrentTop), llngNEColour
    'Left
    pobjObject.picListBar.Line (gconUILeftPos, llngCurrentTop)-(gconUILeftPos, llngCurrentTop + gconUIHeight), llngNEColour
    'Bottom
    pobjObject.picListBar.Line (gconUILeftPos, llngCurrentTop + gconUIHeight)-(gconUILeftPos + gconUIWidth, llngCurrentTop + gconUIHeight), llngSWColour
    'Right
    pobjObject.picListBar.Line (gconUILeftPos + gconUIWidth, llngCurrentTop)-(gconUILeftPos + gconUIWidth, llngCurrentTop + gconUIHeight), llngSWColour
            
End Sub
Sub DrawButton(pobjObject As Form, pintBoxIndex As Integer, _
    plngAdjustment As Long, pintPicIndex As Integer, _
    pstrCaption1 As String, Optional pstrCaption2 As String)

Dim llngCurrentTop As Long
Dim llngExtraHeightAdj As Long
Dim llngLeftPos As Long

    'Write Caption
    llngExtraHeightAdj = (gconUIHeight + gconUIGap) * pintBoxIndex
    
    llngCurrentTop = gconUITopPos + llngExtraHeightAdj
    llngCurrentTop = llngCurrentTop + gconUIHeight '+ 50
    
    llngLeftPos = ((gconUIWidth / 2) - (pobjObject.picListBar.TextWidth(pstrCaption1) / 2)) + gconUILeftPos
        
    pobjObject.picListBar.CurrentY = llngCurrentTop
    pobjObject.picListBar.CurrentX = llngLeftPos
    pobjObject.picListBar.ForeColor = vbWindowBackground
    pobjObject.picListBar.Print pstrCaption1

    llngLeftPos = ((gconUIWidth / 2) - (pobjObject.picListBar.TextWidth(pstrCaption2) / 2)) + gconUILeftPos
    pobjObject.picListBar.CurrentX = llngLeftPos
    pobjObject.picListBar.Print pstrCaption2

    'Paint Icon
    llngExtraHeightAdj = (gconUIHeight + gconUIGap) * pintBoxIndex
    llngCurrentTop = gconUITopPos + llngExtraHeightAdj
    
    On Error Resume Next
    Dim lintStyle As Integer
    
    If glngNumOfColours < 257 Then
         lintStyle = 2
    Else
        lintStyle = 0
    End If

    frmButtons.ImageList16Cols.UseMaskColor = True
    frmButtons.ImageList16Cols.MaskColor = vbRed
    frmButtons.ImageList16Cols.BackColor = pobjObject.picListBar.BackColor

    frmButtons.ImageList16Cols.ListImages(pintPicIndex).Draw pobjObject.picListBar.hdc, _
    gconUILeftPos + plngAdjustment, llngCurrentTop + plngAdjustment, lintStyle

End Sub
Sub DrawScrollButton(pobjObject As Form, pbooDownButton As Boolean, pbooDepressed As Boolean)
Dim llngCurrentTop As Long
Dim llngCurrentLeft As Long
Dim llngNEColour As Long
Dim llngSWColour As Long
Dim llngTriangleTop As Long

    If pbooDepressed = True Then
        llngNEColour = vbBlack
        llngSWColour = vbWhite
    Else
        llngNEColour = vbWhite
        llngSWColour = vbBlack
    End If
    
    If pbooDownButton = True Then
        llngCurrentTop = (pobjObject.picListBar.Height - gconUIButtonHeight) - gconUIPosBottomFromEdgeScroll
    Else
        llngCurrentTop = gconUIPosBottomFromEdgeScroll - 70
    End If
    llngCurrentLeft = (pobjObject.picListBar.Width - gconUIButtonWidth) - gconUIPosLeftFromEdgeScroll
    
    'Top
    pobjObject.picListBar.Line (llngCurrentLeft, llngCurrentTop)- _
        (llngCurrentLeft + gconUIButtonWidth, llngCurrentTop), llngNEColour
    'Left
    pobjObject.picListBar.Line (llngCurrentLeft, llngCurrentTop)- _
        (llngCurrentLeft, llngCurrentTop + gconUIButtonHeight), llngNEColour
    'Bottom
    pobjObject.picListBar.Line (llngCurrentLeft, llngCurrentTop + gconUIButtonHeight)- _
        (llngCurrentLeft + gconUIButtonWidth, llngCurrentTop + gconUIButtonHeight), llngSWColour
    'Right
    pobjObject.picListBar.Line (llngCurrentLeft + gconUIButtonWidth, llngCurrentTop)- _
        (llngCurrentLeft + gconUIButtonWidth, llngCurrentTop + gconUIButtonHeight), llngSWColour
    
    'Grey inner Box
    pobjObject.picListBar.Line (llngCurrentLeft + gconUISpace, llngCurrentTop + gconUISpace)- _
        ((llngCurrentLeft + gconUIButtonWidth) - (gconUISpace * 2), _
            (llngCurrentTop + gconUIButtonHeight) - (gconUISpace * 1)), vbButtonFace, BF
    
    llngCurrentLeft = llngCurrentLeft - 3
    'Black Triangle
    'Must Move slightly when depressed
    If pbooDownButton = True Then
        llngTriangleTop = llngCurrentTop + 85
        Triangle pobjObject.picListBar, (llngCurrentLeft + 50), (llngTriangleTop), _
            (llngCurrentLeft + 199), (llngTriangleTop), _
            (llngCurrentLeft + 125), (llngTriangleTop + 75), True
    Else
        llngTriangleTop = llngCurrentTop + 160
        Triangle pobjObject.picListBar, (llngCurrentLeft + 50), (llngTriangleTop), _
            (llngCurrentLeft + 199), (llngTriangleTop), _
            (llngCurrentLeft + 125), (llngTriangleTop - 75), False
    
    End If

End Sub
Sub ClickScrollButton(pobjObject As Form, pbooDownButton As Boolean)
Dim llngDownVar As Long

    Do
        If pbooDownButton = True Then ' Button at top of bar
            llngDownVar = (pobjObject.picListBar.Height - glngUIBottomPlottedPos) - gconUIBottomPosDefault
            If (gconUITopPos - 50) < gconUIButtonTopPosDefault And (gconUITopPos - 50) < llngDownVar Then
                'Disable button
                gbooUIScrollButtonClicked = False
                Exit Sub
            End If
            gconUITopPos = gconUITopPos - 50
            pobjObject.DrawButtonSet gstrButtonRoute, "TopDepressed"
        Else 'button at bottom of bar
            If glngUIBottomPlottedPos < pobjObject.picListBar.Height And (gconUITopPos + 50) > gconUIButtonTopPosDefault Then
                gbooUIScrollButtonClicked = False
                Exit Sub
            End If
            
            llngDownVar = gconUIButtonTopPosDefault
            If gconUITopPos + 50 > (pobjObject.picListBar.Height - (glngUIBottomPlottedPos + gconUIButtonTopPosDefault)) And _
                gconUITopPos + 50 > llngDownVar Then
                'Disable button
                gbooUIScrollButtonClicked = False
                Exit Sub
            End If
            gconUITopPos = gconUITopPos + 50
            pobjObject.DrawButtonSet gstrButtonRoute, "BottomDepressed"
        End If
        DoEvents
    Loop Until gbooUIScrollButtonClicked = False
    
End Sub
Sub MDILoad(pobjForm As Form, pobjMdiChildStart As Form, Optional pbooSlack As Variant)
Dim lintArrCounter As Integer
Dim lintWindowState As Integer
Dim lstrPicBarVisible As String

    If IsMissing(pbooSlack) Then
        pbooSlack = False
    End If
    
    gintForceAppClose = fcDontClose
    
    gconUITopPos = gconUIButtonTopPosDefault
    gstrButtonRoute = gconstrMainMenu
    
    
    If (Screen.Height / Screen.TwipsPerPixelY) < 600 Or (Screen.Width / Screen.TwipsPerPixelX) < 600 Then
        MsgBox "Your current screen resolution is less than the" & vbCrLf & _
            "minimum requirement of 800 x 600." & vbCrLf & vbCrLf & _
            "The recommended resolution is 1024 x 768!" & vbCrLf & vbCrLf & _
            "Please contact your System Administrator if you are" & vbCrLf & _
            "unable to change your screen resolution."
        If pbooSlack = False Then
            End
        End If
    End If
    
    If pbooSlack = False Then
        pobjForm.Height = 9000
        pobjForm.Width = 12000
    End If
    
    If Not DebugVersion Then
        'Save handle to the form.
        If gHW = 0 Then
            gHW = pobjForm.hwnd
        End If
        'Begin subclassing.
        Hook
    
    End If
    
    Set gstrCurrentLoadedForm = pobjMdiChildStart
    
    If UCase$(App.ProductName) <> "WAREHOUSE" Then
        lstrPicBarVisible = GetSetting(gstrIniAppName, "UI", "PicBarVisible") & ""
        If IsBlank(lstrPicBarVisible) Then lstrPicBarVisible = True
        If lstrPicBarVisible <> True And lstrPicBarVisible <> False Then
            mdiMain.picListBar.Visible = True
            SaveSetting gstrIniAppName, "UI", "PicBarVisible", True
        Else
            mdiMain.picListBar.Visible = lstrPicBarVisible
        End If
        
        lintWindowState = Val(GetSetting(gstrIniAppName, "UI", "AlwaysMaximized"))
        If lintWindowState <> vbMaximized Then
            lintWindowState = vbNormal
            SaveSetting gstrIniAppName, "UI", "AlwaysMaximized", lintWindowState
        End If
        mdiMain.WindowState = lintWindowState
        
        GetFileHistory
    End If
    
End Sub
Sub PicListBarMouseDown(pobjForm As Form, pintButton As Integer, pShift As Integer, psngX As Single, psngY As Single)
Dim lintBoxArrInc As Integer
Dim llngCurrentTop As Long
Dim llngExtraHeightAdj As Long
Dim llngCurrentLeft As Long

    'ListBarMain Buttons
    If psngX > gconUILeftPos And psngX < gconUILeftPos + gconUIWidth Then
        For lintBoxArrInc = 0 To gintUINumberofButtonsDraw  '6
            llngExtraHeightAdj = (gconUIHeight + gconUIGap) * lintBoxArrInc
            llngCurrentTop = gconUITopPos + llngExtraHeightAdj
            
            If psngY > llngCurrentTop And psngY < llngCurrentTop + gconUIHeight Then
                'Your mouse if over a button
                If pintButton = 1 Then
                    ShowBox pobjForm, lintBoxArrInc, True, True
                    DoEvents
                    pobjForm.ButtonSelected lintBoxArrInc
                Else
                    ShowBox pobjForm, lintBoxArrInc, False, True
                End If
            Else
                ShowBox pobjForm, lintBoxArrInc, True, False
            End If
        Next lintBoxArrInc
    End If
    
    'Scroll Down button
    llngCurrentLeft = (pobjForm.picListBar.Width - gconUIButtonWidth) - gconUIPosLeftFromEdgeScroll
    llngCurrentTop = (pobjForm.picListBar.Height - gconUIButtonHeight) - gconUIPosBottomFromEdgeScroll
    If psngX > llngCurrentLeft + gconUISpace And psngX < llngCurrentLeft + gconUIButtonWidth Then
        If psngY > llngCurrentTop + gconUISpace And psngY < (llngCurrentTop + gconUIButtonHeight) - (gconUISpace * 1) Then
            If pintButton = 1 Then
                gbooUIScrollButtonClicked = True
                DrawScrollButton pobjForm, False, True
                DoEvents
                ClickScrollButton pobjForm, False
                pobjForm.DrawButtonSet gstrButtonRoute
            End If
        End If
    End If
    
    'Scroll Up button
    llngCurrentLeft = (pobjForm.picListBar.Width - gconUIButtonWidth) - gconUIPosLeftFromEdgeScroll
    llngCurrentTop = gconUIPosBottomFromEdgeScroll - 70
    If psngX > llngCurrentLeft + gconUISpace And psngX < llngCurrentLeft + gconUIButtonWidth Then
        If psngY > llngCurrentTop + gconUISpace And psngY < (llngCurrentTop + gconUIButtonHeight) - (gconUISpace * 1) Then
            If pintButton = 1 Then
                gbooUIScrollButtonClicked = True
                DrawScrollButton pobjForm, True, True
                DoEvents
                ClickScrollButton pobjForm, True
                pobjForm.DrawButtonSet gstrButtonRoute
            End If
        End If
    End If
    
End Sub
Sub PicListBarMouseMove(pobjForm As Form, pintButton As Integer, pintShift As Integer, psngX As Single, psngY As Single)
Dim lintBoxArrInc As Integer
Dim llngCurrentTop As Long
Dim llngExtraHeightAdj As Long
Dim llngCurrentLeft As Long

    'ListBarMain Buttons
    If psngX > gconUILeftPos And psngX < gconUILeftPos + gconUIWidth Then
        For lintBoxArrInc = 0 To gintUINumberofButtonsDraw  '6
            llngExtraHeightAdj = (gconUIHeight + gconUIGap) * lintBoxArrInc
            llngCurrentTop = gconUITopPos + llngExtraHeightAdj
            
            If psngY > llngCurrentTop And psngY < llngCurrentTop + gconUIHeight Then
                'Your mouse if over a pintButton
                If pintButton = 1 Then
                    ShowBox pobjForm, lintBoxArrInc, True, True
                    gintUILastButtonHighlighted = lintBoxArrInc
                    Exit For
                Else
                    ShowBox pobjForm, lintBoxArrInc, False, True
                    gintUILastButtonHighlighted = lintBoxArrInc
                    Exit Sub
                End If
            Else
                ShowBox pobjForm, lintBoxArrInc, True, False
                gintUILastButtonHighlighted = -1
            End If
        Next lintBoxArrInc
    End If
    
    If gintUILastButtonHighlighted <> -1 And pintButton <> 1 Then
        ShowBox pobjForm, gintUILastButtonHighlighted, True, False
    End If
    
    Exit Sub
    If pintButton = 1 Then
        'Scroll Down pintButton
        llngCurrentLeft = (pobjForm.picListBar.Width - gconUIButtonWidth) - gconUIPosLeftFromEdgeScroll
        llngCurrentTop = (pobjForm.picListBar.Height - gconUIButtonHeight) - gconUIPosBottomFromEdgeScroll
        If psngX > llngCurrentLeft + gconUISpace And psngX < llngCurrentLeft + gconUIButtonWidth Then
            If psngY > llngCurrentTop + gconUISpace And psngY < (llngCurrentTop + gconUIButtonHeight) - (gconUISpace * 1) Then
                gbooUIScrollButtonClicked = True
                ClickScrollButton pobjForm, True
            End If
        End If
        
        'Scroll Up pintButton
        llngCurrentLeft = (pobjForm.picListBar.Width - gconUIButtonWidth) - gconUIPosLeftFromEdgeScroll
        llngCurrentTop = gconUIPosBottomFromEdgeScroll - 70
        If psngX > llngCurrentLeft + gconUISpace And psngX < llngCurrentLeft + gconUIButtonWidth Then
            If psngY > llngCurrentTop + gconUISpace And psngY < (llngCurrentTop + gconUIButtonHeight) - (gconUISpace * 1) Then
                gbooUIScrollButtonClicked = True
                ClickScrollButton pobjForm, False
            End If
        End If
    End If
    
End Sub
Sub FinishDrawingButtonSet(pobjForm As Form, plngDownVar As Long, pstrParam As Variant)

    If gconUITopPos = gconUIButtonTopPosDefault Then
        glngUIBottomPlottedPos = pobjForm.picListBar.CurrentY
    End If
        
    plngDownVar = (pobjForm.picListBar.Height - glngUIBottomPlottedPos) - gconUIBottomPosDefault
    If (gconUITopPos - 50) < gconUIButtonTopPosDefault And (gconUITopPos - 50) < plngDownVar Then
            
    Else
        Select Case pstrParam
        Case "TopDepressed"
            DrawScrollButton pobjForm, False, True
        Case Else
            DrawScrollButton pobjForm, False, False
        End Select
    End If

    If gconUITopPos < 0 Then
        Select Case pstrParam
        Case "BottomDepressed"
            DrawScrollButton pobjForm, True, True
        Case Else
            DrawScrollButton pobjForm, True, False
        End Select
    End If

End Sub
Sub ShowBanner(pfrmForm As Form, Optional pstrRoute As String)
Dim lstrBannerTitle As String
Dim lstrFormName As String
Dim lintBannerPicIndex As Integer
Const gconUISpace = "" '"          "
Const gconUIIniConfig = "Initial Configuration"
Const gconUIRefData = "Reference Data"

    lstrFormName = pfrmForm.Name
        
    Select Case pstrRoute
    Case gconstrAccount
        lstrBannerTitle = "Account Maintenance" & vbCr & gconUISpace
    Case gconstrEntry
        lstrBannerTitle = "Order Entry" & vbCr & gconUISpace
    Case gconstrEnquiry, gconstrOrderModify
        lstrBannerTitle = "Order Enquiry" & vbCr & gconUISpace
    Case gconstrPacking
        lstrBannerTitle = "Packing" & vbCr & gconUISpace
    Case gconstrConsignmentNorm, gconstrThermalPrintRun
        lstrBannerTitle = "Distribution" & vbCr & gconUISpace
    Case gconstrOrdMaint
        lstrBannerTitle = "Order Maintenence" & vbCr & gconUISpace
    Case gconstrFinance
        lstrBannerTitle = "Finance" & vbCr & gconUISpace
    Case gconstrConfigNetInstall, gconstrConfigFilesPaths, gconstrConfigTables
        lstrBannerTitle = "MMOS Configuration" & vbCr & gconUISpace
    End Select
    
    Select Case lstrFormName
    Case "frmMain" ' Admin
        lstrBannerTitle = lstrBannerTitle & "Welcome to Mindwarp Maintenance System"
        lintBannerPicIndex = 11
    Case "frmMainReps" ' Admin
        lstrBannerTitle = lstrBannerTitle & "Welcome to Mindwarp Manager System"
        lintBannerPicIndex = 11
    Case "frmAbout"
        lstrBannerTitle = lstrBannerTitle & "Welcome to Mindwarp Mail Order System"
        lintBannerPicIndex = 11
    Case "frmReports"
        lstrBannerTitle = lstrBannerTitle & "General Reporting"
        lintBannerPicIndex = 16
    Case "frmWarehouse"
        lstrBannerTitle = lstrBannerTitle & "Welcome to Mindwarp Warehouse System"
        lintBannerPicIndex = 11
    Case "frmCustAcctSel"
        lstrBannerTitle = lstrBannerTitle & "Customer Account Select"
        lintBannerPicIndex = 8
    Case "frmAccount"
        Select Case pstrRoute
        Case gconstrAccount
            lstrBannerTitle = lstrBannerTitle & "Account Address Information"
            lintBannerPicIndex = 3
        Case gconstrEntry
            lstrBannerTitle = lstrBannerTitle & "Advice Note && Account Address Information"
            lintBannerPicIndex = 3
        Case gconstrOrderModify, gconstrEnquiry
            lstrBannerTitle = lstrBannerTitle & "Advice Note Address Information"
            lintBannerPicIndex = 3
        End Select
    Case "frmOrder"
        lstrBannerTitle = lstrBannerTitle & "Products && Order Totals"
        lintBannerPicIndex = 1
    Case "frmOrdDetails"
        lstrBannerTitle = lstrBannerTitle & "Details && Handling"
        lintBannerPicIndex = 10
    Case "frmOrdHistory"
        lstrBannerTitle = lstrBannerTitle & "Historical Orders && Products"
        lintBannerPicIndex = 2
    Case "frmPackaging"
        lstrBannerTitle = lstrBannerTitle & "Packing && Product Confirmation"
        lintBannerPicIndex = 5
    Case "frmPForce"
        Select Case pstrRoute
        Case gconstrConsignmentNorm
            lstrBannerTitle = lstrBannerTitle & "Consignment Management"
            lintBannerPicIndex = 6
        Case gconstrThermalPrintRun
            lstrBannerTitle = lstrBannerTitle & "Thermal Label Printing"
            lintBannerPicIndex = 6
        End Select
    Case "frmQAMisc"
        lstrBannerTitle = lstrBannerTitle & "Discrepancy && Order Management"
        lintBannerPicIndex = 7
    Case "frmCheque"
        lstrBannerTitle = lstrBannerTitle & "Cashbook Management"
        lintBannerPicIndex = 4
    Case "frmFolders"
        Select Case pstrRoute
        Case gconstrConfigRoute
            lstrBannerTitle = gconUIIniConfig & vbCr & gconUISpace & "1/9 "
        End Select
        lstrBannerTitle = lstrBannerTitle & "Folders"
        lintBannerPicIndex = 11
    Case "frmStaticFinance"
        Select Case pstrRoute
        Case gconstrAdminRoute
            lstrBannerTitle = gconUIRefData & vbCr & gconUISpace
        Case gconstrConfigRoute
            lstrBannerTitle = gconUIIniConfig & vbCr & gconUISpace & "2/9 "
        End Select
        lstrBannerTitle = lstrBannerTitle & "Financial Details"
        lintBannerPicIndex = 19
    Case "frmStaticCompany"
        Select Case pstrRoute
        Case gconstrAdminRoute
            lstrBannerTitle = gconUIRefData & vbCr & gconUISpace
        Case gconstrConfigRoute
            lstrBannerTitle = gconUIIniConfig & vbCr & gconUISpace & "3/9 "
        End Select
        lstrBannerTitle = lstrBannerTitle & "Company Details"
        lintBannerPicIndex = 13
    Case "frmStaticPForce"
        Select Case pstrRoute
        Case gconstrAdminRoute
            lstrBannerTitle = gconUIRefData & vbCr & gconUISpace
        Case gconstrConfigRoute
            lstrBannerTitle = gconUIIniConfig & vbCr & gconUISpace & "4/9 "
        End Select
        lstrBannerTitle = lstrBannerTitle & "Parcel Force"
        lintBannerPicIndex = 22
    Case "frmStaMultiAccount"
        Select Case pstrRoute
        Case gconstrAdminRoute
            lstrBannerTitle = gconUIRefData & vbCr & gconUISpace
        Case gconstrConfigRoute
            lstrBannerTitle = gconUIIniConfig & vbCr & gconUISpace & "5/9 "
        End Select
        lstrBannerTitle = lstrBannerTitle & "Account Settings"
        lintBannerPicIndex = 29
    Case "frmStaMultiOrder"
        Select Case pstrRoute
        Case gconstrAdminRoute
            lstrBannerTitle = gconUIRefData & vbCr & gconUISpace
        Case gconstrConfigRoute
            lstrBannerTitle = gconUIIniConfig & vbCr & gconUISpace & "6/9 "
        End Select
        lstrBannerTitle = lstrBannerTitle & "Ordering Settings"
        lintBannerPicIndex = 18
    Case "frmStaMultiMarket"
        Select Case pstrRoute
        Case gconstrAdminRoute
            lstrBannerTitle = gconUIRefData & vbCr & gconUISpace
        Case gconstrConfigRoute
            lstrBannerTitle = gconUIIniConfig & vbCr & gconUISpace & "7/9 "
        End Select
        lstrBannerTitle = lstrBannerTitle & "Marketing Settings"
        lintBannerPicIndex = 14
    Case "frmStaMultiConsignment"
        Select Case pstrRoute
        Case gconstrAdminRoute
            lstrBannerTitle = gconUIRefData & vbCr & gconUISpace
        Case gconstrConfigRoute
            lstrBannerTitle = gconUIIniConfig & vbCr & gconUISpace & "8/9 "
        End Select
        lstrBannerTitle = lstrBannerTitle & "Consignment Settings"
        lintBannerPicIndex = 24
    Case "frmDeploy"
        Select Case pstrRoute
        Case gconstrConfigRoute
            lstrBannerTitle = gconUIIniConfig & vbCr & gconUISpace & "9/9 "
        End Select
        lstrBannerTitle = lstrBannerTitle & "Network Installation"
        lintBannerPicIndex = 11
    Case "frmReferenceData"
        lstrBannerTitle = gconUIRefData & vbCr & gconUISpace
        lintBannerPicIndex = 30
    Case "frmStockView"
        lstrBannerTitle = "Stock Management" & vbCr & gconUISpace
        lintBannerPicIndex = 27
    Case "frmSystemOptions"
        lstrBannerTitle = "System Options" & vbCr & gconUISpace
        lintBannerPicIndex = 12
    Case "frmUsers"
        lstrBannerTitle = "System Options" & vbCr & gconUISpace & "User Management"
        lintBannerPicIndex = 25
    Case "frmUpgrade"
        lstrBannerTitle = "System Maintenance" & vbCr & gconUISpace
        lintBannerPicIndex = 20
    Case "frmSummary"
        lstrBannerTitle = "Summary Info" & vbCr & gconUISpace
        lintBannerPicIndex = 15
    Case "frmLabelLayouts"
        lstrBannerTitle = "System Options" & vbCr & gconUISpace & "Label Layouts"
        lintBannerPicIndex = 23
    Case "frmDuplicates"
        lstrBannerTitle = "Duplicate Handling" & vbCr & gconUISpace
        lintBannerPicIndex = 17
    Case "frmLiteValues"
        lstrBannerTitle = "Drop Down Values" & vbCr & gconUISpace
        lintBannerPicIndex = 30
    Case "frmLiteProducts"
        lstrBannerTitle = "Maintain Products" & vbCr & gconUISpace
        lintBannerPicIndex = 27
    Case "frmLiteSettings"
        lstrBannerTitle = "Essential Settings" & vbCr & gconUISpace
        lintBannerPicIndex = 31
    Case "frmConfigure"
        lstrBannerTitle = lstrBannerTitle & "Files && Paths"
        lintBannerPicIndex = 20
    Case "frmTables"
        lstrBannerTitle = lstrBannerTitle & "Database Tables"
        lintBannerPicIndex = 26
    Case "frmNetInstall"
        lstrBannerTitle = lstrBannerTitle & "Network Install"
        lintBannerPicIndex = 13
    End Select
    
    pfrmForm.ctlBanner1.Caption = lstrBannerTitle
    
    Set pfrmForm.ctlBanner1.Picture = frmButtons.ImageList16Cols.ListImages(lintBannerPicIndex).Picture
    
End Sub
Sub Busy(pbooState As Boolean, Optional pobjForm As Variant)

    If Not IsMissing(pobjForm) Then
        Select Case pbooState
        Case True
            pobjForm.MousePointer = vbHourglass
            pobjForm.Enabled = False
        Case False
            pobjForm.MousePointer = vbNormal
            pobjForm.Enabled = True
        End Select
    Else
        Select Case pbooState
        Case True
            Screen.MousePointer = vbHourglass
            Screen.ActiveForm.Enabled = False
        Case False
            Screen.MousePointer = vbNormal
            Screen.ActiveForm.Enabled = True
        End Select
    End If

End Sub
Sub Triangle(pobjObject As Object, plngX1 As Long, plngY1 As Long, _
    plngX2 As Long, plngY2 As Long, plngX3 As Long, plngY3 As Long, pbooDownButton As Boolean)
Dim lintCounter As Integer
Dim lintFactor As Integer

    lintFactor = 2
    Do Until lintCounter = 20
        pobjObject.Line (plngX1, plngY1)-(plngX2, plngY2), vbBlack
        pobjObject.Line (plngX2, plngY2)-(plngX3, plngY3), vbBlack
        pobjObject.Line (plngX3, plngY3)-(plngX1, plngY1), vbBlack
        
        plngX1 = plngX1 + 2 * lintFactor
        plngX2 = plngX2 - 2 * lintFactor
        If pbooDownButton = True Then
            plngY1 = plngY1 + 1 * lintFactor
            plngY2 = plngY2 + 1 * lintFactor
            plngY3 = plngY3 - 2 * lintFactor
        Else
            plngY1 = plngY1 - 1 * lintFactor
            plngY2 = plngY2 - 1 * lintFactor
            plngY3 = plngY3 + 2 * lintFactor
        End If
        
        lintCounter = lintCounter + 1
    Loop
End Sub
Sub ShowStatus(pintStatusID As Integer)

    On Error Resume Next
    Screen.ActiveForm.sbStatusBar.Panels(1).Text = StatusText(pintStatusID)
    mdiMain.sbStatusBar.Panels(1).Text = StatusText(pintStatusID)
    
    Err.Number = 0
    
End Sub
Function StatusText(pintItem As Integer) As String
Const gconUIstrComeOutAndIn = ", Please close the Program and Open again, ASAP..."
Dim lstrYear As String
    
    Select Case pintItem
        Case 0: StatusText = ""
        Case 1: StatusText = "Updating Local DB - List Details..."
        Case 2: StatusText = "Updating Local DB - Lists..."
        Case 3: StatusText = "Updating Local DB - Products..."
        Case 4: StatusText = "Searching for Catalogue Numbers..."
        Case 5: StatusText = "Searching for Item Descriptions..."
        Case 6: StatusText = "Searching for Class Items..."
        Case 7: StatusText = "Getting Customer Account..."
        Case 8: StatusText = "Searching for Customer Numbers..."
        Case 9: StatusText = "Searching for Post Codes..."
        Case 10: StatusText = "Searching for Surnames..."
        Case 11: StatusText = "Enter the Date, press Insert to use the calendar."
        Case 12: StatusText = "Creating Order..."
        Case 13: StatusText = "Written by Mindwarp Consultancy Ltd. 2002"
        Case 14: StatusText = "The network is not available..."
        Case 15: StatusText = "No Products found, try again!"
        Case 16: StatusText = "No Dates needed for this report!"
        Case 17: StatusText = "Dates needed for this report!"
        Case 18: StatusText = "Searching for BPCS Number..."
        Case 19: StatusText = "Getting customer account information..."
        Case 20: StatusText = "Adding New Customer account..."
        Case 21: StatusText = "Retrieving Account number..."
        Case 22: StatusText = "Retrieving Order Number..."
        Case 23: StatusText = "Updating account information..."
        Case 24: StatusText = "Locking account for User..."
        Case 25: StatusText = "Adding Advice Note Data..."
        Case 26: StatusText = "Uploading order lines to Central DB..."
        Case 27: StatusText = "Checking for new stock..."
        Case 28: StatusText = "Adding Remark..."
        Case 29: StatusText = "Adding Refund..."
        Case 30: StatusText = "Copying old Account..."
        Case 31: StatusText = "Updating stock details..."
        Case 32: StatusText = "Deleting Products Master records..."
        Case 33: StatusText = "Updating products Master records..."
        Case 34: StatusText = "Updating Vat on prices..."
        Case 35: StatusText = "Updating reference fields..."
        Case 36: StatusText = "Changing new stock flag ..."
        Case 37: StatusText = "A Program update is ready" & gconUIstrComeOutAndIn
        Case 38: StatusText = "Updating Report Data...."
        Case 39: StatusText = "Press CTRL & A for Quick Address."
        Case 40: StatusText = "New stock is available" & gconUIstrComeOutAndIn
        Case 41: StatusText = "Deploying Static..."
        Case 42: StatusText = "Downloading order lines from master..."
        Case 43: StatusText = "Appending products to order lines..."
        Case 44: StatusText = "Updating order totals..."
        Case 45: StatusText = "Update master despatch quantity..."
        Case 46: StatusText = "Getting customer name..."
        Case 47: StatusText = "Analysing underpayments..."
        Case 48: StatusText = "Updating machine stock time stamp..."
        Case 49: StatusText = "Resetting local products..."
        Case 50: StatusText = "Adding new customer note record..."
        Case 51: StatusText = "Getting remark..."
        Case 52: StatusText = "Retrieving new remark.."
        Case 53: StatusText = "Getting customer note..."
        Case 54: StatusText = "Updating remark..."
        Case 55: StatusText = "Attaching remark to advice note record..."
        Case 56: StatusText = "Updating customer notes..."
        Case 57: StatusText = "Clearing customer account buffer..."
        Case 58: StatusText = "Clearing advice note buffer..."
        Case 59: StatusText = "Calculating order total..."
        Case 60: StatusText = "Clearing remark note buffers..."
        Case 61: StatusText = "Adding sales code to order lines..."
        Case 62: StatusText = "Getting advice note record..."
        Case 63: StatusText = "Updating advice note..."
        Case 64: StatusText = "Updating order status..."
        Case 65: StatusText = "Analysing current advice note status..."
        Case 66: StatusText = "Calculating despatch order totals..."
        Case 67: StatusText = "Adding Cashbook entry..."
        Case 68: StatusText = "Checking for refund scenarios..."
        Case 69: StatusText = "Formatting card number..."
        Case 70: StatusText = "Creating stock summary Delete stage..."
        Case 71: StatusText = "Updating cheque number and printed date..."
        Case 72: StatusText = "Updating bank report print date..."
        Case 73: StatusText = "Getting last current stock number..."
        Case 74: StatusText = "Setting initial stock batch value..."
        Case 75: StatusText = "Updating current stock batch number..."
        Case 76: StatusText = "Initialising databases..."
        Case 77: StatusText = "OK..."
        Case 78: StatusText = "Getting custom report..."
        Case 79: StatusText = "Updating user record..."
        Case 80: StatusText = "Adding new user record..."
        Case 81: StatusText = "Getting user record..."
        Case 82: StatusText = "Checking for new loader program..."
        Case 83: StatusText = "Launching Quick Address..."
        Case 84: StatusText = "Adding parcel force record..."
        Case 85: StatusText = "Get current consignment batch numbers.."
        Case 86: StatusText = "Get consignment range..."
        Case 87: StatusText = "Increment and save batch number..."
        Case 88: StatusText = "Create control file..."
        Case 89: StatusText = "Retrieve advice records..."
        Case 90: StatusText = "Get Awaiting advice records..."
        Case 91: StatusText = "Create PF file..."
        Case 92: StatusText = "Reset PF Static buffers.."
        Case 93: StatusText = "Update advice note with parcel info..."
        Case 94: StatusText = "Updating current consign batch number..."
        Case 95: StatusText = "Update consign status..."
        Case 96: StatusText = "Get Parcel force consignment records..."
        Case 97: StatusText = "Get PF contract details..."
        Case 98: StatusText = "Create PF manifest..."
        Case 99: StatusText = "Copying successful..."
        Case 100: StatusText = "Creating stock summary Append stage..."
        Case 101: StatusText = "Creating stock summary OK..."
        Case 102: StatusText = "Deleting Substitution records..."
        Case 103: StatusText = "Importing Substitution Base Records..."
        Case 104: StatusText = "Transfer Products data..."
        Case 105: StatusText = "Assigning Reference Info Values..."
        Case 106: StatusText = "Decrypting File... Please Wait!"
        Case 107: StatusText = "Calculating total pages in report..."
        Case 108: StatusText = "Displaying report..."
        Case 109: StatusText = "Dates may not necessarily be used..."
        Case 110: StatusText = "Merging Advice Notes record..."
        Case 111: StatusText = "Merging Customer Notes record..."
        Case 112: StatusText = "Merging Cash Book record..."
        Case 113: StatusText = "Merging Order Lines record..."
        Case 114: StatusText = "Merging Parcel Force record..."
        Case 115: StatusText = "Merging Complete..."
        Case 116: StatusText = "Deleting Customer account..."
        Case 117: StatusText = "Process Complete..."
        Case 118: StatusText = "Checking Import File. Please wait..."
        Case 119: StatusText = "Importing data locally. Please wait..."
        Case 120: StatusText = "Clearing Master PAD Tables..."
        Case 121: StatusText = "Uploading Availabity records..."
        Case 122: StatusText = "Uploading Office records..."
        Case 123: StatusText = "Deleting closed offices..."
        Case 124: StatusText = "Uploading Office opening times..."
        
        Case 128: StatusText = "Updating Postcode Outcode..."
        Case 129: StatusText = "Updating PAD Flag..."
        Case 130: StatusText = "Drop PAD Import table..."
        Case 130: StatusText = "Clearing PAD Local Tables..."
        Case 131: StatusText = "Download PAD stage 1 of 3..."
        Case 132: StatusText = "Download PAD stage 2 of 3..."
        Case 133: StatusText = "Download PAD stage 3 of 3..."
        Case 134: StatusText = "Please consider using more than 256 colours!"
        End Select

End Function
Sub SetSelected(pfrmForm As Form)

    On Error Resume Next
    
    pfrmForm.ActiveControl.SelStart = 0
    pfrmForm.ActiveControl.SelLength = Len(pfrmForm.ActiveControl.Text)
        
    On Error GoTo 0

End Sub
Function MCLDebugChoices() As String
Dim lintDebugVersion As Variant
Dim lstrAppHelpFile As String
Dim lstrDepugAppHelpFilePath As String

    If DebugVersion Then
        lintDebugVersion = MsgBox("Mindwarp (Standard ver) " & vbTab & "= Abort" & vbCrLf & _
                                  "My Company              " & vbTab & "= Retry" & vbCrLf & _
                                  "My Company SQL (Testing)" & vbTab & "= Ignore", vbAbortRetryIgnore + vbDefaultButton2)
        Select Case lintDebugVersion
        Case vbRetry
            gstrSystemRoute = srCompanyDebugRoute
            gbooSQLServerInUse = False
        Case vbAbort
            gstrSystemRoute = srStandardRoute
            gbooSQLServerInUse = False
        Case vbIgnore
            gstrSystemRoute = srCompanyDebugRoute
            gbooSQLServerInUse = True
        End Select
    End If
    
    Select Case UCase$(App.ProductName)
    Case "MAINTENANCE"
        lstrAppHelpFile = "Mmaint.chm"
        lstrDepugAppHelpFilePath = "Maint\"
    Case "LITE"
        lstrAppHelpFile = "Mmos.chm"
        lstrDepugAppHelpFilePath = "Lite\"
    Case "CLIENT"
        lstrAppHelpFile = "Mmos.chm"
        lstrDepugAppHelpFilePath = "Client\"
    Case "MANAGER"
        lstrAppHelpFile = "Mmanager.chm"
        lstrDepugAppHelpFilePath = "Manager\"
    End Select
    
    If DebugVersion Then
        gstrHelpFileBase = "D:\HelpBuild\Mos\" & lstrDepugAppHelpFilePath & lstrAppHelpFile
    Else
        gstrHelpFileBase = App.Path & "\" & lstrAppHelpFile
    End If
    
    MCLDebugChoices = lstrAppHelpFile
    
End Function
Public Sub Hook()

    glngPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Public Sub Unhook()
Dim temp As Long

    glngPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, glngPrevWndProc)
    
End Sub
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
Dim MinMax As MINMAXINFO
Dim lintItem As Integer
Dim lstrLog As String

    Select Case uMsg
    Case WM_GETMINMAXINFO
        'Check for request for min/max window sizes.
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.X = 800
        MinMax.ptMinTrackSize.Y = 600
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Case WM_TCARD
        TrainingCards wParam, lParam
        WindowProc = CallWindowProc(glngPrevWndProc, hw, uMsg, wParam, lParam)
    Case Else
        'pass other instructions back to UI
        WindowProc = CallWindowProc(glngPrevWndProc, hw, uMsg, wParam, lParam)
    End Select

    
End Function
Sub xPaintFormMenu(pobjForm As Form)
    
End Sub

Sub MakeVisible(pobjForm As Form, pbooVisible As Boolean)
Dim lintArrInc As Integer
Dim lintArrInc2 As Integer
Dim lvarBackColor As Variant
Dim lintItemCounter As Integer

    On Error Resume Next
    If pbooVisible = False Then
        ReDim mstrControls(0)
        For lintArrInc = 0 To pobjForm.Controls.Count - 1   ' Use the Controls collection
            If Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "tim" And _
                Left$(pobjForm.Controls(lintArrInc).Name, 3) <> "cdg" And _
                Trim$(pobjForm.Controls(lintArrInc).Name) <> "lblBox" And _
                Left$(pobjForm.Controls(lintArrInc).Name, 2) <> "sb" Then
                If pobjForm.Controls(lintArrInc).Visible = True Then
                    pobjForm.Controls(lintArrInc).Visible = False
                    If lintItemCounter > 0 Then ReDim Preserve mstrControls(UBound(mstrControls) + 1)
                    mstrControls(UBound(mstrControls)) = pobjForm.Controls(lintArrInc).Name
                    lintItemCounter = lintItemCounter + 1
                End If
            End If
        Next
    Else
        For lintArrInc2 = 0 To UBound(mstrControls)
            pobjForm.Controls(mstrControls(lintArrInc2)).Visible = True
        Next lintArrInc2
    End If
    
End Sub
Sub AddNewFileHistoryItem(plngCustNum As Long, plngOrderNum As Long, _
    pstrCustName As String, pstrType As String)
Dim lstrHistoryItem As HistroryValue
Dim lintCounter As Integer

    For lintCounter = (UBound(mudtUImnuFileHistory) - 1) To 0 Step -1
        With lstrHistoryItem
            .lngCustNum = Val(GetSetting(gstrIniAppName, "FileHistoryItem" & lintCounter, "CustNum"))
            .lngOrderNum = Val(GetSetting(gstrIniAppName, "FileHistoryItem" & lintCounter, "OrderNum"))
            .strCustName = GetSetting(gstrIniAppName, "FileHistoryItem" & lintCounter, "CustName")
            .strType = GetSetting(gstrIniAppName, "FileHistoryItem" & lintCounter, "Type")
            SaveSetting gstrIniAppName, "FileHistoryItem" & (lintCounter + 1), "CustNum", .lngCustNum
            SaveSetting gstrIniAppName, "FileHistoryItem" & (lintCounter + 1), "OrderNum", .lngOrderNum
            SaveSetting gstrIniAppName, "FileHistoryItem" & (lintCounter + 1), "CustName", .strCustName
            SaveSetting gstrIniAppName, "FileHistoryItem" & (lintCounter + 1), "Type", .strType
        End With
    Next lintCounter
    
    SaveSetting gstrIniAppName, "FileHistoryItem1", "CustNum", plngCustNum
    SaveSetting gstrIniAppName, "FileHistoryItem1", "OrderNum", plngOrderNum
    SaveSetting gstrIniAppName, "FileHistoryItem1", "CustName", pstrCustName
    SaveSetting gstrIniAppName, "FileHistoryItem1", "Type", pstrType

    GetFileHistory
    
End Sub
Sub GetFileHistory()
Dim lstrHistoryItem As HistroryValue
Dim lintArrInc As Integer
Dim lintCounter As Integer

    For lintArrInc = 0 To UBound(mudtUImnuFileHistory)
        With lstrHistoryItem
            .lngCustNum = Val(GetSetting(gstrIniAppName, "FileHistoryItem" & lintArrInc, "CustNum"))
            .lngOrderNum = Val(GetSetting(gstrIniAppName, "FileHistoryItem" & lintArrInc, "OrderNum"))
            .strCustName = GetSetting(gstrIniAppName, "FileHistoryItem" & lintArrInc, "CustName")
            .strType = GetSetting(gstrIniAppName, "FileHistoryItem" & lintArrInc, "Type")
            If .lngCustNum <> 0 Then
                mudtUImnuFileHistory(lintCounter).strItemValue = "(M:" & .lngCustNum & ") " & _
                    .strCustName & " (O:" & .lngOrderNum & ") - " & .strType
                mudtUImnuFileHistory(lintCounter).lngCustNum = .lngCustNum
                mudtUImnuFileHistory(lintCounter).lngOrderNum = .lngOrderNum
                lintCounter = lintCounter + 1
            End If
        End With
    Next lintArrInc
    
End Sub
Sub DBGridLayout(pobjForm As Form, pobjGrid As DBGrid, pstrOperation As String)
Dim lcolColumn As MSDBGrid.Column
Dim llngDefWidth As Long
Dim llngThisWidth As Long
Dim llngRowHeight As Long
Dim lbooDoneOnce As Boolean

    lbooDoneOnce = False
    
    llngDefWidth = pobjGrid.DefColWidth
    For Each lcolColumn In pobjGrid.Columns
        With lcolColumn
            Select Case UCase$(pstrOperation)
            Case "SAVE"
                If lbooDoneOnce = False Then
                    SaveSetting gstrIniAppName, "Grid " & pobjForm.Name, "RowHeight", pobjGrid.RowHeight
                    lbooDoneOnce = True
                End If
                llngThisWidth = .Width
                SaveSetting gstrIniAppName, "Grid " & pobjForm.Name, CStr(.ColIndex), llngThisWidth
            Case "LOAD"
                If lbooDoneOnce = False Then
                    llngRowHeight = Val(GetSetting(gstrIniAppName, "Grid " & pobjForm.Name, "RowHeight"))
                    If llngRowHeight <> 0 Then
                        pobjGrid.RowHeight = llngRowHeight
                    End If
                    lbooDoneOnce = True
                End If
                llngThisWidth = Val(GetSetting(gstrIniAppName, "Grid " & pobjForm.Name, CStr(.ColIndex)))
                If llngThisWidth <> 0 Then
                    .Width = llngThisWidth
                Else 'probably first time, save standard widths
                    llngThisWidth = .Width
                    SaveSetting gstrIniAppName, "Grid " & pobjForm.Name, CStr(.ColIndex) & "DEF", llngThisWidth
                End If
            End Select
        End With
    Next lcolColumn
    
End Sub
Sub ResetGridLayout()
Dim lobjControl As Control
Dim lcolColumn As MSDBGrid.Column
Dim llngDefWidth As Long
Dim llngThisWidth As Long
Dim llngRowHeight As Long
Dim lbooDoneOnce As Boolean

    On Error Resume Next
    For Each lobjControl In mdiMain.ActiveForm
        If UCase$(Left$(lobjControl.Name, 3)) = "DBG" Then
            For Each lcolColumn In lobjControl.Columns
                With lcolColumn
                    If lbooDoneOnce = False Then
                        llngRowHeight = Val(GetSetting(gstrIniAppName, "Grid " & mdiMain.ActiveForm.Name, "RowHeight"))
                        If llngRowHeight <> 0 Then
                            lobjControl.RowHeight = llngRowHeight
                        End If
                        lbooDoneOnce = True
                    End If
                    llngThisWidth = Val(GetSetting(gstrIniAppName, "Grid " & mdiMain.ActiveForm.Name, CStr(.ColIndex) & "DEF"))
                    If llngThisWidth <> 0 Then
                        .Width = llngThisWidth
                    End If
                        
                End With
            Next lcolColumn
            lobjControl.Refresh
        End If
    Next lobjControl
    
    DoEvents
End Sub
Function GetGridFromActive() As Integer
Dim lobjControl As Object

    On Error Resume Next
    For Each lobjControl In mdiMain.ActiveForm
        If UCase$(Left$(lobjControl.Name, 3)) = "DBG" Then
            GetGridFromActive = GetGridFromActive + 1
        End If
    Next lobjControl
        
End Function
Function CheckDB(pstrTableAndFields() As TableAndFields, pobjForm As Form, Optional pvarParam As Variant) As Boolean
Dim datDatabase As Database
Dim tdfLoop As TableDef
Dim lintArrInc As Integer
Dim lintArrInc2 As Integer
Dim prpLoop As Property
Dim rstTables As Recordset
Dim fldRecordset As Field
Dim lintFieldCount As Integer
Dim lstrLastTable As String
Dim lstrErrBuild As String
Dim lstrErrTitleBuild As String

    CheckDB = True
    
    If IsMissing(pvarParam) Then
        pvarParam = ""
    End If

    Select Case gstrUserMode
    Case gconstrTestingMode
        Set datDatabase = OpenDatabase(gstrStatic.strCentralTestingDBFile)
    Case gconstrLiveMode
        Set datDatabase = OpenDatabase(gstrStatic.strCentralDBFile)
    End Select
    
    On Error Resume Next
    For lintArrInc = 0 To UBound(pstrTableAndFields)
        With pstrTableAndFields(lintArrInc)
            If Trim$(.strSourceTable) <> "" Then
                If lstrLastTable <> .strSourceTable Then
                    Set rstTables = datDatabase.OpenRecordset(.strSourceTable)
                    lstrLastTable = .strSourceTable
                End If
                Set fldRecordset = rstTables.Fields(.strName)
                If fldRecordset.AllowZeroLength <> CBool(.strAllowZeroLength) Then
                    lstrErrBuild = lstrErrBuild & Pad(pobjForm, 30, .strSourceTable) & _
                        Pad(pobjForm, 30, .strName) & _
                        Pad(pobjForm, 40, "AllowZeroLength = " & .strAllowZeroLength) & _
                        Pad(pobjForm, 40, fldRecordset.AllowZeroLength) & vbCrLf
                End If
                If fldRecordset.Type <> .strType Then
                    lstrErrBuild = lstrErrBuild & Pad(pobjForm, 30, .strSourceTable) & _
                        Pad(pobjForm, 30, .strName) & _
                        Pad(pobjForm, 40, "Type = " & DataType(CLng(.strType))) & _
                        Pad(pobjForm, 40, DataType(CLng(fldRecordset.Type))) & vbCrLf
                End If
                If fldRecordset.Size <> .strSize Then
                    lstrErrBuild = lstrErrBuild & Pad(pobjForm, 30, .strSourceTable) & _
                        Pad(pobjForm, 30, .strName) & _
                        Pad(pobjForm, 40, "Size = " & .strSize) & _
                        Pad(pobjForm, 40, fldRecordset.Size) & vbCrLf
                End If
                If fldRecordset.DataUpdatable <> .strDataUpdatable Then
                    lstrErrBuild = lstrErrBuild & Pad(pobjForm, 30, .strSourceTable) & _
                        Pad(pobjForm, 30, .strName) & _
                        Pad(pobjForm, 40, "DataUpdatable = " & .strDataUpdatable) & _
                        Pad(pobjForm, 40, fldRecordset.DataUpdatable) & vbCrLf
                End If
                If fldRecordset.DefaultValue <> .strDefaultValue Then
                    If .strDefaultValue = "" Then
                        lstrErrBuild = lstrErrBuild & Pad(pobjForm, 30, .strSourceTable) & _
                            Pad(pobjForm, 30, .strName) & _
                            Pad(pobjForm, 40, "DefaultValue = (Blank)") & _
                            Pad(pobjForm, 40, fldRecordset.DefaultValue) & vbCrLf
                    Else
                        lstrErrBuild = lstrErrBuild & Pad(pobjForm, 30, .strSourceTable) & _
                            Pad(pobjForm, 30, .strName) & _
                            Pad(pobjForm, 40, "DefaultValue = " & .strDefaultValue) & _
                            Pad(pobjForm, 40, fldRecordset.DefaultValue) & vbCrLf
                    End If
                End If
                If fldRecordset.Required <> .strRequired Then
                    lstrErrBuild = lstrErrBuild & Pad(pobjForm, 30, .strSourceTable) & _
                        Pad(pobjForm, 30, .strName) & _
                        Pad(pobjForm, 40, "Required = " & .strRequired) & _
                        Pad(pobjForm, 40, fldRecordset.Required) & vbCrLf
                End If
            End If
        End With
    Next lintArrInc
    
    Busy False
    
    If lstrErrBuild <> "" Then
        lstrErrTitleBuild = Pad(pobjForm, 30, "Table") & _
            Pad(pobjForm, 30, "Field") & _
            Pad(pobjForm, 40, "Correct Value") & _
            Pad(pobjForm, 40, "Current Value") & vbCrLf
        lstrErrTitleBuild = lstrErrTitleBuild & Pad(pobjForm, 30, "=====") & _
            Pad(pobjForm, 30, "=====") & _
            Pad(pobjForm, 40, "=============") & _
            Pad(pobjForm, 40, "=============") & vbCrLf
        
        lstrErrBuild = lstrErrTitleBuild & lstrErrBuild
        If pvarParam <> "LITE" Then
        MsgBox lstrErrBuild & vbCrLf & vbCrLf & "Your database structure has the above Anomalies, please report them!" & vbCrLf & _
            "Please press ALT & 'Print Screen' and paste into Paintbrush to make a copy of this message!", , gconstrTitlPrefix & "DB Check"
        Else
            CheckDB = False
            If DebugVersion Then
                Debug.Print lstrErrBuild
                MsgBox lstrErrBuild & vbCrLf & vbCrLf & "Your database structure has the above Anomalies, please report them!" & vbCrLf & _
                    "This is a Debug message, MCL staff only", , gconstrTitlPrefix & "DB Check"
            Else
                MsgBox "There seems to be a descrepancy in your database structure!" & vbCrLf & _
                    "There are numerous reasons why this has occured!  We recommend" & vbCrLf & _
                    "that you check our web site for more details or resinstall the" & vbCrLf & _
                    "progam!", vbInformation, gconstrTitlPrefix & "DB Check"
            End If
        End If
    Else
        If pvarParam <> "LITE" Then
            MsgBox "Your database structure matches the requirements of this program version!", vbInformation, gconstrTitlPrefix & "DB Check"
        End If
    End If

    datDatabase.Close
    
End Function
Sub xEndProg()

    MsgBox "End prog"
    On Error Resume Next
    Unload mdiMain
    
End Sub
Sub RefreshMenu(pobjForm As Form)
Dim llngUIRetSubMenu As Long
Dim lintArrInc As Integer
Dim lintArrInc2 As Integer
Dim lintAlwaysMaximized As Integer
Dim lstrShowFeatures As String
Dim lbooBatchEnable As Boolean
    
    For lintArrInc2 = 0 To UBound(mudtUImnuFileHistory)
        With mudtUImnuFileHistory(lintArrInc2)
            If .strItemValue <> "" Then
                If mdiMain.ActiveForm.Name = "frmAbout" Then
                    mdiMain.mnuFileHistory(lintArrInc2).Visible = True
                    mdiMain.mnuFileHistory(lintArrInc2).Enabled = True
                    mdiMain.mnuFileHistory(lintArrInc2).Caption = "&" & (lintArrInc2 + 1) & " " & .strItemValue
                Else
                    mdiMain.mnuFileHistory(lintArrInc2).Visible = True
                    mdiMain.mnuFileHistory(lintArrInc2).Enabled = False
                    mdiMain.mnuFileHistory(lintArrInc2).Caption = "&" & (lintArrInc2 + 1) & " " & .strItemValue
                End If
                mdiMain.mnuFileHistorySep.Visible = True
                Dim lbooShowPackOp As Boolean
                Dim lbooShowModOp As Boolean
                Dim lbooShowHistOp As Boolean
                lbooShowPackOp = False: lbooShowModOp = False: lbooShowHistOp = False
                
                Select Case gstrGenSysInfo.lngUserLevel
                Case Is < 20 'Distribution 10
                    lbooShowPackOp = True
                Case Is < 30 'Order Entry
                    lbooShowModOp = True
                    lbooShowHistOp = True
                Case Is < 50 'Accounts
                    lbooShowModOp = True
                    lbooShowHistOp = True
                Case Is < 100 ' Information Systems
                    lbooShowPackOp = True
                    lbooShowModOp = True
                    lbooShowHistOp = True
                End Select
                ShowHistoryOption lbooShowPackOp, lbooShowModOp, lbooShowHistOp
            End If
        End With
    Next lintArrInc2
            
    'View Menu items
    'Show / Hide picture bar
    If UCase$(App.ProductName) <> "WAREHOUSE" Then
        mdiMain.mnuViewShowPicBar.Visible = True
        If mdiMain.picListBar.Visible = True Then
            mdiMain.mnuViewShowPicBar.Checked = True
        Else
            mdiMain.mnuViewShowPicBar.Checked = False
        End If
    End If
    
    If UCase$(App.ProductName) = "LITE" Or UCase$(App.ProductName) = "WAREHOUSE" Or _
        UCase$(App.ProductName) = "CONFIG" Then
        mdiMain.mnuViewSep.Visible = False
    Else
        'Show Features
        lstrShowFeatures = GetSetting(gstrIniAppName, "UI", "ShowFeatures")
        If IsBlank(lstrShowFeatures) Then lstrShowFeatures = True
        If CBool(lstrShowFeatures) = True Then
            mdiMain.mnuViewShowNewFeatures.Visible = True
            mdiMain.mnuViewShowNewFeatures.Checked = True
        Else
            lstrShowFeatures = False
            mdiMain.mnuViewShowNewFeatures.Visible = True
            mdiMain.mnuViewShowNewFeatures.Checked = False
            SaveSetting gstrIniAppName, "UI", "ShowFeatures", lstrShowFeatures
        End If
    End If
        
    'Maximised on startup
    lintAlwaysMaximized = Val(GetSetting(gstrIniAppName, "UI", "AlwaysMaximized"))
    If lintAlwaysMaximized <> vbMaximized Then
        lintAlwaysMaximized = vbNormal
        SaveSetting gstrIniAppName, "UI", "AlwaysMaximized", lintAlwaysMaximized
    End If
    If lintAlwaysMaximized = vbNormal Then
        mdiMain.mnuViewMaxOnStartup.Checked = False
    Else
        mdiMain.mnuViewMaxOnStartup.Checked = True
    End If
        
    Select Case UCase$(App.ProductName)
    Case "LITE", "CLIENT"
        lbooBatchEnable = False
        If mdiMain.Visible Then
            If mdiMain.ActiveForm.Name = "frmAbout" Then
                lbooBatchEnable = True
                Select Case gstrGenSysInfo.lngUserLevel
                Case Is < 20 'Distribution
                    mdiMain.mnuGoItem1.Caption = mnuClientGoPacking
                Case Is < 30 'Order Entry
                    mdiMain.mnuGoItem1.Visible = True
                    mdiMain.mnuGoItem1.Caption = mnuClientGoOrderEntry
                    mdiMain.mnuGoItem2.Visible = True
                    mdiMain.mnuGoItem2.Caption = mnuClientGoEnquiry
                    mdiMain.mnuGoItem3.Visible = True
                    mdiMain.mnuGoItem3.Caption = mnuClientGoAcctMaint
                    mdiMain.mnuGoItem4.Visible = True
                    mdiMain.mnuGoItem4.Caption = mnuClientGoFinance
                Case Is < 40 'Sales
                    mdiMain.mnuGoItem1.Visible = True
                    mdiMain.mnuGoItem1.Caption = mnuClientGoEnquiry
                    mdiMain.mnuGoItem2.Visible = True
                    mdiMain.mnuGoItem2.Caption = mnuClientGoAcctMaint
                Case Is < 50 'Accounts
                    mdiMain.mnuGoItem1.Visible = True
                    mdiMain.mnuGoItem1.Caption = mnuClientGoEnquiry
                    mdiMain.mnuGoItem2.Visible = True
                    mdiMain.mnuGoItem2.Caption = mnuClientGoAcctMaint
                    mdiMain.mnuGoItem3.Visible = True
                    mdiMain.mnuGoItem3.Caption = mnuClientGoFinance
                'Case Is < 99 ' General Managers
                Case Is < 100 ' Information Systems
                    mdiMain.mnuGoItem1.Visible = True
                    mdiMain.mnuGoItem1.Caption = mnuClientGoOrderEntry
                    mdiMain.mnuGoItem2.Visible = True
                    mdiMain.mnuGoItem2.Caption = mnuClientGoEnquiry
                    mdiMain.mnuGoItem3.Visible = True
                    mdiMain.mnuGoItem3.Caption = mnuClientGoAcctMaint
                    mdiMain.mnuGoItem4.Visible = True
                    mdiMain.mnuGoItem4.Caption = mnuClientGoFinance
                    mdiMain.mnuGoItem5.Visible = True
                    mdiMain.mnuGoItem5.Caption = mnuClientGoPacking
                    mdiMain.mnuGoItem6.Visible = True
                    mdiMain.mnuGoItem6.Caption = mnuClientGoOrderMaint
                End Select
            End If
            mdiMain.mnuGoItem1.Enabled = lbooBatchEnable
            mdiMain.mnuGoItem2.Enabled = lbooBatchEnable
            mdiMain.mnuGoItem3.Enabled = lbooBatchEnable
            mdiMain.mnuGoItem4.Enabled = lbooBatchEnable
            mdiMain.mnuGoItem5.Enabled = lbooBatchEnable
            mdiMain.mnuGoItem6.Enabled = lbooBatchEnable
        End If
    Case "MAINTENANCE"
        lbooBatchEnable = False
        If mdiMain.ActiveForm.Name = "frmMain" Then
            lbooBatchEnable = True
            Select Case gstrGenSysInfo.lngUserLevel
            Case Is < 20 'Distribution
            Case Is < 30 'Order Entry
            Case Is < 40 'Sales
                mdiMain.mnuGoItem1.Visible = True
                mdiMain.mnuGoItem1.Caption = mnuMaintGoRefData
            Case Is < 99 ' General Managers
                mdiMain.mnuGoItem1.Visible = True
                mdiMain.mnuGoItem1.Caption = mnuMaintGoRefData
                mdiMain.mnuGoItem2.Visible = True
                mdiMain.mnuGoItem2.Caption = mnuMaintGoStockMan
                mdiMain.mnuGoItem3.Visible = True
                mdiMain.mnuGoItem3.Caption = mnuMaintGoSysOps
            Case Is < 100 ' Information Systems
                mdiMain.mnuGoItem1.Visible = True
                mdiMain.mnuGoItem1.Caption = mnuMaintGoRefData
                mdiMain.mnuGoItem2.Visible = True
                mdiMain.mnuGoItem2.Caption = mnuMaintGoStockMan
                mdiMain.mnuGoItem3.Visible = True
                mdiMain.mnuGoItem3.Caption = mnuMaintGoSysOps
                mdiMain.mnuGoItem4.Visible = True
                mdiMain.mnuGoItem4.Caption = mnuMaintGoSysMain
            End Select
        End If
        mdiMain.mnuGoItem1.Enabled = lbooBatchEnable
        mdiMain.mnuGoItem2.Enabled = lbooBatchEnable
        mdiMain.mnuGoItem3.Enabled = lbooBatchEnable
        mdiMain.mnuGoItem4.Enabled = lbooBatchEnable
    Case "MANAGER"
        lbooBatchEnable = False
        If mdiMain.ActiveForm.Name = "frmMainReps" Then
        'user levels restrictions
            lbooBatchEnable = True
            mdiMain.mnuGoItem1.Visible = True
            mdiMain.mnuGoItem1.Caption = mnuMgrGoGenReps
            mdiMain.mnuGoItem2.Visible = True
            mdiMain.mnuGoItem2.Caption = mnuMgrGoDistribution
            mdiMain.mnuGoItem3.Visible = True
            mdiMain.mnuGoItem3.Caption = mnuMgrGoMarkSets
            mdiMain.mnuGoItem4.Visible = True
            mdiMain.mnuGoItem4.Caption = mnuMgrGoAgentReps
            mdiMain.mnuGoItem5.Visible = True
            mdiMain.mnuGoItem5.Caption = mnuMgrGoSumInfo
            mdiMain.mnuGoItem6.Visible = True
            mdiMain.mnuGoItem6.Caption = mnuMgrGoDupHand
        End If
        mdiMain.mnuGoItem1.Enabled = lbooBatchEnable
        mdiMain.mnuGoItem2.Enabled = lbooBatchEnable
        mdiMain.mnuGoItem3.Enabled = lbooBatchEnable
        mdiMain.mnuGoItem4.Enabled = lbooBatchEnable
        mdiMain.mnuGoItem5.Enabled = lbooBatchEnable
        mdiMain.mnuGoItem6.Enabled = lbooBatchEnable
    Case Else
        mdiMain.mnuGo.Enabled = False
    End Select
    
    'Tools Menu items
    Select Case UCase$(App.ProductName)
    Case "LITE", "CLIENT", "MANAGER", "MAINTENANCE"
        If mdiMain.Visible Then
            If mdiMain.ActiveForm.Name = "frmAbout" Or mdiMain.ActiveForm.Name = "frmMainReps" Or _
                mdiMain.ActiveForm.Name = "frmMain" Then
                mdiMain.mnuToolsMinder.Enabled = True
            Else
                mdiMain.mnuToolsMinder.Enabled = False
            End If
        End If
    Case Else
        mdiMain.mnuToolsMinder.Visible = False
    End Select
        
    If gstrSystemRoute = srStandardRoute Then
        If UCase$(App.ProductName) <> "WAREHOUSE" Then
            If GetGridFromActive > 0 Then
                mdiMain.mnuToolsResetGrid.Visible = True
                mdiMain.mnuToolsResetGrid.Enabled = True
            Else
                mdiMain.mnuToolsResetGrid.Visible = True
                mdiMain.mnuToolsResetGrid.Enabled = False
            End If
        Else
            mdiMain.mnuToolsResetGrid.Visible = False
        End If
    Else
        mdiMain.mnuToolsResetGrid.Visible = False
    End If
                
    If UCase$(App.ProductName) = "LITE" Then
        If mdiMain.ActiveForm.Name = "frmAbout" Then
            mdiMain.mnuToolsConfigureValues.Visible = True
            mdiMain.mnuToolsConfigureValues.Enabled = True
            mdiMain.mnuToolsMaintainProducts.Visible = True
            mdiMain.mnuToolsMaintainProducts.Enabled = True
            mdiMain.mnuToolsEssentialSettings.Visible = True
            mdiMain.mnuToolsEssentialSettings.Enabled = True
        Else
            mdiMain.mnuToolsConfigureValues.Enabled = False
            mdiMain.mnuToolsMaintainProducts.Enabled = False
            mdiMain.mnuToolsEssentialSettings.Enabled = False
        End If
    End If
    If UCase$(App.ProductName) = "MANAGER" Or UCase$(App.ProductName) = "CLIENT" Or _
        UCase$(App.ProductName) = "MAINTENANCE" Then
        mdiMain.mnuToolsChangePassword.Visible = True
    End If
    
    Select Case UCase$(App.ProductName)
    Case "CLIENT"
        mdiMain.mnuToolsSep.Visible = True
        mdiMain.mnuToolsExternalPrograms.Visible = True
        mdiMain.mnuToolsExternalPrograms.Enabled = True
    Case "WAREHOUSE", "CONFIG"
        'Hides Grid layouts option
        mdiMain.mnuTools.Enabled = False
    Case Else
        mdiMain.mnuToolsSep.Visible = False
        mdiMain.mnuToolsExternalPrograms.Visible = False
        mdiMain.mnuToolsExternalPrograms.Enabled = False
    End Select
    
    'Help Menu items
    If UCase$(App.ProductName) = "LITE" Then
        If mdiMain.ActiveForm.Name = "frmAbout" Then
            mdiMain.mnuHelpCFU.Visible = True
            mdiMain.mnuHelpCFU.Enabled = True
        Else
            mdiMain.mnuHelpCFU.Visible = True
            mdiMain.mnuHelpCFU.Enabled = False
        End If
    End If

    If DebugVersion Then
        DebugFormControlSizes pobjForm
    End If
    
End Sub
Function StandardMenuOptions(pstrMenuItem As String, Optional plngIndex As Variant) As Boolean
Dim lintRetVal As Integer
Dim lintAlwaysMaximized As Integer

    StandardMenuOptions = False
    
    Select Case pstrMenuItem
    Case mdiMain.mnuFilePrintSetup.Caption
        StandardMenuOptions = True
        Printer.TrackDefault = False
        On Error GoTo Err_Handler
        With mdiMain.CommonDialog1
            .DialogTitle = "Print Setup"
            .CancelError = True
            .flags = cdlPDPrintSetup
            .ShowPrinter
        End With
    Case mdiMain.mnuFileExit.Caption
        StandardMenuOptions = True
        Unload mdiMain
    Case mdiMain.mnuEditCut.Caption
        On Error Resume Next
        Clipboard.SetText Screen.ActiveControl.Text
        Screen.ActiveControl.Text = ""
        On Error GoTo 0
    Case mdiMain.mnuEditCopy.Caption
        On Error Resume Next
        Clipboard.SetText Screen.ActiveControl.Text
        On Error GoTo 0
    Case mdiMain.mnuEditPaste.Caption
        On Error Resume Next
        Screen.ActiveControl.Text = Clipboard.GetText
        On Error GoTo 0
    Case mdiMain.mnuViewShowPicBar.Caption
        If mdiMain.picListBar.Visible = True Then
            mdiMain.picListBar.Visible = False
        Else
            mdiMain.picListBar.Visible = True
        End If
        SaveSetting gstrIniAppName, "UI", "PicBarVisible", mdiMain.picListBar.Visible
    Case mdiMain.mnuViewShowNewFeatures.Caption
        If mdiMain.ActiveForm.Name = "frmAbout" Or mdiMain.ActiveForm.Name = "frmMain" Or _
            mdiMain.ActiveForm.Name = "frmMainReps" Then
            If mdiMain.ActiveForm.fraFeatures.Visible = True Then
                mdiMain.ActiveForm.fraFeatures.Visible = False
            Else
                mdiMain.ActiveForm.fraFeatures.Visible = True
            End If
            SaveSetting gstrIniAppName, "UI", "ShowFeatures", mdiMain.ActiveForm.fraFeatures.Visible
        End If
    
    Case mdiMain.mnuViewMaxOnStartup.Caption
        lintAlwaysMaximized = Val(GetSetting(gstrIniAppName, "UI", "AlwaysMaximized"))
        If lintAlwaysMaximized = vbNormal Then
            lintAlwaysMaximized = vbMaximized
        Else
            lintAlwaysMaximized = vbNormal
        End If
        SaveSetting gstrIniAppName, "UI", "AlwaysMaximized", lintAlwaysMaximized
    Case mdiMain.mnuToolsMinder.Caption
        lintRetVal = MsgBox("Would you like to run Scandisk and Defrag?", vbYesNo, gconstrTitlPrefix & "Minder")
            If lintRetVal = vbYes Then
                gintForceAppClose = fcCompleteClose
                Unload mdiMain
                
                RunNWait AppPath & "Minder.exe" & " APP"
            End If
    Case mdiMain.mnuToolsResetGrid.Caption
        lintRetVal = MsgBox("Do you wish to reset the grid(s) wisths for this screen?", _
            vbYesNo, gconstrTitlPrefix & "Reset Grid(s) Widths")
        If lintRetVal = vbYes Then
            ResetGridLayout
            MsgBox "The Grid(s) Layout for this screen have been reset!", vbInformation, _
                gconstrTitlPrefix & "Reset Grid(s) Layout"
        End If
    Case mdiMain.mnuHelpContents.Caption
        StandardMenuOptions = True
        glngCurrentHelpHandle = HTMLHelp(mdiMain.hwnd, gstrHelpFileBase, HH_DISPLAY_TOPIC, 0)
    Case mdiMain.mnuHelpWhatsThis.Caption
        StandardMenuOptions = True
        mdiMain.ActiveForm.WhatsThisMode
    Case mdiMain.mnuHelpTutorial.Caption, mdiMain.mnuHelpQuickStart.Caption
        MsgBox "Under Construction!", vbInformation, gconstrTitlPrefix & "Help System"
    Case mdiMain.mnuHelpAbout.Caption
        StandardMenuOptions = True
        frmHelpAbout.Show vbModal
    End Select
    
    Exit Function
Err_Handler:
        Select Case Err.Number
        Case cdlCancel
            Exit Function
        Case Else
            Resume Next
        End Select
    
End Function

Sub ShowHistoryOption(pbooPackVis As Boolean, pbooModVis As Boolean, pbooHistVis As Boolean)

    mdiMain.mnuFileHistoryPackOrder1.Visible = pbooPackVis
    mdiMain.mnuFileHistoryPackOrder2.Visible = pbooPackVis
    mdiMain.mnuFileHistoryPackOrder3.Visible = pbooPackVis
    mdiMain.mnuFileHistoryPackOrder4.Visible = pbooPackVis
    mdiMain.mnuFileHistoryPackOrder5.Visible = pbooPackVis
    
    mdiMain.mnuFileHistoryModOrder1.Visible = pbooModVis
    mdiMain.mnuFileHistoryModOrder2.Visible = pbooModVis
    mdiMain.mnuFileHistoryModOrder3.Visible = pbooModVis
    mdiMain.mnuFileHistoryModOrder4.Visible = pbooModVis
    mdiMain.mnuFileHistoryModOrder5.Visible = pbooModVis
    
    mdiMain.mnuFileHistoryOrdHistory1.Visible = pbooHistVis
    mdiMain.mnuFileHistoryOrdHistory2.Visible = pbooHistVis
    mdiMain.mnuFileHistoryOrdHistory3.Visible = pbooHistVis
    mdiMain.mnuFileHistoryOrdHistory4.Visible = pbooHistVis
    mdiMain.mnuFileHistoryOrdHistory5.Visible = pbooHistVis
    
End Sub
Function DataType(plngType As Long) As String

    Select Case plngType
    Case 16 'dbBigInt
        DataType = "Big Integer"
    Case 9 'dbBinary
        DataType = "Binary"
    Case 1 'dbBoolean
        DataType = "Boolean"
    Case 2 'dbByte
        DataType = "Byte"
    Case 18 'dbChar
        DataType = "Char"
    Case 5 'dbCurrency
        DataType = "Currency"
    Case 8 'dbDate
        DataType = "Date/Time"
    Case 20 'dbDecimal
        DataType = "Decimal"
    Case 7 'dbDouble
        DataType = "Double"
    Case 21 'dbFloat
        DataType = "Float"
    Case 15 'dbGUID
        DataType = "GUID"
    Case 3 'dbInteger
        DataType = "Integer"
    Case 4 'dbLong
        DataType = "Long"
    Case 11 'dbLongBinary
        DataType = "Long Binary (OLE Object)"
    Case 12 'dbMemo
        DataType = "Memo"
    Case 19 'dbNumeric
        DataType = "Numeric"
    Case 6 'dbSingle
        DataType = "Single"
    Case 10 'dbText
        DataType = "Text"
    Case 22 'dbTime
        DataType = "Time"
    Case 23 'dbTimeStamp
        DataType = "Time Stamp"
    Case 17 'dbVarBinary
        DataType = "VarBinary"
    End Select
    
End Function
Sub TrainingCards(wParam As Long, lParam As Long)
On Error Resume Next

    Select Case wParam
    Case mTraLiteGetStad
        Select Case lParam
        Case mTraLiteGetStadValues
            mdiMain.ZOrder 0
            SendKeys "{F10}"
            SendKeys "{RIGHT 4}"
            SendKeys "{DOWN 3}"
        Case mTraLiteGetStadProds
            mdiMain.ZOrder 0
            SendKeys "{F10}"
            SendKeys "{RIGHT 4}"
            SendKeys "{DOWN 4}"
        Case mTraLiteGetStadEssen
            mdiMain.ZOrder 0
            SendKeys "{F10}"
            SendKeys "{RIGHT 4}"
            SendKeys "{DOWN 5}"
        End Select
    End Select
    
End Sub
Sub MdiChildPtrPos(pobjControl As Object)

    PosPtrOnCtl pobjControl, mdiMain.Left + mdiMain.picListBar.Width, 350

End Sub
Sub CopyHelpFile(pstrHelpFile)

    If Not DebugVersion Then
        FileCopyIfNewer gstrStatic.strServerPath & pstrHelpFile, gstrHelpFileBase
    End If
    
End Sub
