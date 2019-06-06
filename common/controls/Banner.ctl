VERSION 5.00
Begin VB.UserControl ctlBanner 
   Alignable       =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8610
   ScaleHeight     =   915
   ScaleWidth      =   8610
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   7800
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image imgButton 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   9580
      Top             =   142
      Width           =   735
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Banner Caption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7410
   End
End
Attribute VB_Name = "ctlBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim llngDarkGrey As Long

'Default Property Values:
Const m_def_Caption = "0"
'Property Variables:
Dim m_Caption As String

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
    Caption = m_Caption
    lblCaption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    lblCaption = New_Caption
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_Caption = m_def_Caption
    
End Sub
Private Sub UserControl_Paint()

    DrawLines
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    lblCaption.Caption = m_Caption
    Set Picture = PropBag.ReadProperty("Picture", Nothing)

    lblCaption.WhatsThisHelpID = PropBag.ReadProperty("WhatIsID", 0)
    picImage.WhatsThisHelpID = PropBag.ReadProperty("WhatIsID", 0)
    imgButton.WhatsThisHelpID = PropBag.ReadProperty("WhatIsID", 0)
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("WhatIsID", lblCaption.WhatsThisHelpID, 0)
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Picture
Public Property Get Picture() As Picture

    Set Picture = imgButton.Picture
    
End Property

Public Property Set Picture(ByVal New_Picture As Picture)

    Set imgButton.Picture = New_Picture
    PropertyChanged "Picture"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,WhatsThisHelpID
Public Property Get WhatIsID() As Long
    WhatIsID = lblCaption.WhatsThisHelpID
End Property

Public Property Let WhatIsID(ByVal New_WhatIsID As Long)

    lblCaption.WhatsThisHelpID() = New_WhatIsID
    picImage.WhatsThisHelpID() = New_WhatIsID
    imgButton.WhatsThisHelpID() = New_WhatIsID
    
    PropertyChanged "WhatIsID"
End Property

Private Sub UserControl_Resize()

    DrawLines

End Sub
Sub DrawLines()
Const lconHVar = 30
Const lconWVar = 8 '10
    
    If gbooJustPreLoading = False Then
	    Cls
	
	    Height = 1050 '1100
	    llngDarkGrey = RGB(127, 127, 127)
	    
	    Line (8, Height - lconHVar)-(Width - lconWVar, Height - lconHVar), llngDarkGrey ' bottom
	    Line (Width - lconWVar, 0)-(Width - lconWVar, Height), vbButtonFace  ' right
	
	    imgButton.Move ((Width - imgButton.Width) - 200), ((Height / 2) - (imgButton.Height / 2) - 30)
	    
    End If
End Sub
