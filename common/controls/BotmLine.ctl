VERSION 5.00
Begin VB.UserControl ctlBottomLine 
   Alignable       =   -1  'True
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   ScaleHeight     =   360
   ScaleWidth      =   3885
   Begin VB.Line linWhite 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3735
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Line linGrey 
      BorderColor     =   &H00808080&
      X1              =   90
      X2              =   3750
      Y1              =   75
      Y2              =   75
   End
End
Attribute VB_Name = "ctlBottomLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub

Private Sub UserControl_Resize()

    If gbooJustPreLoading = False Then
        Cls
        Height = 700 '30
        
        With linGrey
            .X1 = 8
            .Y1 = 0
            .X2 = Width - 8
            .Y2 = 0
        End With
        
        With linWhite
            .X1 = 8
            .Y1 = 10
            .X2 = Width - 8
            .Y2 = 10
        End With
    End If
    
End Sub
