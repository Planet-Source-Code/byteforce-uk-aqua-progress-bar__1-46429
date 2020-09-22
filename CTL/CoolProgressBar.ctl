VERSION 5.00
Begin VB.UserControl CoolProgressBar 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   LockControls    =   -1  'True
   ScaleHeight     =   810
   ScaleWidth      =   2865
   ToolboxBitmap   =   "CoolProgressBar.ctx":0000
   Begin VB.PictureBox picStatusContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00F4EDEA&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   30
      ScaleHeight     =   510
      ScaleWidth      =   2805
      TabIndex        =   1
      Top             =   30
      Width           =   2805
      Begin VB.Image imgRight 
         Height          =   240
         Left            =   2640
         Picture         =   "CoolProgressBar.ctx":0312
         Top             =   165
         Width           =   240
      End
      Begin VB.Image imgSection 
         Height          =   180
         Left            =   165
         Picture         =   "CoolProgressBar.ctx":069C
         Stretch         =   -1  'True
         Top             =   165
         Width           =   2475
      End
      Begin VB.Image imgLeft 
         Height          =   240
         Left            =   -75
         Picture         =   "CoolProgressBar.ctx":076E
         Top             =   165
         Width           =   240
      End
   End
   Begin VB.Image imgCustom_Right 
      Height          =   240
      Index           =   4
      Left            =   1515
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCustom_Section 
      Height          =   180
      Index           =   4
      Left            =   1485
      Top             =   825
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgCustom_Left 
      Height          =   240
      Index           =   4
      Left            =   1245
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCustom_Right 
      Height          =   240
      Index           =   3
      Left            =   1170
      Picture         =   "CoolProgressBar.ctx":0AF8
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCustom_Section 
      Height          =   180
      Index           =   3
      Left            =   1140
      Picture         =   "CoolProgressBar.ctx":0E82
      Top             =   825
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgCustom_Left 
      Height          =   240
      Index           =   3
      Left            =   900
      Picture         =   "CoolProgressBar.ctx":0F54
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCustom_Right 
      Height          =   240
      Index           =   2
      Left            =   810
      Picture         =   "CoolProgressBar.ctx":12DE
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCustom_Section 
      Height          =   180
      Index           =   2
      Left            =   780
      Picture         =   "CoolProgressBar.ctx":1668
      Top             =   825
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgCustom_Left 
      Height          =   240
      Index           =   2
      Left            =   540
      Picture         =   "CoolProgressBar.ctx":173A
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCustom_Right 
      Height          =   240
      Index           =   1
      Left            =   465
      Picture         =   "CoolProgressBar.ctx":1AC4
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCustom_Section 
      Height          =   180
      Index           =   1
      Left            =   420
      Picture         =   "CoolProgressBar.ctx":1E4E
      Top             =   825
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgCustom_Left 
      Height          =   240
      Index           =   1
      Left            =   195
      Picture         =   "CoolProgressBar.ctx":1F20
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCustom_Right 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "CoolProgressBar.ctx":22AA
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCustom_Section 
      Height          =   180
      Index           =   0
      Left            =   90
      Picture         =   "CoolProgressBar.ctx":2634
      Top             =   825
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgCustom_Left 
      Height          =   240
      Index           =   0
      Left            =   -150
      Picture         =   "CoolProgressBar.ctx":2706
      Top             =   825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00EAD9CA&
      Caption         =   "CoolProgressBar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   30
      TabIndex        =   0
      Top             =   555
      Width           =   2805
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00996600&
      Height          =   810
      Left            =   0
      Top             =   0
      Width           =   2865
   End
End
Attribute VB_Name = "CoolProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// Public types
Public Enum cpbSkin
    [Blue] = 0
    [Pink] = 1
    [Red] = 2
    [Lime] = 3
    [Custom...] = 4
End Enum

'//Actual Local Property Values

Private m_lng_BorderColour As OLE_COLOR
Private m_lng_TextBackColor As OLE_COLOR
Private m_lng_TextColour As OLE_COLOR
Private m_lng_ProgressBackColour As OLE_COLOR
Private m_lng_DivisionLineColour As OLE_COLOR
Private m_str_Text As String
Private m_lng_Skin As cpbSkin
Private m_fnt_TextFont As New StdFont
Private m_int_Value As Integer

'//Default Property Constants

Private Const m_def_lng_BorderColour = &H996600
Private Const m_def_lng_TextBackColor = &HEAD9CA
Private Const m_def_lng_TextColour = &H404040
Private Const m_def_lng_ProgressBackColour = &HF4EDEA
Private Const m_def_lng_DivisionLineColour = &HFFFFFF
Private Const m_def_str_Text = "CoolProgressBar"
Private Const m_def_lng_Skin = 0
Private m_def_TextFont As New StdFont
Private Const m_def_int_Value As Integer = 0

Private Sub picStatusContainer_Click()

End Sub

Private Sub UserControl_Initialize()
    
    m_def_TextFont.Bold = False
    m_def_TextFont.Italic = False
    m_def_TextFont.Name = "Tahoma"
    m_def_TextFont.Size = "8"
    m_def_TextFont.Strikethrough = False
    m_def_TextFont.Underline = False

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_lng_BorderColour = PropBag.ReadProperty("BorderColour", m_def_lng_BorderColour)
    m_lng_TextBackColor = PropBag.ReadProperty("TextBackColor", m_def_lng_TextBackColor)
    m_lng_TextColour = PropBag.ReadProperty("TextColour", m_def_lng_TextColour)
    m_lng_ProgressBackColour = PropBag.ReadProperty("ProgressBackColour", m_def_lng_ProgressBackColour)
    m_lng_DivisionLineColour = PropBag.ReadProperty("DivisionLineColour", m_def_lng_DivisionLineColour)
    m_str_Text = PropBag.ReadProperty("Text", m_def_str_Text)
    m_lng_Skin = PropBag.ReadProperty("Skin", m_def_lng_Skin)
    Set m_fnt_TextFont = PropBag.ReadProperty("TextFont", m_def_TextFont)

    Dim img As IPictureDisp
    
    Set img = PropBag.ReadProperty("SkinCustomLeft", Nothing)
    
    If IsEmpty(img) = False Then Set imgCustom_Left(4) = img
    
    Set img = PropBag.ReadProperty("SkinCustomSection", Nothing)
    
    If IsEmpty(img) = False Then Set imgCustom_Section(4) = img
    
    Set img = PropBag.ReadProperty("SkinCustomRight", Nothing)
    
    If IsEmpty(img) = False Then Set imgCustom_Right(4) = img
    
    Call UpdateCtl
    
End Sub

Private Sub UserControl_Resize()
    
    '// Lock height and width
    
    If Not UserControl.Width = 2865 Then UserControl.Width = 2865
    If Not UserControl.Height = 810 Then UserControl.Height = 810
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderColour", m_lng_BorderColour, m_def_lng_BorderColour)
    Call PropBag.WriteProperty("TextBackColor", m_lng_TextBackColor, m_def_lng_TextBackColor)
    Call PropBag.WriteProperty("TextColour", m_lng_TextColour, m_def_lng_TextColour)
    Call PropBag.WriteProperty("ProgressBackColour", m_lng_ProgressBackColour, m_def_lng_ProgressBackColour)
    Call PropBag.WriteProperty("DivisionLineColour", m_lng_DivisionLineColour, m_def_lng_DivisionLineColour)
    Call PropBag.WriteProperty("Text", m_str_Text, m_def_str_Text)
    Call PropBag.WriteProperty("Skin", m_lng_Skin, m_def_lng_Skin)
    Call PropBag.WriteProperty("SkinCustomLeft", imgCustom_Left(4), Empty)
    Call PropBag.WriteProperty("SkinCustomSection", imgCustom_Section(4), Empty)
    Call PropBag.WriteProperty("SkinCustomRight", imgCustom_Right(4), Empty)
    Call PropBag.WriteProperty("TextFont", m_fnt_TextFont, m_def_TextFont)

End Sub

Private Sub UserControl_InitProperties()

    m_lng_BorderColour = m_def_lng_BorderColour
    m_lng_TextBackColor = m_def_lng_TextBackColor
    m_lng_TextColour = m_def_lng_TextColour
    m_lng_ProgressBackColour = m_def_lng_ProgressBackColour
    m_lng_DivisionLineColour = m_def_lng_DivisionLineColour
    m_str_Text = m_def_str_Text
    m_lng_Skin = m_def_lng_Skin
    Set m_fnt_TextFont = m_def_TextFont
    m_int_Value = m_def_int_Value
    
    Call UpdateCtl
    
End Sub

Public Property Let BorderColour(ByVal lng_BorderColour As OLE_COLOR)

    m_lng_BorderColour = lng_BorderColour

    PropertyChanged "BorderColour"

    Call UpdateCtl

End Property

Public Property Get BorderColour() As OLE_COLOR

    BorderColour = m_lng_BorderColour

End Property

Public Property Let TextBackColor(ByVal lng_TextBackColor As OLE_COLOR)

    m_lng_TextBackColor = lng_TextBackColor

    PropertyChanged "TextBackColor"

    Call UpdateCtl

End Property

Public Property Get TextBackColor() As OLE_COLOR

    TextBackColor = m_lng_TextBackColor

End Property

Public Property Let TextColour(ByVal lng_TextColour As OLE_COLOR)

    m_lng_TextColour = lng_TextColour

    PropertyChanged "TextColour"

    Call UpdateCtl

End Property

Public Property Get TextColour() As OLE_COLOR

    TextColour = m_lng_TextColour

End Property

Public Property Let ProgressBackColour(ByVal lng_ProgressBackColour As OLE_COLOR)

    m_lng_ProgressBackColour = lng_ProgressBackColour

    PropertyChanged "ProgressBackColour"

    Call UpdateCtl

End Property

Public Property Get ProgressBackColour() As OLE_COLOR

    ProgressBackColour = m_lng_ProgressBackColour

End Property

Public Property Let DivisionLineColour(ByVal lng_DivisionLineColour As OLE_COLOR)

    m_lng_DivisionLineColour = lng_DivisionLineColour

    PropertyChanged "DivisionLineColour"

    Call UpdateCtl

End Property

Public Property Get DivisionLineColour() As OLE_COLOR

    DivisionLineColour = m_lng_DivisionLineColour

End Property

Public Property Let Text(ByVal str_Text As String)

    m_str_Text = str_Text

    PropertyChanged "Text"

    Call UpdateCtl

End Property

Public Property Get Text() As String

    Text = m_str_Text

End Property

Public Property Let Skin(ByVal lng_SkinID As cpbSkin)

    m_lng_Skin = lng_SkinID

    PropertyChanged "Skin"

    Call UpdateCtl

End Property

Public Property Get Skin() As cpbSkin

    Skin = m_lng_Skin
    
End Property

Public Property Get SkinCustomLeft() As StdPicture
    
    Set SkinCustomLeft = imgCustom_Left(4)
   
End Property

Public Property Set SkinCustomLeft(ByVal picNewValue As IPictureDisp)
    
    Set imgCustom_Left(4) = picNewValue
    
    PropertyChanged "SkinCustomLeft"
    
    Call UpdateCtl
    
End Property

Public Property Get SkinCustomSection() As StdPicture
    
    Set SkinCustomSection = imgCustom_Section(4)
   
End Property

Public Property Set SkinCustomSection(ByVal picNewValue As IPictureDisp)
    
    Set imgCustom_Section(4) = picNewValue
    
    PropertyChanged "SkinCustomSection"
    
    Call UpdateCtl
    
End Property

Public Property Get SkinCustomRight() As StdPicture
    
    Set SkinCustomRight = imgCustom_Right(4)
   
End Property

Public Property Set SkinCustomRight(ByVal picNewValue As IPictureDisp)
    
    Set imgCustom_Right(4) = picNewValue
    
    PropertyChanged "SkinCustomRight"
    
    Call UpdateCtl
    
End Property

Public Property Set TextFont(ByVal fnt_TextFont As StdFont)
    
    Set m_fnt_TextFont = fnt_TextFont
    
    PropertyChanged "TextFont"

    Call UpdateCtl
    
End Property

Public Property Get TextFont() As StdFont
    
    Set TextFont = m_fnt_TextFont
    
End Property

Public Sub UpdateCtl()
    
    '// Update Control to reflect property changes
    
    'Apply the correct skin
    On Error Resume Next
    
    imgLeft.Stretch = False
    Set imgLeft = imgCustom_Left(m_lng_Skin)
    imgLeft.Stretch = True
    
    Set imgSection = imgCustom_Section(m_lng_Skin)
    
    imgRight.Stretch = False
    Set imgRight = imgCustom_Right(m_lng_Skin)
    imgRight.Stretch = True
    
    'Set font
    Set lblCaption.Font = m_fnt_TextFont
    
    'Set label caption
    If Not lblCaption = m_str_Text Then lblCaption = m_str_Text
    
    'Update colours
    If Not shpBorder.BorderColor = m_lng_BorderColour Then shpBorder.BorderColor = m_lng_BorderColour
    If Not picStatusContainer.BackColor = m_lng_ProgressBackColour Then picStatusContainer.BackColor = m_lng_ProgressBackColour
    If Not UserControl.BackColor = m_lng_DivisionLineColour Then UserControl.BackColor = m_lng_DivisionLineColour
    If Not lblCaption.BackColor = m_lng_TextBackColor Then lblCaption.BackColor = m_lng_TextBackColor
    If Not lblCaption.ForeColor = m_lng_TextColour Then lblCaption.ForeColor = m_lng_TextColour
    
End Sub

Public Sub SetProgress(Percent As Integer)
    
    '// Resize images to reflect the percentage change specified
    '// in Percent.
    
    'Make sure only values zero to 100 are handled
    If Percent < 0 Then Err.Raise 380: Exit Sub
    If Percent > 100 Then Err.Raise 380: Exit Sub
    
    If Percent = 0 Then
        
        'Hide all the images, as the percent is set to zero (nothing)
        If Not imgLeft.Visible = False Then imgLeft.Visible = False
        If Not imgSection.Visible = False Then imgSection.Visible = False
        If Not imgRight.Visible = False Then imgRight.Visible = False
        
        'Exit the sub after hiding the image controls
        Exit Sub
        
    End If
    
    'Ensure images are visible as the percentage is 1% or more
    If Not imgLeft.Visible = True Then imgLeft.Visible = True
    If Not imgSection.Visible = True Then imgSection.Visible = True
    If Not imgRight.Visible = True Then imgRight.Visible = True
    
    
    If Percent >= 99 Then
        
        'Make the section maximum size, and cap both sides with the left
        'and right images
        If Not imgLeft.Left = -75 Then imgLeft.Left = -75
        If Not imgSection.Left = imgLeft.Left + imgLeft.Width Then imgSection.Left = imgLeft.Left + imgLeft.Width
        If Not imgSection.Width = 2475 Then imgSection.Width = 2475
        If Not imgRight.Left = imgSection.Left + imgSection.Width Then imgRight.Left = imgSection.Left + imgSection.Width
        
        Exit Sub
        
    End If
    
    'Work out percentage unit, ie, 1% of 2475 = 2475/100
    Dim sng_PercentUnit As Single
    
    sng_PercentUnit = CSng(Round(2475 / 100))
    
    
    'Resize images according to percent
    Dim sng_NewSize As Single
    
    sng_NewSize = Percent * sng_PercentUnit
            
    If Not imgLeft.Left = -75 Then imgLeft.Left = -75
    If Not imgSection.Left = imgLeft.Left + imgLeft.Width Then imgSection.Left = imgLeft.Left + imgLeft.Width
    If Not imgSection.Width = sng_NewSize Then imgSection.Width = sng_NewSize
    If Not imgRight.Left = imgSection.Left + imgSection.Width Then imgRight.Left = imgSection.Left + imgSection.Width
        
End Sub
