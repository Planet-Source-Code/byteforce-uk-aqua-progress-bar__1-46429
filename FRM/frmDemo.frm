VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H00F4EDEA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CoolProgressBar Demo"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDemo 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2730
      Top             =   1365
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   135
      TabIndex        =   1
      Top             =   1125
      Width           =   2865
   End
   Begin ctlCoolProgress.CoolProgressBar cpbDemo 
      Height          =   810
      Left            =   135
      TabIndex        =   0
      Top             =   195
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1429
      Text            =   "www.ispn-online.co.uk"
      SkinCustomLeft  =   "frmDemo.frx":0000
      SkinCustomSection=   "frmDemo.frx":001C
      SkinCustomRight =   "frmDemo.frx":0038
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// CoolProgressBar UserControl
'// ***************************
'//
'// Original design by Matthew Hall 2003
'// ------------------------------------
'//
'// You may use this code project or any related element
'// of the product so long as you include proper reference to
'// the author, Matthew Hall 2003 (www.ispn-online.co.uk).
'//
'// YOU MAY NOT SELL (in a compiled or open format) ANY
'// COMPONENT OF THIS PRODUCT.
'//
'// All Rights Reserved.
'// ©2003 Matthew Hall
'// www.ispn-online.co.uk

'// General Declarations

'Integer variable to hold current percentage value
Private int_Counter%

'// Object Proceedures

Private Sub cmdGo_Click()
    
    cpbDemo.SetProgress 0
    
    Randomize Timer
    
    Dim int_Skin As Integer
    int_Skin = CInt(Int((4 * Rnd()) + 1))
    
    cpbDemo.Skin = int_Skin - 1
    
    '// Start the demo
    
    'Disable command button
    cmdGo.Enabled = False
    
    'Start our timer (if not already started)
    tmrDemo.Enabled = True
    
    'Set progress bar text to equal '0%'
    cpbDemo.Text = "0%"
    
End Sub

Private Sub Form_Load()
    
    'Set progress bar value to equal zero
    cpbDemo.SetProgress 0

End Sub

Private Sub tmrDemo_Timer()
    
    '// Update progress bar
    
    'Increment the counter by 1
    int_Counter% = int_Counter% + 1
    
    'Set the cool progress bar object value
    cpbDemo.SetProgress int_Counter%
    
    'Set the progress bar text value to the current percentage
    cpbDemo.Text = int_Counter% & "%"
        
    'See if we have reached the end (100%)
    If int_Counter% = 100 Then
        
        'We have reached the end so reset everything
        tmrDemo.Enabled = False
        int_Counter% = 0
        cmdGo.Enabled = True
                
    End If

End Sub
