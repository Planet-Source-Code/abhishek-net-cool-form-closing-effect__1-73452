VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cool Form Closing Effect"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
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
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
'                   Cool Form Closing Effect
'Just a snippet i wrote to show fancy form closing effect.
'==============================================================================
'Written By: abhishek
'Email: abhishek007p@hotmail.com
'       binarylife9@yahoo.com
'Add me to your yahoo or msn if u like to chat about programming
'and looking forward to working together on projects.
'==============================================================================

Private Sub Command1_Click()
    Dim i As Long
    
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)

    For i = Me.Left To (Screen.Width / 2) Step 10
        Me.Height = Me.Height - 15
        Me.Width = Me.Width - 20
        Me.Left = Me.Left + 100
        DoEvents
    Next
    
    Unload Me

End Sub

'==============================================================================
'                      Love VB6? then Sign this Petition
'                       15K+ People have already signed
' A PETITION FOR THE DEVELOPMENT OF UNMANAGED (Non .NET) VISUAL BASIC AND VBA
'
'                   http://www.classicvb.org/Petition
'
' Include this message in all your PSC or other submissions if possible.
'==============================================================================
'                   *Disadvantages of .NET*
'1. It is a web platform like java, not optimized for local app development
'2. A huge runtime to run your applications
'   Size of .NET runtime
'    .NET 2.0 ~ 23 MB
'    .NET 3.5 ~ 231 MB
'3. .NET is slower then VB6, Needs more memory
'4. .NET Apps are not complied to native code (machine code) like in VB6 & C++
'   so anyone can decompile your apps and view the source code
'==============================================================================

