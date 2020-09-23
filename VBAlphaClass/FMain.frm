VERSION 5.00
Begin VB.Form FMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alpha Class Test Application"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetAlpha 
      Caption         =   "CHANGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1883
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtAlphaValue 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1103
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "150"
      ToolTipText     =   "Enter a number between 0 and 255 here."
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblDisplay 
      Caption         =   $"FMain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4335
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variable to hold the instance of our class
Dim objAlpha As clsAlpha

Private Sub cmdSetAlpha_Click()
    ' Values less than 100 are being disabled to make sure that
    ' someone does not make the form completely transparent.
    ' Once you have tested it, you can remove this limitation.
    
    If CByte(txtAlphaValue.Text) < 100 Then
        
        ' Update the display to show that we are overriding the
        ' value specified by the user.
        
        MsgBox "Be careful using values less than 100, you might" & _
            vbNewLine & "not be able to see the form!", _
            vbInformation, App.Title
        
        txtAlphaValue.Text = 100
        
        objAlpha.SetLayered Me.hwnd, True, CByte(txtAlphaValue.Text)
    Else
        objAlpha.SetLayered Me.hwnd, True, CByte(txtAlphaValue.Text)
    End If
End Sub

Private Sub Form_Load()
    ' Load an instance of the class object into the variable.
    Set objAlpha = New clsAlpha
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Return the form to the origional state.
    objAlpha.ReleaseDisplay Me.hwnd
    
    ' Cleanup
    Set objAlpha = Nothing
End Sub

Private Sub txtAlphaValue_Change()
    ' Check that the value entered is withing the range of 0 -> 255.
    ' Convert to an integer just incase the value is over 255.
    If Len(txtAlphaValue.Text) = 0 Or Not IsNumeric(txtAlphaValue.Text) Then
        ' Check for null value
        txtAlphaValue.Text = "0"
    Else
        ' Make sure the value is valid
        If CInt(txtAlphaValue.Text) > 255 Then txtAlphaValue.Text = 255
        If CInt(txtAlphaValue.Text) < 0 Then txtAlphaValue.Text = 0
    End If
End Sub
