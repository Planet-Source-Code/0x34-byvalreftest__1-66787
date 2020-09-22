VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reference Example - 0x34"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Call ByRef"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Call ByVal"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton CallFunct 
      Caption         =   "Call Function"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Starting Variable (Z) = 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Example of WHY IS IT SO IMPORTANT to use the right Reference Value!!!  0x34

Private Sub CallFunct_Click()
Dim Z As Long

    Z = 5   ' here's the variable we may or may not want altered by the function
    
    If Option1 Then
        Funct1 Z  'Call ByVal Function
    Else
        Funct2 Z 'Call ByRef Function
    End If
    
    Label1 = " Variable (Z) after Function Call = " & Z

End Sub

Private Function Funct1(ByVal Y As Long) ' Don't modify calling sub's sent variable - ByVal

    Y = Y + 10  'Add 10 to sent variable
    
    Label2 = " Function Value of sent Variable = " & Y

End Function


Private Function Funct2(ByRef Y As Long) ' Modify the calling sub's sent variable - ByRef

    Y = Y + 10  'Add 10 to sent variable
    
    Label2 = " Function Value of sent Variable = " & Y

End Function

'  Both functions are identical, except for the reference to the sent data(register). If the value of Z is
'  sent "ByVal", then the original contents of Z are not altered by any math within the function. If sent
'  "ByRef" then the contents will be altered by any math within the called function.

'  I know this is basic, however there are some who are unfamiliar with the difference.
'  This caused a major bug in one of my programs during an API call. I thought I'd pass it on.

'  VB will default to "ByRef" so, if you don't want the data altered by the function, use "ByVal"
