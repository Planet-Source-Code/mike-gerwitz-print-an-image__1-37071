VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Image Example"
   ClientHeight    =   5385
   ClientLeft      =   2070
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "End"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Image"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   120
      ScaleHeight     =   5115
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin MSComDlg.CommonDialog dlg 
         Left            =   4320
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnd_Click()
    End                 'End the program
End Sub

Private Sub cmdLoad_Click()
    'Open the open dialog
    With dlg
        .DialogTitle = "Load Image"
        .Filter = "Bitmap (*.bmp)|*.bmp|All Files (*.*)|*.*"    'Allow the opening of Bitmap files, or all files
        .ShowOpen
    End With
    
    pic.Picture = LoadPicture(dlg.FileName)             'Load the picture using the info we recieved from dlg
End Sub

Private Sub cmdPrint_Click()
    'Print the image
    Printer.PaintPicture pic.Image, 10, 10
    Printer.EndDoc              'You need this to start printing before the program ends
End Sub
