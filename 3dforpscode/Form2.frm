VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   ClientHeight    =   -600
   ClientLeft      =   3735
   ClientTop       =   2325
   ClientWidth     =   90
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   -600
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Begin VB.Menu menu 
      Caption         =   ""
      Begin VB.Menu about2 
         Caption         =   "about"
      End
      Begin VB.Menu save 
         Caption         =   "SavePicture"
      End
      Begin VB.Menu close 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about2_Click()
about.Timer1.Enabled = True
about.Show
End Sub
Private Sub close_Click()
End
Unload Me
End Sub
Private Sub save_Click()
On Error GoTo l
    Form1.cdisave.Filter = "(*.bmp)| *.bmp"
    Form1.cdisave.ShowSave
    SavePicture Form1.FX.Image, Form1.cdisave.FileName
l:
End Sub
