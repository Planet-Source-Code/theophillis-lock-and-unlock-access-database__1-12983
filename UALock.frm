VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form UALock 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock Unlock Database"
   ClientHeight    =   2430
   ClientLeft      =   3225
   ClientTop       =   3495
   ClientWidth     =   3825
   FillStyle       =   0  'Solid
   Icon            =   "UALock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "UALock.frx":0442
   ScaleHeight     =   2430
   ScaleWidth      =   3825
   Begin MSComDlg.CommonDialog Dlg1 
      Left            =   4200
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Database"
      Filter          =   "Ms Access DB|*.MDB"
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Option 
      Caption         =   "&Options"
      Begin VB.Menu LockDB 
         Caption         =   "&Lock"
      End
      Begin VB.Menu UnlockDB 
         Caption         =   "&Unlock"
      End
   End
End
Attribute VB_Name = "UALock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim objAccess As Object

Private Function AppControl(strFilename As String, booLock As Boolean)
'Checkes if the filename selected is a database
'and that cancel was not selected
   If strFilename = "" Then
        Exit Function
    End If
    If Right(strFilename, 3) <> "mdb" Then
        Exit Function
    End If
    If booLock = True Then
        LockDataBase (strFilename)
    Else
        UnlockDatabase (strFilename)
    End If
End Function
Private Function UnlockDatabase(strDBName)
On Error Resume Next
'This function enables the shift key
'when openning the database
'The shift key is used to bypass the startup options

Set db = DBEngine.Workspaces(0).OpenDatabase(strDBName)
db.Properties.Delete "Allowbypasskey"
db.Properties.Refresh
Set objp = db.CreateProperty("Allowbypasskey", dbBoolean, 1)
db.Properties.Append objp
db.Close
End Function
Private Function LockDataBase(strDBName)
On Error Resume Next
'This function disbles the shift key
Set db = DBEngine.Workspaces(0).OpenDatabase(strDBName)
db.Properties.Delete "Allowbypasskey" 'Deletes the Property from the database object
db.Properties.Refresh
Set objp = db.CreateProperty("Allowbypasskey", dbBoolean, 0)
db.Properties.Append objp 'addes the property to the database object
db.Close
End Function
Private Sub Exit_Click()
    End
End Sub

Private Sub LockDB_Click()
    Dlg1.InitDir = App.Path
    Dlg1.ShowOpen
    Call AppControl(Dlg1.FileName, True)
End Sub

Private Sub UnlockDB_Click()
    Dlg1.InitDir = App.Path
    Dlg1.ShowOpen
    Call AppControl(Dlg1.FileName, False)
End Sub
