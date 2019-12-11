VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "User Creds"
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4350
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBegin_Click()

'Runs external program that goes off to the server and gets the files based on username, password @ branch number passed to it as an argument

    varBranch = ActiveSheet.Name
    Shell ("\\ukhudbdclusfs6.sgbd.co.uk\it\ITServices\Jewson IT Business Services\Application Support Team\G&DRecci\middleware\GandDRecciLiloFetcher.exe lilo " & txtUser.Value & " " & txtPword.Value & " /k8live/live/data/STOCK/ICONMIG " & varBranch)
     
'Checks for token file created by program to determin if the script ran successfully
    x = 0
    
    Do While Dir("c:\null\recci\great.success") = ""
        x = x + 1
        If Dir("c:\null\recci\failed.txt") <> "" Then
            varReturn = MsgBox("Error downloading file, please check details")
            Exit Do
        End If
       
        Application.Wait (Now + TimeValue("0:00:1"))
       
        If x > 20 Then
            varReturn = MsgBox("Time Out Error, please check details")
            Exit Do
        End If
    Loop
    
    If Dir("c:\null\recci\great.success") <> "" Then
        varReturn = MsgBox("Pulled Files Successfully")
        Unload frmLogin
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload frmLogin
End Sub

Private Sub UserForm_Click()

End Sub
