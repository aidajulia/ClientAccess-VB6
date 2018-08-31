VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   6195
   ClientTop       =   3885
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   13605
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   450
      Left            =   9360
      TabIndex        =   2
      Top             =   150
      Width           =   3240
   End
   Begin VB.TextBox txtOutput 
      Height          =   2250
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   615
      Width           =   13170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Connection"
      Height          =   375
      Left            =   255
      TabIndex        =   0
      Top             =   150
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim clientID As Integer
    Dim applicationID As Integer
    Dim UAT As Integer ' If Application Running In SandBox UAT = 1, In Production UAT = 0
    Dim sConnectionString As String
    'Create INI file. Should be removed after INI file created first time
'
'    WriteSetting "ClientAccess", "ClientID", "3"
'    WriteSetting "Application", "UAT", "0"
    ' UAT - 0 - Production
    ' UAT - 1 - Sandbox
'
    'ClientID and UAT can be modified later on in INI file
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    UAT = ReadSettingNumeric("Application", "UAT")
    clientID = ReadSettingNumeric("ClientAccess", "ClientID")
    If UAT = 0 Then
        applicationID = 1000 ' Real Application ID
    Else
        applicationID = 999 ' SandBox Application ID
    End If
    
    If clientID = 0 Then
        txtOutput.Text = "No Client ID" & vbCrLf & "Check the INI file."
    End If
    
    txtOutput.Text = "Loading. Please wait..."
    txtOutput.Refresh
    Dim ret As ReturnData
    ret = GetConnection(clientID, applicationID)
    If ret.ErrorCode = 0 Then
        txtOutput.Text = "SERVER: " & ret.Server & vbCrLf & "UID: " & ret.UID & vbCrLf & "PWD: " & ret.PWD
        sConnectionString = ret.ConnectionString
        txtOutput.Text = txtOutput.Text & vbCrLf & vbCrLf & sConnectionString
    Else
        txtOutput.Text = ret.ErrorDescription
    End If
End Sub

Private Sub Command2_Click()
Dim sSecret As String
txtOutput.Text = "Original: " & "Alex Dev"
sSecret = Encrypt("Alex Dev", "b1082abe-bf8c-4f91-97c6-c47934cbde01")
txtOutput.Text = txtOutput.Text & vbCrLf & "Encrypted: " & sSecret
sSecret = "MtdsPfcXbAw="
txtOutput.Text = txtOutput.Text & vbCrLf & "Decrypted: " & Decrypt(sSecret, "b1082abe-bf8c-4f91-97c6-c47934cbde01")

End Sub
