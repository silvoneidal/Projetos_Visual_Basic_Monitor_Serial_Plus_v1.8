VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconectado"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13485
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   13485
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboChrEspecial 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":169B2
      Left            =   11400
      List            =   "Form1.frx":169C2
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox txtReceive 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   13455
   End
   Begin VB.ComboBox cboSend 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   11295
   End
   Begin MSComctlLib.Slider sldBaudRate 
      Height          =   555
      Left            =   -120
      TabIndex        =   4
      Top             =   6240
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   979
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   8
      SelStart        =   1
      Value           =   1
   End
   Begin MSComctlLib.Slider sldCommPort 
      Height          =   555
      Left            =   -120
      TabIndex        =   3
      Top             =   5640
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   979
      _Version        =   393216
      LargeChange     =   1
      Max             =   16
      SelStart        =   1
      TickStyle       =   1
      Value           =   1
   End
   Begin VB.Label Label25 
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "COM16"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   12840
      TabIndex        =   28
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "COM15"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   12120
      TabIndex        =   27
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "COM14"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   11280
      TabIndex        =   26
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "COM13"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   10560
      TabIndex        =   25
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "COM12"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   9720
      TabIndex        =   24
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "COM11"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   8880
      TabIndex        =   23
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "COM10"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   8040
      TabIndex        =   22
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "COM9"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   7320
      TabIndex        =   21
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "COM8"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "COM7"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   5640
      TabIndex        =   19
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "COM6"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "COM5"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "COM4"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "COM3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "COM2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "COM1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label24 
      Caption         =   "115200"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12600
      TabIndex        =   12
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "57600"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   11
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label22 
      Caption         =   "38400"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   10
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label21 
      Caption         =   "19200"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label20 
      Caption         =   "9600"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label19 
      Caption         =   "4800"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   "2400"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "1200"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6840
      Width           =   615
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mClear 
         Caption         =   "Clear"
         Begin VB.Menu mEnviados 
            Caption         =   "Enviados"
         End
         Begin VB.Menu mRecebidos 
            Caption         =   "Recebidos"
         End
      End
      Begin VB.Menu mScan 
         Caption         =   "Scan"
         Begin VB.Menu mPortaCOM 
            Caption         =   "PortaCOM"
         End
      End
      Begin VB.Menu mFormat 
         Caption         =   "Format"
         Begin VB.Menu mCorEditor 
            Caption         =   "Cor do Editor"
         End
         Begin VB.Menu mCorTexto 
            Caption         =   "Cor do Texto"
         End
         Begin VB.Menu mFonteTexto 
            Caption         =   "Fonte do Texto"
         End
      End
      Begin VB.Menu mGerenciador 
         Caption         =   "Gerênciador"
         Begin VB.Menu mDispositivos 
            Caption         =   "Dispositivos"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Variável global
Dim baudrate(8) As String

Private Sub Form_Load()
   Me.Caption = App.Title & "_v" & App.Major & "." & App.Minor & " by DALÇOQUIO AUTOMAÇÃO"
   writeMensagem ("Desconectado !!!")
   cboChrEspecial.Text = "None"
      
   txtReceive.ToolTipText = "De um Duplo Click para abrir o menu de opções ."
   cboSend.ToolTipText = "Digite o dado a ser enviado, depois pressione Enter para Send."
   cboChrEspecial.ToolTipText = "Selecione o caracter especial para ser enviado no final do dado, ou None para nenhum."
   sldCommPort.ToolTipText = "Selecione a Porta COM para Conectar ou  OFF para Desconectar."
   sldBaudRate.ToolTipText = "Selecione a velocidade de comunicação da serial."
   
   Call scanCommPort
   
   baudrate(1) = 1200
   baudrate(2) = 2400
   baudrate(3) = 4800
   baudrate(4) = 9600
   baudrate(5) = 19200
   baudrate(6) = 38400
   baudrate(7) = 57600
   baudrate(8) = 115200
   sldBaudRate.Value = 4
   sldCommPort.Value = 0
   
End Sub

Private Sub scanCommPort()
   
   Dim i As Integer
   For i = 1 To 16 'Procura portas COM de 1 a 16
      MSComm1.CommPort = i
      On Error Resume Next 'ignora o tratamento de erro
      MSComm1.PortOpen = True 'tenta abrir a porta
      If Err.Number = 0 Then 'a porta está disponível
         Select Case i
            Case 1
               Label1.ForeColor = vbBlue
            Case 2
               Label2.ForeColor = vbBlue
            Case 3
               Label3.ForeColor = vbBlue
            Case 4
               Label4.ForeColor = vbBlue
            Case 5
               Label5.ForeColor = vbBlue
            Case 6
               Label6.ForeColor = vbBlue
            Case 7
               Label7.ForeColor = vbBlue
            Case 8
               Label8.ForeColor = vbBlue
            Case 9
               Label9.ForeColor = vbBlue
            Case 10
               Label10.ForeColor = vbBlue
            Case 11
               Label11.ForeColor = vbBlue
            Case 12
               Label12.ForeColor = vbBlue
            Case 13
               Label13.ForeColor = vbBlue
            Case 14
               Label14.ForeColor = vbBlue
            Case 15
               Label15.ForeColor = vbBlue
            Case 16
               Label16.ForeColor = vbBlue
            Case Else
               'none
         End Select
         MSComm1.PortOpen = False 'fecha a porta
      End If
      On Error GoTo 0 'ativa o tratamento de erro novamente
   Next i

End Sub

Private Sub sldBaudRate_Change()
   MSComm1.Settings = baudrate(sldBaudRate.Value) & ",n,8,1"
   If MSComm1.PortOpen = False Then Exit Sub
   writeMensagem ("Conectado na COM" & sldCommPort.Value & "," & MSComm1.Settings)
   
End Sub

Private Sub sldCommPort_Change()
   On Error GoTo Erro
      ' Desconectar
      If MSComm1.PortOpen = True Then
         MSComm1.PortOpen = False
         writeMensagem ("Desconectado !!!")
      End If
      
      Sleep (1000)
      
      ' Conectar
      If MSComm1.PortOpen = False And sldCommPort.Value <> 0 Then
         MSComm1.Settings = baudrate(sldBaudRate.Value)
         MSComm1.CommPort = sldCommPort.Value
         MSComm1.PortOpen = True
         writeMensagem ("Conectado na COM" & sldCommPort.Value & "," & MSComm1.Settings)
      End If
   Exit Sub
   
Erro:
writeMensagem (Error)
Beep

End Sub

Private Sub MSComm1_OnComm()
   On Error GoTo Erro
      Dim strData As String
      Do While MSComm1.InBufferCount > 0
          If strData = Empty Then Exit Do
          strData = MSComm1.Input(MSComm1.InBufferCount)
      Loop
      txtReceive.Text = txtReceive.Text + MSComm1.Input
      txtReceive.SelStart = Len(txtReceive.Text)
   Exit Sub
   
Erro:
   writeMensagem (Error)
Beep
   
End Sub

Private Sub cboSend_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And MSComm1.PortOpen = True Then
      Call sendDado
      writeMensagem ("Enviado com sucesso...")
      Beep
   End If

End Sub

Private Sub sendDado()
   If cboChrEspecial = "None" Then
      MSComm1.Output = cboSend.Text
   ElseIf cboChrEspecial = "Nova Linha" Then
      MSComm1.Output = cboSend.Text & vbLf 'vbNewLine
   ElseIf cboChrEspecial = "Retorno de Carro" Then
      MSComm1.Output = cboSend.Text & vbCr
   ElseIf cboChrEspecial = "Ambos, NL e CR" Then
      MSComm1.Output = cboSend.Text & vbCrLf
   End If
   
   '------------------------------------------------------
   'atualiza cboSend
   cboSend.AddItem cboSend.Text
   'cboSend.Text = Clear
   
   'remove duplicados
   Dim i As Integer, j As Integer
   For i = 0 To cboSend.ListCount
       For j = i + 1 To cboSend.ListCount
           If cboSend.List(i) = cboSend.List(j) Then
               cboSend.RemoveItem (j)
               j = j - 1
           End If
       Next
   Next
   '------------------------------------------------------

End Sub

Private Sub clearSend()
    Dim i As Integer, j As Integer
    For i = 0 To cboSend.ListCount
        For j = i + 1 To cboSend.ListCount
            cboSend.RemoveItem (i)
        Next
    Next
    cboSend.Text = ""
End Sub

Private Sub txtReceive_DblClick()
   PopupMenu mMenu
  
End Sub

Private Sub mEnviados_Click()
   Call clearSend
   
End Sub

Private Sub mRecebidos_Click()
   txtReceive.Text = Empty
   
End Sub

Private Sub mPortaCOM_Click()
   If MSComm1.PortOpen = True Then
      MSComm1.PortOpen = False
      writeMensagem ("Desconectado !!!")
   End If
   
   sldCommPort.Value = 0
   Label1.ForeColor = &H808080
   Label2.ForeColor = &H808080
   Label3.ForeColor = &H808080
   Label4.ForeColor = &H808080
   Label5.ForeColor = &H808080
   Label6.ForeColor = &H808080
   Label7.ForeColor = &H808080
   Label8.ForeColor = &H808080
   Label9.ForeColor = &H808080
   Label10.ForeColor = &H808080
   Label11.ForeColor = &H808080
   Label12.ForeColor = &H808080
   Label13.ForeColor = &H808080
   Label14.ForeColor = &H808080
   Label15.ForeColor = &H808080
   Label16.ForeColor = &H808080
   writeMensagem ("Scanning...")
   Call scanCommPort
   writeMensagem ("Scan finalizado com sucesso...")
   
End Sub

Private Sub mCorTexto_Click()
    CommonDialog1.ShowColor
    txtReceive.ForeColor = CommonDialog1.Color
End Sub

Private Sub mCorEditor_Click()
    CommonDialog1.ShowColor
    txtReceive.BackColor = CommonDialog1.Color
End Sub

Private Sub mFonteTexto_Click()
    CommonDialog1.Flags = CommonDialog1CFBoth
    CommonDialog1.ShowFont
    txtReceive.Font = CommonDialog1.FontName
    txtReceive.FontBold = CommonDialog1.FontBold
    txtReceive.FontItalic = CommonDialog1.FontItalic
    txtReceive.FontSize = CommonDialog1.FontSize
End Sub

Private Sub mDispositivos_Click()
    Shell ("cmd.exe /c devmgmt.msc")

End Sub

Private Sub writeMensagem(mensagem As String)
   txtReceive.Text = txtReceive.Text & "> " & mensagem & vbCrLf
   txtReceive.SelStart = Len(txtReceive.Text)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MSComm1.PortOpen = True Then
      MSComm1.PortOpen = False
      writeMensagem ("Desconectado !!!")
   End If
   writeMensagem ("Fechando o sistema...")
   DoEvents
   Sleep (1000)
   End

End Sub
