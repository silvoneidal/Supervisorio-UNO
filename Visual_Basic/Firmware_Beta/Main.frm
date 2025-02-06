VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7185
   ClientLeft      =   1350
   ClientTop       =   2970
   ClientWidth     =   6645
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   6720
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame frmInterface 
      Caption         =   "Interface"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   6375
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copiar "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   28
         Top             =   5640
         Width           =   3015
      End
      Begin VB.TextBox txtInformacao 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   3600
         Width           =   6135
      End
      Begin VB.Frame Frame2 
         Height          =   3855
         Left            =   12120
         TabIndex        =   20
         Top             =   120
         Width           =   1455
         Begin VB.ComboBox cboPinos 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   26
            Text            =   "Combo1"
            Top             =   2280
            Width           =   1215
         End
         Begin VB.OptionButton optInput 
            Caption         =   "1.Input"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optOutput 
            Caption         =   "2.Output"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optAnalog 
            Caption         =   "3.Analog"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton optPwm 
            Caption         =   "4.Pwm"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdSendProg 
            Caption         =   "Enviar"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            TabIndex        =   21
            Top             =   2760
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3240
         TabIndex        =   19
         Top             =   5640
         Width           =   3015
      End
      Begin VB.Frame frmPwm 
         Caption         =   "Pwm"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   4800
         TabIndex        =   14
         Top             =   240
         Width           =   1455
         Begin VB.ListBox lstPwm 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2730
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame frmAnalog 
         Caption         =   "Analog"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   1455
         Begin VB.ListBox lstAnalog 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2730
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame frmOutput 
         Caption         =   "Output"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   1455
         Begin VB.ListBox lstOutput 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2730
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame frmInput 
         Caption         =   "Input"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
         Begin VB.ListBox lstInput 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2730
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.Frame frmConnect 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdScanPort 
         Caption         =   "Scanear"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdConnectPort 
         Caption         =   "Conectar"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4680
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboBaudRate 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboCommPort 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmTerminal 
      Caption         =   "Terminal"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6375
      Begin VB.CommandButton cmdSendData 
         Caption         =   "Enviar"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3000
         TabIndex        =   4
         Top             =   5640
         Width           =   1575
      End
      Begin VB.ComboBox cboSendData 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   5640
         Width           =   2775
      End
      Begin VB.CommandButton cmdClearTerminal 
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4680
         TabIndex        =   2
         Top             =   5640
         Width           =   1575
      End
      Begin VB.TextBox txtTerminal 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Shape shpConnect 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   -120
      Top             =   7080
      Width           =   6855
   End
   Begin VB.Menu mTerminal 
      Caption         =   "Terminal"
   End
   Begin VB.Menu mGerenciador 
      Caption         =   "Gerenciador de Dispositivos"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Microsoft Comm Control 6.0
' Microsoft Windows Common Controls 6.0 (SP6)

' Para uso de sleep
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Dim Titulo As String
Dim index As Integer
Dim pwm_value As Integer

Private Sub Form_Load()
    
    Titulo = App.Title & "   " & "v" & App.Major & "." & App.Minor & "." & App.Revision
    Me.Caption = Titulo

On Error GoTo Erro
        
    ' Adiciona lista de baudrate
    cboBaudRate.AddItem "1200"
    cboBaudRate.AddItem "2400"
    cboBaudRate.AddItem "4800"
    cboBaudRate.AddItem "9600"
    cboBaudRate.AddItem "19200"
    cboBaudRate.AddItem "38400"
    cboBaudRate.AddItem "57600"
    cboBaudRate.AddItem "115200"
    cboBaudRate.ListIndex = 3
    
    ' Busca portas disponiveis
    Call cmdScanPort_Click
    
    ' Informação de funções a serem programadas no arduino
    lstInput.ToolTipText = "Selecione um item para verificar a função á ser programada, e alternar o valor a ser enviado para o arduino."
    lstOutput.ToolTipText = "Selecione um item para verificar a função á ser programada, e receber o valor atual do arduino."
    lstAnalog.ToolTipText = "Selecione um item para verificar a função á ser programada, e receber o valor atual do arduino."
    lstPwm.ToolTipText = "Selecione um item para verificar a função á ser programada, e click em + ou - para alterar o valor a ser enviado para o arduino."
    txtInformacao.ToolTipText = "Exemplo de função a ser programada no arduino, click no botão ""Copiar"" para área de transferência."
    
    ' Comandos inicialmente desabilitados
    
    cmdConnectPort.Enabled = True
    frmTerminal.Visible = False
    frmInterface.Visible = True
    frmTerminal.Enabled = False
    frmInterface.Enabled = False
    frmInput.Enabled = True
    frmOutput.Enabled = True
    frmAnalog.Enabled = True
    frmPwm.Enabled = True
    
    ' Ajustes inciais
    shpConnect.BackColor = vbRed
    Call cmdReset_Click
    
    ' Inicia com reset de valores
    Call cmdReset_Click
    
    ' Habilita o formulário para capturar eventos de teclado
    Me.KeyPreview = True
    
Exit Sub

Erro:
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

Private Sub cmdScanPort_Click()
    Dim i As Integer
    
On Error GoTo Erro

    'Detecta portas disponiveis
    cmdScanPort.Caption = "Scan..."
    cboCommPort.Clear
    For i = 1 To 32
        If DetectaPortaCOM(i) <> 0 Then
            cboCommPort.AddItem "COM" & i
        End If
    Next
    cboCommPort.ListIndex = 0
    cmdScanPort.Caption = "Scanear"
        
Exit Sub
    
Erro:
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"

End Sub

Private Sub cmdConnectPort_Click()
On Error GoTo Erro

    ' Conectar
    If cmdConnectPort.Caption = "Conectar" Then
        cmdConnectPort.Caption = "Desconectar"
        MSComm1.CommPort = Mid(cboCommPort.Text, 4, 2)
        MSComm1.Settings = cboBaudRate.Text & "n,8,1"
        MSComm1.RThreshold = 1
        MSComm1.DTREnable = True
        MSComm1.RTSEnable = True
        MSComm1.PortOpen = True
        frmInterface.Enabled = True
        frmTerminal.Enabled = True
        cboCommPort.Enabled = False
        cboBaudRate.Enabled = False
        cmdScanPort.Enabled = False
        shpConnect.BackColor = vbGreen
        Me.Caption = "Conectado na COM" & MSComm1.CommPort & "," & MSComm1.Settings
    ' Desconectar
    Else
        cmdConnectPort.Caption = "Conectar"
        MSComm1.PortOpen = False
        frmInterface.Enabled = False
        frmTerminal.Enabled = False
        cboCommPort.Enabled = True
        cboBaudRate.Enabled = True
        cmdScanPort.Enabled = True
        shpConnect.BackColor = vbRed
        Call cmdReset_Click
        Me.Caption = Titulo
    End If
    
    
Exit Sub

Erro:
    ' Erro relacionados a porta serial
    If Err = 8005 Or Err = 8002 Or Err = 8020 Then
        cmdConnectPort.Caption = "Conectar"
        cboCommPort.Enabled = True
        cboBaudRate.Enabled = True
        cmdScanPort.Enabled = True
        frmTerminal.Enabled = False
        frmInterface.Enabled = False
        shpConnect.BackColor = vbRed
        Me.Caption = Titulo
    End If
    
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

Private Sub lstInput_Click()
    ' Obtém o item selecionado
    Dim selectedItem As String
    selectedItem = lstInput.List(lstInput.ListIndex)
    If selectedItem = Empty Then Exit Sub
    
    ' Verifica o último caractere do item
    Dim state As String
    Dim updatedItem As String
    If Right(selectedItem, 1) = "0" Then
        state = "HIGH"
        updatedItem = Left(selectedItem, Len(selectedItem) - 1) & "1"
    Else
        state = "LOW"
        updatedItem = Left(selectedItem, Len(selectedItem) - 1) & "0"
    End If
    
    ' Atualiza o item na ListBox
    lstInput.List(lstInput.ListIndex) = updatedItem
    
    ' Obtém o item selecionado atualizado
    selectedItem = lstInput.List(lstInput.ListIndex)
    
    ' Atualiza o comando a ser enviado via serial
    Dim pin As Integer
    pin = CInt(Mid(selectedItem, 1, 2))
    
    ' Atualiza informação de função para o arduino
    txtInformacao.Text = "if(Serial.available()){" & vbCrLf & _
                         "  String receivedData = Serial.readStringUntil('\n');" & vbCrLf & _
                         "  receivedData.trim();" & vbCrLf & _
                         "  if (receivedData == """ & selectedItem & """) {" & vbCrLf & _
                         "    digitalWrite(" & pin & "," & state & ");" & vbCrLf & _
                         "  }" & vbCrLf & _
                         "}"
                         
    ' Envia comando via serial para o arduino
    MSComm1.Output = selectedItem & vbLf
    
    ' Limpara a seleção dos demais listbox
    ClearOtherListBoxesSelection "lstInput"
    
End Sub

Private Sub lstOutput_Click()
    ' Obtém o item selecionado
    Dim selectedItem As String
    selectedItem = lstOutput.List(lstOutput.ListIndex)
    If selectedItem = Empty Then Exit Sub

    ' Atualiza informação de função para o arduino
    txtInformacao.Text = "Serial.println(""" & Left(selectedItem, 3) & "0""); // 0 ou 1"
    
    ' Limpara a seleção dos demais listbox
    ClearOtherListBoxesSelection "lstOutput"
    
End Sub

Private Sub lstAnalog_Click()
    ' Obtém o item selecionado
    Dim selectedItem As String
    selectedItem = lstAnalog.List(lstAnalog.ListIndex)
    If selectedItem = Empty Then Exit Sub
    
    Dim pin As String
    pin = Mid(selectedItem, 1, 2)

    ' Atualiza informação de função para o arduino
    txtInformacao.Text = "int value_analog = analogRead(" & pin & ");" & vbCrLf & _
                         "Serial.println(""" & pin & ":"" + String(value_analog);"
                         
    ' Limpara a seleção dos demais listbox
    ClearOtherListBoxesSelection "lstAnalog"

End Sub

Private Sub lstPwm_Click()
    ' Obtém o item selecionado
    Dim selectedItem As String
    selectedItem = lstPwm.List(lstPwm.ListIndex)
    If selectedItem = Empty Then Exit Sub
    
    ' Verifica o último caractere do item selecionado
    Dim updatedItem As String
    pwm_value = Mid(selectedItem, 4, Len(selectedItem))
    
    ' Atualiza informação de função para o arduino
    Dim pin As Integer
    pin = CInt(Mid(selectedItem, 1, 2))
    
    txtInformacao.Text = "if (Serial.available()) {" & vbCrLf & _
                         "  String receivedData = Serial.readStringUntil('\n');" & vbCrLf & _
                         "  receivedData.trim();" & vbCrLf & _
                         "  int index = receivedData.indexOf(':');" & vbCrLf & _
                         "  String pinStr = receivedData.substring(0, index);" & vbCrLf & _
                         "  int pin = pinStr.toInt();" & vbCrLf & _
                         "  String pwmStr = receivedData.substring(index + 1);" & vbCrLf & _
                         "  int pwm_value = pwmStr.toInt();" & vbCrLf & _
                         "  analogWrite(pin, pwm_value);" & vbCrLf & _
                         "}"
                         
    ' Envia comando via serial para o arduino
    MSComm1.Output = selectedItem & vbLf
    
    ' Limpara a seleção dos demais listbox
    ClearOtherListBoxesSelection "lstPwm"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' Verifica se o ListBox PWM está com o foco
    If Me.ActiveControl Is lstPwm Then
        ' Verifica se a tecla "+" foi pressionada
        If KeyAscii = Asc("+") Then
            pwm_value = pwm_value + 1
            If pwm_value >= 255 Then pwm_value = 255
            'Call lstPwm_Click
        End If

        ' Verifica se a tecla "-" foi pressionada
        If KeyAscii = Asc("-") Then
            pwm_value = pwm_value - 1
            If pwm_value <= 0 Then pwm_value = 0
            'Call lstPwm_Click
        End If
    
        ' Obtém o item selecionado
        Dim selectedItem As String
        selectedItem = lstPwm.List(lstPwm.ListIndex)
        
        ' Verifica o último caractere do item selecionado
        Dim updatedItem As String
        updatedItem = Mid(selectedItem, 1, 3) & pwm_value
        
        ' Atualiza o item na ListBox
        lstPwm.List(lstPwm.ListIndex) = updatedItem
        
        ' Envia comando via serial para o arduino
        MSComm1.Output = Mid(selectedItem, 1, 3) & pwm_value & vbLf
    End If
    
End Sub

Private Sub ClearOtherListBoxesSelection(CurrentListBox As String)
    Select Case CurrentListBox
        Case "lstInput"
            lstOutput.ListIndex = -1
            lstAnalog.ListIndex = -1
            lstPwm.ListIndex = -1
        Case "lstOutput"
            lstInput.ListIndex = -1
            lstAnalog.ListIndex = -1
            lstPwm.ListIndex = -1
        Case "lstAnalog"
            lstInput.ListIndex = -1
            lstOutput.ListIndex = -1
            lstPwm.ListIndex = -1
        Case "lstPwm"
            lstInput.ListIndex = -1
            lstOutput.ListIndex = -1
            lstAnalog.ListIndex = -1
    End Select
End Sub

Private Sub MSComm1_OnComm()
    Dim data As String
    
On Error GoTo Erro

    If MSComm1.PortOpen = False Then Exit Sub

    Select Case MSComm1.CommEvent
        Case comEvReceive
            ' Recebe os dados da serial
            data = MSComm1.Input
            
            ' Atualiza valores de Output e Analog
            If Mid(data, 3, 1) = ":" Then
                If Mid(data, 1, 1) = "A" Then
                    Call updateAnalog(data) ' Analog
                Else
                    Call updateOutput(data) ' Input
                End If
            End If
            
            ' Atualiza terminal serial
            With txtTerminal
                .SelStart = Len(txtTerminal.Text)
                .SelText = data
                .SelStart = Len(txtTerminal.Text)
            End With
            
    End Select
    
Exit Sub

Erro:
    If Err = 13 Then Exit Sub
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"

End Sub

Private Sub cmdSendData_Click()
    If MSComm1.PortOpen = False Then Exit Sub
    If cboSendData.Text = "" Then Exit Sub
    
    ' Envia dado pela serial
    MSComm1.Output = cboSendData.Text
    
    'atualiza cboSend
    cboSendData.AddItem cboSendData.Text
    deleteDuplicados
    cboSendData.Text = Empty
    
End Sub

Private Sub updateOutput(data As String)
    Dim searchPrefix As String
    Dim i As Integer
    Dim foundIndex As Integer

    ' Extrai a parte antes do ":" para usar como prefixo de busca
    searchPrefix = Left(data, InStr(data, ":"))

    ' Define um índice inicial para indicar que nenhum item foi encontrado
    foundIndex = -1

    ' Procura no ListBox pelo item que começa com o prefixo
    For i = 0 To lstOutput.ListCount - 1
        If Left(lstOutput.List(i), Len(searchPrefix)) = searchPrefix Then
            ' Encontrou o item correspondente
            foundIndex = i
            Exit For
        End If
    Next i

    ' Verifica se o item foi encontrado
    If foundIndex <> -1 Then
        ' Atualiza o item no ListBox com o valor recebido
        lstOutput.List(foundIndex) = data
    End If

End Sub

Private Sub updateAnalog(data As String)
Dim searchPrefix As String
    Dim i As Integer
    Dim foundIndex As Integer

    ' Extrai a parte antes do ":" para usar como prefixo de busca
    searchPrefix = Left(data, InStr(data, ":"))

    ' Define um índice inicial para indicar que nenhum item foi encontrado
    foundIndex = -1

    ' Procura no ListBox pelo item que começa com o prefixo
    For i = 0 To lstAnalog.ListCount - 1
        If Left(lstAnalog.List(i), Len(searchPrefix)) = searchPrefix Then
            ' Encontrou o item correspondente
            foundIndex = i
            Exit For
        End If
    Next i

    ' Verifica se o item foi encontrado
    If foundIndex <> -1 Then
        ' Atualiza o item no ListBox com o valor recebido
        lstAnalog.List(foundIndex) = data
    End If

End Sub

Private Sub deleteDuplicados()
    Dim i As Integer, j As Integer
    
    For i = 0 To cboSendData.ListCount
        For j = i + 1 To cboSendData.ListCount
            If cboSendData.List(i) = cboSendData.List(j) Then
                cboSendData.RemoveItem (j)
                j = j - 1
            End If
        Next
    Next
    
End Sub

Private Sub clearValue()
    shpInput.BackColor = vbBlack

End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear ' primeiro devo limpar a área de transferência antes de efetuar uma cópia
    Clipboard.SetText txtInformacao.Text ' Copiar o conteúdo do controle TextBox chamado Text1 para a área de transferência
    cmdCopy.Caption = "Copiando..."
    DoEvents
    Sleep (1000)
    cmdCopy.Caption = "Copiar"
    
End Sub

Private Sub cmdReset_Click()
    ' Reset hardware via software
    MSComm1.DTREnable = False
    MSComm1.DTREnable = True
    
    ' Limpa todos os listbox
    lstInput.Clear
    lstOutput.Clear
    lstAnalog.Clear
    lstPwm.Clear
    txtInformacao.Text = Empty
    
    ' Lista de Input
    For i = 2 To 12
        If i < 10 Then
            lstInput.AddItem "0" & i & ":0"
        Else
            lstInput.AddItem i & ":0"
        End If
    Next i
    
    ' Lista de Output
    For i = 2 To 13
        If i < 10 Then
            lstOutput.AddItem "0" & i & ":0"
        Else
            lstOutput.AddItem i & ":0"
        End If
    Next i
    
    ' Lista de Analog
    For i = 0 To 5
        lstAnalog.AddItem "A" & i & ":0"
    Next i
    
    ' Lista Pwm
    lstPwm.AddItem "03:0"
    lstPwm.AddItem "05:0"
    lstPwm.AddItem "06:0"
    lstPwm.AddItem "07:0"
    lstPwm.AddItem "10:0"
    lstPwm.AddItem "11:0"
    
End Sub

Private Sub cmdClearTerminal_Click()
    txtTerminal.Text = Empty
    
End Sub

Private Sub mTerminal_Click()
    If frmInterface.Visible = True Then
        frmInterface.Visible = False
        frmTerminal.Visible = True
        mTerminal.Caption = "Interface"
    Else
        frmTerminal.Visible = False
        frmInterface.Visible = True
        mTerminal.Caption = "Terminal"
    End If
    
End Sub

Private Sub mGerenciador_Click()
    Shell ("cmd.exe /c devmgmt.msc")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Erro
    
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    
Exit Sub

Erro:
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'------------------------------------------------------------------------------------------
'Mid: Retorna o número especificado de caracteres de uma string.
'exemplo: mid(text1.text,1,5) -> retorna as letras 1,2,3,4,5 do text1.
'exemplo: mid(text1.text,20,5) -> retorna  as ultimas 5 letras iniciando da posicai 20 do text1.

'Left:Retorna o número especificado de caracteres a partir do início de uma string.
'exemplo: left(text1.text,3) -> retorna as 3 primeiras letras do text1.

'right:Retorna o número especificado de caracteres a partir do lado direito de uma string.
'exemplo: right(text1.text, 4) -> retorna as quatro últimas letras do text1.


' Função para verificar tempo de processo
'------------------------------------------------------------------------------------------
' Start tempo de processo
'Dim startTime As Double
'Dim endTime As Double
'Dim elapsedTime As Double
'startTime = Timer

' Loop de processo aqui...

' End tempo de processo
'endTime = Timer
'elapsedTime = endTime - startTime
'If elapsedTime > 2 Then txtData = Empty ' limpa txtData, pois houve algum erro.

