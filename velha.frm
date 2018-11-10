VERSION 5.00
Begin VB.Form frmJogoDaVelha 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   Jogo da Velha"
   ClientHeight    =   6810
   ClientLeft      =   4005
   ClientTop       =   2580
   ClientWidth     =   6870
   ForeColor       =   &H00404040&
   Icon            =   "velha.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6870
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   1455
      Left            =   960
      TabIndex        =   10
      Top             =   5160
      Width           =   5055
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   375
         Left            =   2300
         TabIndex        =   16
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "  Placar  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   -60
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Jogador 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Jogador 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label placar1 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Placar2 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   6240
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   5055
      Begin VB.Image image11 
         Height          =   1095
         Left            =   480
         Top             =   480
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   480
         Top             =   2640
         Width           =   4095
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   480
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   3495
         Left            =   3000
         Top             =   480
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   3495
         Left            =   1680
         Top             =   480
         Width           =   255
      End
      Begin VB.Image Image12 
         Height          =   1095
         Left            =   1920
         Top             =   480
         Width           =   1095
      End
      Begin VB.Image Image13 
         Height          =   1095
         Left            =   3240
         Top             =   480
         Width           =   1335
      End
      Begin VB.Image Image23 
         Height          =   855
         Left            =   3240
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Image Image33 
         Height          =   1095
         Left            =   3240
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Image Image32 
         Height          =   1095
         Left            =   1920
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Image Image31 
         Height          =   1095
         Left            =   480
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Image Image22 
         Height          =   855
         Left            =   1920
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Image Image21 
         Height          =   855
         Left            =   480
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   780
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label33 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   3360
         TabIndex        =   9
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label32 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   2040
         TabIndex        =   8
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label31 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   780
         TabIndex        =   7
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   780
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   2040
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   3360
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Menu mnuOpcoes 
      Caption         =   "Opções"
      Begin VB.Menu mnuJogadores 
         Caption         =   "Jogadores"
         Begin VB.Menu mnuHumanoComp 
            Caption         =   "Humano X Computador"
         End
         Begin VB.Menu mnuHumanoHumano 
            Caption         =   "Humano X Humano"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuDificuldade 
         Caption         =   "Dificuldade"
         Begin VB.Menu mnuFacil 
            Caption         =   "Fácil"
         End
         Begin VB.Menu mnuImpossivel 
            Caption         =   "Difícil"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnubarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmJogoDaVelha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim verVitoria(3, 3) As Integer
Dim seqJogada(9) As Integer
Dim rodada As Integer
Dim strSql As String
Public jogador As Integer
Dim rs As New ADODB.Recordset



Private Sub Form_Load()
Me.Caption = Me.Caption & " - Versão " & App.Major & "." & App.Minor
jogador = 0
Call iniciaVerVitoria
mnuHumanoHumano.Checked = False
mnuHumanoComp.Checked = True
mnuImpossivel.Checked = True
Label3.Caption = " Maquina"
End Sub

'Para cada espaço do jogo, existe uma imagem vazia que dependendo da vez de cada
'jogador ao clicar na imagem aparecera um "X" ou uma "O".
Private Sub image11_Click()
If Label11.Caption = "" Then
    If jogador = 0 Then
        Label11.Caption = "X"
        jogador = 1
        verVitoria(1, 1) = 1
        Label11.ForeColor = &H404040
    Else
        Label11.Caption = 0
        jogador = 0
        verVitoria(1, 1) = 0
        Label11.ForeColor = &H404080
    End If
    seqJogada(rodada) = 1 ' Armazena a jogada em um vetor
    rodada = rodada + 1   ' controla a sequencia das jogadas
    Call vencedor         ' verifica se com aquela jogada existe vencedor
End If
End Sub

Private Sub image12_Click()

If Label12.Caption = "" Then
    If jogador = 0 Then
        Label12.Caption = "X"
        jogador = 1
        Label12.ForeColor = &H404040
        verVitoria(1, 2) = 1
    Else
        Label12.Caption = 0
        jogador = 0
        Label12.ForeColor = &H404080
        verVitoria(1, 2) = 0
    End If
    seqJogada(rodada) = 2
    rodada = rodada + 1
    Call vencedor
End If
End Sub

Private Sub image13_Click()

If Label13.Caption = "" Then
    If jogador = 0 Then
        Label13.Caption = "X"
        jogador = 1
        verVitoria(1, 3) = 1
        Label13.ForeColor = &H404040
    Else
        Label13.Caption = 0
        jogador = 0
        verVitoria(1, 3) = 0
        Label13.ForeColor = &H404080
    End If
    seqJogada(rodada) = 3
    rodada = rodada + 1
    Call vencedor
End If
End Sub

Private Sub image21_Click()

If Label21.Caption = "" Then
    If jogador = 0 Then
        Label21.Caption = "X"
        jogador = 1
        verVitoria(2, 1) = 1
        Label21.ForeColor = &H404040
    Else
        Label21.Caption = 0
        jogador = 0
        verVitoria(2, 1) = 0
        Label21.ForeColor = &H404080
    End If
    seqJogada(rodada) = 4
    rodada = rodada + 1
    Call vencedor
End If
End Sub


Private Sub image22_Click()

If Label22.Caption = "" Then
    If jogador = 0 Then
        Label22.Caption = "X"
        jogador = 1
        verVitoria(2, 2) = 1
        Label22.ForeColor = &H404040
    Else
        Label22.Caption = 0
        jogador = 0
        verVitoria(2, 2) = 0
        Label22.ForeColor = &H404080
    End If
    seqJogada(rodada) = 5
    rodada = rodada + 1
    Call vencedor
End If
End Sub


Private Sub image23_Click()

If Label23.Caption = "" Then
    If jogador = 0 Then
        Label23.Caption = "X"
        jogador = 1
        verVitoria(2, 3) = 1
        Label23.ForeColor = &H404040
    Else
        Label23.Caption = 0
        jogador = 0
        verVitoria(2, 3) = 0
        Label23.ForeColor = &H404080
    End If
    seqJogada(rodada) = 6
    rodada = rodada + 1
    Call vencedor
End If
End Sub


Private Sub image31_Click()

If Label31.Caption = "" Then
    If jogador = 0 Then
        Label31.Caption = "X"
        jogador = 1
        verVitoria(3, 1) = 1
        Label31.ForeColor = &H404040
    Else
        Label31.Caption = 0
        jogador = 0
        verVitoria(3, 1) = 0
        Label31.ForeColor = &H404080
    End If
    seqJogada(rodada) = 7
    rodada = rodada + 1
    Call vencedor
End If
End Sub


Private Sub image32_Click()

If Label32.Caption = "" Then
    If jogador = 0 Then
        Label32.Caption = "X"
        jogador = 1
        verVitoria(3, 2) = 1
        Label32.ForeColor = &H404040
    Else
        Label32.Caption = 0
        jogador = 0
        verVitoria(3, 2) = 0
        Label32.ForeColor = &H404080
    End If
    seqJogada(rodada) = 8
    rodada = rodada + 1
    Call vencedor
End If
End Sub


Private Sub image33_Click()

If Label33.Caption = "" Then
    If jogador = 0 Then
        Label33.Caption = "X"
        jogador = 1
        verVitoria(3, 3) = 1
        Label33.ForeColor = &H404040
    Else
        Label33.Caption = 0
        jogador = 0
        verVitoria(3, 3) = 0
        Label33.ForeColor = &H404080
    End If
    seqJogada(rodada) = 9
    rodada = rodada + 1
    Call vencedor
End If
End Sub

Private Sub iniciaVerVitoria()
    Dim X, Y As Integer
    For X = 1 To 3
        For Y = 1 To 3
            verVitoria(X, Y) = 3
        Next Y
    Next X
End Sub

'Função que compara na matriz "verVitoria", se existe uma combinação que leve
'algum jogador a vitoria. Caso ocorra a vitória o jogo é paralizado e o placar
'inclementado para o vitorioso

Private Sub vencedor()
    If (verVitoria(1, 1) = verVitoria(2, 2) And verVitoria(1, 1) = verVitoria(3, 3) And _
    verVitoria(1, 1) <> 3) Then
        If jogador = 0 Then
            Call MsgBox("Vitoria do Jogador 2!", , "Vitória!")
            Placar2 = Placar2.Caption + 1
        ElseIf jogador = 1 Then
            Call MsgBox("Vitoria do Jogador 1!", , "Vitória!")
            placar1 = placar1.Caption + 1
        End If
        Call iniciaTabuleiro
        Call iniciaVerVitoria
    End If
    
    If (verVitoria(1, 1) = verVitoria(2, 1) And verVitoria(1, 1) = verVitoria(3, 1) And _
    verVitoria(1, 1) <> 3) Then
        If jogador = 0 Then
            Call MsgBox("Vitoria do Jogador 2!", , "Vitória!")
            Placar2 = Placar2.Caption + 1
        ElseIf jogador = 1 Then
            Call MsgBox("Vitoria do Jogador 1!", , "Vitória!")
            placar1 = placar1.Caption + 1
        End If
        Call iniciaTabuleiro
        Call iniciaVerVitoria
    End If
    
    
    If (verVitoria(1, 1) = verVitoria(1, 2) And verVitoria(1, 1) = verVitoria(1, 3) And _
    verVitoria(1, 1) <> 3) Then
        If jogador = 0 Then
            Call MsgBox("Vitoria do Jogador 2!", , "Vitória!")
            Placar2 = Placar2.Caption + 1
        ElseIf jogador = 1 Then
            Call MsgBox("Vitoria do Jogador 1!", , "Vitória!")
            placar1 = placar1.Caption + 1
        End If
        Call iniciaTabuleiro
        Call iniciaVerVitoria
    End If


    If (verVitoria(1, 3) = verVitoria(2, 2) And verVitoria(1, 3) = verVitoria(3, 1) And _
    verVitoria(1, 3) <> 3) Then
        If jogador = 0 Then
            Call MsgBox("Vitoria do Jogador 2!", , "Vitória!")
            Placar2 = Placar2.Caption + 1
        ElseIf jogador = 1 Then
            Call MsgBox("Vitoria do Jogador 1!", , "Vitória!")
            placar1 = placar1.Caption + 1
        End If
        Call iniciaTabuleiro
        Call iniciaVerVitoria
    End If
    
    
    If (verVitoria(1, 3) = verVitoria(2, 3) And verVitoria(1, 3) = verVitoria(3, 3) And _
    verVitoria(1, 3) <> 3) Then
        If jogador = 0 Then
            Call MsgBox("Vitoria do Jogador 2!", , "Vitória!")
            Placar2 = Placar2.Caption + 1
        ElseIf jogador = 1 Then
            Call MsgBox("Vitoria do Jogador 1!", , "Vitória!")
            placar1 = placar1.Caption + 1
        End If
        Call iniciaTabuleiro
        Call iniciaVerVitoria
    End If
    
    If (verVitoria(3, 1) = verVitoria(3, 2) And verVitoria(3, 1) = verVitoria(3, 3) And _
    verVitoria(3, 1) <> 3) Then
        If jogador = 0 Then
            Call MsgBox("Vitoria do Jogador 2!", , "Vitória!")
            Placar2 = Placar2.Caption + 1
        ElseIf jogador = 1 Then
            Call MsgBox("Vitoria do Jogador 1!", , "Vitória!")
            placar1 = placar1.Caption + 1
        End If
        Call iniciaTabuleiro
        Call iniciaVerVitoria
    End If
    
    If (verVitoria(1, 2) = verVitoria(2, 2) And verVitoria(1, 2) = verVitoria(3, 2) And _
    verVitoria(1, 2) <> 3) Then
        If jogador = 0 Then
            Call MsgBox("Vitoria do Jogador 2!", , "Vitória!")
            Placar2 = Placar2.Caption + 1
        ElseIf jogador = 1 Then
            Call MsgBox("Vitoria do Jogador 1!", , "Vitória!")
            placar1 = placar1.Caption + 1
        End If
        Call iniciaTabuleiro
        Call iniciaVerVitoria
    End If
    
    If (verVitoria(2, 1) = verVitoria(2, 2) And verVitoria(2, 1) = verVitoria(2, 3) And _
    verVitoria(2, 1) <> 3) Then
        If jogador = 0 Then
            Call MsgBox("Vitoria do Jogador 2!", , "Vitória!")
            Placar2 = Placar2.Caption + 1
        ElseIf jogador = 1 Then
            Call MsgBox("Vitoria do Jogador 1!", , "Vitória!")
            placar1 = placar1.Caption + 1
        End If
        Call iniciaTabuleiro
        Call iniciaVerVitoria
    End If
    
    Call verificaJogador
End Sub

Private Sub iniciaTabuleiro()
    Dim i As Integer
    If mnuHumanoComp.Checked = True Then Call armazenaJogada
    rodada = 0
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = ""
    Label21.Caption = ""
    Label22.Caption = ""
    Label23.Caption = ""
    Label31.Caption = ""
    Label32.Caption = ""
    Label33.Caption = ""
    For i = 0 To 9
        seqJogada(i) = 0
    Next i
End Sub

'Função que seleciona uma jogada qualquer caso sejá selecionado o nivel facil, ou não seja encontrada
'uma jogada na função pensaJogada.
Private Sub escolheJogada()
Dim pensada As Integer
Dim X, aux As Integer
If mnuFacil.Checked = False Then
    pensada = pensaJogada()
End If
    If pensada <> 0 Then
    X = pensada
    If X < 0 Then
    
            If Label11.Caption <> "" Then aux = aux + 1
            If Label12.Caption <> "" Then aux = aux + 1
            If Label13.Caption <> "" Then aux = aux + 1
            If Label22.Caption <> "" Then aux = aux + 1
            If Label31.Caption <> "" Then aux = aux + 1
            If Label33.Caption <> "" Then aux = aux + 1
    
        While X = pensada Or X < 0
            X = Int((9 * Rnd) + 1)
        Wend
    End If
Else

tentanovamente:

    X = Int((9 * Rnd) + 1)
            If Label11.Caption <> "" And Label12.Caption <> "" _
            And Label13.Caption <> "" And Label21.Caption <> "" _
            And Label22.Caption <> "" And Label23.Caption <> "" _
            And Label31.Caption <> "" And Label32.Caption <> "" _
            And Label33.Caption <> "" Then Exit Sub
End If
            
    Select Case X
        Case 1
            If Label11.Caption <> "" Then GoTo tentanovamente
            Call image11_Click
        Case 2
            If Label12.Caption <> "" Then GoTo tentanovamente
            Call image12_Click
        Case 3
            If Label13.Caption <> "" Then GoTo tentanovamente
            Call image13_Click
        Case 4
            If Label21.Caption <> "" Then GoTo tentanovamente
            Call image21_Click
        Case 5
            If Label22.Caption <> "" Then GoTo tentanovamente
            Call image22_Click
        Case 6
            If Label23.Caption <> "" Then GoTo tentanovamente
            Call image23_Click
        Case 7
            If Label31.Caption <> "" Then GoTo tentanovamente
            Call image31_Click
        Case 8
            If Label32.Caption <> "" Then GoTo tentanovamente
            Call image32_Click
        Case 9
            If Label33.Caption <> "" Then GoTo tentanovamente
            Call image33_Click
    End Select
        Timer1.Interval = 1000
End Sub

'Função que verifica se ocorreu empate. Verifica também se é o computador
'que deve fazer a jogada.
Private Sub verificaJogador()
    If mnuHumanoComp.Checked = True Then
        If jogador = 1 Then
            Call escolheJogada
        End If
    End If
    If Label11.Caption <> "" And Label12.Caption <> "" _
    And Label13.Caption <> "" And Label21.Caption <> "" _
    And Label22.Caption <> "" And Label23.Caption <> "" _
    And Label31.Caption <> "" And Label32.Caption <> "" _
    And Label33.Caption <> "" Then
        Call MsgBox("Deu Velha!", , "Empate!")
        rodada = 10
        Call iniciaTabuleiro
        Call iniciaVerVitoria
        Call verificaJogador
    End If
End Sub

Private Sub mnuFacil_Click()
    mnuImpossivel.Checked = False
    mnuFacil.Checked = True
End Sub

Private Sub mnuHumanoComp_Click()
    mnuHumanoHumano.Checked = False
    mnuHumanoComp.Checked = True
    Label3.Caption = " Maquina"
End Sub

Private Sub mnuHumanoHumano_Click()
    mnuHumanoHumano.Checked = True
    mnuHumanoComp.Checked = False
    Label3.Caption = "Jogador 2"
End Sub

'Função para armazenar a Jogada, caso não exista na base de dados
Private Function armazenaJogada()

    Set rs = New ADODB.Recordset
    strSql = "SELECT * FROM "
    If rodada = 10 Then
        strSql = strSql & "Empate "
    ElseIf rodada Mod 2 <> 0 Then
        strSql = strSql & "vitoriasP1 "
    Else
        strSql = strSql & "vitoriasP2 "
    End If
    
    If seqJogada(0) <> 0 Then strSql = strSql & " WHERE J1 LIKE '" & seqJogada(0) & "'"
    If seqJogada(1) <> 0 Then strSql = strSql & " AND J2 LIKE '" & seqJogada(1) & "'"
    If seqJogada(2) <> 0 Then strSql = strSql & " AND J3 LIKE '" & seqJogada(2) & "'"
    If seqJogada(3) <> 0 Then strSql = strSql & " AND J4 LIKE '" & seqJogada(3) & "'"
    If seqJogada(4) <> 0 Then strSql = strSql & " AND J5 LIKE '" & seqJogada(4) & "'"
    If seqJogada(5) <> 0 Then strSql = strSql & " AND J6 LIKE '" & seqJogada(5) & "'"
    If seqJogada(6) <> 0 Then strSql = strSql & " AND J7 LIKE '" & seqJogada(6) & "'"
    If seqJogada(7) <> 0 Then strSql = strSql & " AND J8 LIKE '" & seqJogada(7) & "'"
    If seqJogada(8) <> 0 Then strSql = strSql & " AND J9 LIKE '" & seqJogada(8) & "'"
    
    rs.Open strSql, conexao, adOpenStatic, adLockOptimistic
    If rs.EOF = False Then
        If rodada <> 10 Then rs.Update
        If rodada <> 10 Then rs!Vitoria = rs!Vitoria + 1
        If rodada <> 10 Then rs.Update
        Exit Function
    End If
    If rodada = 10 Then
        strSql = "INSERT INTO Empate ("
    ElseIf rodada Mod 2 <> 0 Then
        strSql = "INSERT INTO vitoriasP1 ("
    Else
        strSql = "INSERT INTO vitoriasP2 ("
    End If
    
    If seqJogada(0) <> 0 Then strSql = strSql & "J1"
    If seqJogada(1) <> 0 Then strSql = strSql & ",J2"
    If seqJogada(2) <> 0 Then strSql = strSql & ",J3"
    If seqJogada(3) <> 0 Then strSql = strSql & ",J4"
    If seqJogada(4) <> 0 Then strSql = strSql & ",J5"
    If seqJogada(5) <> 0 Then strSql = strSql & ",J6"
    If seqJogada(6) <> 0 Then strSql = strSql & ",J7"
    If seqJogada(7) <> 0 Then strSql = strSql & ",J8"
    If seqJogada(8) <> 0 Then strSql = strSql & ",J9"
    strSql = strSql & ")"
    strSql = strSql & " VALUES ("
    If seqJogada(0) <> 0 Then strSql = strSql & seqJogada(0)
    If seqJogada(1) <> 0 Then strSql = strSql & " ," & seqJogada(1)
    If seqJogada(2) <> 0 Then strSql = strSql & " ," & seqJogada(2)
    If seqJogada(3) <> 0 Then strSql = strSql & " ," & seqJogada(3)
    If seqJogada(4) <> 0 Then strSql = strSql & " ," & seqJogada(4)
    If seqJogada(5) <> 0 Then strSql = strSql & " ," & seqJogada(5)
    If seqJogada(6) <> 0 Then strSql = strSql & " ," & seqJogada(6)
    If seqJogada(7) <> 0 Then strSql = strSql & " ," & seqJogada(7)
    If seqJogada(8) <> 0 Then strSql = strSql & " ," & seqJogada(8)
    strSql = strSql & ")"
    
    conexao.Execute (strSql)

End Function

'Função que analisa a melhor jogada à ser feita pelo computador.
Private Function pensaJogada() As Integer

Dim jogadasP1, jogadasP2 As String
Dim pegaJogada As Integer
Dim rsProximo As New ADODB.Recordset
Dim z, obrigatorio As Integer

Set rsProximo = New ADODB.Recordset
Set rs = New ADODB.Recordset

Select Case rodada
  ' Case 0
            'Se o computador ser o primeiro jogador, esta primeira jogada será selecionada
            'randomicamente.
    Case 1
            'Caso o jogador "humano" opte por jogar em uma das diagonais, o sistema dará a
            'resposta no centro, sendo que é a única maneira possível de defesa.
        If verVitoria(1, 1) Or verVitoria(1, 3) Or verVitoria(3, 1) Or verVitoria(3, 3) Then
            pensaJogada = 5
        ElseIf verVitoria(2, 2) Then
            X = Int((4 * Rnd) + 1)
            Select Case X
                Case 1
                    pensaJogada = 1
                Case 2
                    pensaJogada = 3
                Case 3
                    pensaJogada = 7
                Case 4
                    pensaJogada = 9
            End Select
        Else
            'Caso contrário será pesquisado na base de conhecimentos a jogada que lhe da mais vantagens.
            jogadasP2 = "SELECT TOP 1 J2, Count(J2) AS JOGADA From EMPATE WHERE J2<> " & seqJogada(0)
            jogadasP2 = jogadasP2 & " AND J1 =" & seqJogada(0) & " GROUP BY J2"
            rs.Open jogadasP2, conexao
            If rs.EOF = False Then pensaJogada = rs!J2
        End If
        Exit Function
    Case 2
    
        jogadasP1 = "SELECT top 1 j3, count(j3) as vitorias from vitoriasP1  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0) & ""
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 <> " & seqJogada(2) & " group by j3 order by count(j3)desc ;"
        rs.Open jogadasP1, conexao
        If rs.EOF = False Then pensaJogada = rs!J3
        Exit Function
        
    Case 3
        'A partir desta etapa o sistema como segundo jogador, vai procurar primeiramente a jogada
        'Obrigatória que seria para evitar a vitória do oponente, ou derrotar o oponente no caso de
        'um descuido do adversário.
        obrigatorio = preveProximoLance()
        If obrigatorio = 0 Then
        'Caso não seja encontrado nem um lance obrigatório, o sistema vai procurar na base de
        'conhecimentos o caso q vai lhe dar maiores chances de um empate. Será visto também
        'se aquela sequencia não podera leva-lo a uma derrota, caso isso ocorra o sistema vai
        'procurar uma outra jogada.
        z = 0
            jogadasP2 = "SELECT j4, count(j4) as JOGADA from EMPATE  "
            jogadasP2 = jogadasP2 & "where j1 = " & seqJogada(0)
            jogadasP2 = jogadasP2 & "and j2 = " & seqJogada(1)
            jogadasP2 = jogadasP2 & " and j3 = " & seqJogada(2)
            jogadasP2 = jogadasP2 & " and j4 <> " & seqJogada(3) & " group by j4 order by count(j4)desc ;"
            Set rsProximo = New ADODB.Recordset
            rsProximo.Open jogadasP2, conexao
z = 0
tentaNovamente3:
z = z + 1
        If z >= 200 Then Exit Function
        If rsProximo.EOF = False Then
            pegaJogada = rsProximo!j4
        ElseIf z = 1 Then
            jogadasP2 = "SELECT j4, count(j4) as JOGADA from EMPATE  "
            jogadasP2 = jogadasP2 & "where j1 = " & seqJogada(0)
            jogadasP2 = jogadasP2 & "and j2 = " & seqJogada(1)
            jogadasP2 = jogadasP2 & " and j3 = " & seqJogada(2)
            jogadasP2 = jogadasP2 & " and j4 <> " & seqJogada(3) & " group by j4 order by count(j4)desc ;"
            Set rsProximo = New ADODB.Recordset
            rsProximo.Open jogadasP2, conexao
            z = 0
            If rsProximo.EOF = False Then
                rsProximo.MoveNext
            End If
            If rsProximo.EOF = False Then
                pensaJogada = rsProximo!j4
            End If
        Else
            pegaJogada = verTabuleiro()
        End If
        
        jogadasP1 = "SELECT * from vitoriasP1  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & pegaJogada
        'jogadasP1 = jogadasP1 & " and j7 = 0"
        
        Set rs = New ADODB.Recordset
        rs.Open jogadasP1, conexao
        If rs.EOF = True Then
            pensaJogada = pegaJogada
        Else
            If rsProximo.EOF = False Then
                rsProximo.MoveNext
            End If
            GoTo tentaNovamente3
        End If
        Else
        pensaJogada = obrigatorio
        End If
        Exit Function
        
    Case 4
        obrigatorio = preveProximoLance()
        'Como jogador um, o computador procura a melhor maneira de atacar, analisando o seu repertório
        'de vitórias, e verificando se essa tática não pode leva-lo a derrota, caso isso ocorra será
        'selecionado outra forma de ataque.
        If obrigatorio = 0 Then
        jogadasP1 = "SELECT j5, count(j5) as vitorias from vitoriasP1  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j6 = 0"
        jogadasP1 = jogadasP1 & " and j5 <> " & seqJogada(4) & " group by j5 order by count(j5)desc ;"
        rsProximo.Open jogadasP1, conexao
    
        If rsProximo.EOF = False Then
           pensaJogada = rsProximo!j5
           Exit Function
        End If
        Set rsProximo = New ADODB.Recordset
        jogadasP1 = "SELECT j5, count(j5) as vitorias from vitoriasP1  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j5 <> " & seqJogada(4) & " group by j5 order by count(j5)desc ;"
        rsProximo.Open jogadasP1, conexao

z = 0
tentaNovamente4:
z = z + 1
        If z >= 200 Then Exit Function
        If rsProximo.EOF = False Then
            pegaJogada = rsProximo!j5
        Else

            pegaJogada = verTabuleiro()

        End If
        
        jogadasP1 = "SELECT * from vitoriasP2  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j7 = 0"
        jogadasP1 = jogadasP1 & " and j5 = " & pegaJogada
        
        Set rs = New ADODB.Recordset
        rs.Open jogadasP1, conexao
        If rs.EOF = False Then
            If rsProximo.EOF = False Then rsProximo.MoveNext
            GoTo tentaNovamente4
        Else
            pensaJogada = pegaJogada
        End If
        
        Else
        pensaJogada = obrigatorio
        End If
        
        Exit Function
        
    
    Case 5
        obrigatorio = preveProximoLance()
        If obrigatorio = 0 Then
        jogadasP1 = "SELECT top 1 j6, count(j6) as vitorias from vitoriasP2  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j5 = " & seqJogada(4)
        jogadasP1 = jogadasP1 & " and j7 = 0"
        jogadasP1 = jogadasP1 & " and j6 <> " & seqJogada(5) & " group by j6 order by count(j6)desc ;"
        
        rsProximo.Open jogadasP1, conexao
        
        If rsProximo.EOF = False Then
            pensaJogada = rsProximo!j6
            Exit Function
        End If
        
        Set rsProximo = New ADODB.Recordset
        
        jogadasP2 = "SELECT j6, count(j6) as JOGADA from EMPATE  "
        jogadasP2 = jogadasP2 & "where j1 = " & seqJogada(0)
        jogadasP2 = jogadasP2 & "and j2 = " & seqJogada(1)
        jogadasP2 = jogadasP2 & " and j3 = " & seqJogada(2)
        jogadasP2 = jogadasP2 & " and j4 = " & seqJogada(3)
        jogadasP2 = jogadasP2 & " and j5 = " & seqJogada(4)
        jogadasP2 = jogadasP2 & " and j6 <> " & seqJogada(5) & " group by j6 order by count(j6)desc ;"
        rsProximo.Open jogadasP2, conexao
z = 0
tentaNovamente5:
z = z + 1
        If z >= 200 Then Exit Function
        If rsProximo.EOF = False Then
            pegaJogada = rsProximo!j6
        Else
            pegaJogada = verTabuleiro()
        End If
        
        jogadasP1 = "SELECT * from vitoriasP1  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j5 = " & seqJogada(4)
        jogadasP1 = jogadasP1 & " and j6 = " & pegaJogada
        jogadasP1 = jogadasP1 & " and j8 = 0"
        
        Set rs = New ADODB.Recordset
        rs.Open jogadasP1, conexao
        If rs.EOF = True Then
            pensaJogada = pegaJogada
        Else
            If rsProximo.EOF = False Then
                rsProximo.MoveNext
            End If
            GoTo tentaNovamente5
        End If
        
        Else
            pensaJogada = obrigatorio
        End If
        
        Exit Function
        

    Case 6
        obrigatorio = preveProximoLance()
        If obrigatorio = 0 Then
        
        jogadasP1 = "SELECT top 1 j7, count(j7) as vitorias from vitoriasP1  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j5 = " & seqJogada(4)
        jogadasP1 = jogadasP1 & " and j6 = " & seqJogada(5)
        jogadasP1 = jogadasP1 & " and j8 = 0"
        jogadasP1 = jogadasP1 & " and j7 <> " & seqJogada(6) & " group by j7 order by count(j7)desc ;"
        
        rsProximo.Open jogadasP1, conexao
        
        If rsProximo.EOF = False Then
            pensaJogada = rsProximo!j7
            Exit Function
        End If
    
        jogadasP1 = "SELECT top 1 j7, count(j7) as vitorias from vitoriasP1  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j5 = " & seqJogada(4)
        jogadasP1 = jogadasP1 & " and j6 = " & seqJogada(5)
        jogadasP1 = jogadasP1 & " and j7 <> " & seqJogada(6) & " group by j7 order by count(j7)desc ;"
        Set rsProximo = New ADODB.Recordset
        rsProximo.Open jogadasP1, conexao
        If rsProximo.EOF = False Then pegaJogada = rsProximo!j7
  z = 0
tentaNovamente6:
  z = z + 1
        If z >= 200 Then Exit Function
        If rsProximo.EOF = False Then
            pegaJogada = rsProximo!j7
        Else
            pegaJogada = verTabuleiro()
        End If
        
        jogadasP1 = "SELECT * from vitoriasP2  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j5 = " & seqJogada(4)
        jogadasP1 = jogadasP1 & " and j6 = " & seqJogada(5)
        jogadasP1 = jogadasP1 & " and j7 = " & pegaJogada
        
        Set rs = New ADODB.Recordset
        rs.Open jogadasP1, conexao
        If rs.EOF = False Then
            If rsProximo.EOF = False Then rsProximo.MoveNext
            GoTo tentaNovamente6
        Else
            pensaJogada = pegaJogada
            Exit Function
        End If
        Else
            pensaJogada = obrigatorio
            Exit Function
        End If

    Case 7
        obrigatorio = preveProximoLance()
        If obrigatorio = 0 Then
        jogadasP1 = "SELECT top 1 j8, count(j8) as vitorias from vitoriasP2  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j5 = " & seqJogada(4)
        jogadasP1 = jogadasP1 & " and j6 = " & seqJogada(5)
        jogadasP1 = jogadasP1 & " and j7 = " & seqJogada(6)
        jogadasP1 = jogadasP1 & " and j9 = 0"
        jogadasP1 = jogadasP1 & " and j8 <> " & seqJogada(7) & " group by j8 order by count(j8)desc ;"
        
        rsProximo.Open jogadasP1, conexao
        
        If rsProximo.EOF = False Then
            pensaJogada = rsProximo!j8
            Exit Function
        End If
        
        Set rsProximo = New ADODB.Recordset
        
        jogadasP2 = "SELECT j8, count(j8) as JOGADA from EMPATE  "
        jogadasP2 = jogadasP2 & "where j1 = " & seqJogada(0)
        jogadasP2 = jogadasP2 & "and j2 = " & seqJogada(1)
        jogadasP2 = jogadasP2 & " and j3 = " & seqJogada(2)
        jogadasP2 = jogadasP2 & " and j4 = " & seqJogada(3)
        jogadasP2 = jogadasP2 & " and j5 = " & seqJogada(4)
        jogadasP2 = jogadasP2 & " and j6 = " & seqJogada(5)
        jogadasP2 = jogadasP2 & " and j7 = " & seqJogada(6)
        jogadasP2 = jogadasP2 & " and j8 <> " & seqJogada(7) & " group by j8 order by count(j8)desc ;"
        rsProximo.Open jogadasP2, conexao

        If rsProximo.EOF = False Then
            pegaJogada = rsProximo!j8
        Else
            pegaJogada = verTabuleiro()
        End If
        pensaJogada = pegaJogada
        Else
            pensaJogada = obrigatorio
            Exit Function
        End If
    
    Case 8
        jogadasP1 = "SELECT top 1 j9, count(j9) as vitorias from vitoriasP1  "
        jogadasP1 = jogadasP1 & "where j1 = " & seqJogada(0)
        jogadasP1 = jogadasP1 & "and j2 = " & seqJogada(1)
        jogadasP1 = jogadasP1 & " and j3 = " & seqJogada(2)
        jogadasP1 = jogadasP1 & " and j4 = " & seqJogada(3)
        jogadasP1 = jogadasP1 & " and j5 = " & seqJogada(4)
        jogadasP1 = jogadasP1 & " and j6 = " & seqJogada(5)
        jogadasP1 = jogadasP1 & " and j7 = " & seqJogada(6)
        jogadasP1 = jogadasP1 & " and j8 = " & seqJogada(7)
        jogadasP1 = jogadasP1 & " and j9 <> " & seqJogada(8) & " group by j9 order by count(j9)desc ;"
        rs.Open jogadasP1, conexao
        If rs.EOF = False Then pensaJogada = rs!j9
        Exit Function
    
End Select
      


End Function

Private Sub mnuImpossivel_Click()
    mnuImpossivel.Checked = True
    mnuFacil.Checked = False
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub
'Caso sejá necessário sortear alguma jogada, essa parte do sistema tem essa função.
Private Function verTabuleiro() As Integer

Dim X As Integer

tentanovamente:
    X = Int((9 * Rnd) + 1)
            If Label11.Caption <> "" And Label12.Caption <> "" _
            And Label13.Caption <> "" And Label21.Caption <> "" _
            And Label22.Caption <> "" And Label23.Caption <> "" _
            And Label31.Caption <> "" And Label32.Caption <> "" _
            And Label33.Caption <> "" Then Exit Function

    Select Case X
        Case 1
            If Label11.Caption <> "" Then GoTo tentanovamente
            verTabuleiro = 1
        Case 2
            If Label12.Caption <> "" Then GoTo tentanovamente
            verTabuleiro = 2
        Case 3
            If Label13.Caption <> "" Then GoTo tentanovamente
            verTabuleiro = 3
        Case 4
            If Label21.Caption <> "" Then GoTo tentanovamente
            verTabuleiro = 4
        Case 5
            If Label22.Caption <> "" Then GoTo tentanovamente
            verTabuleiro = 5
        Case 6
            If Label23.Caption <> "" Then GoTo tentanovamente
            verTabuleiro = 6
        Case 7
            If Label31.Caption <> "" Then GoTo tentanovamente
            verTabuleiro = 7
        Case 8
            If Label32.Caption <> "" Then GoTo tentanovamente
            verTabuleiro = 8
        Case 9
            If Label33.Caption <> "" Then GoTo tentanovamente
            verTabuleiro = 9
    End Select
End Function

'Preve se o proximo lance pode ser vitória de algum dos jogadores.
Public Function preveProximoLance() As Integer
Dim X, Y As Integer
Dim teste As Integer
teste = 0
novoteste:
For X = 1 To 3
    For Y = 1 To 3
        If verVitoria(X, Y) = 3 Then
            If teste = 0 Then
                verVitoria(X, Y) = 0
            Else
                verVitoria(X, Y) = 1
            End If
            If (verVitoria(1, 1) = verVitoria(2, 2) And verVitoria(1, 1) = verVitoria(3, 3) _
            And verVitoria(1, 1) <> 3) Then
                preveProximoLance = verCasa(X & "," & Y)
                verVitoria(X, Y) = 3
                Exit Function
            End If
    
            If (verVitoria(1, 1) = verVitoria(2, 1) And verVitoria(1, 1) = verVitoria(3, 1) _
            And verVitoria(1, 1) <> 3) Then
                preveProximoLance = verCasa(X & "," & Y)
                verVitoria(X, Y) = 3
                Exit Function
            End If
        
            If (verVitoria(1, 1) = verVitoria(1, 2) And verVitoria(1, 1) = verVitoria(1, 3) _
            And verVitoria(1, 1) <> 3) Then
                preveProximoLance = verCasa(X & "," & Y)
                verVitoria(X, Y) = 3
                Exit Function
            End If

            If (verVitoria(1, 3) = verVitoria(2, 2) And verVitoria(1, 3) = verVitoria(3, 1) _
            And verVitoria(1, 3) <> 3) Then
                preveProximoLance = verCasa(X & "," & Y)
                verVitoria(X, Y) = 3
                Exit Function
            End If
      
            If (verVitoria(1, 3) = verVitoria(2, 3) And verVitoria(1, 3) = verVitoria(3, 3) _
            And verVitoria(1, 3) <> 3) Then
                preveProximoLance = verCasa(X & "," & Y)
                verVitoria(X, Y) = 3
                Exit Function
            End If
    
            If (verVitoria(3, 1) = verVitoria(3, 2) And verVitoria(3, 1) = verVitoria(3, 3) _
            And verVitoria(3, 1) <> 3) Then
                preveProximoLance = verCasa(X & "," & Y)
                verVitoria(X, Y) = 3
                Exit Function
            End If
    
            If (verVitoria(1, 2) = verVitoria(2, 2) And verVitoria(1, 2) = verVitoria(3, 2) _
            And verVitoria(1, 2) <> 3) Then
                preveProximoLance = verCasa(X & "," & Y)
                verVitoria(X, Y) = 3
                Exit Function
            End If
    
            If (verVitoria(2, 1) = verVitoria(2, 2) And verVitoria(2, 1) = verVitoria(2, 3) _
            And verVitoria(2, 1) <> 3) Then
                preveProximoLance = verCasa(X & "," & Y)
                verVitoria(X, Y) = 3
                Exit Function
            End If
        verVitoria(X, Y) = 3
        End If
    Next Y
Next X
If teste = 0 Then
    teste = 1
    GoTo novoteste
End If
End Function


Private Function verCasa(ver As String) As Integer
    If ver = "1,1" Then verCasa = 1
    If ver = "1,2" Then verCasa = 2
    If ver = "1,3" Then verCasa = 3
    If ver = "2,1" Then verCasa = 4
    If ver = "2,2" Then verCasa = 5
    If ver = "2,3" Then verCasa = 6
    If ver = "3,1" Then verCasa = 7
    If ver = "3,2" Then verCasa = 8
    If ver = "3,3" Then verCasa = 9
End Function

