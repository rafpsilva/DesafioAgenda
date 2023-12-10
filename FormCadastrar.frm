VERSION 5.00
Begin VB.Form FormCadastrar 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Castrar Novo Contato"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6045
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtVoltar 
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton SalvarNovoContato 
      BackColor       =   &H80000003&
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      MaskColor       =   &H80000006&
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox TextEmail 
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox TextTelefone 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "(  )    -    "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   360
      MaxLength       =   11
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox TextSobrenome 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   2000
   End
   Begin VB.TextBox TextNome 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label LbEmail 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "EMAIL"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label LbTelefone 
      BackColor       =   &H80000002&
      Caption         =   "TELEFONE"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label LbSobrenome 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "SOBRENOME"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label LbNome 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "NOME"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   570
   End
End
Attribute VB_Name = "FormCadastrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoConectarBanco As New ADODB.Connection
Dim ConectaBanco As String
Private Sub BtVoltar_Click()

'Retorna para a tela principal
Unload Me

End Sub


Private Sub SalvarNovoContato_Click()

Dim cmd As New ADODB.Command
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
    
'Verifica se há campos em branco
If TextNome.Text = "" Or TextSobrenome.Text = "" Or TextTelefone.Text = "" Or TextEmail.Text = "" Then
    MsgBox "Um ou mais campos estão  em branco, preencha novamente", vbInformation
Else
    'Verifica se há problemas de conexão com o banco de dados
    On Error GoTo ErrorConexao
    
    ' Inicia a conexão com o banco de dados PostgreSQL
    ConectaBanco = "DSN=PostgreSQL30;UID=postgres;PWD=1234;"
    conn.Open ConectaBanco
    
    If ValidaEmail(TextEmail) Then
        ' Inserção de valores no banco de dados
        cmd.ActiveConnection = conn
        cmd.CommandText = "INSERT INTO Contatos (nome, sobrenome, telefone, email) VALUES (? , ?, ?, ?)"
        cmd.CommandType = adCmdText
        
        cmd.Parameters.Append cmd.CreateParameter("nomeParam", adVarChar, adParamInput, Len(TextNome.Text), TextNome.Text)
        cmd.Parameters.Append cmd.CreateParameter("sobrenomeParam", adVarChar, adParamInput, Len(TextSobrenome.Text), TextSobrenome.Text)
        cmd.Parameters.Append cmd.CreateParameter("telefoneParam", adVarChar, adParamInput, Len(TextTelefone.Text), TextTelefone.Text)
        cmd.Parameters.Append cmd.CreateParameter("emailParam", adVarChar, adParamInput, Len(TextEmail.Text), TextEmail.Text)
        cmd.Execute
        conn.Close
        MsgBox "Contato Salvo", vbInformation
        
        ' Limpar os TextBox após a inserção
        TextNome.Text = ""
        TextSobrenome.Text = ""
        TextTelefone.Text = ""
        TextEmail.Text = ""
    Else
        MsgBox "Email invalido", vbInformation
        Exit Sub
    End If
    
    'Caso exista problemas de comunicação  é realizado uma nova tentativa de conexão
ErrorConexao:
    If Err.Number = "-2147467259" Then
        MsgBox "Não foi possivel conectar ao banco de dados, entre em contato com o suporte", vbInformation, "Sem Conexão Com o Banco de Dados"
        ConectaBanco = "DSN=PostgreSQL30;UID=postgres;PWD=1234;"
        conn.Open ConectaBanco
            Resume Next
    End If
End If

End Sub
