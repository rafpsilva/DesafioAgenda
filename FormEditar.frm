VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormEditar 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Contato"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   11055
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Excluir Contato"
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
      Left            =   2280
      TabIndex        =   11
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton BtEdSair 
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
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton BtEditar 
      Caption         =   "Salvar Alterações"
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
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox TexId 
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TexEdEmail 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox TexEdTelefone 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   11
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox TexEdSobrenome 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox TexEdNome 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc AdoEditar 
      Height          =   330
      Left            =   9240
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=PostgreSQL30"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PostgreSQL30"
      OtherAttributes =   ""
      UserName        =   "postgres"
      Password        =   "1234"
      RecordSource    =   "select *  from ""contatos"" order by nome"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGridEditar 
      Bindings        =   "FormEditar.frx":0000
      Height          =   4335
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nome"
         Caption         =   "nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "sobrenome"
         Caption         =   "sobrenome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "telefone"
         Caption         =   "telefone"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "email"
         Caption         =   "email"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         Locked          =   -1  'True
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   2654,929
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
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
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label Label2 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   570
   End
End
Attribute VB_Name = "FormEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtEditar_Click()

Dim cmd As New ADODB.Command
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset

'Verifica se há campos em branco
If TexEdNome.Text = "" Or TexEdSobrenome.Text = "" Or TexEdTelefone.Text = "" Or TexEdEmail.Text = "" Then
MsgBox "Todos os Campos Precisam Estar Preenchidos", vbInformation

Else
    'Verifica se há problemas de conexão com o banco de dados
    On Error GoTo ErrorConexao
    ' Inicializar a conexão com o banco de dados PostgreSQL
    ConectaBanco = "DSN=PostgreSQL30;UID=postgres;PWD=1234;"
    conn.Open ConectaBanco
    
    If ValidaEmail(TexEdEmail) Then
        ' Atualização de dados no banco Postgresql
        cmd.ActiveConnection = conn
        cmd.CommandText = "UPDATE Contatos SET (nome, sobrenome, telefone, email) = (? , ?, ?, ?) WHERE id =" & TexId.Text
        cmd.CommandType = adCmdText
        
        cmd.Parameters.Append cmd.CreateParameter("nomeParam", adVarChar, adParamInput, Len(TexEdNome.Text), TexEdNome.Text)
        cmd.Parameters.Append cmd.CreateParameter("sobrenomeParam", adVarChar, adParamInput, Len(TexEdSobrenome.Text), TexEdSobrenome.Text)
        cmd.Parameters.Append cmd.CreateParameter("telefoneParam", adVarChar, adParamInput, Len(TexEdTelefone.Text), TexEdTelefone.Text)
        cmd.Parameters.Append cmd.CreateParameter("emailParam", adVarChar, adParamInput, Len(TexEdEmail.Text), TexEdEmail.Text)
        cmd.Execute
        conn.Close
        MsgBox "Allteração Realizada com Sucesso", vbInformation
        
        ' Limpa os as caixas de texto após a inserção
        TexEdNome.Text = ""
        TexEdSobrenome.Text = ""
        TexEdTelefone.Text = ""
        TexEdEmail.Text = ""
    
    Else
        MsgBox "Email invalido", vbInformation
        TexEdEmail.SetFocus
        Exit Sub
        
    End If
ErrorConexao:
    If Err.Number = "-2147467259" Then
        MsgBox "Não foi possivel conectar ao banco de dados, entre em contato com o suporte", vbInformation, "Sem Conexão Com o Banco de Dados"
        ConectaBanco = "DSN=PostgreSQL30;UID=postgres;PWD=1234;"
        conn.Open ConectaBanco
            Resume Next
    End If
    AdoEditar.Refresh

End If

End Sub
Private Sub BtEdSair_Click()

'Atualiza a tabela DataGridEditar e fecha a tela FormEditar
AdoEditar.Refresh
Unload Me

End Sub
Private Sub Command1_Click()

'Realiza a exclusão
AdoEditar.Recordset.Delete
MsgBox "Excluido com Sucesso", vbInformation

'Limpa as caixas de texto
TexEdNome.Text = ""
TexEdSobrenome.Text = ""
TexEdTelefone.Text = ""
TexEdEmail.Text = ""

End Sub
Private Sub DataGridEditar_Click()

'Exibe os dados da datagrid nas caixas de texto para edição
TexId.Text = DataGridEditar.Columns(0).Text
TexEdNome.Text = DataGridEditar.Columns(1).Text
TexEdSobrenome.Text = DataGridEditar.Columns(2).Text
TexEdTelefone.Text = DataGridEditar.Columns(3).Text
TexEdEmail.Text = DataGridEditar.Columns(4).Text

End Sub

Sub teste()


End Sub

Private Sub Form_Load()

    ' Carregue a imagem na PictureBox
    Picture1.Picture = LoadPicture("C:\Users\rafae\OneDrive\Área de Trabalho\pexels-andrea-piacquadio-3786525.jpg")

    ' Defina o novo tamanho da PictureBox
    'Picture1.Width = 500
    'Picture1.Height = 500
    
    ' Redimensione a imagem para se ajustar ao novo tamanho da PictureBox
    'SetPictureImageSize Picture1, 500, 500
End Sub

Private Sub SetPictureImageSize(ByVal pb As Picture, ByVal Width As Integer, ByVal Height As Integer)
    Dim originalWidth As Integer
    Dim originalHeight As Integer
    Dim ratioWidth As Double
    Dim ratioHeight As Double
    
    ' Verifique as dimensões originais da imagem
   ' originalWidth = pb.Picture.Width
    'originalHeight = pb.Picture.Height
    
    ' Calcule as proporções de redimensionamento
    'ratioWidth = Width / originalWidth
   ' ratioHeight = Height / originalHeight
    
    ' Aplique o redimensionamento à imagem
'    pb.ScaleMode = vbPixels
 '   pb.AutoRedraw = True
  '  pb.PaintPicture pb.Picture, 0, 0, Width, Height, , , , , vbSrcCopy
   ' pb.Refresh
End Sub


    'PictureBox1.Picture = LoadPicture("C:\Users\rafae\OneDrive\Área de Trabalho\pexels-andrea-piacquadio-3786525.jpg")

