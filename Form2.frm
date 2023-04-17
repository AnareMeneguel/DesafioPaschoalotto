VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000002&
   Caption         =   "Form2"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   LinkTopic       =   "Form2"
   ScaleHeight     =   5685
   ScaleWidth      =   14355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton dados 
      Caption         =   "Exportar Dados"
      BeginProperty Font 
         Name            =   "Source Sans Pro Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   5
      ToolTipText     =   "Exporta dados que estão sendo exibidos na tela"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton limpar2 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Source Sans Pro Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Limpa dados de pesquisa"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Pesquisar 
      Caption         =   "Pesquisar"
      BeginProperty Font 
         Name            =   "Source Sans Pro Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "Pesquisa informações relacionadas ao dado"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Pesquisa 
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Text            =   "Pesquise Por CPF, NOME ou DATA INCLUSÃO"
      Top             =   4440
      Width           =   3735
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   7435
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Source Sans Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Voltar 
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "Source Sans Pro Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   0
      ToolTipText     =   "Retorna ao menu inicial"
      Top             =   4560
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form2_Load()
    cmbCampoPesquisar.AddItem "Nome"
    cmbCampoPesquisar

End Sub

Private Sub filtropesquisa_Change()
With Form2.ComboBox1
    .AddItem "Nome"
    .ItemData(.NewIndex) = 1 ' valor para pesquisa por nome
    .AddItem "CPF"
    .ItemData(.NewIndex) = 2 ' valor para pesquisa por CPF
    .AddItem "Telefone"
    .ItemData(.NewIndex) = 3 ' valor para pesquisa por telefone
End With

End Sub
Private Sub dados_Click()
    Dim excelApp As New Excel.Application
    Dim excelWorkbook As Excel.Workbook
    Dim excelWorksheet As Excel.Worksheet
    
    Set excelWorkbook = excelApp.Workbooks.Add
    Set excelWorksheet = excelWorkbook.Worksheets(1)
    
    ' Definir os títulos das colunas
    excelWorksheet.Cells(1, 1) = "Nome"
    excelWorksheet.Cells(1, 2) = "Sobrenome"
    excelWorksheet.Cells(1, 3) = "CPF"
    excelWorksheet.Cells(1, 4) = "Endereço"
    excelWorksheet.Cells(1, 5) = "Telefone"
    excelWorksheet.Cells(1, 6) = "Idade"
    excelWorksheet.Cells(1, 7) = "Mãe"
    excelWorksheet.Cells(1, 8) = "Data Inclusão"
    
    ' Preencher a planilha com os dados da ListView1
    Dim i As Integer
    For i = 1 To ListView1.ListItems.Count
        excelWorksheet.Cells(i + 1, 1) = ListView1.ListItems(i).Text
        excelWorksheet.Cells(i + 1, 2) = ListView1.ListItems(i).SubItems(1)
        excelWorksheet.Cells(i + 1, 3) = ListView1.ListItems(i).SubItems(2)
        excelWorksheet.Cells(i + 1, 4) = ListView1.ListItems(i).SubItems(3)
        excelWorksheet.Cells(i + 1, 5) = ListView1.ListItems(i).SubItems(4)
        excelWorksheet.Cells(i + 1, 6) = ListView1.ListItems(i).SubItems(5)
        excelWorksheet.Cells(i + 1, 8) = ListView1.ListItems(i).SubItems(7)
        excelWorksheet.Cells(i + 1, 7) = ListView1.ListItems(i).SubItems(6)
    Next i
    
    Dim dataHoraAtual As Date
    Dim dataHoraFormatada As String
    Dim local1 As String
    dataHoraAtual = Now()
    dataHoraFormatada = Format(dataHoraAtual, "ddmmyyyyhhmmss")
    
    local1 = "\Dadosfiltro_" & dataHoraFormatada & ".xlsx"
    
    excelWorkbook.SaveAs App.Path & local1
    
    excelApp.Visible = True
    Set excelApp = Nothing
    
    
    
    
    MsgBox "Relatório gerado com sucesso!", vbInformation
End Sub




Private Sub limpar2_Click()
    ' Limpar o campo de pesquisa
    Pesquisa.Text = ""
    
    ' Executar a pesquisa novamente para exibir todos os registros
    pesquisar_Click
End Sub



Private Sub Pesquisa_GotFocus()
    Pesquisa.Text = ""
End Sub


Private Sub pesquisar_Click()
    Dim conn As New ADODB.Connection
    conn.ConnectionString = "DRIVER={PostgreSQL ANSI};SERVER=localhost;PORT=5432;DATABASE=vb6_paschoalotto;UID=postgres;PWD=anare"
    conn.Open
    
    Dim sql As String
    Dim valor As String
    
    ' Obter o valor para pesquisa
    valor = "%" & Pesquisa.Text & "%"
    
    sql = "SELECT * FROM cadastro WHERE nome ILIKE ? OR cpf ILIKE ? OR dt_inclusao ILIKE ?;"
    
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    
    ' o que for pesquisar
    cmd.Parameters.Append cmd.CreateParameter("nome", adVarChar, adParamInput, 255, valor)
    cmd.Parameters.Append cmd.CreateParameter("cpf", adVarChar, adParamInput, 255, valor)
    cmd.Parameters.Append cmd.CreateParameter("dt_inclusao", adVarChar, adParamInput, 255, valor)
    Dim rs As New ADODB.Recordset
    rs.Open cmd
    
   
    ListView1.ListItems.Clear
    
    ' adicionar ao ListView'
    While Not rs.EOF
        Dim registro As ListItem
        Set registro = ListView1.ListItems.Add(, , rs("nome").Value)
        registro.SubItems(1) = rs("sobrenome").Value
        registro.SubItems(2) = rs("cpf").Value
        registro.SubItems(3) = rs("endereco").Value
        registro.SubItems(4) = rs("idade").Value
        registro.SubItems(5) = rs("mae").Value
        registro.SubItems(6) = rs("dt_inclusao").Value
        rs.MoveNext
    Wend
    
    rs.Close
    conn.Close
End Sub



Private Function validarCPF(cpf As String) As Boolean
    ' valida cpf
    Dim soma1 As Integer
    Dim soma2 As Integer
    Dim digito1 As Integer
    Dim digito2 As Integer
    Dim resto1 As Integer
    Dim resto2 As Integer
    Dim i As Integer
    
    ' remove qualquer caractere que não seja um número
    cpf = Replace(Replace(Replace(cpf, ".", ""), "-", ""), "/", "")
    
    If Len(cpf) <> 11 Then '
        validarCPF = False
    Else
        soma1 = 0
        soma2 = 0
        
        For i = 1 To 9
            soma1 = soma1 + (Mid(cpf, i, 1) * (11 - i))
        Next i
        resto1 = (soma1 * 10) Mod 11
        If resto1 = 10 Then
            resto1 = 0
        End If
        
        
        For i = 1 To 10
            soma2 = soma2 + (Mid(cpf, i, 1) * (12 - i))
        Next i
        resto2 = (soma2 * 10) Mod 11
        If resto2 = 10 Then
            resto2 = 0
        End If
        
        ' verifica se os dígitos verificadores são iguais aos dígitos informados pelo usuário
        digito1 = Mid(cpf, 10, 1)
        digito2 = Mid(cpf, 11, 1)
        validarCPF = (resto1 = digito1) And (resto2 = digito2)
    End If
End Function


Private Sub Voltar_Click()
Unload Me
Form1.Show
End Sub

