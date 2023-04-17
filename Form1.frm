VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Source Sans Pro Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar"
      BeginProperty Font 
         Name            =   "Source Sans Pro Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      ToolTipText     =   "Copia os CPF/CNPJ para a area de transferencia"
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "CPF/CNPJ Duplicados"
      Height          =   2895
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   2175
      Begin MSComctlLib.ListView Importados 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   4260
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton gerar 
      Caption         =   "Exportar Relatorio Geral"
      BeginProperty Font 
         Name            =   "Source Sans Pro Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Exporta um relatorio com todas as informações contidas no Banco de dados"
      Top             =   2640
      Width           =   5055
   End
   Begin VB.CommandButton mostrar 
      Caption         =   "Pesquisar Informações"
      BeginProperty Font 
         Name            =   "Source Sans Pro Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Mostra todas as Informações e extração de relatorio"
      Top             =   1440
      Width           =   5055
   End
   Begin VB.CommandButton importar 
      Caption         =   "Importar"
      BeginProperty Font 
         Name            =   "Source Sans Pro Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Importar documento excel para o banco de dados"
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public ListBox2 As ListBox


Private Sub exportar_Click()
    ' conexão com o bd
    Dim conexao As ADODB.Connection
    Set conexao = New ADODB.Connection
    conn.ConnectionString = "DRIVER={PostgreSQL ANSI};SERVER=localhost;PORT=5432;DATABASE=vb6_paschoalotto;UID=postgres;PWD=anare"
    conexao.Open
    
    '  Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conexao
    
    '  gerar o arquivo
    Dim xlApp As Excel.Application
    Set xlApp = New Excel.Application
    xlApp.Visible = True
    Dim xlWb As Excel.Workbook
    Set xlWb = xlApp.Workbooks.Add
    Dim xlWs As Excel.Worksheet
    Set xlWs = xlWb.Sheets(1)
    
    ' Cria cabeçalhos
    xlWs.Range("A1").Value = "Nome"
    xlWs.Range("B1").Value = "CPF"
    
    '   procurar por CPFs duplicados
    Dim i As Integer
    Dim j As Integer
    Dim cpf As String
    Dim duplicado As Boolean
    For i = 0 To ListBox1.ListCount - 1
        cpf = ListBox1.List(i, 1)
        duplicado = False
        
        ' Verifica se o CPF está duplicado
        For j = 0 To ListBox1.ListCount - 1
            If i <> j And ListBox1.List(j, 1) = cpf Then
                duplicado = True
                Exit For
            End If
        Next j
        
        If duplicado Then
            ' Faz uma consulta ao banco de dados para obter as informações de nome e CPF
            rs.Open "SELECT Nome FROM Tabela WHERE CPF = '" & cpf & "'", conexao
            If Not rs.EOF Then
                ' Adiciona as informações no arquivo Excel
                xlWs.Range("A" & xlWs.Cells.Rows.Count).End(xlUp).Offset(1, 0).Value = rs.Fields("Nome").Value
                xlWs.Range("B" & xlWs.Cells.Rows.Count).End(xlUp).Offset(1, 0).Value = cpf
            End If
            rs.Close
        End If
    Next i
    
    ' Salva o arquivo
    xlWb.SaveAs App.Path & "\duplicados.xlsx"
    
    ' Fecha o bd
    conexao.Close
    
    ' Exibe mensagem
    MsgBox "Arquivo de duplicados gerado com sucesso!", vbInformation
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Dim cpf As String
    
    
    For i = 1 To Importados.ListItems.Count
        cpf = cpf & Importados.ListItems(i).Text & vbCrLf
    Next i
    
    ' Limpar a área para colar os CPF'
    Clipboard.Clear
    Clipboard.SetText cpf
    
  
    MsgBox "CPF(s) copiado(s) para a área de transferência!"
End Sub



Private Sub Command2_Click()
    Unload Me
    End
End Sub


Private Sub Gerar_Click()

    ' Conectar ao bd
    Dim conn As New ADODB.Connection
    conn.ConnectionString = "DRIVER={PostgreSQL ANSI};SERVER=localhost;PORT=5432;DATABASE=vb6_paschoalotto;UID=postgres;PWD=anare"
    conn.Open
    
    Dim sql As String
    sql = "SELECT * FROM cadastro order by nome;"
    Dim rs As New ADODB.Recordset
    rs.Open sql, conn
    
    
    Dim xlApp As Object
    Set xlApp = CreateObject("Excel.Application")
    
    

    ' Cworkbook
    Dim xlWb As Excel.Workbook
    Set xlWb = xlApp.Workbooks.Add
    
    ' worksheet
    Dim xlWs As Excel.Worksheet
    Set xlWs = xlWb.Worksheets.Add
    
    ' Nome coluna'
    xlWs.Range("A1").Value = "Nome"
    xlWs.Range("B1").Value = "Sobrenome"
    xlWs.Range("C1").Value = "CPF"
    xlWs.Range("D1").Value = "Endereço"
    xlWs.Range("E1").Value = "Idade"
    xlWs.Range("F1").Value = "Mãe"
    xlWs.Range("G1").Value = "Data Inclusão"
    ' copia e cola excel'
    Dim row As Integer
    row = 2
    While Not rs.EOF
        xlWs.Range("A" & row).Value = rs("nome").Value
        xlWs.Range("B" & row).Value = rs("sobrenome").Value
        xlWs.Range("C" & row).Value = rs("cpf").Value
        xlWs.Range("D" & row).Value = rs("endereco").Value
        xlWs.Range("E" & row).Value = rs("idade").Value
        xlWs.Range("F" & row).Value = rs("mae").Value
        xlWs.Range("G" & row).Value = rs("dt_inclusao").Value
        row = row + 1
        rs.MoveNext
    Wend
    'salva'
    Dim dataHoraAtual As Date
    Dim dataHoraFormatada As String
    Dim local1 As String
    dataHoraAtual = Now()
    dataHoraFormatada = Format(dataHoraAtual, "ddmmyyyyhhmmss")
    
    local1 = "\relatorio_" & dataHoraFormatada & ".xlsx"
    
    xlWb.SaveAs App.Path & local1
    
   
    xlApp.Visible = True
    
    rs.Close
    conn.Close
    

    MsgBox "Relatório gerado com sucesso!", vbInformation
  
    
End Sub


Private Sub importar_Click()
    Dim objDialog As Object: Set objDialog = CreateObject("MSComDlg.CommonDialog")
    objDialog.Filter = "Arquivos do Excel (*.xls; *.xlsx)|*.xls;*.xlsx"
    objDialog.ShowOpen
    
    Importados.View = lvwReport
    Importados.ColumnHeaders.Add , , "CPF duplicado"

    If objDialog.fileName <> "" Then
        Dim excelApp As Object
        Dim excelWorkbook As Object
        Dim excelWorksheet As Object
        
        Set excelApp = CreateObject("Excel.Application")
        Set excelWorkbook = excelApp.Workbooks.Open(objDialog.fileName)
        Set excelWorksheet = excelWorkbook.Worksheets(1)
        
       

        
        For i = 2 To excelWorksheet.UsedRange.Rows.Count
            
            If WorksheetFunction.CountA(excelWorksheet.Rows(i)) > 0 Then
                Dim nome As String
                Dim sobrenome As String
                Dim cpf As Variant
                Dim endereco As String
                Dim telefone As String
                Dim idade As Variant
                Dim mae As String
                Dim dataAtual As String
                dataAtual = Format(Date, "dd/MM/yyyy")

                nome = excelWorksheet.Cells(i, 1)
                sobrenome = excelWorksheet.Cells(i, 2)
                cpf = excelWorksheet.Cells(i, 3)
                endereco = excelWorksheet.Cells(i, 4)
                telefone = excelWorksheet.Cells(i, 5)
                idade = excelWorksheet.Cells(i, 6)
                mae = excelWorksheet.Cells(i, 7)
                
                'Conectar ao bd '
                Dim conn As New ADODB.Connection
                conn.ConnectionString = "DRIVER={PostgreSQL ANSI};SERVER=localhost;PORT=5432;DATABASE=vb6_paschoalotto;UID=postgres;PWD=anare"
                conn.Open
                
                'Verifica se dados já foram inseridos no bd'
                Dim sqlVerificar As String
                sqlVerificar = "SELECT COUNT(*) FROM cadastro WHERE nome = '" & nome & "' AND sobrenome = '" & sobrenome & "' AND cpf = '" & cpf & "' AND endereco = '" & endereco & "' AND telefone = '" & telefone & "' AND idade = '" & CStr(idade) & "' AND mae = '" & mae & "';"
                Dim rsVerificar As New ADODB.Recordset
                rsVerificar.Open sqlVerificar, conn
                
                If rsVerificar.Fields(0).Value = 0 Then 'Verifica se não existe registro no banco
                    'Inserir os dados no bd'
                    Dim sql As String
                    sql = "INSERT INTO cadastro (nome, sobrenome, cpf, endereco, telefone, idade, mae, dt_inclusao) VALUES ('" & nome & "', '" & sobrenome & "', '" & cpf & "', '" & endereco & "', '" & telefone & "', '" & idade & "', '" & mae & "', '" & dataAtual & "');"

                    conn.Execute sql
                Else ' adiciona CPF duplicado na lista'
                    duplicados = True
                    listaCPFDuplicados = listaCPFDuplicados & cpf & vbNewLine
                End If
                
                rsVerificar.Close
                
                'Fechar  bd'
                conn.Close
            End If
        Next i

        
        excelWorkbook.Close
        excelApp.Quit
        Set excelWorksheet = Nothing
        Set excelWorkbook = Nothing

        
        
    If objDialog.fileName <> "" Then
        If MsgBox("Deseja prosseguir com a importação?!", vbOKCancel, "importação") = vbOK Then
            MsgBox "Arquivo importado com sucesso!", vbInformation, "Importação"
             Else
                 MsgBox "Importação cancelada pelo usuário.", vbInformation, "Importação"
        End If
  
    End If
        If duplicados = True Then
            MsgBox "Atenção: Existem informações na planilha que já estão inseridas no banco observe a lista ao lado" & vbNewLine & vbNewLine & "CPF(s) duplicado(s):" & vbNewLine & listaCPFDuplicados, vbInformation, "Importação"
    
        End If
       
    
   Dim cpfArray() As String
    cpfArray = Split(listaCPFDuplicados, vbNewLine)
    
    For i = 0 To UBound(cpfArray)
        Importados.ListItems.Add , , cpfArray(i)
    Next i
   
End If


   
End Sub

Private Sub Limpar_Click()
ListBox1.Clear
End Sub


Private Sub mostrar_Click()
    
    Dim conn As New ADODB.Connection
    conn.ConnectionString = "DRIVER={PostgreSQL ANSI};SERVER=localhost;PORT=5432;DATABASE=vb6_paschoalotto;UID=postgres;PWD=anare"
    conn.Open
    
   
    Dim sql As String
    sql = "SELECT * FROM cadastro;"
    Dim rs As New ADODB.Recordset
    rs.Open sql, conn
    
    ' Limpar o ListV
    Form2.ListView1.ListItems.Clear
    Form2.ListView1.ColumnHeaders.Clear
    
    ' Adicionar as colunas no ListView do Form2
    Form2.ListView1.View = lvwReport
    Form2.ListView1.ColumnHeaders.Add , , "Nome", 1000
    Form2.ListView1.ColumnHeaders.Add , , "Sobrenome", 1000
    Form2.ListView1.ColumnHeaders.Add , , "CPF", 1000
    Form2.ListView1.ColumnHeaders.Add , , "Endereço", 2000
    Form2.ListView1.ColumnHeaders.Add , , "Telefone", 2000
    Form2.ListView1.ColumnHeaders.Add , , "Idade", 500
    Form2.ListView1.ColumnHeaders.Add , , "Mãe", 1500
    Form2.ListView1.ColumnHeaders.Add , , "Data Inclusão", 2500
    
    ' Adicionar as linhas
    Form2.ListView1.Gridlines = True
    Form2.ListView1.FullRowSelect = True
    
    '  adicionar os registros ao ListView
    While Not rs.EOF
        Dim registro As ListItem
        Set registro = Form2.ListView1.ListItems.Add(, , rs("nome").Value)
        registro.SubItems(1) = rs("sobrenome").Value
        registro.SubItems(2) = rs("cpf").Value
        registro.SubItems(3) = rs("endereco").Value
        registro.SubItems(4) = rs("telefone").Value
        registro.SubItems(5) = rs("idade").Value
        registro.SubItems(6) = rs("mae").Value
        registro.SubItems(7) = rs("dt_inclusao").Value
        rs.MoveNext
    Wend
    
    
    rs.Close
    conn.Close
  
    Form2.Show
End Sub





