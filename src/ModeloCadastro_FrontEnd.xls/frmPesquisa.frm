VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPesquisa 
   Caption         =   "Pesquisa"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8160
   OleObjectBlob   =   "frmPesquisa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modelo de Aplicativo de Cadastro em VBA no Microsoft Excel
'Autor: Tomás Vásquez
'http://www.tomasvasquez.com.br
'http://tomas.vasquez.blog.uol.com.br
'março de 2008

Option Explicit

'constantes para auxiliar na verificação do código
Private Const Ascendente As Byte = 0
Private Const Descendente As Byte = 1
Private caminhoArquivoDados As String

Private Sub btnExportar_Click()
    Call Exportar
End Sub

Private Sub btnFiltrar_Click()
    Call PopulaListBox(txtNomeEmpresa.Text, txtNomeContato.Text, txtEndereco.Text, txtTelefone.Text, txtRegiao.Text)
End Sub

Private Sub lstLista_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstLista.ListIndex > 0 Then
        Dim indiceRegistro As Long
        indiceRegistro = frmCadastro.ProcuraIndiceRegistroPodId(lstLista.List(lstLista.ListIndex, 0))
        If indiceRegistro <> -1 Then
            Call frmCadastro.CarregaRegistroPorIndice(indiceRegistro)
        End If
        Unload Me
    Else
        lblMensagens.Caption = "É preciso selecionar um item válido na lista"
    End If
End Sub

Private Sub DefinePlanilhaDados()
    Dim wb As Workbook
    Dim caminhoCompleto As String
    Dim ARQUIVO_DADOS As String
    Dim PASTA_DADOS As String
    
    ThisWorkbook.Activate
    
    ARQUIVO_DADOS = Range("ARQUIVO_DADOS").Value
    PASTA_DADOS = Range("PASTA_DADOS").Value
    
    If ThisWorkbook.name <> ARQUIVO_DADOS Then
        'monta a string do caminho completo
        If PASTA_DADOS = vbNullString Or PASTA_DADOS = "" Then
            caminhoCompleto = Replace(ThisWorkbook.FullName, ThisWorkbook.name, vbNullString) & ARQUIVO_DADOS
        Else
            If Right(PASTA_DADOS, 1) = "\" Then
                caminhoCompleto = PASTA_DADOS & ARQUIVO_DADOS
            Else
                caminhoCompleto = PASTA_DADOS & "\" & ARQUIVO_DADOS
            End If
        End If
    End If
    
    caminhoArquivoDados = caminhoCompleto
    
End Sub

Private Sub UserForm_Initialize()
'preenche o cboDirecao e seleciona o primeiro item
    cboDirecao.Clear
    cboDirecao.AddItem "Ascendente"
    cboDirecao.AddItem "Descendente"
    cboDirecao.ListIndex = 0

    Call DefinePlanilhaDados
    Call PopulaCidades
    Call PopulaListBox(vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
End Sub

Private Sub Exportar()
    Dim i As Integer
    Dim NewWorkBook As Workbook
    Dim rst As ADODB.Recordset
    ' Preenche o RecordSet com os filtros atuais
    Set rst = PreecheRecordSet(txtNomeEmpresa.Text, txtNomeContato.Text, txtEndereco.Text, txtTelefone.Text, txtRegiao.Text)
    'cria um novo Workbook
    Set NewWorkBook = Application.Workbooks.Add
    ' Efetua loop em todos os campos, retornando os nomes de campos
    ' à planilha.
    For i = 0 To rst.Fields.Count - 1
        NewWorkBook.Sheets(1).Range("A1").Offset(0, i).Value = rst.Fields(i).name
    Next i

    NewWorkBook.Sheets(1).Range("A2").CopyFromRecordset rst
    NewWorkBook.Activate
End Sub


Private Sub PopulaCidades()
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim sql As String

    Set conn = New ADODB.Connection
    With conn
        .Provider = "Microsoft.JET.OLEDB.4.0"
        .ConnectionString = "Data Source=" & caminhoArquivoDados & ";Extended Properties=Excel 8.0;"
        .Open
    End With

    sql = "SELECT DISTINCT Cidade FROM [Fornecedores$]"

    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = conn
        .Open sql, conn, adOpenDynamic, _
              adLockBatchOptimistic
    End With

    Do While Not rst.EOF
        If Not IsNull(rst(0).Value) Then
            lstCidades.AddItem rst(0).Value
        End If
        rst.MoveNext
    Loop

    ' Fecha o conjunto de registros.
    Set rst = Nothing
    ' Fecha a conexão.
    conn.Close

End Sub

Private Sub PopulaListBox(ByVal NomeEmpresa As String, _
                          ByVal NomeContato As String, _
                          ByVal Endereco As String, _
                          ByVal Telefone As String, _
                          ByVal Regiao As String)

    On Error GoTo TrataErro

    Dim rst As ADODB.Recordset
    Dim campo As Field
    Dim myArray() As Variant
    Dim i As Integer

    Set rst = PreecheRecordSet(NomeEmpresa, NomeContato, Endereco, Telefone, Regiao)

    'pega o número de registros para atribuí-lo ao listbox
    lstLista.ColumnCount = rst.Fields.Count

    'preenche o combobox com os nomes dos campos
    'persiste o índice
    Dim indiceTemp As Long
    indiceTemp = cboOrdenarPor.ListIndex
    cboOrdenarPor.Clear
    For Each campo In rst.Fields
        cboOrdenarPor.AddItem campo.name
    Next
    'recupera o índice selecionado
    cboOrdenarPor.ListIndex = indiceTemp

    'coloca as linhas do RecordSet num Array, se houver linhas neste
    If Not rst.EOF And Not rst.BOF Then
        myArray = rst.GetRows
        'troca linhas por colunas no Array
        myArray = Array2DTranspose(myArray)
        'atribui o Array ao listbox
        lstLista.List = myArray
        'adiciona a linha de cabeçalho da coluna
        lstLista.AddItem , 0
        'preenche o cabeçalho
        For i = 0 To rst.Fields.Count - 1
            lstLista.List(0, i) = rst.Fields(i).name
        Next i
        'seleciona o primeiro item da lista
        lstLista.ListIndex = 0
    Else
        lstLista.Clear
    End If

    'atualiza o label de mensagens
    If lstLista.ListCount <= 0 Then
        lblMensagens.Caption = lstLista.ListCount & " registros encontrados"
    Else
        lblMensagens.Caption = lstLista.ListCount - 1 & " registros encontrados"
    End If

    ' Fecha o conjunto de registros.
    Set rst = Nothing
    ' Fecha a conexão.
    'conn.Close

TrataSaida:
    Exit Sub
TrataErro:
    Debug.Print Err.Description & vbNewLine & Err.Number & vbNewLine & Err.Source
    Resume TrataSaida
End Sub

Private Function PreecheRecordSet(ByVal NomeEmpresa As String, _
                                  ByVal NomeContato As String, _
                                  ByVal Endereco As String, _
                                  ByVal Telefone As String, _
                                  ByVal Regiao As String) As Recordset
    On Error GoTo TrataErro

    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim sql As String
    Dim sqlWhere As String
    Dim sqlOrderBy As String
    Dim i As Integer
    Dim campo As Field
    Dim myArray() As Variant

    Set conn = New ADODB.Connection
    With conn
        .Provider = "Microsoft.JET.OLEDB.4.0"
        .ConnectionString = "Data Source=" & caminhoArquivoDados & ";Extended Properties=Excel 8.0;"
        .Open
    End With

    sql = "SELECT * FROM [Fornecedores$]"

    'monta a cláusula WHERE
    'NomeDaEmpresa
    Call MontaClausulaWhere(txtNomeEmpresa.name, "NomeDaEmpresa", sqlWhere)

    'NomeDoContato
    Call MontaClausulaWhere(txtNomeContato.name, "NomeDoContato", sqlWhere)

    'Endereço
    Call MontaClausulaWhere(txtEndereco.name, "Endereço", sqlWhere)

    'Cidade
    For i = 1 To lstCidades.ListCount
        'verifica se o item está selecionado
        If lstCidades.Selected(i - 1) Then
            'Monta a cláusula WHERE com OR
            Debug.Print lstCidades.List(i - 1) & " selecionado"
            If sqlWhere <> vbNullString Then
                sqlWhere = sqlWhere & " OR"
            End If
            sqlWhere = sqlWhere & " UCASE(Cidade) LIKE UCASE('%" & Trim(lstCidades.List(i - 1)) & "%')"
        End If
    Next

    'Telefone
    Call MontaClausulaWhere(txtTelefone.name, "Telefone", sqlWhere)

    'Região
    Call MontaClausulaWhere(txtRegiao.name, "Região", sqlWhere)

    'faz a união da string SQL com a cláusula WHERE
    If sqlWhere <> vbNullString Then
        sql = sql & " WHERE " & sqlWhere
    End If

    'faz a união da string SQL com a cláusula ORDER BY
    If cboOrdenarPor.ListIndex <> -1 Then
        sqlOrderBy = " ORDER BY " & cboOrdenarPor.List(cboOrdenarPor.ListIndex, 0)
        'define a direção
        Select Case cboDirecao.ListIndex
        Case Ascendente
            sqlOrderBy = sqlOrderBy & " ASC"
        Case Descendente
            sqlOrderBy = sqlOrderBy & " DESC"
        End Select
        'une a query order ao sql
        sql = sql & sqlOrderBy
    End If

    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    With rst
        .ActiveConnection = conn
        .Open sql, conn, adOpenForwardOnly, _
              adLockBatchOptimistic
    End With

    Set rst.ActiveConnection = Nothing

    ' Fecha a conexão.
    conn.Close

    Set PreecheRecordSet = rst
    Exit Function
TrataErro:
    Set rst = Nothing
End Function

Private Sub MontaClausulaWhere(ByVal NomeControle As String, ByVal NomeCampo As String, ByRef sqlWhere As String)
'NomeDoContato
    If Trim(Me.Controls(NomeControle).Text) <> vbNullString Then
        If sqlWhere <> vbNullString Then
            sqlWhere = sqlWhere & " AND"
        End If
        sqlWhere = sqlWhere & " UCASE(" & NomeCampo & ") LIKE UCASE('%" & Trim(Me.Controls(NomeControle).Text) & "%')"
    End If
End Sub

'Faz a transpasição de um array, transformando linhas em colunas
Private Function Array2DTranspose(avValues As Variant) As Variant
    Dim lThisCol As Long, lThisRow As Long
    Dim lUb2 As Long, lLb2 As Long
    Dim lUb1 As Long, lLb1 As Long
    Dim avTransposed As Variant

    If IsArray(avValues) Then
        On Error GoTo ErrFailed
        lUb2 = UBound(avValues, 2)
        lLb2 = LBound(avValues, 2)
        lUb1 = UBound(avValues, 1)
        lLb1 = LBound(avValues, 1)

        ReDim avTransposed(lLb2 To lUb2, lLb1 To lUb1)
        For lThisCol = lLb1 To lUb1
            For lThisRow = lLb2 To lUb2
                avTransposed(lThisRow, lThisCol) = avValues(lThisCol, lThisRow)
            Next
        Next
    End If

    Array2DTranspose = avTransposed
    Exit Function

ErrFailed:
    Debug.Print Err.Description
    Debug.Assert False
    Array2DTranspose = Empty
    Exit Function
    Resume
End Function

