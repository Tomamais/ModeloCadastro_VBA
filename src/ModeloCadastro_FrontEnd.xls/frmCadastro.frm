VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadastro 
   Caption         =   "Cadastro"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   OleObjectBlob   =   "frmCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadastro"
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

Const colCodigoDoFornecedor As Integer = 1
Const colNomeDaEmpresa As Integer = 2
Const colNomeDoContato As Integer = 3
Const colCargoDoContato As Integer = 4
Const colEndereco As Integer = 5
Const colCidade As Integer = 6
Const colRegiao As Integer = 7
Const colCEP As Integer = 8
Const colPais As Integer = 9
Const colTelefone As Integer = 10
Const colFax As Integer = 11
Const colHomePage As Integer = 12
Const indiceMinimo As Byte = 2
Const corDisabledTextBox As Long = -2147483633
Const corEnabledTextBox As Long = -2147483643
Const nomePlanilhaCadastro As String = "Fornecedores"

Private wsCadastro As Worksheet
Private wbCadastro As Workbook
Private indiceRegistro As Long

Private Sub btnCancelar_Click()
    btnOK.Enabled = False
    btnCancelar.Enabled = False
    Call DesabilitaControles
    Call CarregaDadosInicial
    Call HabilitaBotoesAlteracao
End Sub

Private Sub btnOK_Click()
    Dim proximoId As Long

    'Altera
    If optAlterar.Value Then
        Call SalvaRegistro(CLng(txtCodigoFornecedor.Text), indiceRegistro)
        lblMensagem.Caption = "Registro salvo com sucesso"
    End If
    'Novo
    If optNovo.Value Then
        proximoId = PegaProximoId
        'pega a próxima linha
        Dim proximoIndice As Long
        'atualiza o arquivo para pegar o próximo registro atualizado
        Call AtualizarArquivo(False)
        proximoIndice = wsCadastro.UsedRange.Rows.Count + 1
        Call SalvaRegistro(proximoId, proximoIndice)
        txtCodigoFornecedor = proximoId
        lblMensagem.Caption = "Registro salvo com sucesso"
    End If
    'Excluir
    If optExcluir.Value Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Deseja excluir o registro nº " & txtCodigoFornecedor.Text & " ?", vbYesNo, "Confirmação")

        If result = vbYes Then
            'abre o arquivo para escrita
            Call AtualizarArquivo(False)
            wsCadastro.Range(wsCadastro.Cells(indiceRegistro, colCodigoDoFornecedor), wsCadastro.Cells(indiceRegistro, colCodigoDoFornecedor)).EntireRow.Delete
            'salva
            wbCadastro.Save
            'abre somente leitura
            Call AtualizarArquivo(True)
            Call CarregaDadosInicial
            lblMensagem.Caption = "Registro excluído com sucesso"
        End If
    End If

    Call HabilitaBotoesAlteracao
    Call DesabilitaControles

End Sub

Private Sub btnPesquisar_Click()
    frmPesquisa.Show
End Sub

Private Sub optAlterar_Click()
    If txtCodigoFornecedor.Text <> vbNullString And txtCodigoFornecedor.Text <> "" Then
        Call HabilitaControles
        Call DesabilitaBotoesAlteracao
        'dá o foco ao primeiro controle de dados
        txtNomeEmpresa.SetFocus
    Else
        lblMensagem.Caption = "Não há registro a ser alterado"
    End If
End Sub

Private Sub optExcluir_Click()
    If txtCodigoFornecedor.Text <> vbNullString And txtCodigoFornecedor.Text <> "" Then
        Call DesabilitaBotoesAlteracao
        lblMensagem.Caption = "Modo de exclusão. Confira o dados do registro antes de excluí-lo"
    Else
        lblMensagem.Caption = "Não há registro a ser excluído"
    End If
End Sub

Private Sub optNovo_Click()
    Call LimpaControles
    Call HabilitaControles
    Call DesabilitaBotoesAlteracao
    'dá o foco ao primeiro controle de dados
    txtNomeEmpresa.SetFocus
End Sub

Private Sub UserForm_Initialize()
    
    Call DefinePlanilhaDados
    Call HabilitaBotoesAlteracao
    Call CarregaDadosInicial
    Call DesabilitaControles
End Sub

Private Sub btnAnterior_Click()
    If indiceRegistro > indiceMinimo Then
        indiceRegistro = indiceRegistro - 1
    End If
    If indiceRegistro > 1 Then
        Call CarregaRegistro
    End If
End Sub

Private Sub btnPrimeiro_Click()
    indiceRegistro = indiceMinimo
    If indiceRegistro > 1 Then
        Call CarregaRegistro
    End If
End Sub

Private Sub btnProximo_Click()
    If indiceRegistro < wsCadastro.UsedRange.Rows.Count Then
        indiceRegistro = indiceRegistro + 1
    End If
    If indiceRegistro > 1 Then
        Call CarregaRegistro
    End If
End Sub

Private Sub btnUltimo_Click()
    indiceRegistro = wsCadastro.UsedRange.Rows.Count
    If indiceRegistro > 1 Then
        Call CarregaRegistro
    End If
End Sub

Private Sub CarregaDadosInicial()
    indiceRegistro = 2
    Call CarregaRegistro
End Sub

Private Sub CarregaRegistro()
'carrega os dados do primeiro registro
    With wsCadastro
        If Not IsEmpty(.Cells(indiceRegistro, colCodigoDoFornecedor)) Then
            Me.txtCodigoFornecedor.Text = .Cells(indiceRegistro, colCodigoDoFornecedor).Value
            Me.txtNomeEmpresa.Text = .Cells(indiceRegistro, colNomeDaEmpresa).Value
            Me.txtNomeContato.Text = .Cells(indiceRegistro, colNomeDoContato).Value
            Me.txtCargoContato.Text = .Cells(indiceRegistro, colCargoDoContato).Value
            Me.txtEndereco.Text = .Cells(indiceRegistro, colEndereco).Value
            Me.txtCidade.Text = .Cells(indiceRegistro, colCidade).Value
            Me.txtRegiao.Text = .Cells(indiceRegistro, colRegiao).Value
            Me.txtCEP.Text = .Cells(indiceRegistro, colCEP).Value
            Me.txtPais.Text = .Cells(indiceRegistro, colPais).Value
            Me.txtTelefone.Text = .Cells(indiceRegistro, colTelefone).Value
            Me.txtFax.Text = .Cells(indiceRegistro, colFax).Value
            Me.txtHomePage.Text = .Cells(indiceRegistro, colHomePage).Value
        End If
    End With

    Call AtualizaRegistroCorrente
End Sub

Public Sub CarregaRegistroPorIndice(ByVal indice As Long)
'carrega os dados do registro baseado no índice
    indiceRegistro = indice

    Call CarregaRegistro
End Sub

Private Sub AtualizarArquivo(ByVal ReadOnly As Boolean)
    Dim caminhoCompleto As String
    'fecha o arquivo de dados e tenta abrí-lo
    'guarda o caminho
    caminhoCompleto = wbCadastro.FullName
    wbCadastro.Saved = True
    wbCadastro.Close SaveChanges:=False
    
    'abre o arquivo em modo escrita
    Set wbCadastro = Workbooks.Open(fileName:=caminhoCompleto, ReadOnly:=ReadOnly)
    
    'oculta a janela
    wbCadastro.Windows(1).Visible = False
    
    'reatribui a planilha de cadastro
    Set wsCadastro = wbCadastro.Worksheets(nomePlanilhaCadastro)
End Sub

Private Sub SalvaRegistro(ByVal id As Long, ByVal indice As Long)
    'tenta abrir o arquivo em modo escrita
    Call AtualizarArquivo(False)
    
    With wsCadastro
        .Cells(indice, colCodigoDoFornecedor).Value = id
        .Cells(indice, colNomeDaEmpresa).Value = Me.txtNomeEmpresa.Text
        .Cells(indice, colNomeDoContato).Value = Me.txtNomeContato.Text
        .Cells(indice, colCargoDoContato).Value = Me.txtCargoContato.Text
        .Cells(indice, colEndereco).Value = Me.txtEndereco.Text
        .Cells(indice, colCidade).Value = Me.txtCidade.Text
        .Cells(indice, colRegiao).Value = Me.txtRegiao.Text
        .Cells(indice, colCEP).Value = Me.txtCEP.Text
        .Cells(indice, colPais).Value = Me.txtPais.Text
        .Cells(indice, colTelefone).Value = Me.txtTelefone.Text
        .Cells(indice, colFax).Value = Me.txtFax.Text
        .Cells(indice, colHomePage).Value = Me.txtHomePage.Text
    End With
    
    'salva o arquivo
    Call wbCadastro.Save
    
    'abre o arquivo novamente em modo leitura
    Call AtualizarArquivo(True)

    Call AtualizaRegistroCorrente
End Sub

Private Function PegaProximoId() As Long
    Dim rangeIds As Range
    'pega o range que se refere a toda a coluna do código (id)
    Set rangeIds = wsCadastro.Range(wsCadastro.Cells(indiceMinimo, colCodigoDoFornecedor), wsCadastro.Cells(wsCadastro.UsedRange.Rows.Count, colCodigoDoFornecedor))
    PegaProximoId = WorksheetFunction.Max(rangeIds) + 1
End Function

Private Sub AtualizaRegistroCorrente()
    lblNavigator.Caption = indiceRegistro - 1 & " de " & wsCadastro.UsedRange.Rows.Count - 1
    lblMensagem.Caption = ""
End Sub

Private Sub LimpaControles()
    Me.txtCodigoFornecedor.Text = ""
    Me.txtNomeEmpresa.Text = ""
    Me.txtNomeContato.Text = ""
    Me.txtCargoContato.Text = ""
    Me.txtEndereco.Text = ""
    Me.txtCidade.Text = ""
    Me.txtRegiao.Text = ""
    Me.txtCEP.Text = ""
    Me.txtPais.Text = ""
    Me.txtTelefone.Text = ""
    Me.txtFax.Text = ""
    Me.txtHomePage.Text = ""
End Sub

Private Sub HabilitaControles()
'Me.txtCodigoFornecedor.Locked = False
    Me.txtNomeEmpresa.Locked = False
    Me.txtNomeContato.Locked = False
    Me.txtCargoContato.Locked = False
    Me.txtEndereco.Locked = False
    Me.txtCidade.Locked = False
    Me.txtRegiao.Locked = False
    Me.txtCEP.Locked = False
    Me.txtPais.Locked = False
    Me.txtTelefone.Locked = False
    Me.txtFax.Locked = False
    Me.txtHomePage.Locked = False

    Me.txtNomeEmpresa.BackColor = corEnabledTextBox
    Me.txtNomeContato.BackColor = corEnabledTextBox
    Me.txtCargoContato.BackColor = corEnabledTextBox
    Me.txtEndereco.BackColor = corEnabledTextBox
    Me.txtCidade.BackColor = corEnabledTextBox
    Me.txtRegiao.BackColor = corEnabledTextBox
    Me.txtCEP.BackColor = corEnabledTextBox
    Me.txtPais.BackColor = corEnabledTextBox
    Me.txtTelefone.BackColor = corEnabledTextBox
    Me.txtFax.BackColor = corEnabledTextBox
    Me.txtHomePage.BackColor = corEnabledTextBox
End Sub

Private Sub DesabilitaControles()
'Me.txtCodigoFornecedor.Locked = True
    Me.txtNomeEmpresa.Locked = True
    Me.txtNomeContato.Locked = True
    Me.txtCargoContato.Locked = True
    Me.txtEndereco.Locked = True
    Me.txtCidade.Locked = True
    Me.txtRegiao.Locked = True
    Me.txtCEP.Locked = True
    Me.txtPais.Locked = True
    Me.txtTelefone.Locked = True
    Me.txtFax.Locked = True
    Me.txtHomePage.Locked = True

    Me.txtNomeEmpresa.BackColor = corDisabledTextBox
    Me.txtNomeContato.BackColor = corDisabledTextBox
    Me.txtCargoContato.BackColor = corDisabledTextBox
    Me.txtEndereco.BackColor = corDisabledTextBox
    Me.txtCidade.BackColor = corDisabledTextBox
    Me.txtRegiao.BackColor = corDisabledTextBox
    Me.txtCEP.BackColor = corDisabledTextBox
    Me.txtPais.BackColor = corDisabledTextBox
    Me.txtTelefone.BackColor = corDisabledTextBox
    Me.txtFax.BackColor = corDisabledTextBox
    Me.txtHomePage.BackColor = corDisabledTextBox
End Sub

Private Sub HabilitaBotoesAlteracao()
'habilita os botões de alteração
    optAlterar.Enabled = True
    optExcluir.Enabled = True
    optNovo.Enabled = True
    btnPesquisar.Enabled = True
    btnOK.Enabled = False
    btnCancelar.Enabled = False

    'limpa os valores dos controles
    optAlterar.Value = False
    optExcluir.Value = False
    optNovo.Value = False
End Sub

Private Sub DesabilitaBotoesAlteracao()
'desabilita os botões de alteração
    optAlterar.Enabled = False
    optExcluir.Enabled = False
    optNovo.Enabled = False
    btnPesquisar.Enabled = False
    btnOK.Enabled = True
    btnCancelar.Enabled = True
End Sub

Public Function ProcuraIndiceRegistroPodId(ByVal id As Long) As Long
    Dim i As Long
    Dim retorno As Long
    Dim encontrado As Boolean

    i = indiceMinimo
    With wsCadastro
        Do While Not IsEmpty(.Cells(i, colCodigoDoFornecedor))
            If .Cells(i, colCodigoDoFornecedor).Value = id Then
                retorno = i
                encontrado = True
                Exit Do
            End If
            i = i + 1
        Loop
    End With

    'caso não encontre o registro, retorna -1
    If Not encontrado Then
        retorno = -1
    End If

    ProcuraIndiceRegistroPodId = i
End Function

Private Sub DefinePlanilhaDados()
    Dim abrirArquivo As Boolean
    Dim wb As Workbook
    Dim caminhoCompleto As String
    Dim ARQUIVO_DADOS As String
    Dim PASTA_DADOS As String
    
    abrirArquivo = True
    
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
        
        'verifica se o arquivo não está aberto
        For Each wb In Application.Workbooks
            If wb.name = ARQUIVO_DADOS Then
                abrirArquivo = False
                Exit For
            End If
        Next
        
        'atribui o arquivo
        If abrirArquivo Then
            Set wbCadastro = Workbooks.Open(fileName:=caminhoCompleto, ReadOnly:=True)
        Else
            Set wbCadastro = Workbooks(ARQUIVO_DADOS)
        End If
    Else
        Set wbCadastro = ThisWorkbook
    End If
    
    Set wsCadastro = wbCadastro.Worksheets(nomePlanilhaCadastro)
    
    'oculta o arquivo de dados
    wbCadastro.Windows(1).Visible = False
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'fecha a planilha de dados, se estiver aberta
    If Not wbCadastro Is Nothing Then
        wbCadastro.Saved = True
        wbCadastro.Close SaveChanges:=False
    End If
    
    Set wbCadastro = Nothing
End Sub
