Public con As ADODB.Connection
Public rs As ADODB.Recordset

Public Type Clientes
    codigo As String
    nome As String
    bairro As String
    cidade As String
    dia As String
    mes As String
    ano As String
    fone As String
End Type


Function abreConexao() As Boolean
    Dim usuario, senha, Banco As String
    
    On Error GoTo ErrorConectado:
        If con.State Then
            'MsgBox "Banco ja Aberto !!!"
            abreConexao = True
            Exit Function
        End If
ErrorConectado:
    abreConexao = False
    Set con = New ADODB.Connection
    
    usuario = "*****"
    senha = "*****"
    Banco = "C:\BANCO.fdb"
    
    On Error GoTo ErrorHandler:
        With con
            .ConnectionString = "DSN=BANCO;"
            .ConnectionString = .ConnectionString & "Driver=Firebird/InterBase(r) driver;"
            .ConnectionString = .ConnectionString & "Dbname=" & Banco & ";"
            .ConnectionString = .ConnectionString & "CHARSET=NONE;"
            .ConnectionString = .ConnectionString & "PWD=" & senha & ";"
            .ConnectionString = .ConnectionString & "UID=" & usuario & ";"
            .ConnectionString = .ConnectionString & "Client=C:\Arquivos de programas\Firebird\Firebird_2_0\bin\fbclient.dll;"
            .Open
        End With
        abreConexao = True
        Exit Function
ErrorHandler:
     MsgBox "Deu merda: Erro " & Err.Number & ": " & Err.Description
End Function

Sub fechaConexao()
    On Error GoTo ErrorHandler:
        If con.State Then
            con.Close
        End If
        'MsgBox "Conexao Fechada !!!"
        Exit Sub
ErrorHandler:
     MsgBox "Deu merda: Erro " & Err.Number & ": " & Err.Description
End Sub

Sub abreRS(sql As String)
    On Error GoTo ErrorHandler:
        Set rs = New ADODB.Recordset
        rs.Open sql, con
        'MsgBox "Record Executado !!!"
        Exit Sub
ErrorHandler:
     MsgBox "Deu merda: Erro " & Err.Number & ": " & Err.Description
End Sub

Sub fechaRS()
    On Error GoTo ErrorHandler:
        If rs.Status Then
            rs.Close
        End If
        'MsgBox "Record Fechado !!!"
        Exit Sub
ErrorHandler:
     MsgBox "Deu merda: Erro " & Err.Number & ": " & Err.Description
End Sub


Function gerCodigo() As String
    Dim sql As String
    
    If abreConexao Then
        sql = "select max(codigo)+1 from tbcliente"
        abreRS (sql)
        If Not rs.EOF Then
            gerCodigo = rs(0)
        End If
        fechaRS
    End If
End Function

Sub insertCliente(novo As Clientes)
    Dim sql As String

    If abreConexao Then
        sql = "insert into tbcliente("
        sql = sql & " codigo,"
        sql = sql & " nome,"
        sql = sql & " bairro,"
        sql = sql & " cidade,"
        sql = sql & " dia,"
        sql = sql & " mes,"
        sql = sql & " ano,"
        sql = sql & " fone)"
        sql = sql & " values(" & novo.codigo & ","
        sql = sql & " '" & novo.nome & "',"
        sql = sql & " '" & novo.bairro & "',"
        sql = sql & " '" & novo.cidade & "',"
        sql = sql & " '" & novo.dia & "',"
        sql = sql & " '" & novo.mes & "',"
        sql = sql & " '" & novo.ano & "',"
        sql = sql & " '" & novo.fone & "')"
        con.Execute (sql)
        MsgBox " Cadastrado com sucesso !!!"
    End If
    
End Sub


Sub listarClientes(busca As String)
    Dim sql As String
    Dim i As Integer
    busca = UCase(busca)
    
    If abreConexao Then
        sql = "select"
        sql = sql & " cli.codigo,"
        sql = sql & " cli.data,"
        sql = sql & " cli.nome,"
        sql = sql & " cli.bairro,"
        sql = sql & " cli.cidade,"
        sql = sql & " cli.dia,"
        sql = sql & " cli.mes,"
        sql = sql & " cli.ano,"
        sql = sql & " cli.fone,"
        sql = sql & " cli.fone1"
        sql = sql & " from tbcliente cli"
        sql = sql & " where cli.nome like '%" & busca & "%'"
        sql = sql & " order by cli.codigo desc"
        'MsgBox sql
        abreRS (sql)
        'OBS: BOF fica TRUE se estiver antes do primeiro registro
        'EOF fica TRUE se estiver depois do ultimo registro
        i = 4
        If Not rs.EOF Then
            Do While Not rs.EOF
                Cells(i, 1).Value = rs(0).Value
                Cells(i, 2).Value = rs(1).Value
                Cells(i, 3).Value = rs(2).Value
                Cells(i, 4).Value = rs(3).Value
                Cells(i, 5).Value = rs(4).Value
                Cells(i, 6).Value = rs(5).Value
                Cells(i, 7).Value = rs(6).Value
                Cells(i, 8).Value = rs(7).Value
                Cells(i, 9).Value = rs(8).Value
                Cells(i, 10).Value = rs(9).Value
                rs.MoveNext
                i = i + 1
            Loop
        End If
    End If
End Sub
