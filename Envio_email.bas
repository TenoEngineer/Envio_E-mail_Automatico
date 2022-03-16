'******************* ENVIO DE E-MAIL    *********************
'#AUTHOR: HEITOR TENO MÜLLER | ISMAEL HENRIQUE BUTTENBENDER
'#DATE: 25/11/2021
'#CONTACT: heitor.muller@herval.com.br | heitortmuller@gmail.com

Public resultado As VbMsgBoxResult
Public E_MAIL_1 As String
Public E_MAIL_2 As String
Public E_MAIL_3 As String
Public E_MAIL_4 As String
Public CONFERENCIA As String
Public COTAS As String
Public ID As String
Public ERRO As Integer
Public USUARIO As String
Public CAMINHO As String
Public ASSUNTO As String


Sub ENVIAR_EMAILS_DIV()
     
    Run "Seguranca"
    If resultado = vbNo Then
        Exit Sub
    End If

    'DESOCULTA ABAS
    Sheets("e-mail DIV").Visible = True
    Sheets("ASSINATURA").Visible = True
    
    'INSERE OS CAMINHOS DA PLANILHA E DECLARA ALGUMAS VARIAVEIS
    ERRO = 1
    USUARIO = Environ("Username")
    CAMINHO = "C:\Users\" & USUARIO
    ASSUNTO = Sheets("PAINEL").Range("H2").Value

    'COMEÇA A SELECIONAR OS QUE VÃO MANDAR E-MAIL
    Sheets("DIVERGÊNCIAS_").Visible = True
    Sheets("DIVERGÊNCIAS_").Select
    Range("U1", Range("U1").End(xlDown)).Clear
    Range("C1").Value = "Filiais"
    Range("C1").AutoFilter
    Range("C1", Range("C1").End(xlDown)).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
    Range("C1", Range("C1").End(xlDown)).Copy
    Range("U1").PasteSpecial xlPasteValues
    Range("C1").AutoFilter
    X = Range("U2", Range("U2").End(xlDown)).Rows.Count
    If X > 10000 Then X = 1
    For i = 1 To X
        On Error Resume Next
        Workbooks("E-mail.xls").Close
        Application.DisplayAlerts = True
        Workbooks("Envio de e-mails.xlsm").Activate
        
        cadastro_email = Sheets("CADASTRO E-MAIL").Range("A:E", Sheets("CADASTRO E-MAIL").Range("A:E").End(xlDown))
        Sheets("DIVERGÊNCIAS_").Select
        
        ID = Range("U" & i + 1).Value
        If ID = "" Or ID < 1 Then GoTo 6
        
        E_MAIL_1 = Application.VLookup(Range("U" & i + 1), cadastro_email, 2, 0)
        E_MAIL_2 = Application.VLookup(Range("U" & i + 1), cadastro_email, 3, 0)
        E_MAIL_3 = Application.VLookup(Range("U" & i + 1), cadastro_email, 4, 0)
        E_MAIL_4 = Application.VLookup(Range("U" & i + 1), cadastro_email, 5, 0)
        CONFERENCIA = Sheets("CADASTRO E-MAIL").Range("G2").Value

        GoTo 5
            
3       If E_MAIL_1 <> "" Then
            On Error GoTo 6
            Run "Gerar_Enviar"
        Else
            On Error GoTo 6
            Exit For
        End If
        
        On Error Resume Next
        Workbooks("E-mail.xls").Close
        Application.DisplayAlerts = True
        Workbooks("Envio de e-mails.xlsm").Activate
        
11  Next
    
    Run "Concluido"
    GoTo 8
    
    'Limpa os dados do e-mail anterior
5    Sheets("e-mail DIV").Rows("19:500").Clear
    
    'Filtra os dados da filial
    Sheets("DIVERGÊNCIAS").Select
    Range("C2").Select
    If Sheets("DIVERGÊNCIAS").AutoFilterMode = True Then Selection.AutoFilter
    Sheets("DIVERGÊNCIAS").Range("A2:N2", Sheets("CAIXA").Range("A2:N2").End(xlDown)).Select
    Selection.AutoFilter
    Sheets("DIVERGÊNCIAS").Range("A2:N2", Range("A2:N2").End(xlDown)).AutoFilter Field:=14, Criteria1:="S"
    Sheets("DIVERGÊNCIAS").Range("A2:N2", Range("A2:N2").End(xlDown)).AutoFilter Field:=3, Criteria1:=ID
    Range("C2:K2", Range("C2:K2").End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select

    ASSIN = Application.WorksheetFunction.Subtotal(2, Selection)
    If ASSIN > 10000 Then GoTo 11
    Selection.Copy Worksheets("e-mail DIV").Range("A18")

    'Insere a assinatura do e-mail
    Sheets("e-mail DIV").Range("B:I").Columns.AutoFit
    Worksheets("ASSINATURA").Range("A1:E6").Copy Worksheets("e-mail DIV").Cells(20 + ASSIN / 2, 1)
    Sheets("DIVERGÊNCIAS").Range("A1").AutoFilter
    
    'Gera um excel novo
    Sheets("e-mail DIV").Activate
    Sheets("e-mail DIV").Copy After:=Sheets(3)
    Sheets("e-mail DIV (2)").Name = ID
    
    Sheets(ID).Select
    Sheets(ID).Move
    
    Application.DisplayAlerts = False
    
    'Salva o excel em um caminho especifico
    Run "Salva_Excel"
    Sheets(ID).Select
    GoTo 3

6   On Error Resume Next
    Workbooks("E-mail.xls").Close
    Application.DisplayAlerts = True
    Workbooks("Envio de e-mails.xlsm").Activate
    Sheets("ERROS").Activate
    Range("A" & ERRO).Value = ID
    ERRO = ERRO + 1
    GoTo 11

8   On Error Resume Next
    Workbooks("E-mail.xls").Close
    Application.DisplayAlerts = True
    Workbooks("Envio de e-mails.xlsm").Activate
    Sheets("e-mail DIV").Select
    Sheets("DIVERGÊNCIAS_").Visible = False
    Sheets("e-mail DIV").Visible = False
    Sheets("ASSINATURA").Visible = False
    Sheets("PAINEL").Select
   
End Sub

Sub ENVIAR_EMAILS_CAIXA()
     
    Run "Seguranca"
    If resultado = vbNo Then
        Exit Sub
    End If

    'DESOCULTA ABAS
    Sheets("e-mail CAIXA").Visible = True
    Sheets("ASSINATURA").Visible = True
    
    'INSERE OS CAMINHOS DA PLANILHA E DECLARA ALGUMAS VARIAVEIS
    ERRO = 1
    USUARIO = Environ("Username")
    CAMINHO = "C:\Users\" & USUARIO
    ASSUNTO = Sheets("PAINEL").Range("H5").Value

    'COMEÇA A SELECIONAR OS QUE VÃO MANDAR E-MAIL
    Sheets("CAIXA_").Visible = True
    Sheets("CAIXA_").Select
    Range("U1", Range("U1").End(xlDown)).Clear
    Range("B1").Value = "Filiais"
    Range("B1").AutoFilter
    Range("B1", Range("B1").End(xlDown)).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
    Range("B1", Range("B1").End(xlDown)).Copy
    Range("U1").PasteSpecial xlPasteValues
    Range("B1").AutoFilter
    X = Range("U2", Range("U2").End(xlDown)).Rows.Count
    If X > 10000 Then X = 1
    For i = 1 To X
        On Error Resume Next
        Workbooks("E-mail.xls").Close
        Application.DisplayAlerts = True
        Workbooks("Envio de e-mails.xlsm").Activate
        
        cadastro_email = Sheets("CADASTRO E-MAIL").Range("A:E", Sheets("CADASTRO E-MAIL").Range("A:E").End(xlDown))
        Sheets("CAIXA_").Select
        
        ID = Range("U" & i + 1).Value
        If ID = "" Or ID < 1 Then GoTo 6
        
        E_MAIL_1 = Application.VLookup(Range("U" & i + 1), cadastro_email, 2, 0)
        E_MAIL_2 = Application.VLookup(Range("U" & i + 1), cadastro_email, 3, 0)
        E_MAIL_3 = Application.VLookup(Range("U" & i + 1), cadastro_email, 4, 0)
        E_MAIL_4 = Application.VLookup(Range("U" & i + 1), cadastro_email, 5, 0)
        CONFERENCIA = Sheets("CADASTRO E-MAIL").Range("G2").Value
        
        GoTo 5
            
3       If E_MAIL_1 <> "" Then
            On Error GoTo 6
            Run "Gerar_Enviar"
        Else
            Exit For
        End If
        
        On Error Resume Next
        Workbooks("E-mail.xls").Close
        Application.DisplayAlerts = True
        Workbooks("Envio de e-mails.xlsm").Activate
        
11  Next
    
    Run "Concluido"
    GoTo 8
    
    'Limpa os dados do e-mail anterior
5   Sheets("e-mail CAIXA").Rows("5:300").Clear
    
    'Filtra os dados da filial
    Sheets("CAIXA").Activate
    Range("B1").Select
    If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
    Range("A1:H1", Range("A1:H1").End(xlDown)).Select
    Selection.AutoFilter
    ActiveSheet.Range("A1:H1", Range("A1:H1").End(xlDown)).AutoFilter Field:=7, Criteria1:="S"
    ActiveSheet.Range("A1:H1", Range("A1:H1").End(xlDown)).AutoFilter Field:=2, Criteria1:=ID
    Range("A1:F1", Range("A1:F1").End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    
    ASSIN = Application.WorksheetFunction.Subtotal(2, Selection)
    If ASSIN > 10000 Then GoTo 11
    Selection.Copy Worksheets("e-mail CAIXA").Range("A5")
    
    
    'Insere a assinatura do e-mail
    Sheets("e-mail CAIXA").Range("B:I").Columns.AutoFit
    Worksheets("ASSINATURA").Range("A1:E6").Copy Worksheets("e-mail CAIXA").Cells(7 + ASSIN / 2, 1)
    Sheets("CAIXA").Range("A1").AutoFilter
    
    'Gera um excel novo
    Sheets("e-mail CAIXA").Activate
    Sheets("e-mail CAIXA").Copy After:=Sheets(3)
    Sheets("e-mail CAIXA (2)").Name = ID
    
    Sheets(ID).Select
    Sheets(ID).Move
    
    Application.DisplayAlerts = False
    
    'Salva o excel em um caminho especifico
    Run "Salva_Excel"
    Sheets(ID).Select
    GoTo 3

6   On Error Resume Next
    Workbooks("E-mail.xls").Close
    Application.DisplayAlerts = True
    Workbooks("Envio de e-mails.xlsm").Activate
    Sheets("ERROS").Activate
    Range("A" & ERRO).Value = ID
    ERRO = ERRO + 1

8   On Error Resume Next
    Workbooks("E-mail.xls").Close
    Application.DisplayAlerts = True
    Workbooks("Envio de e-mails.xlsm").Activate
    Sheets("CAIXA_").Visible = False
    Sheets("e-mail CAIXA").Visible = False
    Sheets("ASSINATURA").Visible = False
    Sheets("PAINEL").Select
   
End Sub

Sub ENVIAR_EMAILS_COTAS()
     
    Run "Seguranca"
    If resultado = vbNo Then
        Exit Sub
    End If

    'DESOCULTA ABAS
    Sheets("e-mail COTAS").Visible = True
    Sheets("ASSINATURA").Visible = True

    'INSERE OS CAMINHOS DA PLANILHA E DECLARA ALGUMAS VARIAVEIS
    ERRO = 1
    USUARIO = Environ("Username")
    CAMINHO = "C:\Users\" & USUARIO
    ASSUNTO = Sheets("PAINEL").Range("H8").Value

    'COMEÇA A SELECIONAR OS QUE VÃO MANDAR E-MAIL
    Sheets("COTAS_").Visible = True
    Sheets("COTAS_").Activate
    Range("U1", Range("U1").End(xlDown)).Clear
    Range("B1").Value = "Filiais"
    Range("B1").AutoFilter
    Range("B1", Range("B1").End(xlDown)).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
    Range("B1", Range("B1").End(xlDown)).Copy
    Range("U1").PasteSpecial xlPasteValues
    Range("B1").AutoFilter
    X = Range("U2", Range("U2").End(xlDown)).Rows.Count
    
    If X > 10000 Then X = 1
    For i = 1 To X
        On Error Resume Next
        Workbooks("E-mail.xls").Close
        Application.DisplayAlerts = True
        Workbooks("Envio de e-mails.xlsm").Activate
        
        cadastro_email = Sheets("CADASTRO E-MAIL").Range("A:E", Sheets("CADASTRO E-MAIL").Range("A:E").End(xlDown))
        Sheets("COTAS_").Activate
        
        ID = Range("U" & i + 1).Value
        If ID = "" Or ID < 1 Then GoTo 6
        
        E_MAIL_1 = Application.VLookup(Range("U" & i + 1), cadastro_email, 2, 0)
        E_MAIL_2 = Application.VLookup(Range("U" & i + 1), cadastro_email, 3, 0)
        E_MAIL_3 = Application.VLookup(Range("U" & i + 1), cadastro_email, 4, 0)
        E_MAIL_4 = Application.VLookup(Range("U" & i + 1), cadastro_email, 5, 0)
        COTAS = Sheets("CADASTRO E-MAIL").Range("F2").Value

        GoTo 5
            
3       If E_MAIL_1 <> "" Then
            On Error GoTo 6
            Run "Gerar_Enviar_Cotas"
        Else
            Exit For
        End If
        
        On Error Resume Next
        Workbooks("E-mail.xls").Close
        Application.DisplayAlerts = True
        Workbooks("Envio de e-mails.xlsm").Activate
        
11  Next
    
    Run "Concluido"
    GoTo 8
    
    'Limpa os dados do e-mail anterior
5   Sheets("e-mail COTAS").Rows("7:300").Clear

    'Filtra os dados da filial
    Sheets("COTAS").Select
    Range("B1").Select
    If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
    Range("A1:F1", Range("A1:F1").End(xlDown)).Select
    Selection.AutoFilter
    ActiveSheet.Range("A1:F1", Range("A1:F1").End(xlDown)).AutoFilter Field:=6, Criteria1:="S"
    ActiveSheet.Range("A1:F1", Range("A1:F1").End(xlDown)).AutoFilter Field:=2, Criteria1:=ID
    Range("B1:E1", Range("B1:E1").End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    
    ASSIN = Application.WorksheetFunction.Subtotal(2, Selection)
    If ASSIN > 10000 Then GoTo 11
    Selection.Copy Worksheets("e-mail COTAS").Range("A7")

    'Insere a assinatura do e-mail
    Sheets("e-mail COTAS").Range("B:I").Columns.AutoFit
    Worksheets("ASSINATURA").Range("A1:E6").Copy Worksheets("e-mail COTAS").Cells(9 + ASSIN / 2, 1)
    Sheets("COTAS").Range("A1").AutoFilter

    'Gera um excel novo
    Sheets("e-mail COTAS").Activate
    Sheets("e-mail COTAS").Copy After:=Sheets(3)
    Sheets("e-mail COTAS (2)").Name = ID
    
    Sheets(ID).Select
    Sheets(ID).Move
    
    Application.DisplayAlerts = False
    
    'Salva o excel em um caminho especifico
    Run "Salva_Excel"
    Sheets(ID).Select
    GoTo 3

6   On Error Resume Next
    Workbooks("E-mail.xls").Close
    Application.DisplayAlerts = True
    Workbooks("Envio de e-mails.xlsm").Activate
    Sheets("ERROS").Activate
    Range("A" & ERRO).Value = ID
    ERRO = ERRO + 1

8   On Error Resume Next
    Workbooks("E-mail.xls").Close
    Application.DisplayAlerts = True
    Workbooks("Envio de e-mails.xlsm").Activate
    Sheets("COTAS_").Visible = False
    Sheets("e-mail COTAS").Visible = False
    Sheets("ASSINATURA").Visible = False
    Sheets("PAINEL").Select
   
End Sub
Sub Salva_Excel()
    ChDir CAMINHO
    ActiveWorkbook.SaveAs Filename:=CAMINHO & "\" & "E-mail.xls", _
        FileFormat:=xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
End Sub
Sub Concluido()
    MsgBox "Envio de E-mails Concluído!"
    If ERRO > 1 Then MsgBox "Houve Erro! Filiais com erro salvo na aba ERROS"
End Sub
Sub Seguranca()
    'Mensagem de segurança
    resultado = MsgBox("Deseja enviar os e-mails?", vbYesNo, "Tomando uma decisão")
    If resultado = vbYes Then
    Else
         MsgBox "Envio de e-mails cancelado"
         Exit Sub
    End If
End Sub
Sub Gerar_Enviar()
    Workbooks("E-mail.xls").EnvelopeVisible = True
    With ActiveSheet.MailEnvelope
        .Item.To = E_MAIL_1 & ";" & E_MAIL_2 & ";" & E_MAIL_3 & ";" & E_MAIL_4 & ";" & CONFERENCIA
        .Item.Subject = ASSUNTO & " - " & ID
        .Introduction = " "
        .Item.Send
    End With
    Workbooks("E-mail.xls").EnvelopeVisible = False
    
    Workbooks("E-mail.xls").Close
    Application.DisplayAlerts = True
End Sub
Sub Gerar_Enviar_Cotas()
    Workbooks("E-mail.xls").EnvelopeVisible = True
    With ActiveSheet.MailEnvelope
        .Item.To = E_MAIL_1 & ";" & E_MAIL_2 & ";" & E_MAIL_3 & ";" & E_MAIL_4 & ";" & COTAS
        .Item.Subject = ASSUNTO & " - " & ID
        .Introduction = " "
        .Item.Send
    End With
    Workbooks("E-mail.xls").EnvelopeVisible = False
    
    Workbooks("E-mail.xls").Close
    Application.DisplayAlerts = True
End Sub
