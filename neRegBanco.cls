VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neRegBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Set rRs = RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 2)****************************ok
'         Set rRs = RegBanco_Critica_Estrutura(ADOConnectRodbel, cFields, 6)***************ok
'         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 7)*********************ok
'         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)*********************OK
'              Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)****************OK
'              Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)*******************OK
'         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 15)********************OK
'              Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 16)****************OK
'         Set rRs = RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 44)*******************OK
'              Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 4)********************ok
'              Call RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 5) ********************ok
'              Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 13)****************OK





'         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 18)********************FALTA TOTAL
'              Call RegBanco_Critica_Cred_Acesso(ADOConnectRodbel, cFields, 19)



'




' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS NOVATOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Novatos(ByVal sBdRm As String, _
                                     ByVal sBdRodBel As String, _
                                     ByVal c_Data_Admissao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nGrvOkPessoa As Integer 'verificar existencia na tabela pessoa
Dim nGrvOkCredPessoa As Integer 'verificar existencia na tabela cred_pessoa
Dim nGrvOkCredenciais As Integer 'verificar existencia na tabela credenciais
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set que pesquisara se existe o funcionario no Rodbel
Dim nx As Integer 'VARIAVEL CONTAGEM LOOP
Dim nAre_Numero As Integer 'CAMPO PARA GRAVACAO DO ATRIBUTO DA TABELA CRED-ACESSO

On Error GoTo Erro

Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura

'Set ADOConnectRm = cConect.Coneccao_ODBC(sBdRodBel, "A")
Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

Set cFields = New Collection
cFields.Add c_Data_Admissao, "c_Data_Admissao"

Set Funcionarios_Novatos = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 1)


nGravou = 0
nLidos = 0

If Funcionarios_Novatos.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO NAO EXISTA NO CADASTRO
   
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Novatos.RecordCount

   Funcionarios_Novatos.MoveFirst
   
   While Not Funcionarios_Novatos.EOF

'         If Val(Funcionarios_Novatos!CHAPA) = 3284 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
        
         Rem CRITICA PARA SABER SE O PIS EXISTE CASO NAO, SERÁ CRITICADO
         
         If Len(Trim(Funcionarios_Novatos!PISPARAFGTS)) = 0 Or IsNull(Funcionarios_Novatos!PISPARAFGTS) Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            sMsg = "Funcionario com o Numero do PIS em brancos"
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            GoTo PROXIMO_REGISTRO
         End If
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Novatos!CHAPA
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS e
         Rem PESSOAS, CASO EXISTA, IGNORAR REGISTRO
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         
         If rRs.RecordCount > 0 Then
            GoTo PROXIMO_REGISTRO
         End If
         
         Set cFields = New Collection
         cFields.Add IIf(IsNull(Funcionarios_Novatos!INTEGRGERENCIAL), "0", Funcionarios_Novatos!INTEGRGERENCIAL)
         
         Rem VERIFICA A EXISTENCIA DO CENTRO DE CUSTO DA RM NO BANCO DA RODBEL
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Estrutura(ADOConnectRodbel, cFields, 6)
         
        
         If rRs!ENCONTRADO = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            sMsg = "Funcionario com o Centro de custo inexistente -  " & cFields(1)
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 1
         End If
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Novatos!CHAPA

         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. CRED_PESSOAS, CASO EXISTA, ERRO.
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 2)
         nGrvOkPessoa = 0
         
         If rRs.RecordCount = 1 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            sMsg = "Tentando incluir Func. em PESSOAS,já cadastrado"
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 0 ' modificado para continuar verificando as outras tabelas que está faltando ser atualizada
            nGrvOkPessoa = 1 ' acrescentado para nao gravar um novo registro
         End If
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Novatos!CHAPA
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS, CASO EXISTA, ERRO.
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 7)
         nGrvOkCredPessoa = 0
         
         If rRs.RecordCount = 1 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            sMsg = "Existe na TAB.(CRED_PESSOA) e não existe na (PESSOAS)"
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 0 ' modificado para continuar verificando as outras tabelas que está faltando ser atualizada
            nGrvOkCredPessoa = 1 ' acrescentado para nao gravar um novo registro
         
         End If
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Novatos!CHAPA

         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. CREDENCIAIS, CASO EXISTA, ERRO.
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 44)
         nGrvOkCredenciais = 0 ' rRs!CD_ESTRUTURA_ORGANIZACIONAL
         
         If rRs.RecordCount = 1 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            sMsg = "Tentando incluir Func. em CREDENCIAIS,já cadastrado"
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 0 ' modificado para continuar verificando as outras tabelas que está faltando ser atualizada
            nGrvOkCredenciais = 1 ' acrescentado para nao gravar um novo registro
         End If
         
        
         If nGrvOk = 0 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS E CRED_PESSOAS *****************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            If nGrvOkPessoa = 0 Then
               Set cFields = New Collection
               cFields.Add Funcionarios_Novatos!CHAPA 'NUMERO 1
               cFields.Add IIf(IsNull(Funcionarios_Novatos!NOME), "", Funcionarios_Novatos!NOME) 'NOME 2
               cFields.Add IIf(IsNull(Funcionarios_Novatos!CPF), "", Funcionarios_Novatos!CPF) 'CPF 3
               cFields.Add IIf(IsNull(Funcionarios_Novatos!INTEGRGERENCIAL), 1, Funcionarios_Novatos!INTEGRGERENCIAL) 'OBSERVACAO 4
               cFields.Add "" 'E-MAIL 5
               cFields.Add IIf(IsNull(Funcionarios_Novatos!CARTIDENTIDADE), "", Funcionarios_Novatos!CARTIDENTIDADE) 'RG 6
               cFields.Add IIf(IsNull(Funcionarios_Novatos!INTEGRGERENCIAL), "", Funcionarios_Novatos!INTEGRGERENCIAL) 'NUMERO 7
               cFields.Add "0" 'FLAGREENTRADA 8
               cFields.Add "0" 'MASTER 9
               cFields.Add "Nao_Disponivel.JPG" 'FOTO 10
               cFields.Add "" 'SENHA 11
               cFields.Add "" 'RAMAL 12
               cFields.Add "" 'CAMPOPERSO1 13
               cFields.Add " " 'CAMPOPERSO2 14
               cFields.Add " " 'CAMPOPERSO3 15
               cFields.Add " " 'CAMPOPERSO4 16
               cFields.Add " " 'CAMPOPERSO5 17
               cFields.Add "8" 'PESSIT_NUMERO - de 2 para 8 18
               cFields.Add "4" 'tipo acesso 19
               cFields.Add "" 'DTSITFINAL 20
               cFields.Add " " 'CRED_PROV 21
               cFields.Add "" 'TOK_NUMEROSERIE 22
               cFields.Add IIf(IsNull(Funcionarios_Novatos!PISPARAFGTS), "", Funcionarios_Novatos!PISPARAFGTS) 'PISPARAFGTS 23
               cFields.Add "11" 'CD_SITUACAO_PESSOA 24
               cFields.Add "1" 'CD_ESTRUTURA_ORGANIZACIONAL 25
               cFields.Add "1" 'CD_ESTRUTURA_ORG_EMPRESA 26
               cFields.Add "0" 'TP_MASTER 27
               
               Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 3)
            End If
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAL  ***************
            If nGrvOkCredenciais = 0 Then
               Set cFields = New Collection
               cFields.Add "" '       [CD_CREDENCIAL]
               cFields.Add "1" '      ,[CD_TIPO_CREDENCIAL]
               cFields.Add "1" '      ,[CD_ESTRUTURA_ORGANIZACIONAL]
               cFields.Add "1" '      ,[TP_TECNOLOGIA]
               cFields.Add Funcionarios_Novatos!CHAPA '      ,[NU_CREDENCIAL]
               cFields.Add "1" '      ,[FL_PERMANENTE]
               cFields.Add "0" '      ,[FL_BLOQUEADA]
               cFields.Add "" '      ,[DS_MOTIVO_BLOQUEIO]
               cFields.Add "" '      ,[DT_INICIO_VALIDADE]
               cFields.Add "" '      ,[DT_FIM_VALIDADE]
               cFields.Add "" '      ,[DS_SENHA]
               cFields.Add "0" '      ,[TP_MASTER]
               cFields.Add "0" '      ,[FL_REENTRADA]
               cFields.Add "" '      ,[CD_PERFIL_ACESSO]
               cFields.Add "" '      ,[CD_AREA]
               cFields.Add "" '      ,[DT_ULTIMA_AUTORIZACAO_ACESSO]
               cFields.Add "" '      ,[CD_EQP_ULT_AUTORIZACAO_ACESSO]
               cFields.Add "1" '      ,[FL_DISPONIVEL_USO]
               cFields.Add "0" '      ,[FL_DISPENSA_SENHA]
               cFields.Add "0" '      ,[FL_SUPERVISOR_EQUIPAMENTO]
               cFields.Add "0" '      ,[FL_CREDENCIAL_PUBLICA]
               cFields.Add "" '      ,[HR_RECOLHIMENTO_COFRE]
               cFields.Add "0" '      ,[NU_PORCENTAGEM_SORTEIO]
               cFields.Add "" '      ,[DT_CREDENCIAL_SORTEADA]
               cFields.Add "" '      ,[BN_QRCODE]
               
               Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 4)
               
            End If
            
            
            Rem *************  CREDENCIAL_PESSOAS  ***************
            If nGrvOkCredPessoa = 0 Then
               
               Rem PEGAR O CAMPO CD _CREDENCIAL RECEM CRIADA NA TABELA CREDENCIAL.
               Set cFields = New Collection
               cFields.Add Funcionarios_Novatos!CHAPA
               Set rRs = New ADODB.Recordset
               Set rRs = RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 61)
               
               Set cFields = New Collection
               cFields.Add rRs!CD_CREDENCIAL 'CRED_NUMERO
               cFields.Add rRs!CD_PESSOA   'PES_NUMERO
               cFields.Add Format(Now(), "DD/MM/YYYY HH:MM:SS") 'CRPES_DATAENTRADA
               cFields.Add ""
               cFields.Add ""
               cFields.Add ""
               cFields.Add ""
               cFields.Add "1"
               cFields.Add "1"
            
               Call RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 5)
             End If
         
           
            Rem ************************************************************************************************
            
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            nGravou = nGravou + 1
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            sMsg = "Inclusão Novatos, a partir de " & Mid$(c_Data_Admissao, 8, 2) & "/" & Mid$(c_Data_Admissao, 6, 2) & "/" & Mid$(c_Data_Admissao, 2, 4)
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            
'''            Set cFields = New Collection
'''            cFields.Add Funcionarios_Novatos!CHAPA
'''
'''            Rem LEITURA DA TABELA DE GRUPOS_AREAS, VERIFICANDO OS RELOGIOS E SUAS EXISTENCIAS
'''
'''            Set rRs = New ADODB.Recordset
'''            Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 18)
'''
'''
'''            If rRs.RecordCount = 0 Then
'''               ADOConnectRodbel.CommitTrans
'''               ADOConnectRodbel.BeginTrans
'''               GoTo PROXIMO_REGISTRO
'''            End If
'''
'''            Rem VERIFICAR OS REGISTROS DE GRUPOS-AREAS SELECIONADOS E GRAVAR OS REGISTROS NA TABELA DE CRED_ACESSO
'''
'''            While Not rRs.EOF
'''                  nGrvOk = 0
'''                  If rRs!GRP_NUMERO = 6 Or rRs!GRP_NUMERO = 12 Or rRs!GRP_NUMERO = 13 Then
'''                     nGrvOk = 1
'''                  Else
'''                     If rRs!GRP_NUMERO = 1 Or rRs!GRP_NUMERO = 7 Or rRs!GRP_NUMERO = 8 Or _
'''                        rRs!GRP_NUMERO = 10 Or rRs!GRP_NUMERO = 11 Or rRs!GRP_NUMERO = 14 Then
'''                        If Funcionarios_Novatos!CARGO <> "00006" And _
'''                           Funcionarios_Novatos!CARGO <> "00008" And _
'''                           Funcionarios_Novatos!CARGO <> "00016" And _
'''                           Funcionarios_Novatos!CARGO <> "00020" And _
'''                           Funcionarios_Novatos!CARGO <> "00025" And _
'''                           Funcionarios_Novatos!CARGO <> "00044" And _
'''                           Funcionarios_Novatos!CARGO <> "00054" Then
'''                           nGrvOk = 1
'''                        End If
'''                     End If
'''                  End If
'''
'''                  Rem GRAVAR REGISTRO DE CRED-ACESSO DE ACORDO COM A REGRA ACIMA, PASSANDO OS 2,3,4,5(LIVRES)
'''
'''                  If nGrvOk = 0 Then
'''                     nAre_Numero = rRs!ARE_NUMEROENTRADA
'''
'''                     For nx = 1 To 2
'''                         Rem *************  CRED-ACESSO  ***************
'''                         Set cFields = New Collection
'''                         cFields.Add "'" & Funcionarios_Novatos!CHAPA & "'" 'CRED_NUMERO
'''                         cFields.Add rRs!GRP_NUMERO 'GRP_NUMERO
'''                         cFields.Add nAre_Numero 'ARE_NUMERO
'''                         cFields.Add "1" 'CRACTI_NUMERO
'''                         cFields.Add "" 'CRAC_DIFPORGRUPO
'''                         cFields.Add "" 'JORN_NUMERO
'''                         cFields.Add "" 'TURN_NUMERO
'''                         cFields.Add "999" 'CRAC_QTDTOTALACESSOS
'''                         cFields.Add "" 'CRAC_QTDACESSOS
'''                         cFields.Add "0" 'CRAC_VERFSENHA
'''                         Rem VERIFICAR O GRUPO DOS RELOGIOS SE SAO DIGITAIS, SE SIM=1 SE NAO = 0
'''                         If rRs!GRP_NUMERO = 1 Or _
'''                            rRs!GRP_NUMERO = 2 Or _
'''                            rRs!GRP_NUMERO = 4 Or _
'''                            rRs!GRP_NUMERO = 6 Then
'''                            cFields.Add "0" 'CRAC_VERFDIGITAL
'''                         Else
'''                            cFields.Add "1" 'CRAC_VERFDIGITAL
'''                         End If
'''                         cFields.Add "20000101" 'CRAC_ULTPASSAGEM
'''
'''                         Call RegBanco_Critica_Cred_Acesso(ADOConnectRodbel, cFields, 19)
'''
'''                         nAre_Numero = rRs!ARE_NUMEROSAIDA
'''
'''                     Next
'''
'''                  End If
'''
'''                  rRs.MoveNext
'''
'''            Wend
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Novatos.EOF Then
            Funcionarios_Novatos.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
               
         Rem caso tenha gravado os registros normal de um funcionario, este será atualizado
         
         ADOConnectRodbel.CommitTrans
         ADOConnectRodbel.BeginTrans
   
   Wend

End If

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Novatos"
   sCodFun = "0000"
   sMsg = "Nao ha Novatos no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If


ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
   sCodFun = Funcionarios_Novatos!CHAPA
Else
   sStatusMsg = "2"
   sCodFun = "0000"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Novatos"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS DESLIGADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Desligados(ByVal sBdRm As String, _
                                        ByVal sBdRodBel As String, _
                                        ByVal c_Data_Demissao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add c_Data_Demissao, "c_Data_Demissao"

Set Funcionarios_Desligados = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 9)

nGravou = 0
nLidos = 0

If Funcionarios_Desligados.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO NAO EXISTA NO CADASTRO
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Desligados.RecordCount

   Funcionarios_Desligados.MoveFirst
   
   While Not Funcionarios_Desligados.EOF
'         If Val(Funcionarios_Desligados!CHAPA) = 1429 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
         
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Desligados!CHAPA
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS e PESSOAS,
         Rem CASO EXISTA, VERIFICAR SE JA FOI DEMITIDO CASO SIM IGNORAR REGISTRO
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         If rRs.RecordCount = 0 Then
            GoTo PROXIMO_REGISTRO
         Else
            If rRs!pessit_numero = 8 Then 'de 3 para 9
               GoTo PROXIMO_REGISTRO
            End If
         End If
         
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Desligados!CHAPA
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. PESSOAS, NAO CASO NAO EXISTA, ERRO.
         If RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 2)!ENCONTRADO = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Desligados"
            sCodFun = Funcionarios_Desligados!CHAPA
            sMsg = "Funcionarios Existe na TAB.CREDENCIAIS"
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 1
         
         End If
         
         If nGrvOk = 0 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS ********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            Rem  ATUALIZACAO DO CAMPO PESSIST_NUMERO sStatusMsg = 3
            Set cFields = New Collection
            cFields.Add Funcionarios_Desligados!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add "8" 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAIS  ***************
            Set cFields = New Collection
            cFields.Add Funcionarios_Desligados!CHAPA 'CRED_NUMERO
            cFields.Add "" 'CRED_FLAGREENTRADA
            cFields.Add "" 'CRED_AREANOMOMENTO
            cFields.Add "" 'CRED_DTVALIDINICIAL
            cFields.Add "" 'CRED_DTVALIDFINAL
            cFields.Add "1" 'CRED_BLOQUEADA
            cFields.Add "" 'CRED_FLAGCATRACA
            cFields.Add "" 'CRTI_NUMERO
            cFields.Add "" 'CRED_MASTER
            cFields.Add "" 'CRED_ULTPASSAGEM
            cFields.Add "" 'CRED_EQPINUMERO
            cFields.Add "" 'CRED_SENHA
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'CRED_PROVDEVOLV
            cFields.Add "" 'CRED_NUMRA
            
            Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
            
            Rem ************************************************************************************************
            
            nGravou = nGravou + 1
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Desligados"
            sCodFun = Funcionarios_Desligados!CHAPA
            sMsg = "Desligamento funcionario, data Periodo de " & c_Data_Demissao
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Desligados.EOF Then
            Funcionarios_Desligados.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Desligados"
   sCodFun = "0000"
   sMsg = "Nao ha Desligamento no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

'If nGravou > 0 Then
'   If Err.Number > 9999 And Err.Number < 10001 Then 'Erro de consistencia de dados para atualizacao do RodBel
'
'   Else
'
'   End If
'End If

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Desligados"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS DESLIGADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_AfastLongoTempo(ByVal sBdRm As String, _
                                        ByVal sBdRodBel As String, _
                                        ByVal c_Data_Demissao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add c_Data_Demissao, "c_Data_Demissao"

Set Funcionarios_AfastLongoTempo = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 62)

nGravou = 0
nLidos = 0

If Funcionarios_AfastLongoTempo.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO NAO EXISTA NO CADASTRO
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_AfastLongoTempo.RecordCount

   Funcionarios_AfastLongoTempo.MoveFirst
   
   While Not Funcionarios_AfastLongoTempo.EOF
'         If Val(Funcionarios_AfastLongoTempo!CHAPA) = 1429 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
         
         
         Set cFields = New Collection
         cFields.Add Funcionarios_AfastLongoTempo!CHAPA
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS e PESSOAS,
         Rem CASO EXISTA, VERIFICAR SE JA FOI DEMITIDO CASO SIM IGNORAR REGISTRO
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         If rRs.RecordCount = 0 Then
            GoTo PROXIMO_REGISTRO
         Else
            If rRs!pessit_numero = 8 Then 'de 3 para 9
               GoTo PROXIMO_REGISTRO
            End If
         End If
         
         
         Set cFields = New Collection
         cFields.Add Funcionarios_AfastLongoTempo!CHAPA
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. PESSOAS, NAO CASO NAO EXISTA, ERRO.
         If RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 2)!ENCONTRADO = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Desligados"
            sCodFun = Funcionarios_AfastLongoTempo!CHAPA
            sMsg = "Funcionarios Existe na TAB.CREDENCIAIS"
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 1
         
         End If
         
         If nGrvOk = 0 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS ********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            Rem  ATUALIZACAO DO CAMPO PESSIST_NUMERO sStatusMsg = 3
            Set cFields = New Collection
            cFields.Add Funcionarios_AfastLongoTempo!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add "8" 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAIS  ***************
            Set cFields = New Collection
            cFields.Add Funcionarios_AfastLongoTempo!CHAPA 'CRED_NUMERO
            cFields.Add "" 'CRED_FLAGREENTRADA
            cFields.Add "" 'CRED_AREANOMOMENTO
            cFields.Add "" 'CRED_DTVALIDINICIAL
            cFields.Add "" 'CRED_DTVALIDFINAL
            cFields.Add "1" 'CRED_BLOQUEADA
            cFields.Add "" 'CRED_FLAGCATRACA
            cFields.Add "" 'CRTI_NUMERO
            cFields.Add "" 'CRED_MASTER
            cFields.Add "" 'CRED_ULTPASSAGEM
            cFields.Add "" 'CRED_EQPINUMERO
            cFields.Add "" 'CRED_SENHA
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'CRED_PROVDEVOLV
            cFields.Add "" 'CRED_NUMRA
            
            Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
            
            Rem ************************************************************************************************
            
            nGravou = nGravou + 1
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Desligados"
            sCodFun = Funcionarios_AfastLongoTempo!CHAPA
            sMsg = "Desligamento Temporario, data Periodo de " & c_Data_Demissao
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_AfastLongoTempo.EOF Then
            Funcionarios_AfastLongoTempo.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Desligados"
   sCodFun = "0000"
   sMsg = "Nao ha Desligamento no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

'If nGravou > 0 Then
'   If Err.Number > 9999 And Err.Number < 10001 Then 'Erro de consistencia de dados para atualizacao do RodBel
'
'   Else
'
'   End If
'End If

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Desligados"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS EM FERIAS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Ferias_Histo(ByVal sBdRm As String, _
                                          ByVal sBdRodBel As String, _
                                          ByVal c_Dt_Ferias As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset
Dim c_Data_Ferias As String

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

Set cFields = New Collection
cFields.Add c_Dt_Ferias, "c_Data_Ferias"

Set Funcionarios_Ferias_Histo = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 122)

nGravou = 0
nLidos = 0
c_Data_Ferias = CDate(Format(Now(), "dd/mm/yyyy"))


If Funcionarios_Ferias_Histo.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO NAO EXISTA NO CADASTRO
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Ferias_Histo.RecordCount
   
   Funcionarios_Ferias_Histo.MoveFirst
   
   While Not Funcionarios_Ferias_Histo.EOF
   
'         If Val(Funcionarios_Ferias_Histo!CHAPA) = 124 Then
''             GoTo PROXIMO_REGISTRO
'         End If

         nGrvOk = 0
         
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Ferias_Histo!CHAPA
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS e PESSOAS,
         Rem CASO EXISTA, VERIFICAR SE JA FOI DEMITIDO(3) OU JA ESTEJA AFASTADO(4) CASO SIM IGNORAR REGISTRO
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         If rRs.RecordCount = 0 Then
            nsStatusMsg = 0
         Else
            Rem VERIFICAR O INTERVALO DO PERIODO DE FERIAS E TAMBEM
            Rem O sStatusMsg DO FUNCIONARIO, CASO SATISFACA, ENTRA NO AFASTAMENTO
            Rem NO CASO ENTRANDO DE FERIAS sStatusMsg = 4
            nsStatusMsg = 0
            
            If Funcionarios_Ferias_Histo!INICPROGFERIAS1 <= CDate(c_Data_Ferias) And _
               Funcionarios_Ferias_Histo!FIMPROGFERIAS1 >= CDate(c_Data_Ferias) Then 'CASO A DATA DE FERIAS ESTEJA NO INTERVALO ENCONTRADO, ENTAO ESTA DE FERIAS
               If rRs!pessit_numero = 8 Then
                  nsStatusMsg = 0
               Else
                If rRs!pessit_numero = 11 Then 'SO PODERA ENTRAR DE FERIAS CASO ESTA DESBLOQUEADO, CASO ESTEJA COM BLOQUEO DE OUTRA ACAO, NADA FARA.
                   nsStatusMsg = 8
                End If
               End If
            End If
            
            Rem VERIFICAR O INTERVALO DO PERIODO DE FERIAS E TAMBEM
            Rem O sStatusMsg DO FUNCIONARIO, NO CASO VOLTANDO DE FERIAS sStatusMsg = 2
            If Funcionarios_Ferias_Histo!FIMPROGFERIAS1 < CDate(c_Data_Ferias) Then
               If rRs!pessit_numero = 8 Then
                  nsStatusMsg = 11
               End If
            End If
            
         End If
         
         Rem NO CASO O sStatusMsg PARA ATUALIZAR OS DADOS SERA 2 E 4 RESPECTIVAMENTE
         If nsStatusMsg = 8 Or nsStatusMsg = 11 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS********************************************
            Rem ************************************************************************************************
            
'            Rem *************  PESSOAS  ***************
'            Set cFields = New Collection
'            cFields.Add Funcionarios_Ferias_Histo!CHAPA
'            cFields.Add "" 'NOME
'            cFields.Add "" 'CPF
'            cFields.Add "" 'OBSERVACAO
'            cFields.Add "" 'E-MAIL
'            cFields.Add "" 'IDENTIDADE
'            cFields.Add "" 'CENTRO DE CUSTO
'            cFields.Add "" 'FLAGREENTRADA
'            cFields.Add "" 'MASTER
'            cFields.Add "" 'FOTO
'            cFields.Add "" 'SENHA
'            cFields.Add "" 'RAMAL
'            cFields.Add "" 'CAMPOPERSO1
'            cFields.Add "" 'CAMPOPERSO2
'            cFields.Add "" 'CAMPOPERSO3
'            cFields.Add "" 'CAMPOPERSO4
'            cFields.Add "" 'CAMPOPERSO5
'            cFields.Add nsStatusMsg 'PESSIT_NUMERO
'            If nGrvOk = 2 Then
'               cFields.Add "" 'DTSITINICIAL
'               cFields.Add "" 'DTSITFINAL
'            Else
'               cFields.Add Mid$(Funcionarios_Ferias_Histo!INICPROGFERIAS1, 7, 4) & Mid$(Funcionarios_Ferias_Histo!INICPROGFERIAS1, 4, 2) & Mid$(Funcionarios_Ferias_Histo!INICPROGFERIAS1, 1, 2) 'DTSITINICIAL
'               cFields.Add Mid$(Funcionarios_Ferias_Histo!FIMPROGFERIAS1, 7, 4) & Mid$(Funcionarios_Ferias_Histo!FIMPROGFERIAS1, 4, 2) & Mid$(Funcionarios_Ferias_Histo!FIMPROGFERIAS1, 1, 2) 'DTSITFINAL
'            End If
'            cFields.Add "" 'CRED_PROV
'            cFields.Add "" 'TOK_NUMEROSERIE
'
'            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 13)
            Set cFields = New Collection
            cFields.Add Funcionarios_Ferias_Histo!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add nsStatusMsg 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAIS  ***************
            Set cFields = New Collection
            cFields.Add Funcionarios_Ferias_Histo!CHAPA 'CRED_NUMERO
            cFields.Add "" 'CRED_FLAGREENTRADA
            cFields.Add "" 'CRED_AREANOMOMENTO
            cFields.Add "" 'CRED_DTVALIDINICIAL
            cFields.Add "" 'CRED_DTVALIDFINAL
            If nsStatusMsg = 8 Then
               cFields.Add "1" 'CRED_BLOQUEADA
            Else
               cFields.Add "0" 'CRED_BLOQUEADA
            End If
            cFields.Add "" 'CRED_FLAGCATRACA
            cFields.Add "" 'CRTI_NUMERO
            cFields.Add "" 'CRED_MASTER
            cFields.Add "" 'CRED_ULTPASSAGEM
            cFields.Add "" 'CRED_EQPINUMERO
            cFields.Add "" 'CRED_SENHA
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'CRED_PROVDEVOLV
            cFields.Add "" 'CRED_NUMRA
            
            Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
            
            Rem ************************************************************************************************
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Ferias"
            sCodFun = Funcionarios_Ferias_Histo!CHAPA
            If nsStatusMsg = 8 Then
               sMsg = "Retorno de Ferias Periodo de " & Funcionarios_Ferias_Histo!INICPROGFERIAS1 & " a " & Funcionarios_Ferias_Histo!FIMPROGFERIAS1
            Else
               sMsg = "Entrando de Ferias Periodo de " & Funcionarios_Ferias_Histo!INICPROGFERIAS1 & " a " & Funcionarios_Ferias_Histo!FIMPROGFERIAS1
            End If
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            nGravou = nGravou + 1
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Ferias_Histo.EOF Then
            Funcionarios_Ferias_Histo.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If


ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Ferias"
   sCodFun = "0000"
   sMsg = "Nao ha Ferias no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing



Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Ferias"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS EM FERIAS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Ferias(ByVal sBdRm As String, _
                                    ByVal sBdRodBel As String, _
                                    ByVal c_Dt_Ferias As String, _
                                    ByVal d_Data_Abono As Date, _
                                    ByVal n_dias As Integer) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset
Dim c_Data_Ferias As String
Dim dFIMPROGFERIAS1 As Date

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

Set cFields = New Collection
cFields.Add c_Dt_Ferias, "c_Data_Ferias"

Set Funcionarios_Ferias = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 12)

nGravou = 0
nLidos = 0

c_Data_Ferias = CDate(Format(Now(), "dd/mm/yyyy"))


If Funcionarios_Ferias.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO NAO EXISTA NO CADASTRO
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Ferias.RecordCount
   
   Funcionarios_Ferias.MoveFirst
   
   While Not Funcionarios_Ferias.EOF
   
'         If Val(Funcionarios_Ferias!CHAPA) = 1777 Then
'             MsgBox ""
'         End If

         nGrvOk = 0
         Rem aqui será comparado a data do do abono na tela para ver se antecipa o fim das ferias
         Rem com numero de dias informado na tela. No caso a data deve estar no intervalo que o funcionário
         Rem estará de ferias. Caso case o intervalo, sera antecipado o final das ferias pelo numero de dias informado.
         Rem
         dFIMPROGFERIAS1 = Funcionarios_Ferias!FIMPROGFERIAS1
         If (d_Data_Abono >= Funcionarios_Ferias!INICPROGFERIAS1) And (d_Data_Abono <= Funcionarios_Ferias!FIMPROGFERIAS1) Then
            If n_dias > 0 Then
               dFIMPROGFERIAS1 = Funcionarios_Ferias!FIMPROGFERIAS1 - n_dias
            End If
         End If
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Ferias!CHAPA
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS e PESSOAS,
         Rem CASO EXISTA, VERIFICAR SE JA FOI DEMITIDO(3) OU JA ESTEJA AFASTADO(4) CASO SIM IGNORAR REGISTRO
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         
         If rRs.RecordCount = 0 Then
            nsStatusMsg = 0
         Else
            Rem VERIFICAR O INTERVALO DO PERIODO DE FERIAS E TAMBEM
            Rem O sStatusMsg DO FUNCIONARIO, CASO SATISFACA, ENTRA NO AFASTAMENTO
            Rem NO CASO ENTRANDO DE FERIAS sStatusMsg = 4
            nsStatusMsg = 0
            
            If Funcionarios_Ferias!INICPROGFERIAS1 <= CDate(c_Data_Ferias) And _
               dFIMPROGFERIAS1 >= CDate(c_Data_Ferias) Then 'CASO A DATA DE FERIAS ESTEJA NO INTERVALO ENCONTRADO, ENTAO ESTA DE FERIAS
               If rRs!pessit_numero = 8 Then
                  nsStatusMsg = 0
               Else
                If rRs!pessit_numero = 11 Then 'SO PODERA ENTRAR DE FERIAS CASO ESTA DESBLOQUEADO, CASO ESTEJA COM BLOQUEO DE OUTRA ACAO, NADA FARA.
                   nsStatusMsg = 8
                End If
               End If
            Else
               If rRs!pessit_numero = 8 Then
                  nsStatusMsg = 11
               End If
            End If
            
            Rem VERIFICAR O INTERVALO DO PERIODO DE FERIAS E TAMBEM
            Rem O sStatusMsg DO FUNCIONARIO, NO CASO VOLTANDO DE FERIAS sStatusMsg = 2
            If dFIMPROGFERIAS1 < CDate(c_Data_Ferias) Then
               If rRs!pessit_numero = 8 Then
                  nsStatusMsg = 11
               End If
            End If
            
         End If
         
         Rem NO CASO O sStatusMsg PARA ATUALIZAR OS DADOS SERA 2 E 4 RESPECTIVAMENTE
         If nsStatusMsg = 8 Or nsStatusMsg = 11 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS********************************************
            Rem ************************************************************************************************
            
'            Rem *************  HISTORICO  ***************
'            Set cFields = New Collection
'            cFields.Add rRs!CD_PESSOA
'            cFields.Add "" 'NOME
'            cFields.Add "" 'CPF
'            cFields.Add "" 'OBSERVACAO
'            cFields.Add "" 'E-MAIL
'            cFields.Add "" 'IDENTIDADE
'            cFields.Add "" 'CENTRO DE CUSTO
'            cFields.Add "" 'FLAGREENTRADA
'            cFields.Add "" 'MASTER
'            cFields.Add "" 'FOTO
'            cFields.Add "" 'SENHA
'            cFields.Add "" 'RAMAL
'            cFields.Add "" 'CAMPOPERSO1
'            cFields.Add "" 'CAMPOPERSO2
'            cFields.Add "" 'CAMPOPERSO3
'            cFields.Add "" 'CAMPOPERSO4
'            cFields.Add "" 'CAMPOPERSO5
'            cFields.Add nsStatusMsg 'PESSIT_NUMERO
'            If nGrvOk = 2 Then
'               cFields.Add "" 'DTSITINICIAL
'               cFields.Add "" 'DTSITFINAL
'            Else
'               cFields.Add Funcionarios_Ferias!INICPROGFERIAS1 'DTSITINICIAL
'               cFields.Add Funcionarios_Ferias!FIMPROGFERIAS1 'DTSITFINAL
'            End If
'
'            cFields.Add "" 'CRED_PROV
'            cFields.Add "" 'TOK_NUMEROSERIE
'
'            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 13)
            
            Rem *************  PESSOAS  ***************
            Rem  ATUALIZACAO DO CAMPO PESSIST_NUMERO sStatusMsg = 5
            Set cFields = New Collection
            cFields.Add Funcionarios_Ferias!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add nsStatusMsg 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAIS  ***************
            Set cFields = New Collection
            cFields.Add Funcionarios_Ferias!CHAPA 'CRED_NUMERO
            cFields.Add "" 'CRED_FLAGREENTRADA
            cFields.Add "" 'CRED_AREANOMOMENTO
            cFields.Add "" 'CRED_DTVALIDINICIAL
            cFields.Add "" 'CRED_DTVALIDFINAL
            If nsStatusMsg = 8 Then
               cFields.Add "1" 'CRED_BLOQUEADA
            Else
               cFields.Add "0" 'CRED_BLOQUEADA
            End If
            cFields.Add "" 'CRED_FLAGCATRACA
            cFields.Add "" 'CRTI_NUMERO
            cFields.Add "" 'CRED_MASTER
            cFields.Add "" 'CRED_ULTPASSAGEM
            cFields.Add "" 'CRED_EQPINUMERO
            cFields.Add "" 'CRED_SENHA
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'CRED_PROVDEVOLV
            cFields.Add "" 'CRED_NUMRA
            
            Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
            
            Rem ************************************************************************************************
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Ferias"
            sCodFun = Funcionarios_Ferias!CHAPA
            If nsStatusMsg = 8 Then
               sMsg = "Retorno de Ferias Periodo de " & Funcionarios_Ferias!INICPROGFERIAS1 & " a " & Funcionarios_Ferias!FIMPROGFERIAS1
            Else
               sMsg = "Entrando de Ferias Periodo de " & Funcionarios_Ferias!INICPROGFERIAS1 & " a " & Funcionarios_Ferias!FIMPROGFERIAS1
            End If
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            nGravou = nGravou + 1
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Ferias.EOF Then
            Funcionarios_Ferias.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If


ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Ferias"
   sCodFun = "0000"
   sMsg = "Nao ha Ferias no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing



Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Ferias"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DA MUDANÇA DE SEÇÃO
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionario_Historico_Secao(ByVal sBdRm As String, _
                                            ByVal sBdRodBel As String, _
                                            ByVal c_Data_Hoje As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add c_Data_Hoje, "c_Data_Hoje"

Set Funcionario_Historico_Secao = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 14)

nGravou = 0
nLidos = 0

If Funcionario_Historico_Secao.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO EXISTA NO CADASTRO
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionario_Historico_Secao.RecordCount

   Funcionario_Historico_Secao.MoveFirst
   
   While Not Funcionario_Historico_Secao.EOF
'         If Val(Funcionario_Historico_Secao!CHAPA) = 1429 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
         
         
         Set cFields = New Collection
         cFields.Add Funcionario_Historico_Secao!CHAPA
         cFields.Add Funcionario_Historico_Secao!CODSECAO
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE ESTRUTURA OUVE MUDANCA DE SECAO NESTE FUNCIONARIO
         Rem CASO TENHA, ATUALIZAR O CTAB. PESSOAS O CAMPO EST_NUMERO COM A NOVA SECAO
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 15)
         
         If rRs.RecordCount = 0 Then 'CASO NAO ACHE NENHUM REGISTRO LER O PROXIMO
            GoTo PROXIMO_REGISTRO
         End If
         
         Set cFields = New Collection
         cFields.Add Funcionario_Historico_Secao!CHAPA
         
         Rem CRITICA NO CASO DA NÃO EXISTENCIA DO NO SETOR NO CADASTRO DA RODBEL
         If rRs!ENCONTRA_EST = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Mudança Setor"
            sCodFun = Funcionario_Historico_Secao!CHAPA
            sMsg = "Secao do RM nao Cadastada na RodBel  - " & Funcionario_Historico_Secao!CODSECAO
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 1
         
         End If
         
         Rem NO CASO DA SECAO DO FUNCIONARIO SER A MESMA, LER OUTRO REGISTRO
         If rRs!ENCONTRADO = 0 Then
            GoTo PROXIMO_REGISTRO
         End If
         
         
         If nGrvOk = 0 Then
            
            Rem ************************************************************************************************
            Rem ********* atualizar AS TABELA DE Funcionario_Historico_Secao            ********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************

            Set cFields = New Collection
            cFields.Add Funcionario_Historico_Secao!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add Funcionario_Historico_Secao!CODSECAO 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add "" 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 16)
            
            nGravou = nGravou + 1
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Mudança Setor"
            sCodFun = Funcionario_Historico_Secao!CHAPA
            sMsg = "Mudanca de Setor para " & Str(Funcionario_Historico_Secao!CODSECAO)
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionario_Historico_Secao.EOF Then
            Funcionario_Historico_Secao.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Mudança Setor"
   sCodFun = "0000"
   sMsg = "Nao ha Mudancas de Secao no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Mudança Setor"
sCodFun = Funcionario_Historico_Secao!CHAPA
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS VERIFICACAO NO GED BATIDAS E CRIAR NOVA
'   SOLICITACAO CASO NÃO TENHA
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Atualizar_Horas_Extras_Ged(ByVal sBdRm As String, _
                                           ByVal sBdGed As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectGed As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset
Dim rRsWork As ADODB.Recordset
Dim rRsGravar As ADODB.Recordset
Dim sDataMov As String
Dim sDataUlt As String
Dim nNumSolicitacao As Double

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectGed = New ADODB.Connection
Set cConect = New daAbertura


Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectGed = cConect.Coneccao_OLEDB(sBdGed, "A")

Set cFields = New Collection
cFields.Add "GED_TAB_ITEM_SOLICITACAO"
cFields.Add "ULT_DATA_HEXTRA_REALIZADA"
cFields.Add "0"

Rem Pegar a ultima data do registro atualizado no ged
Set Atualizar_Horas_Extras_Ged = RegBanco_dados(ADOConnectGed, cFields, 21)

If Atualizar_Horas_Extras_Ged.RecordCount = 0 Or Len(Trim(Atualizar_Horas_Extras_Ged!TPR_CONTEUDO)) = 0 Then
   sStatusMsg = "1"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Horas Extras Realizadas"
   sCodFun = "0001"
   sMsg = "Data do Parametro nao encontrado no GED."
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   Call Gerar_Situacao_Log(cFields)
   GoTo SAIDA
End If

Rem GUARDAR A DATA DA ULTIMA ATUALIZACAO, FORMATO YYYY-MM-DD
sDataMov = DateAdd("d", -(Val(frmNotfIc!txt_Dias_Antecedencia.Text)), Atualizar_Horas_Extras_Ged!TPR_CONTEUDO)
sDataUlt = "01/01/2014"

Rem atualizacao em 29-03-2012, deletar os registros de solicitacao e itens. Registros selecionados com a data vinda do parametro
Set cFields = New Collection

cFields.Add "1", "COLIGADA"
cFields.Add Format(sDataMov, "YYYYMMDD"), "sDataMov"


Set Atualizar_Horas_Extras_Ged = RegBanco_dados(ADOConnectGed, cFields, 29)


Rem Verificar no Rm se existe movimentação a partir da data recebida do ged.
Set cFields = New Collection
cFields.Add "1", "COLIGADA"
cFields.Add Format(sDataMov, "YYYY-MM-DD"), "sDataMov"
cFields.Add "", "CHAPA"

Set Atualizar_Horas_Extras_Ged = RegBanco_dados(ADOConnectRm, cFields, 22)


nGravou = 0
nLidos = 0

' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO GED AS HORAS REALIZADAS
If Atualizar_Horas_Extras_Ged.RecordCount > 0 Then
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Atualizar_Horas_Extras_Ged.RecordCount

   Atualizar_Horas_Extras_Ged.MoveFirst
   
   While Not Atualizar_Horas_Extras_Ged.EOF
   
'         If Val(Atualizar_Horas_Extras_Ged!CHAPA) = 3417 Then
'             nGrvOk = 0
'         Else
'             GoTo PROXIMO_REGISTRO
'         End If
         
         nGrvOk = 0
         
         Rem critica para saber se existe h.extra, tem que ter torerancoa de até 10 minutos
         If Atualizar_Horas_Extras_Ged!NUMHORAS < 11 Then
            GoTo PROXIMO_REGISTRO
         End If
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE SOLICITACAO A EXISTENCIA DESTE FUNCIONARIO
         Rem CASO TENHA, ATUALIZAR O REGISTRO DA SOLICITACAO COM AS HORAS REALIZADAS
         Set cFields = New Collection
         cFields.Add Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD"), "sDataMov"
         cFields.Add "1", "COLIGADA"
         cFields.Add Atualizar_Horas_Extras_Ged!CHAPA 'Atualizar_Horas_Extras_Ged!CHAPA
         cFields.Add DateAdd("n", Atualizar_Horas_Extras_Ged!NUMHORAS, "00:00:00")
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_dados(ADOConnectGed, cFields, 23)
         
         If rRs.Fields(0) = 0 Then 'CASO NAO ACHE NENHUM REGISTRO CRIAR UMA NOVA SOLICITACAO DESTE FUNCIONARIO
            
            Rem com o setor pegar no workflow, os funcionarios que liberam a solitacao
            Set cFields = New Collection
            cFields.Add "1", "COLIGADA"
            cFields.Add Atualizar_Horas_Extras_Ged!INTEGRGERENCIAL 'setor do funcionario
            Set rRsWork = New ADODB.Recordset
            Set rRsWork = RegBanco_dados(ADOConnectGed, cFields, 26)
            
            If rRsWork.RecordCount = 0 Then
               sStatusMsg = "1"
               sData = Format(Now(), "dd/mm/yyyy")
               sHora = Format(Now(), "HH:MM")
               sTipo = "Horas Realizadas sem Solicitar"
               sCodFun = Atualizar_Horas_Extras_Ged!CHAPA
               sMsg = "WorkFlow do funcionario não cadastrado - " & Atualizar_Horas_Extras_Ged!INTEGRGERENCIAL
               Set cFields = New Collection
               cFields.Add sStatusMsg & ";" & _
                                  sData & ";" & _
                                  sHora & ";" & _
                                  sCodFun & ";" & _
                                  sTipo & ";" & _
                                  sMsg
               Call Gerar_Situacao_Log(cFields)
            Else
               Rem gravar a solicitacao do funcionario sem ser requisitada
               Set cFields = New Collection
               
               cFields.Add "-1", "SOL_CODIGO_TREINAMENTO"
               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & "'", "SOL_DATA_EMISSAO"
               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & "'", "SOL_DATA_LIBERACAO_NIVEL1"
               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & "'", "SOL_DATA_LIBERACAO_NIVEL2"
               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & "'", "SOL_DATA_LIBERACAO_NIVEL3"
               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & "'", "SOL_DATA_LIBERACAO_NIVEL4"
               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & "'", "SOL_DATA_REALIZACAO_TRABALHO_EXTRA"
               cFields.Add "NULL", "SOL_IS_FIRST"
               cFields.Add "'" & "Hora Extra sem planejamento" & "'", "SOL_JUSTIFICATIVA_MOTIVO"
               cFields.Add "'" & Format(DateAdd("m", -1, Atualizar_Horas_Extras_Ged!Data), "MM/YYYY") & "'", "SOL_MES_ANO"
               cFields.Add "NULL", "SOL_NOME_TREINAMENTO"
               cFields.Add "'" & "60%" & "'", "SOL_PERCENTUAL_HORA"
               cFields.Add "'" & "F" & "'", "SOL_SITUACAO"
               cFields.Add "1", "SOL_EMP_CODIGO"
               cFields.Add "1", "SOL_TFF_CODIGO"
               cFields.Add "60", "SOL_TMS_CODIGO"
               cFields.Add "'" & Atualizar_Horas_Extras_Ged!INTEGRGERENCIAL & "'", "SOL_TSE_CODIGO"
               cFields.Add IIf(IsNull(rRsWork!TAW_USU_NIVEL1), "Null", rRsWork!TAW_USU_NIVEL1), "TAW_USU_NIVEL1"
               cFields.Add IIf(IsNull(rRsWork!TAW_USU_NIVEL2), "Null", rRsWork!TAW_USU_NIVEL2), "TAW_USU_NIVEL2"
               cFields.Add IIf(IsNull(rRsWork!TAW_USU_NIVEL3), "Null", rRsWork!TAW_USU_NIVEL3), "TAW_USU_NIVEL3"
               cFields.Add IIf(IsNull(rRsWork!TAW_USU_NIVEL4), "Null", rRsWork!TAW_USU_NIVEL4), "TAW_USU_NIVEL4"
               cFields.Add "1", "SOL_USU_CODIGO_SOLICITANTE"
               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & " " & DateAdd("n", 0, "00:00:00") & "'", "SOL_HORA_EMISSAO"
'               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & " " & DateAdd("n", Atualizar_Horas_Extras_Ged!NUMHORAS, "00:00:00") & "'", "SOL_HORA_EMISSAO"
               cFields.Add "'" & "S" & "'", "SOL_IS_RETARDATARIO"
               Set rRsGravar = New ADODB.Recordset
               Set rRsGravar = RegBanco_dados(ADOConnectGed, cFields, 27)
               
               Rem gravar o item desta solicitacao que veio do registro da solicitação.
               Set cFields = New Collection
               cFields.Add "NULL", "ITS_ADIANTAMENTO_DECIMO_TERCEIRO"
               cFields.Add "'" & Atualizar_Horas_Extras_Ged!CHAPA & "'", "ITS_USU_CHAPA"
               cFields.Add "NULL", "ITS_CHAPA_DESAUTORIZADOR"
               cFields.Add "NULL", "ITS_DATA_INICIO_FERIAS"
               cFields.Add "NULL", "ITS_DATA_FIM_FERIAS"
               cFields.Add "'" & Atualizar_Horas_Extras_Ged!NOME_FUNCAO & "'", "ITS_USU_FUNCAO"
               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & " 00:00:00" & "'", "ITS_HORA_EXTRA_ENTRADA"
               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & " 00:00:00" & "'", "ITS_HORA_EXTRA_SAIDA"
'               cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & " " & DateAdd("n", Atualizar_Horas_Extras_Ged!NUMHORAS, "00:00:00") & "'", "ITS_HORA_EXTRA_SAIDA"
               cFields.Add "'" & "Hora Extra sem planejamento" & "'", "ITS_JUSTIFICATIVA"
               cFields.Add "'" & "N" & "'", "ITS_LANCHE"
               cFields.Add "'" & Atualizar_Horas_Extras_Ged!NOME_FUNCIONARIO & "'", "ITS_USU_NOME"
               cFields.Add "'" & "N" & "'", "ITS_REFEICAO"
               If Atualizar_Horas_Extras_Ged!NUMHORAS = 1440 Then
                  cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & " " & "00:00:00" & "'", "ITS_HORA_EXTRA_REALIZADA"
               Else
                  cFields.Add "'" & Format(Atualizar_Horas_Extras_Ged!Data, "YYYY-MM-DD") & " " & DateAdd("n", Atualizar_Horas_Extras_Ged!NUMHORAS, "00:00:00") & "'", "ITS_HORA_EXTRA_REALIZADA"
               End If
               cFields.Add "'1º Turno'", "ITS_USU_TURNO"
               cFields.Add "NULL", "ITS_CONDICAO_FERIAS"
               nNumSolicitacao = rRsGravar!SOL_CODIGO
               cFields.Add nNumSolicitacao, "ITS_SOL_CODIGO"
               cFields.Add "'N'", "ITS_CEIA"
               cFields.Add "'N'", "ITS_DESJEJUM"
               cFields.Add "'N'", "ITS_JANTA"
               
               cFields.Add "0", "ITS_DIFERENCA_SAIDA_ENTRADA"
               cFields.Add "' '", "ITS_BAIRRO"
               cFields.Add "' '", "ITS_CIDADE"
               cFields.Add "' '", "ITS_ENDERECO"
               
               Set rRsGravar = New ADODB.Recordset
               Set rRsGravar = RegBanco_dados(ADOConnectGed, cFields, 28)
            
               sStatusMsg = "1"
               sData = Format(Now(), "dd/mm/yyyy")
               sHora = Format(Now(), "HH:MM")
               sTipo = "Horas Realizadas sem Solicitar"
               sCodFun = Atualizar_Horas_Extras_Ged!CHAPA
               sMsg = "Incluido Solicitação do func. no GED Setor - " & Atualizar_Horas_Extras_Ged!INTEGRGERENCIAL
               Set cFields = New Collection
               cFields.Add sStatusMsg & ";" & _
                                  sData & ";" & _
                                  sHora & ";" & _
                                  sCodFun & ";" & _
                                  sTipo & ";" & _
                                  sMsg
               Call Gerar_Situacao_Log(cFields)
            
            End If
            nGrvOk = 1
         
         End If
         
         If nGrvOk = 0 Then
            Set rRs = New ADODB.Recordset
            Set rRs = RegBanco_dados(ADOConnectGed, cFields, 24)
   
           Rem guardar a ultima data que houve atualizacao de ponto realizado
            If CDate(sDataUlt) < Atualizar_Horas_Extras_Ged!Data Then
               sDataUlt = Format(Atualizar_Horas_Extras_Ged!Data, "DD/MM/YYYY")
            End If
            nGravou = nGravou + 1
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Horas Extras Realizadas"
            sCodFun = Atualizar_Horas_Extras_Ged!CHAPA
            sMsg = "Atualizado," & DateAdd("n", Atualizar_Horas_Extras_Ged!NUMHORAS, "00:00") & " H.Extras em " & Format(cFields(1), "DD-MM-YYYY")
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectGed.CommitTrans
            ADOConnectGed.BeginTrans
            
         End If

PROXIMO_REGISTRO:
         
         If Not Atualizar_Horas_Extras_Ged.EOF Then
            Atualizar_Horas_Extras_Ged.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend
    
   Rem atualizar o registro de paramentros no GED com a ultima data da atualizacao das H.extras
   Set cFields = New Collection
   cFields.Add "GED_TAB_ITEM_SOLICITACAO"
   cFields.Add "ULT_DATA_HEXTRA_REALIZADA"
   cFields.Add "0"
   cFields.Add sDataUlt
   
   Set Atualizar_Horas_Extras_Ged = RegBanco_dados(ADOConnectGed, cFields, 25)
   
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Ultimo dia da Atu.GED"
   sCodFun = "00001"
   sMsg = "Atualizado GED," & sDataUlt & " H.Extras dos Funcionários."
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
   
   
End If

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Horas Extras Realizadas"
   sCodFun = "0000"
   sMsg = "Nao ha Registro de H.Extras no periodo de " & sDataMov & ", executado em " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If


SAIDA:

ADOConnectRm.CommitTrans
ADOConnectGed.CommitTrans
ADOConnectRm.Close
ADOConnectGed.Close

Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Horas Extras Realizadas"
sCodFun = "0001"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS AFASTADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Afastado(ByVal sBdRm As String, _
                                      ByVal sBdRodBel As String, _
                                      ByVal c_Data_Afastado As String, _
                                      ByVal d_Data_Abono As Date, _
                                      ByVal n_dias As Integer) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date
Dim dDTFINAL As Date

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add c_Data_Afastado, "c_Data_Afastado"

Set Funcionarios_Afastado = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 17)

nGravou = 0
nLidos = 0

If Funcionarios_Afastado.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR TABELA PESSOAS NO BANCO RODBEL CASO SATISFACA CONDICOES
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Afastado.RecordCount
   
   Funcionarios_Afastado.MoveFirst
   
   While Not Funcionarios_Afastado.EOF
         If Val(Funcionarios_Afastado!CHAPA) = 2560 Then
             nGrvOk = 0
         End If
         
         nGrvOk = 0
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Afastado!CHAPA
         
         Rem SELECIONAR O REGISTRO PESSOAS DA RODBEL PARA VERIFICAR O STATUS DO CAMPO PESSIT_NUMERO E
         Rem CRED_BLOQUEADA DA TABELA CREDENCIAIS, E TESTAR A SITUACAO QUE VEIO DA RM
         Rem PESSIT_NUMERO=11-> liberada/8-> Bloqueada PESSOA
         Rem FL_BLOQUEADA = 0-LIBERADA/1-BLOQUEADA = CREDENCIAL
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         If rRs.RecordCount = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Afastados"
            sCodFun = Funcionarios_Afastado!CHAPA
            sMsg = "Tentativa de Afastar um funcionario nao Cadastrado na RodBel."
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            GoTo PROXIMO_REGISTRO
         Else
            nsStatusMsg = 0
            Rem VERIFICAR A DATA FINAL :
            Rem CASO SEJA NULL OU > QUE A DATA DE HOJE, ESTARA AFASTADO
            Rem NO CASO DA DATA FINAL < DATA DE HOJE, OUVE RETORNO
            Rem NO RESULTADO ACIMA COMPARAR O CAMPO PESSIT_NUMERO E ATUALIZAR COMFORME A CRITICA ABAIXO
            dDataCritica = CDate(Trim(Mid$(c_Data_Afastado, 8, 2) & "/" & Mid$(c_Data_Afastado, 6, 2) & "/" & Mid$(c_Data_Afastado, 2, 4)))
            
            Rem aqui será comparado a data do do abono na tela para ver se antecipa o fim das lic.renumeradas
            Rem com numero de dias informado na tela. No caso a data deve estar no intervalo que o funcionário
            Rem estará de Licença. Caso esteja no intervalo, sera antecipado o final das da Licença Renumerada pelo numero de dias informado.
            Rem
            If Not IsNull(Funcionarios_Afastado!DTFINAL) Then
               dDTFINAL = Funcionarios_Afastado!DTFINAL
               If (d_Data_Abono >= Funcionarios_Afastado!dtinicio) And _
                  (d_Data_Abono <= Funcionarios_Afastado!DTFINAL) Then
                  If n_dias > 0 And Funcionarios_Afastado!TIPO = "R" Then
                     dDTFINAL = Funcionarios_Afastado!DTFINAL - n_dias
                  End If
               End If
            End If
            
            If IsNull(Funcionarios_Afastado!DTFINAL) Then
               If rRs!pessit_numero = 11 Then 'CASO JA ESTEJA COM A OPCAO 11 , BLOQUEA-LO
                  nsStatusMsg = 8
               Else
                  If rRs!pessit_numero = 8 Then 'CASO ESTEJA LIBERADO, SERA BLOQUADO COMA OPCAO 5
                     nsStatusMsg = 0
                  End If
               End If
           ElseIf (dDTFINAL > dDataCritica And _
                   CDate(Format(Mid$(Funcionarios_Afastado!dtinicio, 1, 10), "DD/MM/YYYY")) <= dDataCritica) Then
               If rRs!pessit_numero = 11 Then 'CASO JA ESTEJA COM A OPCAO 11 , JA ESTA BLOQUEADO
                  nsStatusMsg = 0
               Else
                  If rRs!pessit_numero = 8 Then 'CASO ESTEJA LIBERADO, SERA BLOQUADO COMA OPCAO 5
                     nsStatusMsg = 11
                  End If
               End If
            ElseIf dDTFINAL < dDataCritica Then
               If rRs!pessit_numero = 8 Then
                  nsStatusMsg = 11
               Else
                  If (rRs!pessit_numero = 11) Then 'CASO ESTEJA EM BLOQUEIO DE AFASTAMENTO, SERA LIBERADO COM A OPCAO 8
                     nsStatusMsg = 8
                  End If
'                  If (rRs!pessit_numero = 12) Then 'CASO ESTEJA EM BLOQUEIO DE AFASTAMENTO, SERA LIBERADO COM A OPCAO 11
'                     nsStatusMsg = 8
'                  End If
               End If
            End If
         
         End If
         
         Rem NO CASO O sStatusMsg PARA ATUALIZAR OS DADOS SERRA 2 E 5 RESPECTIVAMENTE
         If nsStatusMsg = 8 Or nsStatusMsg = 11 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            Rem  ATUALIZACAO DO CAMPO PESSIST_NUMERO sStatusMsg = 5
            Set cFields = New Collection
            cFields.Add Funcionarios_Afastado!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add nsStatusMsg 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAIS  ***************
            Set cFields = New Collection
            cFields.Add Funcionarios_Afastado!CHAPA 'CRED_NUMERO
            cFields.Add "" 'CRED_FLAGREENTRADA
            cFields.Add "" 'CRED_AREANOMOMENTO
            cFields.Add "" 'CRED_DTVALIDINICIAL
            cFields.Add "" 'CRED_DTVALIDFINAL
            If nsStatusMsg = 8 Then
               cFields.Add "1" 'CRED_BLOQUEADA
            Else
               cFields.Add "0" 'CRED_BLOQUEADA
            End If
            cFields.Add "" 'CRED_FLAGCATRACA
            cFields.Add "" 'CRTI_NUMERO
            cFields.Add "" 'CRED_MASTER
            cFields.Add "" 'CRED_ULTPASSAGEM
            cFields.Add "" 'CRED_EQPINUMERO
            cFields.Add "" 'CRED_SENHA
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'CRED_PROVDEVOLV
            cFields.Add "" 'CRED_NUMRA
            
            Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
            
            Rem ************************************************************************************************
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Afastados"
            sCodFun = Funcionarios_Afastado!CHAPA
            If nsStatusMsg = 11 Then
               sMsg = "Retorno de Afastamento Periodo de " & Funcionarios_Afastado!dtinicio & " a " & Funcionarios_Afastado!DTFINAL
            Else
               sMsg = "Afastando no Inicio de " & Funcionarios_Afastado!dtinicio & IIf(IsNull(Funcionarios_Afastado!DTFINAL), " ", " a " & Funcionarios_Afastado!DTFINAL)
            End If
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            nGravou = nGravou + 1
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Afastado.EOF Then
            Funcionarios_Afastado.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If


ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Afastados"
   sCodFun = "0000"
   sMsg = "Nao ha Afastamento no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Afastados"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS DE ATESTADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Atestado(ByVal sBdRm As String, _
                                      ByVal sBdRodBel As String, _
                                      ByVal c_Data_Atestado As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add c_Data_Atestado, "c_Data_Atestado"

Set Funcionarios_Atestado = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 20)

nGravou = 0
nLidos = 0

If Funcionarios_Atestado.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR TABELA PESSOAS NO BANCO RODBEL CASO SATISFACA CONDICOES
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Atestado.RecordCount
   
   Funcionarios_Atestado.MoveFirst
   
   While Not Funcionarios_Atestado.EOF
'         If Val(Funcionarios_Atestado!CHAPA) = 2541 Then
'             nGrvOk = 0
'         Else
'             GoTo PROXIMO_REGISTRO
'         End If
         
         nGrvOk = 0
         
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Atestado!CHAPA
         Rem SELECIONAR O REGISTRO PESSOAS DA RODBEL PARA VERIFICAR O STATUS DO CAMPO PESSIT_NUMERO E
         Rem CRED_BLOQUEADA DA TABELA CREDENCIAIS, E TESTAR A SITUACAO QUE VEIO DA RM
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         If rRs.RecordCount = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Atestados"
            sCodFun = Funcionarios_Atestado!CHAPA
            sMsg = "Tentativa de Afastar um funcionario nao cadastrado na RodBel."
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            GoTo PROXIMO_REGISTRO
         Else
            nsStatusMsg = 0
            Rem VERIFICAR O A DATA FINAL :
            Rem CASO SEJA NULL OU > QUA A DATA DE HOJE, ESTARA Atestado
            Rem NO CASO DA DATA FINAL < DATA DE HOJE, OUVE RETORNO
            Rem NO RESULTADO ACIMA COMPARAR O CAMPO PESSIT_NUMERO E ATUALIZAR COMFORME A CRITICA ABAIXO
            dDataCritica = CDate(Trim(Mid$(c_Data_Atestado, 8, 2) & "/" & Mid$(c_Data_Atestado, 6, 2) & "/" & Mid$(c_Data_Atestado, 2, 4)))
            
            If IsNull(Funcionarios_Atestado!DTFINAL) Then
               If Funcionarios_Atestado!dtinicio <= dDataCritica Then
                  If rRs!pessit_numero = 11 Then 'CASO JA ESTEJA COM A OPCAO 12 , JA ESTA BLOQUEADO
                     nsStatusMsg = 0
                  Else
                     If rRs!pessit_numero = 8 Then 'CASO ESTEJA LIBERADO, SERA BLOQUEADO COMA OPCAO 8
                        nsStatusMsg = 11
                     End If
                  End If
               Else
                  If rRs!pessit_numero = 11 Then 'CASO JA ESTEJA COM A OPCAO 12 , JA ESTA BLOQUEADO
                     nsStatusMsg = 8
                  End If
               End If
            Else
               If CDate(Mid$(Funcionarios_Atestado!dtinicio, 1, 10)) <= dDataCritica And _
                                       Funcionarios_Atestado!DTFINAL >= dDataCritica Then
                  If rRs!pessit_numero = 11 Then 'CASO JA ESTEJA COM A OPCAO 12 , JA ESTA BLOQUEADO
                     nsStatusMsg = 0
                  Else
                     If rRs!pessit_numero = 8 Then 'CASO ESTEJA LIBERADO, SERA BLOQUEADO COMA OPCAO 8
                        nsStatusMsg = 11
                     End If
                  End If
               Else
                  If Funcionarios_Atestado!DTFINAL < dDataCritica Then
                     If rRs!pessit_numero = 11 Then
                        nsStatusMsg = 0
                     Else
                        If rRs!pessit_numero = 8 Then 'CASO ESTEJA EM BLOQUEIO DE Atestado, SERA LIBERADO COM A OPCAO 2
                           nsStatusMsg = 11
                        End If
                     End If
                  End If
               End If
            End If
            
'            If IsNull(Funcionarios_Atestado!DTFINAL) Or _
'               Funcionarios_Atestado!DTFINAL > dDataCritica Then
'               If rRs!pessit_numero = 12 Then 'CASO JA ESTEJA COM A OPCAO 12 , JA ESTA BLOQUEADO
'                  nsStatusMsg = 0
'               Else
'                  If rRs!pessit_numero = 8 Then 'CASO ESTEJA LIBERADO, SERA BLOQUEADO COMA OPCAO 8
'                     nsStatusMsg = 12
'                  End If
'               End If
'            ElseIf Funcionarios_Atestado!DTFINAL < dDataCritica Then
'               If rRs!pessit_numero = 8 Then
'                  nsStatusMsg = 0
'               Else
'                  If rRs!pessit_numero = 12 Then 'CASO ESTEJA EM BLOQUEIO DE Atestado, SERA LIBERADO COM A OPCAO 2
'                     nsStatusMsg = 8
'                  End If
'               End If
'            End If
         
         End If
         
         Rem NO CASO O sStatusMsg PARA ATUALIZAR OS DADOS SERRA 2 E 4 RESPECTIVAMENTE
         If nsStatusMsg = 8 Or nsStatusMsg = 11 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            Rem  ATUALIZACAO DO CAMPO PESSIST_NUMERO sStatusMsg = 6
            Set cFields = New Collection
            cFields.Add Funcionarios_Atestado!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add nsStatusMsg 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAIS  ***************
            Set cFields = New Collection
            cFields.Add Funcionarios_Atestado!CHAPA 'CRED_NUMERO
            cFields.Add "" 'CRED_FLAGREENTRADA
            cFields.Add "" 'CRED_AREANOMOMENTO
            cFields.Add "" 'CRED_DTVALIDINICIAL
            cFields.Add "" 'CRED_DTVALIDFINAL
            If nsStatusMsg = 8 Then
               cFields.Add "1" 'CRED_BLOQUEADA
            Else
               cFields.Add "0" 'CRED_BLOQUEADA
            End If
            cFields.Add "" 'CRED_FLAGCATRACA
            cFields.Add "" 'CRTI_NUMERO
            cFields.Add "" 'CRED_MASTER
            cFields.Add "" 'CRED_ULTPASSAGEM
            cFields.Add "" 'CRED_EQPINUMERO
            cFields.Add "" 'CRED_SENHA
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'CRED_PROVDEVOLV
            cFields.Add "" 'CRED_NUMRA
            
            Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
            
            Rem ************************************************************************************************
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Atestados"
            sCodFun = Funcionarios_Atestado!CHAPA
            If nsStatusMsg = 8 Then
               sMsg = "Retorno de Atestado Periodo de " & Funcionarios_Atestado!dtinicio & " a " & Funcionarios_Atestado!DTFINAL
            Else
               sMsg = "Atestado no Inicio de " & Funcionarios_Atestado!dtinicio & IIf(IsNull(Funcionarios_Atestado!DTFINAL), " ", " a " & Funcionarios_Atestado!DTFINAL)
            End If
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            nGravou = nGravou + 1
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Atestado.EOF Then
            Funcionarios_Atestado.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If


ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Atestados"
   sCodFun = "0000"
   sMsg = "Nao ha Atestado no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing



Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Atestados"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS ATUALIZAR O PELO CADASTRO DE FUNCIONARIOS DO RM OS VALES NO SISTEMA DO VALE TRANSPORTE
'   E AS SOLICITACOES DO GED QUE TENHAM VALES A PAGAR E CRIAR AS SOLICITACOES NO SIST. VT.
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Atualizar_Vale_Transporte(ByVal sBdRm As String, _
                                          ByVal sBdGed As String, _
                                          ByVal sBdVt, _
                                          ByVal c_Data_Admissao As String) As ADODB.Recordset
On Error GoTo Erro


Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectVt As ADODB.Connection
Dim ADOConnectGed As ADODB.Connection

Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset
Dim rRsWork As ADODB.Recordset
Dim rRsGravar As ADODB.Recordset
Dim rRsWork_FUN As ADODB.Recordset
Dim sDataMov As String
Dim sDataUlt As String
Dim nNumSolicitacao As Double
Dim sCodigoHorario As String
Dim stxt_feriado As String
Dim nQtd1 As Double
Dim nVt1 As String
Dim nQtd2 As Double
Dim nVt2 As String

On Error GoTo Erro

Set ADOConnectRm = New ADODB.Connection
Set ADOConnectGed = New ADODB.Connection
Set ADOConnectVt = New ADODB.Connection
Set cConect = New daAbertura


Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectGed = cConect.Coneccao_OLEDB(sBdGed, "A")
Set ADOConnectVt = cConect.Coneccao_OLEDB(sBdVt, "A")

Rem GUARDAR A DATA DA ULTIMA ATUALIZACAO, FORMATO YYYY-MM-DD
sDataMov = c_Data_Admissao
nGravou = 0
nLidos = 0


'Atualizacao para os funcionario que nao pegam vale nem onibus da empresa
'***************************************************************
' leitura de todos os funcionarios que não estejam demitidos
'***************************************************************
    
Rem Consultar CADASTRO DE PESSOAL VERIFICANDO OS QUE NAO ESTEJAM DEMITIDOS
Set cFields = New Collection

cFields.Add "1", "COLIGADA"
cFields.Add Format(sDataMov, "YYYY-MM-DD"), "sDataMov"

Set Atualizar_Vale_Transporte = RegBanco_dados(ADOConnectRm, cFields, 30)

    
Atualizar_Vale_Transporte.MoveFirst
    
nGravou = 0
nLidos = 0
frmNotfIc!Pr_Prog.Visible = True
frmNotfIc!Pr_Prog.Min = 0
frmNotfIc!Pr_Prog.Max = Atualizar_Vale_Transporte.RecordCount

    
While Not Atualizar_Vale_Transporte.EOF
'***************************************************************
' inclui os funcionarios que não existam no cadastro de cadvale (zerado)
'***************************************************************
    Set rRsWork = New ADODB.Recordset
    
    Set cFields = New Collection
    cFields.Add Atualizar_Vale_Transporte!CODCOLIGADA, "COLIGADA"
    cFields.Add Atualizar_Vale_Transporte!CHAPA, "Chapa"
    
    Set rRsWork = RegBanco_dados(ADOConnectVt, cFields, 31)
    
    If rRsWork.RecordCount = 0 Then
        Set cFields = New Collection
        cFields.Add Atualizar_Vale_Transporte!CODCOLIGADA
        cFields.Add Atualizar_Vale_Transporte!CHAPA
        cFields.Add "'*'"
        cFields.Add 0
        cFields.Add "'*'"
        cFields.Add 0
        cFields.Add "'*'"
        cFields.Add 0
        cFields.Add "'*'"
        cFields.Add 0
        Call RegBanco_dados(ADOConnectVt, cFields, 32)
    
        sStatusMsg = "0"
        sData = Format(Now(), "dd/mm/yyyy")
        sHora = Format(Now(), "HH:MM")
        sTipo = "ValeTransporte"
        sCodFun = Atualizar_Vale_Transporte!CHAPA
        sMsg = "Incluido novo cadastro de funcionario no Vale Transporte"
        
        Set cFields = New Collection
        cFields.Add sStatusMsg & ";" & _
                           sData & ";" & _
                           sHora & ";" & _
                           sCodFun & ";" & _
                           sTipo & ";" & _
                           sMsg
        
        Call Gerar_Situacao_Log(cFields)
     
        ADOConnectVt.CommitTrans
        ADOConnectVt.BeginTrans
        nGravou = nGravou + 1
    
    End If
    Atualizar_Vale_Transporte.MoveNext
    nLidos = nLidos + 1
    frmNotfIc!Pr_Prog.Value = nLidos
    
Wend
        
'***************************************************************
' leitura da tabela fixa no RM dos funcionarios que recebem vales opcionais ref a verba 0532
'***************************************************************
Set cFields = New Collection
cFields.Add "0532", "Codevento"

Set rRsWork = RegBanco_dados(ADOConnectRm, cFields, 33)

rRsWork.MoveFirst
'Atualizacao das pessoas que pegam o onibus da empresa
'***************************************************************
' Atualizar o cadvale , apenas os que foram inclusos e os que nao recebem e tornam a receber
'***************************************************************
nLidos = 0
frmNotfIc!Pr_Prog.Visible = True
frmNotfIc!Pr_Prog.Min = 0
frmNotfIc!Pr_Prog.Max = rRsWork.RecordCount
    


While Not rRsWork.EOF
    'Checa se já existe em CadVale
    
    Set cFields = New Collection
    cFields.Add rRsWork!CODCOLIGADA, "COLIGADA"
    cFields.Add rRsWork!CHAPA, "Chapa"
    
    Set Atualizar_Vale_Transporte = RegBanco_dados(ADOConnectVt, cFields, 34)
    
    If Atualizar_Vale_Transporte.RecordCount = 1 Then
       Set cFields = New Collection
       cFields.Add rRsWork!CODCOLIGADA
       cFields.Add rRsWork!CHAPA
       cFields.Add "'B'"
       cFields.Add 1
       cFields.Add "'*'"
       cFields.Add 0
       cFields.Add "'B'"
       cFields.Add 2
       cFields.Add "'*'"
       cFields.Add 0
       
       Call RegBanco_dados(ADOConnectVt, cFields, 35)
   
       sStatusMsg = "0"
       sData = Format(Now(), "dd/mm/yyyy")
       sHora = Format(Now(), "HH:MM")
       sTipo = "ValeTransporte"
       sCodFun = rRsWork!CHAPA
       sMsg = "Atualizado VT-532 no cadastro de funcionario no Vale Transporte"
       
       Set cFields = New Collection
       cFields.Add sStatusMsg & ";" & _
                          sData & ";" & _
                          sHora & ";" & _
                          sCodFun & ";" & _
                          sTipo & ";" & _
                          sMsg
       
       Call Gerar_Situacao_Log(cFields)
    
       ADOConnectVt.CommitTrans
       ADOConnectVt.BeginTrans
       nGravou = nGravou + 1
    End If
    
    rRsWork.MoveNext
    nLidos = nLidos + 1
    frmNotfIc!Pr_Prog.Value = nLidos
Wend
    
    'Atualizacao das pessoas que pegam vale transporte
'***************************************************************
' Leitura da tabela de pfvaletr - RECEBER REGISTRO DA TABELA PfValeTr, P/ ATUALIZAR VALES DOS NOVATOS
'***************************************************************
Set cFields = New Collection
cFields.Add " "

Set rRsWork = RegBanco_dados(ADOConnectRm, cFields, 36)

rRsWork.MoveFirst

nGravou = 0
nLidos = 0
frmNotfIc!Pr_Prog.Visible = True
frmNotfIc!Pr_Prog.Min = 0
frmNotfIc!Pr_Prog.Max = rRsWork.RecordCount
    

'verifica se a chapa já existe no CAdVale, Se existir, atualiza, senao inclui

While Not rRsWork.EOF
    'Checa se já existe em CadVale
    
    Set cFields = New Collection
    cFields.Add rRsWork!CODCOLIGADA, "COLIGADA"
    cFields.Add rRsWork!CHAPA, "Chapa"
    
    Set rRsWork_FUN = New ADODB.Recordset
    Set rRsWork_FUN = RegBanco_dados(ADOConnectRm, cFields, 39)
    
    If rRsWork_FUN.RecordCount = 0 Then GoTo PROXIMO_REG1
    
    
    Set Atualizar_Vale_Transporte = RegBanco_dados(ADOConnectVt, cFields, 43)
    
    If Atualizar_Vale_Transporte.RecordCount = 1 Then
    
'       If rRsWork!CodLinha <> Atualizar_Vale_Transporte!TipoVale1 Then
          Set cFields = New Collection
          cFields.Add rRsWork!CODCOLIGADA
          cFields.Add rRsWork!CHAPA
          cFields.Add ""
          cFields.Add 0
          cFields.Add "'" & rRsWork!CodLinha & "'"
          cFields.Add 1
          cFields.Add ""
          cFields.Add 0
          cFields.Add "'" & rRsWork!CodLinha & "'"
          cFields.Add 2
          
          Call RegBanco_dados(ADOConnectVt, cFields, 35)
      
          sStatusMsg = "0"
          sData = Format(Now(), "dd/mm/yyyy")
          sHora = Format(Now(), "HH:MM")
          sTipo = "ValeTransporte"
          sCodFun = rRsWork!CHAPA
          sMsg = "INCLUIDO VT-Normal na tabela de cadvale no Vale Transporte"
          
          Set cFields = New Collection
          cFields.Add sStatusMsg & ";" & _
                             sData & ";" & _
                             sHora & ";" & _
                             sCodFun & ";" & _
                             sTipo & ";" & _
                             sMsg
          
          Call Gerar_Situacao_Log(cFields)
       
          ADOConnectVt.CommitTrans
          ADOConnectVt.BeginTrans
          nGravou = nGravou + 1
       End If
'    End If

PROXIMO_REG1:

    rRsWork.MoveNext
    nLidos = nLidos + 1
    frmNotfIc!Pr_Prog.Value = nLidos

Wend


'***************************************************************
' Leitura da tabela de SOLICITACOES, QUE CONTEM AS SOLICITACOES PENDENTES DE RECEBEREM VT
'***************************************************************
Set cFields = New Collection
cFields.Add Mid$(sDataMov, 2, 4) & "-" & Mid$(sDataMov, 6, 2) & "-" & Mid$(sDataMov, 8, 2), "sDataMov"

Set Atualizar_Vale_Transporte = RegBanco_dados(ADOConnectGed, cFields, 37)

rRsWork.MoveFirst

nGravou = 0
nLidos = 0

frmNotfIc!Pr_Prog.Visible = True
frmNotfIc!Pr_Prog.Min = 0
frmNotfIc!Pr_Prog.Max = Atualizar_Vale_Transporte.RecordCount

'verifica se a chapa já existe no CAdVale, Se existir, atualiza, senao inclui

While Not Atualizar_Vale_Transporte.EOF
    'Checa se já existe NO MOVIMENTO MOVVALE
    
'      If Val(Atualizar_Vale_Transporte!ITS_USU_CHAPA) = 382 Then
'         MsgBox ""
'      End If
      Set cFields = New Collection
      cFields.Add Atualizar_Vale_Transporte!SOL_EMP_CODIGO, "COLIGADA"
      cFields.Add Atualizar_Vale_Transporte!ITS_USU_CHAPA, "Chapa"
      cFields.Add Format(Atualizar_Vale_Transporte!SOL_DATA_REALIZACAO_TRABALHO_EXTRA, "YYYY-MM-DD"), "SOL_DATA_REALIZACAO_TRABALHO_EXTRA"
      
      Set rRsWork = RegBanco_dados(ADOConnectVt, cFields, 38)
      
      
      Rem VERICAR SE JA RECEBEU VALES NO ARQUIVO DE MOVIMENTO NESTA DATA
      If rRsWork.RecordCount = 0 Then
'         stxtCODFUNC = Atualizar_Vale_Transporte!ITS_USU_CHAPA
         
         ' Verificar a situação do funcionário, SE PODE GERAR MOVIMENTACAO DE VALE
         Set rRsWork = RegBanco_dados(ADOConnectRm, cFields, 39)
         Rem AQUI MARCOS 28/08/2014, FALTA DO FERIADO PARA BAIXO.
         Rem DIA
         Rem 21 - DE  9:30 - 19:00 8:30
         Rem 22 - DE 14:00 - 17:00 3:00
         Rem 25 - DE  8:30 - 19:00 9:30
         Rem 26 - DE 14:00 - 19:00 5:00
         Rem 28 - de  9:30 - 19:00 8:30
         Rem 01 - de  8:00 - 17:00 8:00
         Rem 03 - de  10:00 - 17:00 6:00
'                                   48:30 = 4.222,00
         
         If rRsWork.RecordCount > 0 Then
   'Verificando a tabela AHORARIO de e exibindo o conteudo do campo DESCRICAO(Turno)
             Set cFields = New Collection
             cFields.Add Format(Atualizar_Vale_Transporte!SOL_DATA_REALIZACAO_TRABALHO_EXTRA, "YYYY-MM-DD"), "SOL_DATA_REALIZACAO_TRABALHO_EXTRA"

             
             sCodigoHorario = rRsWork!CODHORARIO
             
             Set rRsWork = New ADODB.Recordset
             Set rRsWork = RegBanco_dados(ADOConnectVt, cFields, 41)
   
    Rem VERIFICAR O FERIADO NESTE DIA
    
'                  Set cFields = New Collection
'                  Set rRsWork = New ADODB.Recordset
'                  cFields.Add Format(Atualizar_Vale_Transporte!SOL_DATA_REALIZACAO_TRABALHO_EXTRA, "DD/MM/YYYY"), "SOL_DATA_REALIZACAO_TRABALHO_EXTRA"
'                  Set rRsWork = RegBanco_dados(ADOConnectVt, cFields, 42)
             
 'Verificação do dia da Semana para tratar os campos dos vales tranportes
    
             If rRsWork.RecordCount > 0 Then
                stxt_feriado = rRsWork!DESCRICAO
                sCodigoHorario = 10
             Else
                stxt_feriado = ""
             End If


 'Verificação cadastrado na tabela CadVale
             Set cFields = New Collection
             cFields.Add Atualizar_Vale_Transporte!SOL_EMP_CODIGO, "COLIGADA"
             cFields.Add Atualizar_Vale_Transporte!ITS_USU_CHAPA, "Chapa"
             
             Set rRsWork = RegBanco_dados(ADOConnectVt, cFields, 31)
   
             If rRsWork.RecordCount > 0 Then
                frmNotfIc!DT_Pesquisa.Value = Format(Atualizar_Vale_Transporte!SOL_DATA_REALIZACAO_TRABALHO_EXTRA, "DD/MM/YYYY")
                nQtd1 = 0
                nVt1 = ""
                nQtd2 = 0
                nVt2 = ""
             
             
             
              'Verificar a tabela de feriado para dobrar os vales
                If sCodigoHorario = 9 Or sCodigoHorario = 14 Or sCodigoHorario = 10 Or _
                   (frmNotfIc!DT_Pesquisa.DayOfWeek = 7 Or frmNotfIc!DT_Pesquisa.DayOfWeek = 1) Then
        
                       If rRsWork!TipoVale3 <> "*" Then
                           nVt1 = rRsWork!TipoVale3
                           If rRsWork!QtdVale3 = 0 Then
                              nQtd1 = 0
                           Else
                              nQtd1 = rRsWork!QtdVale3
                           End If
                       Else
                           nVt1 = "A"
                           nQtd1 = 0
                       End If
        
                       If rRsWork!QtdVale4 <> 0 And _
                          rRsWork!TipoVale4 <> "" Then
                           nVt2 = rRsWork!TipoVale4
                           nQtd2 = rRsWork!QtdVale4
                       Else
                           nVt2 = "*"
                           nQtd2 = 0
                       End If
                Else
                
                       If rRsWork!TipoVale1 <> "*" Then
                           nVt1 = rRsWork!TipoVale1
                           If rRsWork!QtdVale1 = 0 Then
                              nQtd1 = "0"
                           Else
                              nQtd1 = rRsWork!QtdVale1
                           End If
                       Else
                           nVt1 = "A"
                           nQtd1 = 0
                       End If
        
                       If rRsWork!QtdVale2 <> "" And _
                          rRsWork!TipoVale2 <> "" Then
                           nVt2 = rRsWork!TipoVale2
                           nQtd2 = rRsWork!QtdVale2
                       Else
                           nVt2 = "*"
                           nQtd2 = 0
                       End If
                       
                End If
        
                If nQtd1 > 0 Or nQtd2 > 0 Then
                   
                   Set cFields = New Collection
                   cFields.Add Atualizar_Vale_Transporte!SOL_EMP_CODIGO, "SOL_EMP_CODIGO"
                   cFields.Add Atualizar_Vale_Transporte!ITS_USU_CHAPA, "Chapa"
                   cFields.Add "'S'"
                   cFields.Add "'E'"
                   cFields.Add "'" & Format(Atualizar_Vale_Transporte!SOL_DATA_REALIZACAO_TRABALHO_EXTRA, "YYYYMMDD") & "'"
                   cFields.Add "'" & nVt1 & "'"
                   cFields.Add nQtd1
                   cFields.Add "'" & nVt2 & "'"
                   cFields.Add nQtd2
                   cFields.Add "'N'"
                   cFields.Add "Null"
                   cFields.Add "Null"
                   cFields.Add "'MAURO'"
                   cFields.Add "'MAURO'"
                   cFields.Add "Null"
                   cFields.Add "Null"
                   cFields.Add "'" & Format(Now(), "yyyy-mm-dd HH:MM:SS") & "'"
                   Set rRsWork = RegBanco_dados(ADOConnectVt, cFields, 42)
                
                   sStatusMsg = "0"
                   sData = Format(Now(), "dd/mm/yyyy")
                   sHora = Format(Now(), "HH:MM")
                   sTipo = "ValeTransporte"
                   sCodFun = Atualizar_Vale_Transporte!ITS_USU_CHAPA
                   sMsg = "Incluido VT-Ged Normal no Movvale, em " & Format(Atualizar_Vale_Transporte!SOL_DATA_REALIZACAO_TRABALHO_EXTRA, "YYYYMMDD")
                   
                   Set cFields = New Collection
                   cFields.Add sStatusMsg & ";" & _
                                      sData & ";" & _
                                      sHora & ";" & _
                                      sCodFun & ";" & _
                                      sTipo & ";" & _
                                      sMsg
                   
                   Call Gerar_Situacao_Log(cFields)
                
                   ADOConnectVt.CommitTrans
                   ADOConnectVt.BeginTrans
                   nGravou = nGravou + 1
                
                End If
             End If
         End If
      End If
      Atualizar_Vale_Transporte.MoveNext
      nLidos = nLidos + 1
      frmNotfIc!Pr_Prog.Value = nLidos
Wend

SAIDA:

ADOConnectRm.CommitTrans
ADOConnectGed.CommitTrans
ADOConnectVt.CommitTrans
ADOConnectRm.Close
ADOConnectGed.Close
ADOConnectVt.Close

Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set ADOConnectVt = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If

sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "ValeTransporte"
sCodFun = "0001"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

End Function
Public Function Atualizar_Cad_Funcao_Setor_Ged(ByVal sBdRm As String, _
                                               ByVal sBdGed As String, _
                                               ByVal c_Data_Atestado As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectGed As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date
Dim bSetor As Boolean

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectGed = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectGed = cConect.Coneccao_OLEDB(sBdGed, "A")

cFields.Add c_Data_Atestado, "c_Data_Atestado"

Set Atualizar_Cad_Funcao_Setor_Ged = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 45)

bSetor = True
nGravou = 0
nLidos = 0

If Atualizar_Cad_Funcao_Setor_Ged.RecordCount > 0 Then ' CASO ACHE ATUALIZAR OS SETORES DO GED VINDO DO RM
   
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Atualizar_Cad_Funcao_Setor_Ged.RecordCount
   
   Atualizar_Cad_Funcao_Setor_Ged.MoveFirst
   
   While Not Atualizar_Cad_Funcao_Setor_Ged.EOF
         
         nGrvOk = 0
         
         Set cFields = New Collection
         cFields.Add Atualizar_Cad_Funcao_Setor_Ged!CODCOLIGADA
         Rem Quando for da empresa 2 (Manuas), por um ponto no final na hora de consultar e incluir o registro que faltar.
         If Atualizar_Cad_Funcao_Setor_Ged!CODCOLIGADA = 1 Then
            cFields.Add Atualizar_Cad_Funcao_Setor_Ged!CODINTERNO
         Else
            cFields.Add Atualizar_Cad_Funcao_Setor_Ged!CODINTERNO & "."
         End If
         cFields.Add Atualizar_Cad_Funcao_Setor_Ged!DESCRICAO
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectGed, cFields, 46)
         
         If rRs.RecordCount = 0 Then
            
            Call RegBanco_Critica_Pessoas(ADOConnectGed, cFields, 47)
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "CadFunSet"
            sCodFun = Str(Atualizar_Cad_Funcao_Setor_Ged!CODCOLIGADA) & Atualizar_Cad_Funcao_Setor_Ged!CODINTERNO
            sMsg = "Novo setor no GED " & Atualizar_Cad_Funcao_Setor_Ged!DESCRICAO
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            ADOConnectGed.CommitTrans
            ADOConnectGed.BeginTrans
            nGravou = nGravou + 1
         End If

PROXIMO_REGISTRO:
         
         If Not Atualizar_Cad_Funcao_Setor_Ged.EOF Then
            Atualizar_Cad_Funcao_Setor_Ged.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If


Set Atualizar_Cad_Funcao_Setor_Ged = New ADODB.Recordset
Set cFields = New Collection
cFields.Add c_Data_Atestado, "c_Data_Atestado"

Set Atualizar_Cad_Funcao_Setor_Ged = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 48)

nLidos = 0
bSetor = False

If Atualizar_Cad_Funcao_Setor_Ged.RecordCount > 0 Then ' CASO ACHE ATUALIZAR AS FUNCOES DO GED VINDO DO RM
   
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Atualizar_Cad_Funcao_Setor_Ged.RecordCount
   
   Atualizar_Cad_Funcao_Setor_Ged.MoveFirst
   
   While Not Atualizar_Cad_Funcao_Setor_Ged.EOF
         
         nGrvOk = 0

         Set cFields = New Collection
         cFields.Add Atualizar_Cad_Funcao_Setor_Ged!CODCOLIGADA
         If Atualizar_Cad_Funcao_Setor_Ged!CODCOLIGADA = 1 Then
            cFields.Add Atualizar_Cad_Funcao_Setor_Ged!CODIGO
         Else
            cFields.Add Atualizar_Cad_Funcao_Setor_Ged!CODIGO & "."
         End If
         
         cFields.Add Atualizar_Cad_Funcao_Setor_Ged!NOME
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectGed, cFields, 49)
         
         If rRs.RecordCount = 0 Then
            
            Call RegBanco_Critica_Pessoas(ADOConnectGed, cFields, 50)
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "CadFunSet"
            sCodFun = Str(Atualizar_Cad_Funcao_Setor_Ged!CODCOLIGADA) & Atualizar_Cad_Funcao_Setor_Ged!CODIGO
            sMsg = "Nova Funcao no GED " & Atualizar_Cad_Funcao_Setor_Ged!NOME
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            ADOConnectGed.CommitTrans
            ADOConnectGed.BeginTrans
            nGravou = nGravou + 1
         End If

PROXIMO_REGISTROF:
         
         If Not Atualizar_Cad_Funcao_Setor_Ged.EOF Then
            Atualizar_Cad_Funcao_Setor_Ged.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If


ADOConnectRm.CommitTrans
ADOConnectGed.CommitTrans
ADOConnectRm.Close
ADOConnectGed.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "CadFunSet"
   sCodFun = "0000"
   sMsg = "Nao ha Funcao/Setor no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing



Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "CadFunSet"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear
'
'If Err.Number <> 0 Then
'   If bSetor Then
'      GoTo PROXIMO_REGISTRO
'   Else
'      GoTo PROXIMO_REGISTROF
'   End If
'End If

'ADOConnectRm.RollbackTrans
'ADOConnectGed.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Atualizar_Status_NFE(ByVal sBdNFE As String, _
                                     ByVal c_Data_Dia As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectNFE As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date
Dim bSetor As Boolean

On Error GoTo Erro


Set ADOConnectNFE = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectNFE = cConect.Coneccao_OLEDB(sBdNFE, "A")

cFields.Add c_Data_Dia, "c_Data_Dia"

Set Atualizar_Status_NFE = RegBanco_Critica_Pessoas(ADOConnectNFE, cFields, 59)

bSetor = True
nGravou = 0
nLidos = 0

If Atualizar_Status_NFE.RecordCount > 0 Then ' CASO ACHE ATUALIZAR AS NFE PARA O STATUS 99, PARA VALIDAR NA SEFAZ
   
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Atualizar_Status_NFE.RecordCount
   
   Atualizar_Status_NFE.MoveFirst
   
   While Not Atualizar_Status_NFE.EOF
         
         nGrvOk = 0
         
         Set cFields = New Collection
         cFields.Add Atualizar_Status_NFE!NR_DOCUMENTO_FISCAL
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectNFE, cFields, 60)
            
         sStatusMsg = "0"
         sData = Format(Now(), "dd/mm/yyyy")
         sHora = Format(Now(), "HH:MM")
         sTipo = "Ajuste_NFE"
         sCodFun = "9999"
         sMsg = "NFE Atualizada para Validar, " & Str(Atualizar_Status_NFE!NR_DOCUMENTO_FISCAL)
         
         Set cFields = New Collection
         cFields.Add sStatusMsg & ";" & _
                            sData & ";" & _
                            sHora & ";" & _
                            sCodFun & ";" & _
                            sTipo & ";" & _
                            sMsg
         
         Call Gerar_Situacao_Log(cFields)
         ADOConnectNFE.CommitTrans
         ADOConnectNFE.BeginTrans
         nGravou = nGravou + 1

PROXIMO_REGISTRO:
         
         If Not Atualizar_Status_NFE.EOF Then
            Atualizar_Status_NFE.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If

ADOConnectNFE.CommitTrans
ADOConnectNFE.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "CadFunSet"
   sCodFun = "0000"
   sMsg = "Nao ha NFE no periodo de " & c_Data_Dia
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

Set ADOConnectNFE = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing



Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "CadFunSet"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear
'
'If Err.Number <> 0 Then
'   If bSetor Then
'      GoTo PROXIMO_REGISTRO
'   Else
'      GoTo PROXIMO_REGISTROF
'   End If
'End If

'ADOConnectNFE.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Atualizar_Horas_Bat_Ajuste_Ged(ByVal sBdRm As String, _
                                               ByVal sBdGed As String, _
                                               ByVal c_Data_Atestado As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectGed As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date
Dim bSetor As Boolean

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectGed = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectGed = cConect.Coneccao_OLEDB(sBdGed, "A")

Set cFields = New Collection
cFields.Add "GED_TAB_ITEM_SOLICITACAO"
cFields.Add "ULT_DATA_HEXTRA_REALIZADA"
cFields.Add "0"

Rem Pegar a ultima data do registro atualizado no ged
Set Atualizar_Horas_Bat_Ajuste_Ged = RegBanco_dados(ADOConnectGed, cFields, 21)

If Atualizar_Horas_Bat_Ajuste_Ged.RecordCount = 0 Or Len(Trim(Atualizar_Horas_Bat_Ajuste_Ged!TPR_CONTEUDO)) = 0 Then
   sStatusMsg = "1"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Horas Extras Realizadas"
   sCodFun = "0001"
   sMsg = "Data do Parametro nao encontrado no GED."
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   Call Gerar_Situacao_Log(cFields)
   GoTo SAIDA
End If

Set cFields = New Collection
cFields.Add Format(CDate(Atualizar_Horas_Bat_Ajuste_Ged!TPR_CONTEUDO) - 30, "YYYY-MM-DD")
cFields.Add "1", "COLIGADA"

Rem SELECIONAR NO GED AS SOLICITAÇÕES DE TRINTA DIAS ANTES DA ULTIMA ATUALIZAÇÃO DA DATA PARAMENTRO DO GED.
Set Atualizar_Horas_Bat_Ajuste_Ged = RegBanco_Critica_Pessoas(ADOConnectGed, cFields, 51)

bSetor = True
nGravou = 0
nLidos = 0

If Atualizar_Horas_Bat_Ajuste_Ged.RecordCount > 0 Then ' CASO ACHE ATUALIZAR OS SETORES DO GED VINDO DO RM
   
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Atualizar_Horas_Bat_Ajuste_Ged.RecordCount
   
   Atualizar_Horas_Bat_Ajuste_Ged.MoveFirst
   
   While Not Atualizar_Horas_Bat_Ajuste_Ged.EOF
         
         nGrvOk = 0
         
         
        Rem Verificar no Rm se existe movimentação a partir da data recebida do ged.
        Set cFields = New Collection
        cFields.Add "1", "COLIGADA"
        cFields.Add Format(Atualizar_Horas_Bat_Ajuste_Ged!ITS_HORA_EXTRA_REALIZADA, "YYYY-MM-DD"), "sDataMov"
        cFields.Add Atualizar_Horas_Bat_Ajuste_Ged!ITS_USU_CHAPA, "CHAPA"
        Set rRs = New ADODB.Recordset
        Set rRs = RegBanco_dados(ADOConnectRm, cFields, 52)
         
        If rRs.RecordCount = 0 Then
           GoTo PROXIMO_REGISTRO
        End If
        
        Rem VERIFICAR SE HÁ DIFERRENCA NAS HORAS REALIZADAS DO RM E O GED
        
        If Format(rRs!Data, "YYYY-MM-DD") & " " & DateAdd("n", rRs!NUMHORAS, "00:00:00") = Format(Atualizar_Horas_Bat_Ajuste_Ged!ITS_HORA_EXTRA_REALIZADA, "YYYY-MM-DD HH:MM:SS") Then
           GoTo PROXIMO_REGISTRO
        End If
         
         
        Set cFields = New Collection
        cFields.Add Atualizar_Horas_Bat_Ajuste_Ged!ITS_CODIGO
        
        If rRs!NUMHORAS = 1440 Then
           cFields.Add "'" & Format(rRs!Data, "YYYY-MM-DD") & " " & "00:00:00" & "'", "ITS_HORA_EXTRA_REALIZADA"
        Else
           cFields.Add "'" & Format(rRs!Data, "YYYY-MM-DD") & " " & DateAdd("n", rRs!NUMHORAS, "00:00:00") & "'", "ITS_HORA_EXTRA_REALIZADA"
        End If
        
        Call RegBanco_dados(ADOConnectGed, cFields, 53)
        
        sStatusMsg = "0"
        sData = Format(Now(), "dd/mm/yyyy")
        sHora = Format(Now(), "HH:MM")
        sTipo = "SolBatAjuste"
        sCodFun = Atualizar_Horas_Bat_Ajuste_Ged!ITS_USU_CHAPA
        sMsg = "Ajuste Ponto de " & Format(rRs!Data, "YYYY-MM-DD") & " " & DateAdd("n", rRs!NUMHORAS, "00:00:00") & " / " & Format(Atualizar_Horas_Bat_Ajuste_Ged!ITS_HORA_EXTRA_REALIZADA, "YYYY-MM-DD HH:MM:SS")
        
        Set cFields = New Collection
        cFields.Add sStatusMsg & ";" & _
                           sData & ";" & _
                           sHora & ";" & _
                           sCodFun & ";" & _
                           sTipo & ";" & _
                           sMsg
        
        Call Gerar_Situacao_Log(cFields)
        ADOConnectGed.CommitTrans
        ADOConnectGed.BeginTrans
        nGravou = nGravou + 1

PROXIMO_REGISTRO:
         
         If Not Atualizar_Horas_Bat_Ajuste_Ged.EOF Then
            Atualizar_Horas_Bat_Ajuste_Ged.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If

ADOConnectRm.CommitTrans
ADOConnectGed.CommitTrans
ADOConnectRm.Close
ADOConnectGed.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "SolBatAjuste"
   sCodFun = "0000"
   sMsg = "Nao ha ajustes de Batidas no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

SAIDA:

Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If

sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "CadFunSet"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear
'
'If Err.Number <> 0 Then
'   If bSetor Then
'      GoTo PROXIMO_REGISTRO
'   Else
'      GoTo PROXIMO_REGISTROF
'   End If
'End If

'ADOConnectRm.RollbackTrans
'ADOConnectGed.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Enviar_Email_TresdiasUteis_Ged(ByVal sBdRm As String, _
                                               ByVal sBdGed As String, _
                                               ByVal c_Data_Atestado As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectGed As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date
Dim nSeq As Integer
Dim bPulaData As Boolean
Dim nTipo60 As Boolean
Dim nTipoNormal As Boolean
Dim sNOME_N1 As String
Dim nx As Integer
Dim sEmail As String
Dim sEmail2 As String
Dim bTemDiferenca As Boolean

Rem VARIAVEIS PARA FORMATACAO DO TEXTO PARA OS SUPERIORES

Dim sSetor As String * 5
Dim sChapa As String * 5
Dim sNome As String * 30
Dim sFuncao As String * 20
Dim sData As String * 10
Dim sTotal_solicitada As String * 6
Dim sTotal_real As String * 6
Dim sDiferenca  As String * 6
Dim sTextoMsg As String

On Error GoTo Erro

Set ADOConnectRm = New ADODB.Connection
Set ADOConnectGed = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectGed = cConect.Coneccao_OLEDB(sBdGed, "A")

Rem VERIFICAR A DATA CORRESPONDENTE DE ACORDO COM O PLANEJAMENTO.
Rem DE TRES DIAS ANTERIORES QUE SEJA UTEIS.
Rem CASO SEJA UM DIA DE SAB/DOM/FER, NÃO CONTAR E PEGAR UM DIA ANTERIOR.
Rem CASO SEJA O DIA HOJE : SA/DOM/FER, NÃO ENVIAR NADA
dDataCritica = Format(Now(), "DD/MM/YYYY")

Rem VERIFICAR O DIA DA SEMANA
If Retorna_DiaSemana(dDataCritica) = "SABADO" Or _
   Retorna_DiaSemana(dDataCritica) = "DOMINGO" Then
   sStatusMsg = "1"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0001"
   sMsg = "No SAB/DOM não enviar e-mail de HE. extrapoladas/não autorizadas "
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   Call Gerar_Situacao_Log(cFields)
   GoTo SAIDA
End If

Rem verificar os feriados caso haja no dia
Set cFields = New Collection
cFields.Add "'" & Format(dDataCritica, "YYYY-MM-DD") & "'"

Set Enviar_Email_TresdiasUteis_Ged = RegBanco_dados(ADOConnectRm, cFields, 55)

If Enviar_Email_TresdiasUteis_Ged.RecordCount = 1 Then
   sStatusMsg = "1"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0001"
   sMsg = "No feriado não enviar e-mail de HE. Extrapoladas/Não autorizadas "
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   Call Gerar_Situacao_Log(cFields)
   GoTo SAIDA
End If

Rem anteceder 5 dias uteis, caso seja feriado , sabado ou domingo, retroagir a data até que seja um dia util.
nSeq = 1
bPulaData = False

While nSeq < 6
      dDataCritica = dDataCritica - 1
      Rem verificar se é sab/dom.
'      If Retorna_DiaSemana(dDataCritica) = "SABADO" Or _
'         Retorna_DiaSemana(dDataCritica) = "DOMINGO" Then
'         bPulaData = True
'      Else
'         Rem verificar os feriados caso haja no dia
'         Set cFields = New Collection
'         cFields.Add "'" & Format(dDataCritica, "YYYY-MM-DD") & "'"
'
'         Set Enviar_Email_TresdiasUteis_Ged = RegBanco_dados(ADOConnectRm, cFields, 55)
'
'         If Enviar_Email_TresdiasUteis_Ged.RecordCount = 1 Then
'            bPulaData = True
'         End If
'      End If
'
'      If Not bPulaData Then
         nSeq = nSeq + 1
'      End If
      
'      bPulaData = False
Wend

Set cFields = New Collection
cFields.Add "01", "COLIGADA"
cFields.Add "'" & Format(dDataCritica, "YYYY-MM-DD") & "'"
cFields.Add "1", "classifica"

Rem Pegar as solicitações dos dias antecipados sem contar os dias uteis,feriado e fim de semanas
Set Enviar_Email_TresdiasUteis_Ged = RegBanco_dados(ADOConnectGed, cFields, 54)



If Enviar_Email_TresdiasUteis_Ged.RecordCount = 0 Then
   sStatusMsg = "1"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0001"
   sMsg = "Sem Solicitações neste dia.não enviar e-mail de HE. extrapoladas/Não autorizadas"
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   Call Gerar_Situacao_Log(cFields)
   GoTo SAIDA
End If

nGravou = 0
nLidos = 0
nTipo60 = 0
nTipoNormal = 0

If Enviar_Email_TresdiasUteis_Ged.RecordCount > 0 Then ' CASO ACHE ENVIAR EMAILS PARA OS GESTORES
   
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Enviar_Email_TresdiasUteis_Ged.RecordCount
   
   Enviar_Email_TresdiasUteis_Ged.MoveFirst
   
   sNOME_N1 = Enviar_Email_TresdiasUteis_Ged!NOME_N1
   
   Set cFields = New Collection
'   sEmail = "consultmp@hotmail.com"
'   sEmail2 = "mauro@musashi.com.br"
   sEmail = Enviar_Email_TresdiasUteis_Ged!EMAIL_N1
   sEmail2 = Enviar_Email_TresdiasUteis_Ged!EMAIL_N2
   
   cFields.Add "<strong>Divergência de horas extras solicitadas vs realizadas</strong> <br/>"
   cFields.Add "<style>table {border-collapse:collapse;}table, th, td {border: 1px solid black;}</style>"
   cFields.Add "<table><tr><td>Setor</td><td>Chapa</td><td>Nome</td><td>Função</td><td>Data</td><td>He.Sol</td><td>He.Rea</td><td>He.Dif</td><tr>"
   
'   cFields.Add " Setor Chapa Nome                           Funcao               Data       H.Sol. H.Rea. H.Dif" & Chr$(13) + Chr$(10)
'   cFields.Add " ==============================================================================================" & Chr$(13) + Chr$(10)
'   cFields.Add " " & Chr$(13) + Chr$(10)
   
   While Not Enviar_Email_TresdiasUteis_Ged.EOF
         
         nGrvOk = 0
         
         Rem VERIFICAR SE FOR DO TIPO 60, INCLUIDA PELO RODBEL, ENVIAR MENSAGEM
         
         sSetor = Enviar_Email_TresdiasUteis_Ged!SOL_TSE_CODIGO
         sChapa = Enviar_Email_TresdiasUteis_Ged!ITS_USU_CHAPA
         sNome = Enviar_Email_TresdiasUteis_Ged!NOME_USU
         sFuncao = Enviar_Email_TresdiasUteis_Ged!NOME_FUNCAO
         sData = Format(Enviar_Email_TresdiasUteis_Ged!Data, "DD/MM/YYYY") 'Format(dDataCritica, "DD/MM/YYYY")
         sTotal_solicitada = Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO
         sTotal_real = Mid$(Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO, 1, 5)
         sDiferenca = "00:00"
'         sDiferenca = Seg2Tempo(DateDiff("n", CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO), CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO & ":00")))
         If CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO) > _
            CDate("01/01/2000" & " " & Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO & ":00") Then
            sDiferenca = Mid$(Seg2Tempo(DateDiff("s", CDate("01/01/2000") & " " & Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO & ":00", CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO))), 1, 5)
         End If
         
         If Trim(sNOME_N1) = Trim(Enviar_Email_TresdiasUteis_Ged!NOME_N1) Then
         
            Rem VERIFICAR SE HÁ DIFERRENCA NAS HORAS REALIZADAS OU TIPO 60, QUE significa NÃO SOLICITADO
            bTemDiferenca = False
            If Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO = "00:00" Then
               bTemDiferenca = True
            ElseIf Mid$(Trim(sDiferenca), 1, 2) <> "00" Or Val(Mid$(Trim(sDiferenca), 4, 2)) > 20 Then
               bTemDiferenca = True
            End If
               
            
            Rem VERIFICAR SE HÁ DIFERRENCA NAS HORAS REALIZADAS OU TIPO 60, QUE É NÃO SOLICITADO
            If bTemDiferenca Or Enviar_Email_TresdiasUteis_Ged!SOL_TMS_CODIGO = 60 Then
            
               cFields.Add "<tr> <td>" & sSetor & "</td>" & _
                             "<td>" & sChapa & " </td>" & _
                             "<td>" & sNome & "</td>" & _
                             "<td>" & sFuncao & "</td>" & _
                             "<td>" & sData & "</td>" & _
                             "<td>" & sTotal_solicitada & "</td>" & _
                             "<td>" & sTotal_real & "</td>" & _
                             "<td>" & sDiferenca & "<tr>"
            End If
'****************************************************************************************************************
         Else
            If cFields.Count > 3 Then ' caso haja movimento a ser enviado
               Rem gravar o aquivo de email para envio aos superiores
               sTextoMsg = ""
               Rem atualizar a varialvel com os dados do corpo do email
               For nx = 1 To cFields.Count
                   sTextoMsg = sTextoMsg & cFields(nx)
               Next
               
               sTextoMsg = sTextoMsg & " </table>"
               
               Set cFields = New Collection
               cFields.Add "Email de horas extras não autorizadas", "TE_ASSUNTO"
               cFields.Add "1", "TE_ENVIADO"
               cFields.Add sTextoMsg, "TE_MENSAGEM"
               cFields.Add sEmail, "TE_PARA"
               cFields.Add "", "TE_ANEXO"
               cFields.Add Format(Now(), "YYYY-MM-DD"), "TE_DATA_CRIACAO"
               cFields.Add "", "TE_DATA_ENVIO"
               cFields.Add "0", "TE_ENVIADO_ST"
                
               Set rRs = New ADODB.Recordset
               Set rRs = RegBanco_dados(ADOConnectGed, cFields, 56)
                  
               Set cFields = New Collection
               cFields.Add "Email de horas extras não autorizadas", "TE_ASSUNTO"
               cFields.Add "1", "TE_ENVIADO"
               cFields.Add sTextoMsg, "TE_MENSAGEM"
               cFields.Add sEmail2, "TE_PARA"
               cFields.Add "", "TE_ANEXO"
               cFields.Add Format(Now(), "YYYY-MM-DD"), "TE_DATA_CRIACAO"
               cFields.Add "", "TE_DATA_ENVIO"
               cFields.Add "0", "TE_ENVIADO_ST"
                
               Set rRs = New ADODB.Recordset
               Set rRs = RegBanco_dados(ADOConnectGed, cFields, 56)
               
               sStatusMsg = "1"
               sData = Format(Now(), "dd/mm/yyyy")
               sHora = Format(Now(), "HH:MM")
               sTipo = "Enviar_Email"
               sCodFun = "0001"
               sMsg = "Email enviado de Solic C/Dif ou Tipo 60, para " & sEmail
               Set cFields = New Collection
               cFields.Add sStatusMsg & ";" & _
                                  sData & ";" & _
                                  sHora & ";" & _
                                  sCodFun & ";" & _
                                  sTipo & ";" & _
                                  sMsg
               Call Gerar_Situacao_Log(cFields)
               ADOConnectGed.CommitTrans
               ADOConnectGed.BeginTrans
               nGravou = nGravou + 1
            End If
            
            Rem novo email, outro email a ser enviado
            Set cFields = New Collection
            
            cFields.Add "<strong>Divergência de horas extras solicitadas vs realizadas</strong> <br/>"
            cFields.Add "<style>table {border-collapse:collapse;}table, th, td {border: 1px solid black;}</style>"
            cFields.Add "<table><tr><td>Setor</td><td>Chapa</td><td>Nome</td><td>Função</td><td>Data</td><td>He.Sol</td><td>He.Rea</td><td>He.Dif</td><tr>"
'            cFields.Add " Setor Chapa Nome                           Funcao               Data       H.Sol. H.Rea. H.Dif" & Chr$(13) + Chr$(10)
'            cFields.Add " ==============================================================================================" & Chr$(13) + Chr$(10)
'            cFields.Add " " & Chr$(13) + Chr$(10)
            
            sSetor = Enviar_Email_TresdiasUteis_Ged!SOL_TSE_CODIGO
            sChapa = Enviar_Email_TresdiasUteis_Ged!ITS_USU_CHAPA
            sNome = Enviar_Email_TresdiasUteis_Ged!NOME_USU
            sFuncao = Enviar_Email_TresdiasUteis_Ged!NOME_FUNCAO
            sData = Format(dDataCritica, "DD/MM/YYYY")
            sTotal_solicitada = Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO
            sTotal_real = Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO
            sNOME_N1 = Enviar_Email_TresdiasUteis_Ged!NOME_N1
'            sEmail = "consultmp@hotmail.com"
'            sEmail2 = "mauro@musashi.com.br"
            sEmail = Enviar_Email_TresdiasUteis_Ged!EMAIL_N1
            sEmail2 = Enviar_Email_TresdiasUteis_Ged!EMAIL_N2
            If CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO) < CDate("01/01/2000") & " " & Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO & ":00" Then
               sDiferenca = Mid$(Seg2Tempo(DateDiff("s", CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO), Format("01/01/2000" & " " & Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO & ":00", "DD/MM/YYYY HH:MM:SS"))), 1, 5)
            Else
               sDiferenca = Mid$(Seg2Tempo(DateDiff("s", CDate("01/01/2000") & " " & Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO & ":00", CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO))), 1, 5)
            End If
            Rem VERIFICAR SE HÁ DIFERRENCA NAS HORAS REALIZADAS OU TIPO 60, QUE É NÃO SOLICITADO
            bTemDiferenca = False
            If Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO = "00:00" Then
               bTemDiferenca = True
            ElseIf Mid$(Trim(sDiferenca), 1, 2) <> "00" Or Val(Mid$(Trim(sDiferenca), 4, 2)) > 20 Then
               bTemDiferenca = True
            End If
            
            If bTemDiferenca Or Enviar_Email_TresdiasUteis_Ged!SOL_TMS_CODIGO = 60 Then
               cFields.Add "<tr> <td>" & sSetor & "</td>" & _
                             "<td>" & sChapa & " </td>" & _
                             "<td>" & sNome & "</td>" & _
                             "<td>" & sFuncao & "</td>" & _
                             "<td>" & sData & "</td>" & _
                             "<td>" & sTotal_solicitada & "</td>" & _
                             "<td>" & sTotal_real & "</td>" & _
                             "<td>" & sDiferenca & "<tr>"
            End If
        End If

PROXIMO_REGISTRO:
         
         If Not Enviar_Email_TresdiasUteis_Ged.EOF Then
            Enviar_Email_TresdiasUteis_Ged.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

   If cFields.Count > 3 Then ' caso haja movimento a ser enviado
      Rem gravar o aquivo de email para envio aos superiores
      sTextoMsg = ""
      Rem atualizar a varialvel com os dados do corpo do email
      For nx = 1 To cFields.Count
          sTextoMsg = sTextoMsg & cFields(nx)
      Next
      sTextoMsg = sTextoMsg & " </table>"
      
      Set cFields = New Collection
      cFields.Add "Email de horas extras não autorizadas", "TE_ASSUNTO"
      cFields.Add "1", "TE_ENVIADO"
      cFields.Add sTextoMsg, "TE_MENSAGEM"
      cFields.Add sEmail, "TE_PARA"
      cFields.Add "", "TE_ANEXO"
      cFields.Add Format(Now(), "YYYY-MM-DD"), "TE_DATA_CRIACAO"
      cFields.Add "", "TE_DATA_ENVIO"
      cFields.Add "0", "TE_ENVIADO_ST"
       
      Set rRs = New ADODB.Recordset
      Set rRs = RegBanco_dados(ADOConnectGed, cFields, 56)
         
      sStatusMsg = "1"
      sData = Format(Now(), "dd/mm/yyyy")
      sHora = Format(Now(), "HH:MM")
      sTipo = "Enviar_Email"
      sCodFun = "0001"
      sMsg = "Email enviado de Solic C/Dif ou Tipo 60, para " & sEmail
      Set cFields = New Collection
      cFields.Add sStatusMsg & ";" & _
                         sData & ";" & _
                         sHora & ";" & _
                         sCodFun & ";" & _
                         sTipo & ";" & _
                         sMsg
      Call Gerar_Situacao_Log(cFields)
      ADOConnectGed.CommitTrans
      ADOConnectGed.BeginTrans
      nGravou = nGravou + 1
   End If


End If

ADOConnectRm.CommitTrans
ADOConnectGed.CommitTrans
ADOConnectRm.Close
ADOConnectGed.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0000"
   sMsg = "Nao Houve envio de emails nesta data, para os Gestores N1 e N2."
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

SAIDA:

Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If

sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Enviar_Email"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear
Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing
'
'If Err.Number <> 0 Then
'   If bSetor Then
'      GoTo PROXIMO_REGISTRO
'   Else
'      GoTo PROXIMO_REGISTROF
'   End If
'End If

'ADOConnectRm.RollbackTrans
'ADOConnectGed.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Enviar_Email_RH(ByVal sBdRm As String, _
                                ByVal sBdGed As String, _
                                ByVal c_Data_Atestado As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectGed As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date
Dim dData1 As Date
Dim dData2 As Date
Dim nSeq As Integer
Dim bPulaData As Boolean
Dim nTipo60 As Boolean
Dim nTipoNormal As Boolean
Dim sNOME_N1 As String
Dim nx As Integer
Dim rs As New ADODB.Recordset
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim sHoraArquivo As String
Dim bTemDiferenca As Boolean

Rem VARIAVEIS PARA FORMATACAO DO TEXTO PARA OS SUPERIORES

Dim sSetor As String * 5
Dim sChapa As String * 5
Dim sNome As String * 30
Dim sFuncao As String * 20
Dim sData As String * 10
Dim sTotal_solicitada As String * 6
Dim sTotal_real As String * 6
Dim sDiferenca  As String * 7
Dim sTextoMsg As String

On Error GoTo Erro

Set ADOConnectRm = New ADODB.Connection
Set ADOConnectGed = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectGed = cConect.Coneccao_OLEDB(sBdGed, "A")

Rem VERIFICAR A DATA CORRESPONDENTE DE ACORDO COM O PLANEJAMENTO.

dData1 = CDate("16/" & Mid$(DateAdd("m", -1, c_Data_Atestado), 4, 10))
dData2 = CDate("15/" & Mid$(c_Data_Atestado, 4, 10))


Set cFields = New Collection
cFields.Add "01", "COLIGADA"
cFields.Add "'" & Format(dData1, "YYYY-MM-DD") & "'"
cFields.Add "'" & Format(dData2, "YYYY-MM-DD") & "'"

Rem Pegar as solicitações de tres dias antecipados em dias uteis sem feriado e fim de semanas
Set Enviar_Email_RH = RegBanco_dados(ADOConnectGed, cFields, 58)

If Enviar_Email_RH.RecordCount = 0 Then
   sStatusMsg = "1"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0001"
   sMsg = "Sem Solicitações.Não enviar e-mail de HE.Extrapoladas/Não autorizadas RH"
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   Call Gerar_Situacao_Log(cFields)
   GoTo SAIDA
End If

nGravou = 0
nLidos = 0
nTipo60 = 0
nTipoNormal = 0
   
rs.Fields.Append "sSetor", adVarChar, 5
rs.Fields.Append "sChapa", adVarChar, 5
rs.Fields.Append "sNome", adVarChar, 30
rs.Fields.Append "sFuncao", adVarChar, 20
rs.Fields.Append "sData", adVarChar, 10
rs.Fields.Append "sTotal_solicitada", adVarChar, 6
rs.Fields.Append "sTotal_real", adVarChar, 6
rs.Fields.Append "sDiferenca", adVarChar, 6
rs.Open
   
If Enviar_Email_RH.RecordCount > 0 Then ' CASO ACHE ENVIAR EMAILS PARA OS GESTORES
   
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Enviar_Email_RH.RecordCount
   
   Enviar_Email_RH.MoveFirst
   
   While Not Enviar_Email_RH.EOF
         
         nGrvOk = 0
         
         Rem VERIFICAR SE FOR DO TIPO 60, INCLUIDA PELO RODBEL, ENVIAR MENSAGEM
         
         sSetor = Enviar_Email_RH!SOL_TSE_CODIGO
         sChapa = Enviar_Email_RH!ITS_USU_CHAPA
         sNome = Enviar_Email_RH!NOME_USU
         sFuncao = Enviar_Email_RH!NOME_FUNCAO
         sData = Format(Enviar_Email_RH!Data, "DD/MM/YYYY")
         sTotal_solicitada = Enviar_Email_RH!TOTAL_SOLICITADO
         sTotal_real = Mid$(Enviar_Email_RH!TOTAL_REALIZADO, 1, 5)
         sDiferenca = "00:00"
         
         If CDate("01/01/2000 " & Enviar_Email_RH!TOTAL_REALIZADO) > CDate("01/01/2000") & " " & Enviar_Email_RH!TOTAL_SOLICITADO & ":00" Then
'            sDiferenca = Mid$(Seg2Tempo(DateDiff("s", CDate("01/01/2000 " & Enviar_Email_RH!TOTAL_REALIZADO), Format("01/01/2000" & " " & Enviar_Email_RH!TOTAL_SOLICITADO & ":00", "DD/MM/YYYY HH:MM:SS"))), 1, 5)
'         Else
            sDiferenca = Mid$(Seg2Tempo(DateDiff("s", CDate("01/01/2000") & " " & Enviar_Email_RH!TOTAL_SOLICITADO & ":00", CDate("01/01/2000 " & Enviar_Email_RH!TOTAL_REALIZADO))), 1, 5)
         End If
         
         bTemDiferenca = False
         If Trim(sDiferenca) <> "00:00" Then ' POR TOLERANCIA DE 20 MINUTOS, CASO SEJA MAIOR QUE ISSO, DEVERÁ SER ENVIADO O ENAIL
            If (DateDiff("s", CDate("01/01/2000") & " " & Enviar_Email_RH!TOTAL_SOLICITADO & ":00", CDate("01/01/2000 " & Enviar_Email_RH!TOTAL_REALIZADO))) > 1200 Then
               bTemDiferenca = True
            End If
         End If
            
         
         Rem VERIFICAR SE HÁ DIFERRENCA NAS HORAS REALIZADAS OU TIPO 60, QUE É NÃO SOLICITADO
         If bTemDiferenca Or Enviar_Email_RH!SOL_TMS_CODIGO = 60 Then
            rs.AddNew
            rs.Fields("sSetor").Value = sSetor
            rs.Fields("sChapa").Value = sChapa
            rs.Fields("sNome").Value = sNome
            rs.Fields("sFuncao").Value = sFuncao
            rs.Fields("sData").Value = sData
            rs.Fields("sTotal_solicitada").Value = sTotal_solicitada
            rs.Fields("sTotal_real").Value = sTotal_real
            rs.Fields("sDiferenca").Value = Mid$(sDiferenca, 1, 6)
            rs.Update
         End If

PROXIMO_REGISTRO:
         
         If Not Enviar_Email_RH.EOF Then
            Enviar_Email_RH.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend
 
   Rem  tratar o arquivo para gerar o pdf no cristal reports.
   
   sHoraArquivo = Format(Now(), "YYYYMMDD_HHMM")
   
   If rs.RecordCount > 0 Then
      Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Solicitacao_RH.rpt")
      
      CrystalReport1.Database.SetDataSource rs
      With CrystalReport1
         .ParameterFields(1).AddCurrentValue "(" & dData1 & " a " & dData2 & ")"
         .ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue
'         .EnableParameterPrompting = False
         .ExportOptions.FormatType = crEFTPortableDocFormat
         .ExportOptions.DestinationType = crEDTDiskFile
         .ExportOptions.PDFExportAllPages = True
         .ExportOptions.DiskFileName = App.Path & "\relatorios\Solicitações_" & sHoraArquivo & ".pdf"
         .Export False
      End With
      
      rs.Clone
      
      CrystalReport1.PrintOutEx False
   End If

   Rem selecionar os emails dos funcionarios do rh para enviar o email aos destinatarios
   Set cFields = New Collection
   cFields.Add "01", "COLIGADA"
   cFields.Add "651"
   
   Rem Pegar os emails dos niveis 1 e 2 do setor 651
   Set Enviar_Email_RH = RegBanco_dados(ADOConnectGed, cFields, 57)

'   Rem Enviar os emails nivel 1
   Set cFields = New Collection
   cFields.Add "Email de horas extras não autorizadas", "TE_ASSUNTO"
   cFields.Add "1", "TE_ENVIADO"
   cFields.Add "Em anexo, Aquivo com as Horas extras extrapolados ou sem autorizações.", "TE_MENSAGEM"
'   cFields.Add "consultmp@hotmail.com"
   cFields.Add Enviar_Email_RH!EMAIL_N1, "TE_PARA"
   cFields.Add App.Path & "\relatorios\Solicitações_" & sHoraArquivo & ".PDF", "TE_ANEXO"
   cFields.Add Format(Now(), "YYYY-MM-DD"), "TE_DATA_CRIACAO"
   cFields.Add "", "TE_DATA_ENVIO"
   cFields.Add "0", "TE_ENVIADO_ST"

   Set rRs = New ADODB.Recordset
   Set rRs = RegBanco_dados(ADOConnectGed, cFields, 56)

   Rem Enviar os emails nivel 2
   Set cFields = New Collection
   cFields.Add "Email de horas extras não autorizadas", "TE_ASSUNTO"
   cFields.Add "1", "TE_ENVIADO"
   cFields.Add "Em anexo, Aquivo com as Horas extras extrapolados ou sem autorizações.", "TE_MENSAGEM"
'   cFields.Add "mauro@musashi.com.br"
   cFields.Add Enviar_Email_RH!EMAIL_N2, "TE_PARA"
   cFields.Add App.Path & "\relatorios\Solicitações_" & sHoraArquivo & ".PDF", "TE_ANEXO"
   cFields.Add Format(Now(), "YYYY-MM-DD"), "TE_DATA_CRIACAO"
   cFields.Add "", "TE_DATA_ENVIO"
   cFields.Add "0", "TE_ENVIADO_ST"

   Set rRs = New ADODB.Recordset
   Set rRs = RegBanco_dados(ADOConnectGed, cFields, 56)
   
   sStatusMsg = "1"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0001"
   sMsg = "Email ataxado enviado de Solic C/Dif ou Tipo 60 "
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   Call Gerar_Situacao_Log(cFields)
   ADOConnectGed.CommitTrans
   ADOConnectGed.BeginTrans
   nGravou = nGravou + 1
End If

ADOConnectRm.CommitTrans
ADOConnectGed.CommitTrans
ADOConnectRm.Close
ADOConnectGed.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0000"
   sMsg = "Nao Houve envio de emails nesta data, para os Gestores N1 e N2."
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

SAIDA:

Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If

sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Enviar_Email"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear
Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing
'
'If Err.Number <> 0 Then
'   If bSetor Then
'      GoTo PROXIMO_REGISTRO
'   Else
'      GoTo PROXIMO_REGISTROF
'   End If
'End If

'ADOConnectRm.RollbackTrans
'ADOConnectGed.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Enviar_Email_N1N2(ByVal sBdRm As String, _
                                  ByVal sBdGed As String, _
                                  ByVal c_Data_Atestado As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectGed As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date
Dim dData1 As Date
Dim dData2 As Date
Dim nSeq As Integer
Dim bPulaData As Boolean
Dim nTipo60 As Boolean
Dim nTipoNormal As Boolean
Dim sNOME_N1 As String
Dim nx As Integer
Dim rs As New ADODB.Recordset
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim sHoraArquivo As String
Dim bTemDiferenca As Boolean

Rem VARIAVEIS PARA FORMATACAO DO TEXTO PARA OS SUPERIORES

Dim sSetor As String * 5
Dim sChapa As String * 5
Dim sNome As String * 30
Dim sFuncao As String * 20
Dim sData As String * 10
Dim sTotal_solicitada As String * 6
Dim sTotal_real As String * 6
Dim sDiferenca  As String * 7
Dim sTextoMsg As String
Dim nSetor_Ant As String
Dim sEmailN1 As String
Dim sEmailN2 As String

On Error GoTo Erro

Set ADOConnectRm = New ADODB.Connection
Set ADOConnectGed = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectGed = cConect.Coneccao_OLEDB(sBdGed, "A")

Rem VERIFICAR A DATA CORRESPONDENTE DE ACORDO COM O PLANEJAMENTO.

dData1 = CDate("15/" & Mid$(DateAdd("m", -1, c_Data_Atestado), 4, 10))
dData2 = CDate("16/" & Mid$(c_Data_Atestado, 4, 10))
sHoraArquivo = Format(Now(), "HHMMSS")

Set cFields = New Collection
cFields.Add "01", "COLIGADA"
cFields.Add "'" & Format(dData1, "YYYY-MM-DD") & "'"
cFields.Add "'" & Format(dData2, "YYYY-MM-DD") & "'"

Rem Pegar as solicitações do mes anterior
Set Enviar_Email_N1N2 = RegBanco_dados(ADOConnectGed, cFields, 58)

If Enviar_Email_N1N2.RecordCount = 0 Then
   sStatusMsg = "1"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0001"
   sMsg = "Sem Solicitações.Não enviar e-mail de HE.Extrapoladas/Não autorizadas RH"
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   Call Gerar_Situacao_Log(cFields)
   GoTo SAIDA
End If

nGravou = 0
nLidos = 0
nTipo60 = 0
nTipoNormal = 0
   
rs.Fields.Append "sSetor", adVarChar, 5
rs.Fields.Append "sChapa", adVarChar, 5
rs.Fields.Append "sNome", adVarChar, 30
rs.Fields.Append "sFuncao", adVarChar, 20
rs.Fields.Append "sData", adVarChar, 10
rs.Fields.Append "sTotal_solicitada", adVarChar, 6
rs.Fields.Append "sTotal_real", adVarChar, 6
rs.Fields.Append "sDiferenca", adVarChar, 6
rs.Open
   
If Enviar_Email_N1N2.RecordCount > 0 Then ' CASO ACHE ENVIAR EMAILS PARA OS GESTORES
   
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Enviar_Email_N1N2.RecordCount
   
   Enviar_Email_N1N2.MoveFirst
   nSetor_Ant = Enviar_Email_N1N2!SOL_TSE_CODIGO
   sEmailN1 = Enviar_Email_N1N2!EMAIL_N1
   sEmailN2 = Enviar_Email_N1N2!EMAIL_N2
   
   While Not Enviar_Email_N1N2.EOF
         
         nGrvOk = 0
         
         Rem VERIFICAR SE FOR DO TIPO 60, INCLUIDA PELO RODBEL, ENVIAR MENSAGEM
         
         sSetor = Enviar_Email_N1N2!SOL_TSE_CODIGO
         sChapa = Enviar_Email_N1N2!ITS_USU_CHAPA
         sNome = Enviar_Email_N1N2!NOME_USU
         sFuncao = Enviar_Email_N1N2!NOME_FUNCAO
         sData = Format(Enviar_Email_N1N2!Data, "DD/MM/YYYY")
         sTotal_solicitada = Enviar_Email_N1N2!TOTAL_SOLICITADO
         sTotal_real = Mid$(Enviar_Email_N1N2!TOTAL_REALIZADO, 1, 5)
         sDiferenca = "00:00"
         
         If CDate("01/01/2000 " & Enviar_Email_N1N2!TOTAL_REALIZADO) > CDate("01/01/2000") & " " & Enviar_Email_N1N2!TOTAL_SOLICITADO & ":00" Then
'            sDiferenca = Mid$(Seg2Tempo(DateDiff("s", CDate("01/01/2000 " & Enviar_Email_N1N2!TOTAL_REALIZADO), Format("01/01/2000" & " " & Enviar_Email_N1N2!TOTAL_SOLICITADO & ":00", "DD/MM/YYYY HH:MM:SS"))), 1, 5)
'         Else
            sDiferenca = Mid$(Seg2Tempo(DateDiff("s", CDate("01/01/2000") & " " & Enviar_Email_N1N2!TOTAL_SOLICITADO & ":00", CDate("01/01/2000 " & Enviar_Email_N1N2!TOTAL_REALIZADO))), 1, 5)
         End If
         
         bTemDiferenca = False
         If Trim(sDiferenca) <> "00:00" Then ' POR TOLERANCIA DE 20 MINUTOS, CASO SEJA MAIOR QUE ISSO, DEVERÁ SER ENVIADO O ENAIL
            If (DateDiff("s", CDate("01/01/2000") & " " & Enviar_Email_N1N2!TOTAL_SOLICITADO & ":00", CDate("01/01/2000 " & Enviar_Email_N1N2!TOTAL_REALIZADO))) > 1200 Then
               bTemDiferenca = True
            End If
         End If
            
         
         Rem VERIFICAR SE HÁ DIFERRENCA NAS HORAS REALIZADAS OU TIPO 60, QUE É NÃO SOLICITADO
         If bTemDiferenca Or Enviar_Email_N1N2!SOL_TMS_CODIGO = 60 Then
            rs.AddNew
            rs.Fields("sSetor").Value = sSetor
            rs.Fields("sChapa").Value = sChapa
            rs.Fields("sNome").Value = sNome
            rs.Fields("sFuncao").Value = sFuncao
            rs.Fields("sData").Value = sData
            rs.Fields("sTotal_solicitada").Value = sTotal_solicitada
            rs.Fields("sTotal_real").Value = sTotal_real
            rs.Fields("sDiferenca").Value = Mid$(sDiferenca, 1, 6)
            rs.Update
         End If

PROXIMO_REGISTRO:
         
         If Not Enviar_Email_N1N2.EOF Then
            Enviar_Email_N1N2.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
         
         If Not Enviar_Email_N1N2.EOF Then
            If nSetor_Ant <> Enviar_Email_N1N2!SOL_TSE_CODIGO And rs.RecordCount > 0 Then
               Rem verificar se é do mesmo setor, caso não seja gerar  arquivo para emissão do pdf e email.
               Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Solicitacao_RH.rpt")
               
               CrystalReport1.Database.SetDataSource rs
               With CrystalReport1
                  .ParameterFields(1).AddCurrentValue "(" & dData1 & " a " & dData2 & ")"
                  .ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue
'                  .EnableParameterPrompting = False
                  .ExportOptions.FormatType = crEFTPortableDocFormat
                  .ExportOptions.DestinationType = crEDTDiskFile
                  .ExportOptions.PDFExportAllPages = True
                  .ExportOptions.DiskFileName = App.Path & "\relatorios\Solicitações_" & sHoraArquivo & "_" & nSetor_Ant & ".pdf"
                  .Export False
               End With
               
               rs.Clone
               
               CrystalReport1.PrintOutEx False
            
               Rem Enviar os emails nivel 1
               Set cFields = New Collection
               cFields.Add "Email de horas extras não autorizadas - " & nSetor_Ant, "TE_ASSUNTO"
               cFields.Add "1", "TE_ENVIADO"
               cFields.Add "Em anexo, Aquivo com as Horas extras extrapolados ou sem autorizações.", "TE_MENSAGEM"
'               cFields.Add "consultmp@hotmail.com" ' sEmailN1, "TE_PARA"
               cFields.Add sEmailN1, "TE_PARA"
               cFields.Add App.Path & "\relatorios\Solicitações_" & sHoraArquivo & "_" & nSetor_Ant & ".pdf", "TE_ANEXO"
               cFields.Add Format(Now(), "YYYY-MM-DD"), "TE_DATA_CRIACAO"
               cFields.Add "", "TE_DATA_ENVIO"
               cFields.Add "0", "TE_ENVIADO_ST"
               Set rRs = New ADODB.Recordset
               Set rRs = RegBanco_dados(ADOConnectGed, cFields, 56)
                  
               Rem Enviar os emails nivel 2
               Set cFields = New Collection
               cFields.Add "Email de horas extras não autorizadas - " & nSetor_Ant, "TE_ASSUNTO"
               cFields.Add "1", "TE_ENVIADO"
               cFields.Add "Em anexo, Aquivo com as Horas extras extrapolados ou sem autorizações.", "TE_MENSAGEM"
'               cFields.Add "consultmp@hotmail.com" ' sEmailN2, "TE_PARA"
               cFields.Add sEmailN2, "TE_PARA"
               cFields.Add App.Path & "\relatorios\Solicitações_" & sHoraArquivo & "_" & nSetor_Ant & ".pdf", "TE_ANEXO"
               cFields.Add Format(Now(), "YYYY-MM-DD"), "TE_DATA_CRIACAO"
               cFields.Add "", "TE_DATA_ENVIO"
               cFields.Add "0", "TE_ENVIADO_ST"
               Set rRs = New ADODB.Recordset
               Set rRs = RegBanco_dados(ADOConnectGed, cFields, 56)
            
               rs.Close
               Set rs = New ADODB.Recordset
               rs.Fields.Append "sSetor", adVarChar, 5
               rs.Fields.Append "sChapa", adVarChar, 5
               rs.Fields.Append "sNome", adVarChar, 30
               rs.Fields.Append "sFuncao", adVarChar, 20
               rs.Fields.Append "sData", adVarChar, 10
               rs.Fields.Append "sTotal_solicitada", adVarChar, 6
               rs.Fields.Append "sTotal_real", adVarChar, 6
               rs.Fields.Append "sDiferenca", adVarChar, 6
               rs.Open
            End If
            If nSetor_Ant <> Enviar_Email_N1N2!SOL_TSE_CODIGO Then
               nSetor_Ant = Enviar_Email_N1N2!SOL_TSE_CODIGO
               sEmailN1 = Enviar_Email_N1N2!EMAIL_N1
               sEmailN2 = Enviar_Email_N1N2!EMAIL_N2
            End If
        End If
        
   Wend
 
   Rem  tratar o arquivo para gerar o pdf no cristal reports.
   
   If rs.RecordCount > 0 Then
      
      Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Solicitacao_RH.rpt")
      
      CrystalReport1.Database.SetDataSource rs
      With CrystalReport1
         .ParameterFields(1).AddCurrentValue "(" & dData1 & " a " & dData2 & ")"
         .ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue
'         .EnableParameterPrompting = False
         .ExportOptions.FormatType = crEFTPortableDocFormat
         .ExportOptions.DestinationType = crEDTDiskFile
         .ExportOptions.PDFExportAllPages = True
         .ExportOptions.DiskFileName = App.Path & "\relatorios\Solicitações_" & sHoraArquivo & "_" & nSetor_Ant & ".pdf"
         .Export False
      End With
      
      rs.Clone
      
      CrystalReport1.PrintOutEx False
      
      Rem Enviar os emails nivel 1
      Set cFields = New Collection
      cFields.Add "Email de horas extras não autorizadas - " & nSetor_Ant, "TE_ASSUNTO"
      cFields.Add "1", "TE_ENVIADO"
      cFields.Add "Em anexo, Aquivo com as Horas extras extrapolados ou sem autorizações.", "TE_MENSAGEM"
'      cFields.Add "consultmp@hotmail.com" ' sEmailN1, "TE_PARA"
      cFields.Add sEmailN1, "TE_PARA"
      cFields.Add App.Path & "\relatorios\Solicitações_" & sHoraArquivo & "_" & nSetor_Ant & ".pdf", "TE_ANEXO"
      cFields.Add Format(Now(), "YYYY-MM-DD"), "TE_DATA_CRIACAO"
      cFields.Add "", "TE_DATA_ENVIO"
      cFields.Add "0", "TE_ENVIADO_ST"
      Set rRs = New ADODB.Recordset
      Set rRs = RegBanco_dados(ADOConnectGed, cFields, 56)
         
      Rem Enviar os emails nivel 2
      Set cFields = New Collection
      cFields.Add "Email de horas extras não autorizadas - " & nSetor_Ant, "TE_ASSUNTO"
      cFields.Add "1", "TE_ENVIADO"
      cFields.Add "Em anexo, Aquivo com as Horas extras extrapolados ou sem autorizações.", "TE_MENSAGEM"
'      cFields.Add "consultmp@hotmail.com" ' sEmailN2, "TE_PARA"
      cFields.Add sEmailN2, "TE_PARA"
      cFields.Add App.Path & "\relatorios\Solicitações_" & sHoraArquivo & "_" & nSetor_Ant & ".pdf", "TE_ANEXO"
      cFields.Add Format(Now(), "YYYY-MM-DD"), "TE_DATA_CRIACAO"
      cFields.Add "", "TE_DATA_ENVIO"
      cFields.Add "0", "TE_ENVIADO_ST"
      Set rRs = New ADODB.Recordset
      Set rRs = RegBanco_dados(ADOConnectGed, cFields, 56)
      
   End If
   
   sStatusMsg = "1"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0001"
   sMsg = "Email c/anexo enviado de Solic C/Dif ou Tipo 60 "
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   Call Gerar_Situacao_Log(cFields)
   ADOConnectGed.CommitTrans
   ADOConnectGed.BeginTrans
   nGravou = nGravou + 1
End If

ADOConnectRm.CommitTrans
ADOConnectGed.CommitTrans
ADOConnectRm.Close
ADOConnectGed.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Enviar_Email"
   sCodFun = "0000"
   sMsg = "Nao Houve envio de emails nesta data, para os Gestores N1 e N2."
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

SAIDA:

Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If

sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Enviar_Email"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear
Set ADOConnectRm = Nothing
Set ADOConnectGed = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing
End Function


' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA CRITICA E PREPARA OS DADOS PARA MANU
'   TENCAO NO BANCO, EM ESPECIAL A TABELA DE PESSOAS NO RODBEL
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_Critica_Pessoas(ByVal cCONECCAO As Connection, _
                                         ByVal cFields As Collection, _
                                         ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro

Dim nx As Integer
Dim sSQL As String
Dim sName As String


If nAction > 1 And nAction <> 9 And nAction <> 12 And nAction < 14 Then
   'Critica dos Campos dos Vendedor

    If nAction = 3 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Codigo do funcionario Invalido"
    End If
    If nAction = 3 And Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50000, Err.Source, "Nome do funcionario invalido"
    End If
    If nAction = 3 And Len(Trim(cFields(3))) = 0 Then
       Err.Raise 50000, Err.Source, "CPF do funcionario invalido"
    End If
    If nAction = 3 And Len(Trim(cFields(6))) = 0 Then
       Err.Raise 50000, Err.Source, "carteira de identidade do funcionario invalido"
    End If
    If nAction = 3 And Len(Trim(cFields(7))) = 0 Then
       Err.Raise 50000, Err.Source, "Centro de Custo do funcionario invalido"
    End If
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 27
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
        
            Select Case nx
                'Altera campos alfanuméricos
                Case 2, 3, 4, 5, 6, 10 To 17, 21, 22
                    sSQL = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1

                ' Altera campos tipo data
                Case 20
                    sSQL = "'" & Format(cFields(nx), "mm/dd/yyyy") & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       sSQL = "1"
                   Else
                       sSQL = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add sSQL, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        sSQL = Replace(cFields(nx), ".", "")
                        sSQL = Replace(sSQL, ",", ".")
                        cFields.Remove nx
                        cFields.Add sSQL, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Pessoas = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA CRITICA E PREPARA OS DADOS PARA MANU-
'   TENCAO NO BANCO, EM ESPECIAL A TABELA DE RELACIONAMENTO
'   CRED_PESSOAS NO RODBEL
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_Critica_Rel(ByVal cCONECCAO As Connection, _
                                     ByVal cFields As Collection, _
                                     ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos Vendedor
'    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
'       Err.Raise 50000, Err.Source, "Data inicial Inválido"
'    End If
'    If Len(Trim(cFields(2))) = 0 Then
'       Err.Raise 50000, Err.Source, "dataFechamento Inválido"
'    End If
'    If Len(Trim(cFields(4))) = 0 Then
'       Err.Raise 50000, Err.Source, "CPF / CGC do VENDEDOR não Digitado"
'    End If
   
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 9
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 0
                    sSQL = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1

                ' Altera campos tipo data
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 3, 4, 5, 6
                    sSQL = "'" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       sSQL = "1"
                   Else
                       sSQL = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add sSQL, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        sSQL = Replace(cFields(nx), ".", "")
                        sSQL = Replace(sSQL, ",", ".")
                        cFields.Remove nx
                        cFields.Add sSQL, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Rel = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function RegBanco_Critica_Estrutura(ByVal cCONECCAO As Connection, _
                                           ByVal cFields As Collection, _
                                           ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos Vendedor
'    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
'       Err.Raise 50000, Err.Source, "Data inicial Inválido"
'    End If
'    If Len(Trim(cFields(2))) = 0 Then
'       Err.Raise 50000, Err.Source, "dataFechamento Inválido"
'    End If
'    If Len(Trim(cFields(4))) = 0 Then
'       Err.Raise 50000, Err.Source, "CPF / CGC do VENDEDOR não Digitado"
'    End If
   
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 15
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 0
                    sSQL = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1

                ' Altera campos tipo data
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       sSQL = "1"
                   Else
                       sSQL = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add sSQL, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        sSQL = Replace(cFields(nx), ".", "")
                        sSQL = Replace(sSQL, ",", ".")
                        cFields.Remove nx
                        cFields.Add sSQL, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Estrutura = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description

End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA CRITICA E PREPARA OS DADOS PARA MANU
'   TENCAO NO BANCO, EM ESPECIAL A TABELA DE CREDENCIAIS  NO
'   RODBEL
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_Critica_Cred(ByVal cCONECCAO As Connection, _
                                      ByVal cFields As Collection, _
                                      ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction <= 11 Then
   
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 25
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 8, 11
                    sSQL = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1

                ' Altera campos tipo data
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 9, 10, 16, 22, 24
                    sSQL = "'" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       sSQL = "1"
                   Else
                       sSQL = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add sSQL, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        sSQL = Replace(cFields(nx), ".", "")
                        sSQL = Replace(sSQL, ",", ".")
                        cFields.Remove nx
                        cFields.Add sSQL, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Cred = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA CRITICA E PREPARA OS DADOS PARA MANU
'   TENCAO NO BANCO, EM ESPECIAL A TABELA DE CRED_ACESSO   NO
'   RODBEL
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_Critica_Cred_Acesso(ByVal cCONECCAO As Connection, _
                                             ByVal cFields As Collection, _
                                             ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction = 19 Then
   'Critica dos Campos dos Vendedor
'    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
'       Err.Raise 50000, Err.Source, "Data inicial Inválido"
'    End If
'    If Len(Trim(cFields(2))) = 0 Then
'       Err.Raise 50000, Err.Source, "dataFechamento Inválido"
'    End If
'    If Len(Trim(cFields(4))) = 0 Then
'       Err.Raise 50000, Err.Source, "CPF / CGC do VENDEDOR não Digitado"
'    End If
   
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 5 To 9
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 12
                    sSQL = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1

                ' Altera campos tipo data
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       sSQL = "1"
                   Else
                       sSQL = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add sSQL, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        sSQL = Replace(cFields(nx), ".", "")
                        sSQL = Replace(sSQL, ",", ".")
                        cFields.Remove nx
                        cFields.Add sSQL, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Cred_Acesso = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_dados(ByVal cCONECCAO As Connection, _
                               ByVal cFields As Collection, _
                               ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro

Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim nSeq As String
Dim sVirgula As String

Select Case nAction

    Case 1 'Consultar CADASTRO DE PESSOAL VERIFICANDO A DATA DE ADMISSAO
           sSQL = "SELECT " & _
                  "CONVERT(BIGINT,PFUNC.CHAPA) AS CHAPA, " & _
                  "PFUNC.NOME, " & _
                  "PFUNC.INTEGRGERENCIAL, " & _
                  "PPESSOA.CPF, " & _
                  "PPESSOA.CARTIDENTIDADE, " & _
                  "PFUNC.CODFUNCAO, " & _
                  "PFUNCAO.CARGO, " & _
                  "PFUNC.PISPASEP AS PISPARAFGTS "

           sSQL = sSQL & _
                  "FROM PFUNC " & _
                  "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO " & _
                  "INNER JOIN PFUNCAO ON PFUNC.CODFUNCAO = PFUNCAO.CODIGO AND PFUNC.CODCOLIGADA = PFUNCAO.CODCOLIGADA"
       
           sSQL = sSQL & _
                  " WHERE PFUNC.CODCOLIGADA = 1" & _
                  " AND PFUNC.DATAADMISSAO >= " & cFields(1) & ""
         
    Case 2 'Consultar CADASTRO DA RODBEL VERIFICANDO SE HA O REGISTRO DO FUNCIONARIO A SER ADMITIDO
           sSQL = "SELECT " & _
                  "NU_MATRICULA AS ENCONTRADO "
                  
           sSQL = sSQL & _
                  "FROM PESSOA "
       
           sSQL = sSQL & _
                  " WHERE NU_MATRICULA = " & cFields(1)
         
    Case 3 'INSERIR NA TABELA DE PESSOAS DA RODBEL OS NOVOS FUNCIONARIOS ADMITIDOS PELO SISTEMA RM
            
            sSQL = "INSERT INTO PESSOA (" & _
                   "NU_MATRICULA, " & _
                   "NM_PESSOA, " & _
                   "CD_SITUACAO_PESSOA, " & _
                   "CD_ESTRUTURA_ORGANIZACIONAL, " & _
                   "CD_ESTRUTURA_ORG_EMPRESA, " & _
                   "NU_CPF, " & _
                   "NU_RG, " & _
                   "NU_PIS, " & _
                   "DS_EMAIL, " & _
                   "NU_RAMAL, "

            sSQL = sSQL & _
                   "DS_OBSERVACAO, " & _
                   "DS_SENHA, " & _
                   "TP_MASTER, " & _
                   "FL_REENTRADA, " & _
                   "FL_QUEBRA_SEQUENCIA, " & _
                   "FL_DISPENSA_SENHA, " & _
                   "FL_DISPENSA_DIGITAL, " & _
                   "FL_AUTORIZADOR_EXCEPCIONAL, " & _
                   "CD_PERFIL_ACESSO, " & _
                   "TX_CAMPO01, "
            
            sSQL = sSQL & _
                   "TX_CAMPO02, " & _
                   "TX_CAMPO03, " & _
                   "TX_CAMPO04, " & _
                   "TX_CAMPO05, " & _
                   "TX_CAMPO06, " & _
                   "TX_CAMPO07, " & _
                   "TX_CAMPO08, " & _
                   "TX_CAMPO09, " & _
                   "TX_CAMPO10, " & _
                   "CD_GRUPO_PESSOA, "

            sSQL = sSQL & _
                   "DT_ULTIMA_AUTORIZACAO_ACESSO, " & _
                   "CD_EQP_ULT_AUTORIZACAO_ACESSO, " & _
                   "CD_AREA, " & _
                   "CD_CREDENCIAL, " & _
                   "CD_CREDENCIAL_REP, " & _
                   "FL_BLOQ_POR_PROVISORIO, " & _
                   "TP_CONDICAO_SITUACAO, " & _
                   "DT_RECOLHE_FIM_VALIDADE, " & _
                   "DT_BLOQ_POR_PROVISORIO, " & _
                   "FL_REQUER_AUTORIZADOR, "
            
            sSQL = sSQL & _
                   "TP_CARREGA_BIOMETRIA, " & _
                   "CD_CREDENCIAL_FACE, " & _
                   "NU_CELULAR, " & _
                   "NU_PORCENTAGEM_SORTEIO, " & _
                   "DT_PESSOA_SORTEADA, " & _
                   "DT_PES_TEMPO_TOL_NOVO_SORTEIO "
            
            sSQL = sSQL & ") VALUES (" & _
                            cFields(1) & ", " & _
                            cFields(2) & ", " & _
                            cFields(24) & ", " & _
                            cFields(25) & ", " & _
                            cFields(26) & ", " & _
                            cFields(3) & ", " & _
                            cFields(6) & ", " & _
                            cFields(23) & ", " & _
                            cFields(5) & ", " & _
                            cFields(12) & ", "
            sSQL = sSQL & _
                            cFields(4) & ", " & _
                            cFields(11) & ", " & _
                            cFields(27) & ", " & _
                            cFields(27) & ", " & _
                            cFields(27) & ", " & _
                            cFields(27) & ", " & _
                            cFields(27) & ", " & _
                            cFields(27) & ", " & _
                            cFields(19) & ", " & _
                            cFields(13) & ", "

            sSQL = sSQL & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", "
            
            sSQL = sSQL & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(27) & ", " & _
                            cFields(27) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(27) & ", "

            sSQL = sSQL & _
                            cFields(27) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ", " & _
                            cFields(27) & ", " & _
                            cFields(13) & ", " & _
                            cFields(13) & ")"

    Case 4 'INSERIR NA TABELA DE CREDENCIAIS DA RODBEL OS NOVOS FUNCIONARIOS ADMITIDOS PELO SISTEMA RM
           
           sSQL = "INSERT INTO CREDENCIAL (" & _
                  "CD_TIPO_CREDENCIAL, " & _
                  "CD_ESTRUTURA_ORGANIZACIONAL, " & _
                  "TP_TECNOLOGIA, " & _
                  "NU_CREDENCIAL, " & _
                  "FL_PERMANENTE, " & _
                  "FL_BLOQUEADA, " & _
                  "DS_MOTIVO_BLOQUEIO, " & _
                  "DT_INICIO_VALIDADE, " & _
                  "DT_FIM_VALIDADE, " & _
                  "DS_SENHA, " & _
                  "TP_MASTER, " & _
                  "FL_REENTRADA, " & _
                  "CD_PERFIL_ACESSO, " & _
                  "CD_AREA, " & _
                  "DT_ULTIMA_AUTORIZACAO_ACESSO, " & _
                  "CD_EQP_ULT_AUTORIZACAO_ACESSO, " & _
                  "FL_DISPONIVEL_USO, " & _
                  "FL_DISPENSA_SENHA, " & _
                  "FL_SUPERVISOR_EQUIPAMENTO, " & _
                  "FL_CREDENCIAL_PUBLICA, " & _
                  "HR_RECOLHIMENTO_COFRE, " & _
                  "NU_PORCENTAGEM_SORTEIO, " & _
                  "DT_CREDENCIAL_SORTEADA, " & _
                  "BN_QRCODE "
           
           sSQL = sSQL & ") VALUES (" & _
                           cFields(2) & ", " & _
                           cFields(3) & ", " & _
                           cFields(4) & ", " & _
                           cFields(5) & ", " & _
                           cFields(7) & ", " & _
                           cFields(6) & ", " & _
                           cFields(8) & ", " & _
                           cFields(9) & ", " & _
                           cFields(10) & ", " & _
                           cFields(11) & ", " & _
                           cFields(12) & ", " & _
                           cFields(13) & ", " & _
                           cFields(14) & ", " & _
                           cFields(15) & ", " & _
                           cFields(16) & ", " & _
                           cFields(17) & ", " & _
                           cFields(18) & ", " & _
                           cFields(19) & ", " & _
                           cFields(20) & ", " & _
                           cFields(21) & ", " & _
                           cFields(22) & ", " & _
                           cFields(23) & ", " & _
                           cFields(24) & ", " & _
                           cFields(25) & " )"

    Case 5 'INSERIR NA TABELA DE CRED_PESSOAS DA RODBEL OS NOVOS FUNCIONARIOS ADMITIDOS PELO SISTEMA RM
           sSQL = "INSERT INTO CREDENCIAL_PESSOA (" & _
                  "CD_CREDENCIAL, " & _
                  "CD_PESSOA, " & _
                  "DT_INICIO, " & _
                  "DT_FIM, " & _
                  "HR_RECOLHIMENTO_COFRE, " & _
                  "DT_BAIXA_CREDENCIAL, " & _
                  "CD_CREDENCIAL_ORIGINAL, " & _
                  "FL_PERMANENTE, " & _
                  "TP_TECNOLOGIA "
                  
           sSQL = sSQL & ") VALUES (" & _
                           cFields(1) & ", " & _
                           cFields(2) & ", " & _
                           cFields(3) & ", " & _
                           cFields(4) & ", " & _
                           cFields(5) & ", " & _
                           cFields(6) & ", " & _
                           cFields(7) & ", " & _
                           cFields(8) & ", " & _
                           cFields(9) & " )"

    Case 6 'Consultar TABELA DE ESTRUTURA DA RODBEL VERIFICANDO SE HA O REGISTRO DO CENTRO DE CUSTO CADASTRADO
           sSQL = "SELECT " & _
                  "COUNT(*) AS ENCONTRADO "
                  
           sSQL = sSQL & _
                  "FROM ESTRUTURA_ORGANIZACIONAL "
       
           sSQL = sSQL & _
                  " WHERE NU_ESTRUTURA = '" & cFields(1) & "'"
    
    Case 7 'Consultar TABELA DE CRED_PESSOAS DA RODBEL VERIFICANDO SE HA O REGISTRO CADASTRADO
           sSQL = "SELECT " & _
                  "CD_PESSOA AS ENCONTRADO "
                  
           sSQL = sSQL & _
                  "FROM CREDENCIAL_PESSOA "
       
           sSQL = sSQL & _
                  " WHERE CD_PESSOA = " & cFields(1)
    
    Case 8 'Consultar TABELA DE RELACIONAMENTO DE CRED_PESSOAS E PESSOAS DA RODBEL
           'VERIFICANDO SE HA O REGISTRO CADASTRADO
           
           sSQL = "SELECT " & _
                  "COUNT(*) AS ENCONTRADO," & _
                  "CREDENCIAL.FL_BLOQUEADA, " & _
                  "PESSOA.CD_SITUACAO_PESSOA  AS PESSIT_NUMERO,PESSOA.CD_PESSOA "
                  
           sSQL = sSQL & _
                  "FROM CREDENCIAL_PESSOA " & _
                  "INNER JOIN CREDENCIAL " & _
                  "ON CREDENCIAL.CD_CREDENCIAL = CREDENCIAL_PESSOA.CD_CREDENCIAL " & _
                  "INNER JOIN PESSOA " & _
                  "ON CREDENCIAL_PESSOA.CD_PESSOA = PESSOA.CD_PESSOA "
           
           sSQL = sSQL & _
                  " WHERE PESSOA.NU_MATRICULA = " & cFields(1)

           sSQL = sSQL & _
                  " GROUP BY CREDENCIAL.FL_BLOQUEADA, PESSOA.CD_SITUACAO_PESSOA,PESSOA.CD_PESSOA"
    
    Case 9 'Consultar CADASTRO DE PESSOAL VERIFICANDO A DATA DE DEMISSAO
           sSQL = "SELECT " & _
                  "CONVERT(BIGINT,PFUNC.CHAPA) AS CHAPA, " & _
                  "PFUNC.NOME, " & _
                  "PFUNC.INTEGRGERENCIAL, " & _
                  "PPESSOA.CPF, " & _
                  "PPESSOA.CARTIDENTIDADE "
           
           sSQL = sSQL & _
                  "FROM PFUNC " & _
                  "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO"
       
           sSQL = sSQL & _
                  " WHERE PFUNC.CODCOLIGADA = 1" & _
                  " AND PFUNC.DATADEMISSAO >= " & cFields(1) & ""
                  
    Case 10 'ATUALIZAR A TABELA DE PESSOAS O CAMPO PESSIST_NUMERO COM O sStatusMsg VINDO DA ACAO
           sSQL = "UPDATE PESSOA SET " & _
                  "CD_SITUACAO_PESSOA = " & cFields(18)
                  
           sSQL = sSQL & _
                  " WHERE NU_MATRICULA = " & cFields(1)
                  
    Case 11 'ATUALIZAR A TABELA DE CREDENCIAIS O CAMPO CRED_BLOQUEADA COM O sStatusMsg VINDO DA ACAO
           sSQL = "UPDATE CREDENCIAL SET " & _
                  "FL_BLOQUEADA = " & cFields(6)
       
           sSQL = sSQL & _
                  " WHERE NU_CREDENCIAL = " & cFields(1)

    Case 12 'Consultar CADASTRO DE PESSOAL VERIFICANDO A DATA DE PROGRAMACAO DE FERIAS
           sSQL = "SELECT " & _
                  "CONVERT(BIGINT,PFUNC.CHAPA) AS CHAPA, " & _
                  "PFUNC.INICPROGFERIAS1, " & _
                  "PFUNC.FIMPROGFERIAS1 "

           sSQL = sSQL & _
                  "FROM PFUNC " & _
                  "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO "

           sSQL = sSQL & _
                  " WHERE PFUNC.CODCOLIGADA = 1" & _
                  " AND PFUNC.INICPROGFERIAS1 > 0 "

    Case 122 'Consultar CADASTRO DE PESSOAL VERIFICANDO A DATA DE PROGRAMACAO DE FERIAS pelo historico
           sSQL = "SELECT CONVERT(BIGINT,PFUNC.CHAPA) AS CHAPA, " & _
                  "DTINIGOZO AS INICPROGFERIAS1, " & _
                  "DTFIMGOZO AS FIMPROGFERIAS1 " & _
                  "FROM PFUNC " & _
                  "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO " & _
                  "INNER JOIN PFHSTFER ON  PFHSTFER.CHAPA = PFUNC.CHAPA AND PFHSTFER.CODCOLIGADA = PFUNC.CODCOLIGADA " & _
                  ""
                  
           sSQL = sSQL & _
                  "WHERE  PFUNC.CODCOLIGADA = 1 AND DTINIGOZO>0 " & _
                  "AND   DTINIGOZO >= " & cFields(1) & " " & _
                  "AND   CODSITUACAO <> 'D' "
           sSQL = sSQL & _
                  "AND   PFUNC.CHAPA IN(SELECT PFUNC.CHAPA FROM PFHSTFER " & _
                  "INNER JOIN PFUNC ON  PFHSTFER.CHAPA = PFUNC.CHAPA "
           sSQL = sSQL & _
                  "WHERE DTINIGOZO >= " & cFields(1) & " " & _
                  "AND   CODSITUACAO <> 'D' " & _
                  "GROUP BY PFUNC.CHAPA) "

    Case 13 'ATUALIZAR A TABELA DE PESSOAS O CAMPO PESSIST_NUMERO,INICPROGFERIAS1,FIMPROGFERIAS1
            'COM OS DADOS VINDO DA ACAO
'de           sSQL = "UPDATE PESSOA SET " & _
'                  "CD_SITUACAO_PESSOA = " & cFields(18) & _
'                  ",PES_DTSITINICIAL = " & cFields(19) & _
'                  ",PES_DTSITFINAL = " & cFields(20)
'
'           sSQL = sSQL & _
'                  " WHERE NU_MATRICULA = " & cFields(1)
'para
            sSQL = "INSERT INTO HISTORICO_SITUACAO_PESSOA (" & _
                   "CD_PESSOA, " & _
                   "CD_SITUACAO_PESSOA, " & _
                   "DT_ATRIBUICAO, " & _
                   "DT_SITUACAO_INICIAL, " & _
                   "DT_SITUACAO_FINAL "
           
           sSQL = sSQL & ") VALUES (" & _
                           cFields(1) & ", " & _
                           cFields(18) & ",'" & _
                           Format(Now(), "MM/DD/YYYY HH:MM:SS") & "', " & _
                           cFields(19) & ", " & _
                           cFields(20) & " )"

    Case 14 'Consultar TABELA DE MUDANCA DE SECOES EM UMA DATA DEFINIDA NO BANCO DA RM
           sSQL = "SELECT " & _
                  "CONVERT(BIGINT, PFHSTSEC.CHAPA) AS CHAPA, " & _
                  " PFHSTSEC.MOTIVO, " & _
                  " SUBSTRING(CODSECAO, 10, 3) AS CODSECAO "
           
           sSQL = sSQL & _
                  "FROM  PFHSTSEC "
       
           sSQL = sSQL & _
                  " WHERE PFHSTSEC.CODCOLIGADA = 1" & _
                  " AND PFHSTSEC.DTMUDANCA = " & cFields(1) & ""

    Case 15 'Consultar TABELA DE ESTRUTURA DA RODBEL VERIFICANDO SE HA O FUNCIONARIO E COM A SECAO DIFERENTE. aTUALIZANDO DEPOIS
           sSQL = "SELECT " & _
                  "COUNT(*) AS ENCONTRADO, "
                  
           sSQL = sSQL & _
                  "(SELECT COUNT(*) FROM ESTRUTURA_ORGANIZACIONAL WHERE CD_ESTRUTURA_ORGANIZACIONAL = " & cFields(2) & ") AS ENCONTRA_EST "
           
           sSQL = sSQL & _
                  "FROM PESSOA "
       
           sSQL = sSQL & _
                  " WHERE NU_MATRICULA = " & cFields(1) & _
                  " AND   CD_ESTRUTURA_ORGANIZACIONAL <> " & cFields(2)
    
    Case 16 'ATUALIZAR A TABELA DE PESSOAS O CAMPO EST_NUMERO, COM A NOVA SECAO DO FUNCIONARIO
           sSQL = "UPDATE PESSOAS SET " & _
                  "CD_ESTRUTURA_ORGANIZACIONAL = " & cFields(7)
                  
           sSQL = sSQL & _
                  " WHERE NU_MATRICULA = " & cFields(1)


    Case 17 'SELECIONAR FUNCIONARIOS COM AFASTAMENTO NO PERIODO ULTIMO DIA REGISTRADO
    
           sSQL = "SELECT " & _
                  "PFUNC.CHAPA , PFHSTAFT.TIPO, dtinicio , dtfinal "
           sSQL = sSQL & _
                  "From PFHSTAFT " & _
                  "INNER JOIN PFUNC " & _
                  "   ON PFHSTAFT.CHAPA = PFUNC.CHAPA " & _
                  "  AND PFHSTAFT.CODCOLIGADA = PFUNC.CODCOLIGADA "
           sSQL = sSQL & _
                  "WHERE (PFUNC.CODSITUACAO <> 'D') " & _
                  "  AND (PFUNC.CODCOLIGADA = 1) " & _
                  "  AND (dtfinal >= ( getdate() - 180)  " & _
                  "   OR  dtfinal is null  " & _
                  "   OR  dtfinal = 0 ) ORDER BY PFUNC.CHAPA"

    
    
'''           sSQL = "SELECT " & _
'''                  "PFUNC.CHAPA,PFHSTAFT.TIPO, "
'''           sSQL = sSQL & _
'''                  "(SELECT TOP 1 AFAST.DTINICIO" & _
'''                  " FROM PFHSTAFT AFAST " & _
'''                  " WHERE AFAST.CODCOLIGADA = 1 " & _
'''                  " AND   AFAST.CHAPA = PFUNC.CHAPA" & _
'''                  " ORDER BY AFAST.DTINICIO DESC) AS DTINICIO, "
'''          sSQL = sSQL & _
'''                 "(SELECT TOP 1 AFAST2.DTFINAL" & _
'''                  " FROM   PFHSTAFT AFAST2 " & _
'''                  " WHERE  AFAST2.CODCOLIGADA = 1 " & _
'''                  " AND    AFAST2.CHAPA = PFUNC.CHAPA " & _
'''                  " ORDER  BY AFAST2.DTINICIO DESC) AS DTFINAL "
'''
'''          sSQL = sSQL & _
'''                  "FROM PFHSTAFT " & _
'''                  " INNER JOIN PFUNC " & _
'''                  " ON  PFHSTAFT.CHAPA = PFUNC.CHAPA " & _
'''                  " AND PFHSTAFT.CODCOLIGADA = PFUNC.CODCOLIGADA "
'''
'''         sSQL = sSQL & _
'''                  "WHERE (PFUNC.CODSITUACAO <> 'D') AND (PFUNC.CODCOLIGADA = 1) " & _
'''                  " GROUP BY PFUNC.CHAPA,PFHSTAFT.TIPO " & _
'''                  " ORDER BY PFUNC.CHAPA "
                  
'                  "WHERE (PFUNC.CODSITUACAO <> 'D') AND (PFUNC.CODCOLIGADA = 1) AND PFHSTAFT.CHAPA = '01881'" & _

    Case 18 'SELECIONAR OS GRUPOS DE ACESSO CADASTRADO NA TABELA DE GRUPOS_AREAS
    
           sSQL = "SELECT " & _
                  "GRP_NUMERO, " & _
                  "ARE_NUMEROENTRADA, " & _
                  "ARE_NUMEROSAIDA, " & _
                  "GRAR_TOLERANCIADESTINO, " & _
                  "GRAR_TOLERANCIAORIGEM, " & _
                  "GRAR_TIPOTOLERANCIA "
           sSQL = sSQL & _
                  "From GRUPOS_AREAS "
                  
    Case 19 'INCLUIR DADOS NA TABELA DE CRED_ACESSO
           
           sSQL = "INSERT INTO CRED_ACESSO (" & _
                  "CRED_NUMERO, " & _
                  "GRP_NUMERO, " & _
                  "ARE_NUMERO, " & _
                  "CRACTI_NUMERO, " & _
                  "CRAC_DIFPORGRUPO, " & _
                  "JORN_NUMERO, " & _
                  "TURN_NUMERO, " & _
                  "CRAC_QTDTOTALACESSOS, " & _
                  "CRAC_QTDACESSOS, " & _
                  "CRAC_VERFSENHA, " & _
                  "CRAC_VERFDIGITAL, " & _
                  "CRAC_ULTPASSAGEM"
           
           sSQL = sSQL & ") VALUES (" & _
                           cFields(1) & ", " & _
                           cFields(2) & ", " & _
                           cFields(3) & ", " & _
                           cFields(4) & ", " & _
                           cFields(5) & ", " & _
                           cFields(7) & ", " & _
                           cFields(6) & ", " & _
                           cFields(8) & ", " & _
                           cFields(9) & ", " & _
                           cFields(10) & ", " & _
                           cFields(11) & ", " & _
                           cFields(12) & " )"
                  
    Case 20 'SELECIONAR FUNCIONARIOS COM ATESTADO NO PERIODO ULTIMO DIA REGISTRADO
    
           sSQL = "SELECT " & _
                  "PFUNC.CHAPA,PFUNC.CODPESSOA, "
           sSQL = sSQL & _
                  "(SELECT TOP 1 ATEST.DTINICIO" & _
                  " FROM VATESTADO ATEST " & _
                  " WHERE ATEST.CODCOLIGADA = 1 " & _
                  " AND   ATEST.CODPESSOA = PFUNC.CODPESSOA" & _
                  " ORDER BY ATEST.DTINICIO DESC) AS DTINICIO, "
          sSQL = sSQL & _
                 "(SELECT TOP 1 ATEST2.DTFINAL" & _
                  " FROM   VATESTADO ATEST2 " & _
                  " WHERE  ATEST2.CODCOLIGADA = 1 " & _
                  " AND    ATEST2.CODPESSOA = PFUNC.CODPESSOA " & _
                  " ORDER  BY ATEST2.DTINICIO DESC) AS DTFINAL "
                  
          sSQL = sSQL & _
                  "FROM VATESTADO " & _
                  " INNER JOIN PFUNC " & _
                  " ON  VATESTADO.CODPESSOA = PFUNC.CODPESSOA " & _
                  " AND VATESTADO.CODCOLIGADA = PFUNC.CODCOLIGADA "
                  
         sSQL = sSQL & _
                  "WHERE (PFUNC.CODSITUACAO <> 'D') AND (PFUNC.CODCOLIGADA = 1) " & _
                  "AND (PFUNC.CODSITUACAO <> 'T') AND (PFUNC.CODSITUACAO <> 'P') " & _
                  "AND (PFUNC.CODSITUACAO <> 'O') AND (PFUNC.CODSITUACAO <> 'E') " & _
                  " GROUP BY PFUNC.CHAPA,PFUNC.CODPESSOA " & _
                  " ORDER BY PFUNC.CHAPA,PFUNC.CODPESSOA "
                  
    Case 21 'SELECIONAR O REGISTRO NA TABELA GED_TAB_PARAMETRO A ULTIMA DATA PROCESSADA DAS HORAS EXTRAS VINDAS DO RM
    
           sSQL = "SELECT " & _
                  "TPR_CONTEUDO "
           sSQL = sSQL & _
                  "FROM GED_TAB_PARAMETRO "
           sSQL = sSQL & _
                  "WHERE TPR_NOME_TAB ='" & cFields(1) & "' " & _
                  "AND   TPR_ATRIBUTO ='" & cFields(2) & "' " & _
                  "AND   TPR_VALOR = " & cFields(3)

    Case 22 'SELECIONAR O REGISTRO NA TABELA AMOVFUNDIA COM AS HORAS BATIDAS NO PONTO ELETRONICO A PARTIR DE UMA DATA
           sSQL = "SELECT " & _
                  "AAFHTFUN.CHAPA, " & _
                  "AAFHTFUN.DATA, " & _
                  "AAFHTFUN.EXTRAAUTORIZADO AS NUMHORAS, "
           sSQL = sSQL & _
                  "(SELECT TOP 1 SUBSTRING(CODSECAO,10,3)  FROM PFHSTSEC " & _
                  " WHERE CODCOLIGADA = " & cFields(1) & " " & _
                  " AND   CHAPA = AAFHTFUN.CHAPA " & _
                  " AND   DTMUDANCA <=  AAFHTFUN.DATA " & _
                  " ORDER BY DTMUDANCA DESC) AS INTEGRGERENCIAL, "
           sSQL = sSQL & _
                  "(SELECT NOME FROM PFUNC WHERE PFUNC.CHAPA = AAFHTFUN.CHAPA AND PFUNC.CODCOLIGADA = '" & cFields(1) & "') " & " AS NOME_FUNCIONARIO, "
           sSQL = sSQL & _
                  "(SELECT PFUNCAO.NOME FROM PFUNCAO,PFUNC WHERE PFUNC.CODFUNCAO = PFUNCAO.CODIGO AND PFUNC.CODCOLIGADA = PFUNCAO.CODCOLIGADA AND PFUNC.CHAPA = AAFHTFUN.CHAPA AND PFUNCAO.CODCOLIGADA = '" & cFields(1) & "') " & " AS NOME_FUNCAO "
           sSQL = sSQL & _
                  "FROM AAFHTFUN "
           sSQL = sSQL & _
                  " WHERE AAFHTFUN.CODCOLIGADA = '" & cFields(1) & "' " & _
                  " AND   CONVERT(NVARCHAR(10),AAFHTFUN.DATA, 121) >= '" & cFields(2) & "' " & _
                  " AND   CONVERT(NVARCHAR(10),AAFHTFUN.DATA, 121) < '" & Format(Now(), "YYYY-MM-DD") & "'" & _
                  " AND   EXTRAAUTORIZADO <> 0 "
           sSQL = sSQL & _
                  " ORDER BY CHAPA, DATA DESC "
    Case 23 ''VERIFICAR EXISTENCIA DO REGISTRO NA TABELA GED_TAB_ITEM_SOLICITACAO, DAS HORAS EXTRAS
    
           sSQL = "SELECT COUNT(*) "
           sSQL = sSQL & "FROM GED_TAB_SOLICITACAO AS SOL, GED_TAB_ITEM_SOLICITACAO AS ISO "
           sSQL = sSQL & _
                  "WHERE SOL.SOL_CODIGO = ISO.ITS_SOL_CODIGO " & _
                  "AND   SOL.SOL_TFF_CODIGO = 1 " & _
                  "AND   SOL.SOL_SITUACAO = 'F' " & _
                  "AND   CONVERT(NVARCHAR(10),ITS_HORA_EXTRA_ENTRADA, 121) ='" & cFields(1) & "' " & _
                  "AND   SOL.SOL_EMP_CODIGO = " & cFields(2) & " " & _
                  "AND   ISO.ITS_USU_CHAPA ='" & cFields(3) & "' "

    
    Case 24 'ATUALIZAR O REGISTRO NO BANCO GET, TABELA GED_TAB_ITEM_SOLICITACAO O VALOR REALIZADO DA HORA EXTRA
           sSQL = "UPDATE GED_TAB_ITEM_SOLICITACAO " & _
                  "SET ITS_HORA_EXTRA_REALIZADA = '" & Replace(cFields(1), "-", "") & " " & cFields(4) & "' "
           
           sSQL = sSQL & "FROM GED_TAB_SOLICITACAO AS SOL, GED_TAB_ITEM_SOLICITACAO AS ISO "
           
           sSQL = sSQL & _
                  "WHERE SOL.SOL_CODIGO = ISO.ITS_SOL_CODIGO " & _
                  "AND   SOL.SOL_TFF_CODIGO = 1 " & _
                  "AND   SOL.SOL_SITUACAO = 'F' " & _
                  "AND   CONVERT(NVARCHAR(10),ITS_HORA_EXTRA_ENTRADA, 121) ='" & cFields(1) & "' " & _
                  "AND   SOL.SOL_EMP_CODIGO = " & cFields(2) & " " & _
                  "AND   ISO.ITS_USU_CHAPA ='" & cFields(3) & "' "
    
    Case 25 'ATUALIZAR O REGISTRO NA TABELA GED_TAB_PARAMETRO COM A ULTIMA DATA PROCESSADA DAS HORAS EXTRAS VINDAS DO RM
    
           sSQL = "UPDATE GED_TAB_PARAMETRO SET TPR_CONTEUDO = '" & cFields(4) & "' "
           sSQL = sSQL & _
                  "WHERE TPR_NOME_TAB ='" & cFields(1) & "' " & _
                  "AND   TPR_ATRIBUTO ='" & cFields(2) & "' " & _
                  "AND   TPR_VALOR = " & cFields(3)

    Case 26 ' RECEBER REGISTRO DA TABELA DO WORKFLOW DOS USUARIOS CORRESPONDENTE AO SETOR COMO PARAMETRO
    
           sSQL = "SELECT " & _
                  "TAW_EMP_CODIGO, " & _
                  "TAW_SETOR_CODIGO, " & _
                  "TAW_USU_NIVEL1, " & _
                  "TAW_USU_NIVEL2, " & _
                  "TAW_USU_NIVEL3, " & _
                  "TAW_USU_NIVEL4 "
           sSQL = sSQL & _
                  "FROM GED_TAB_APROVACAO_WORKFLOW "
           sSQL = sSQL & _
                  "WHERE TAW_SETOR_CODIGO ='" & cFields(2) & "' " & _
                  "AND   TAW_EMP_CODIGO =" & cFields(1)

    Case 27 ' INSERIR NA TABELA DE SOLICITAÇOES, AS MESMAS REALIZADAS SEM SER SOLICITADAS,
    
           sSQL = "INSERT INTO GED_TAB_SOLICITACAO (" & _
                  "SOL_CODIGO_TREINAMENTO, " & _
                  "SOL_DATA_EMISSAO, " & _
                  "SOL_DATA_LIBERAÇAO_NIVEL1, " & _
                  "SOL_DATA_LIBERAÇAO_NIVEL2, " & _
                  "SOL_DATA_LIBERAÇAO_NIVEL3, " & _
                  "SOL_DATA_LIBERAÇAO_NIVEL4, " & _
                  "SOL_DATA_REALIZACAO_TRABALHO_EXTRA, " & _
                  "SOL_IS_FIRST, " & _
                  "SOL_JUSTIFICATIVA_MOTIVO, " & _
                  "SOL_MES_ANO, " & _
                  "SOL_NOME_TREINAMENTO, " & _
                  "SOL_PERCENTUAL_HORA, " & _
                  "SOL_SITUACAO, " & _
                  "SOL_EMP_CODIGO," & _
                  "SOL_TFF_CODIGO, " & _
                  "SOL_TMS_CODIGO, " & _
                  "SOL_TSE_CODIGO, " & _
                  "SOL_USU_CODIGO_NIVEL1, " & _
                  "SOL_USU_CODIGO_NIVEL2, " & _
                  "SOL_USU_CODIGO_NIVEL3, " & _
                  "SOL_USU_CODIGO_NIVEL4, " & _
                  "SOL_USU_CODIGO_SOLICITANTE, " & _
                  "SOL_HORA_EMISSAO, " & _
                  "SOL_IS_RETARDATARIO "
           
           sSQL = sSQL & ") VALUES (" & _
                           cFields(1) & " , " & _
                           cFields(2) & " , " & _
                           cFields(3) & " , " & _
                           cFields(4) & " , " & _
                           cFields(5) & " , " & _
                           cFields(7) & " , " & _
                           cFields(6) & " , " & _
                           cFields(8) & " , " & _
                           cFields(9) & " , " & _
                           cFields(10) & " , " & _
                           cFields(11) & " , " & _
                           cFields(12) & " , " & _
                           cFields(13) & " , " & _
                           cFields(14) & " , " & _
                           cFields(15) & " , " & _
                           cFields(16) & " , " & _
                           cFields(17) & " , " & _
                           cFields(18) & " , " & _
                           cFields(19) & " , " & _
                           cFields(20) & " , " & _
                           cFields(21) & " , " & _
                           cFields(22) & " , " & _
                           cFields(23) & " , " & _
                           cFields(24) & " )"
           
           Set rs = New ADODB.Recordset
           cCONECCAO.CursorLocation = adUseClientBatch
           rs.Open sSQL, cCONECCAO
           cCONECCAO.CommitTrans
           cCONECCAO.BeginTrans
           
           sSQL = "SELECT TOP 1 SOL_CODIGO FROM GED_TAB_SOLICITACAO WHERE SOL_IS_RETARDATARIO='S' AND SOL_SITUACAO='F' AND SOL_TMS_CODIGO=60 ORDER BY SOL_CODIGO DESC"


    Case 28 ' INSERIR NA TABELA DE ITENS DA SOLICITAÇOES, AS MESMAS REALIZADAS SEM SER SOLICITADAS,
    
           sSQL = "INSERT INTO GED_TAB_ITEM_SOLICITACAO (" & _
                  "ITS_ADIANTAMENTO_DECIMO_TERCEIRO, " & _
                  "ITS_USU_CHAPA, " & _
                  "ITS_CHAPA_DESAUTORIZADOR, " & _
                  "ITS_DATA_INICIO_FERIAS, " & _
                  "ITS_DATA_FIM_FERIAS, " & _
                  "ITS_USU_FUNCAO, " & _
                  "ITS_HORA_EXTRA_ENTRADA, " & _
                  "ITS_HORA_EXTRA_SAIDA, " & _
                  "ITS_JUSTIFICATIVA, " & _
                  "ITS_LANCHE, " & _
                  "ITS_USU_NOME, " & _
                  "ITS_REFEICAO," & _
                  "ITS_HORA_EXTRA_REALIZADA, " & _
                  "ITS_USU_TURNO, " & _
                  "ITS_CONDICAO_FERIAS, " & _
                  "ITS_SOL_CODIGO, " & _
                  "ITS_CEIA, " & _
                  "ITS_DESJEJUM, " & _
                  "ITS_JANTA, " & _
                  "ITS_DIFERENCA_SAIDA_ENTRADA, " & _
                  "ITS_BAIRRO, " & _
                  "ITS_CIDADE, " & _
                  "ITS_ENDERECO "
           
           sSQL = sSQL & ") VALUES (" & _
                           cFields(1) & ", " & _
                           cFields(2) & ", " & _
                           cFields(3) & ", " & _
                           cFields(4) & ", " & _
                           cFields(5) & ", " & _
                           cFields(6) & ", " & _
                           cFields(7) & ", " & _
                           cFields(8) & ", " & _
                           cFields(9) & ", " & _
                           cFields(10) & ", " & _
                           cFields(11) & ", " & _
                           cFields(12) & ", " & _
                           cFields(13) & ", " & _
                           cFields(14) & ", " & _
                           cFields(15) & ", " & _
                           cFields(16) & ", " & _
                           cFields(17) & ", " & _
                           cFields(18) & ", " & _
                           cFields(19) & ", " & _
                           cFields(20) & ", " & _
                           cFields(21) & ", " & _
                           cFields(22) & ", " & _
                           cFields(23) & " )"

    Case 29 ' DELETAR REGISTRO DA TABELA DE SOLICITACAO E ITENS NO PERIODO SOLICITADO EM 29-03-2012
    
           sSQL = "DELETE FROM GED_TAB_ITEM_SOLICITACAO WHERE " & _
                  "       ITS_SOL_CODIGO IN ( SELECT SOL_CODIGO FROM GED_TAB_SOLICITACAO " & _
                  "WHERE  SOL_DATA_EMISSAO > '" & cFields(2) & "' " & _
                  "AND    SOL_TMS_CODIGO = 60 )"

           Set rs = New ADODB.Recordset
           cCONECCAO.CursorLocation = adUseClientBatch
           rs.Open sSQL, cCONECCAO
           cCONECCAO.CommitTrans
           cCONECCAO.BeginTrans

           sSQL = "DELETE FROM GED_TAB_SOLICITACAO " & _
                  "WHERE  SOL_DATA_EMISSAO > '" & cFields(2) & "' " & _
                  "AND    SOL_TMS_CODIGO = 60 "
                  
                  
    Case 30 'Consultar CADASTRO DE PESSOAL VERIFICANDO OS QUE NAO ESTEJAM DEMITIDOS
           sSQL = "SELECT * FROM PFUNC "
       
           sSQL = sSQL & _
                  " WHERE PFUNC.CODCOLIGADA = " & cFields(1) & _
                  " AND CodSituacao <> 'D'"
                  
    Case 31 ' RECEBER REGISTRO DA TABELA DO CADASTRO DE VALE DOS FUNCIONARIOS NO VALE TRANSPORTE
    
           sSQL = "SELECT * FROM CADVALE "
           
           sSQL = sSQL & _
                  "WHERE COLIGADA =" & cFields(1) & " " & _
                  "AND   CODFUNC =" & cFields(2) & ""
                  
    Case 32 'INCLUIR DADOS NA TABELA DE CADVALE OS NOVOS FUNCIONARIOS NO VALE TRANSPORTE
           
           sSQL = "INSERT INTO CadVale (" & _
                  "Coligada, " & _
                  "CodFunc, " & _
                  "TipoVale1, " & _
                  "QtdVale1, " & _
                  "TipoVale2, " & _
                  "QtdVale2, " & _
                  "TipoVale3, " & _
                  "QtdVale3, " & _
                  "TipoVale4, " & _
                  "QtdVale4 "
           
           sSQL = sSQL & ") VALUES (" & _
                           cFields(1) & ", " & _
                           cFields(2) & ", " & _
                           cFields(3) & ", " & _
                           cFields(4) & ", " & _
                           cFields(5) & ", " & _
                           cFields(6) & ", " & _
                           cFields(7) & ", " & _
                           cFields(8) & ", " & _
                           cFields(9) & ", " & _
                           cFields(10) & " )"
                  
    Case 33 ' RECEBER REGISTRO DA TABELA PfCodFix QUE CONTE  VERBA 0532, CORRESPONDENTE AOS ONIBUA OPCIONAIS
    
           sSQL = "SELECT * FROM PfCodFix "
           
           sSQL = sSQL & _
                  "WHERE Codevento ='" & cFields(1) & "'"
        
    Case 34 ' RECEBER REGISTRO DA TABELA DO CADASTRO DE VALE DOS FUNCIONARIOS NO VALE TRANSPORTE QUE ACABARAM DE SER INCLUSOS
    
           sSQL = "SELECT * FROM CADVALE "
           
           sSQL = sSQL & _
                  "WHERE COLIGADA =" & cFields(1) & " " & _
                  "AND   CODFUNC =" & cFields(2) & " " & _
                  "AND   TipoVale1 ='*' " & _
                  "AND   TipoVale2 ='*' " & _
                  "AND   TipoVale3 ='*' " & _
                  "AND   TipoVale4 ='*' "

    Case 35 'ATUALIZAR DADOS NA TABELA DE CADVALE OS NOVOS FUNCIONARIOS NO VALE TRANSPORTE
           
           sSQL = "UPDATE CadVale SET "
           sVirgula = ""
           
           If Len(Trim(cFields(3))) > 0 Then
              sSQL = sSQL & " TipoVale1 = " & cFields(3) & _
                    ",QtdVale1 = " & cFields(4)
              sVirgula = ","
           End If
           If Len(Trim(cFields(5))) > 0 Then
              sSQL = sSQL & sVirgula & "TipoVale2 = " & cFields(5) & _
                    ",QtdVale2 = " & cFields(6)
              sVirgula = ","
           End If
           If Len(Trim(cFields(7))) > 0 Then
              sSQL = sSQL & sVirgula & "TipoVale3 = " & cFields(7) & _
                    ",QtdVale3 = " & cFields(8)
              sVirgula = ","
           End If
           If Len(Trim(cFields(9))) > 0 Then
              sSQL = sSQL & sVirgula & "TipoVale4 = " & cFields(9) & _
                    ",QtdVale4 = " & cFields(10)
              sVirgula = ","
           End If

           sSQL = sSQL & " WHERE Coligada = " & cFields(1) & _
                         " AND CodFunc = " & cFields(2) & ""

    Case 36 ' RECEBER REGISTRO DA TABELA PfValeTr, AP/ ATUALIZAR VALES DOS NOVATOS
    
           sSQL = "SELECT * FROM PfValeTr "

    Case 37 ' RECEBER REGISTRO DA TABELA GED, SOLICITACOES
    
    sSQL = "SELECT   " & _
           "ITS.ITS_USU_CHAPA,SOL_EMP_CODIGO, " & _
           "ITS.ITS_USU_NOME, " & _
           "TSE.TSE_NOME, " & _
           "TSE.TSE_CODIGO, " & _
           "SOL.SOL_DATA_REALIZACAO_TRABALHO_EXTRA " & _
           "FROM GED_TAB_ITEM_SOLICITACAO AS ITS, " & _
           "     GED_TAB_SOLICITACAO AS SOL, " & _
           "     GED_TAB_SETOR AS TSE "
           
    sSQL = sSQL & _
           "WHERE ITS.ITS_SOL_CODIGO = SOL.SOL_CODIGO "
    sSQL = sSQL & _
           "AND SOL.SOL_TSE_CODIGO = TSE.TSE_CODIGO "
    sSQL = sSQL & _
           "AND SOL_TMS_CODIGO <> 60 "
    sSQL = sSQL & _
           "AND SOL_SITUACAO = 'F' "
    sSQL = sSQL & _
           "AND ITS_CHAPA_DESAUTORIZADOR IS NULL "
    
    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " AND CONVERT(NVARCHAR(10),SOL.SOL_DATA_REALIZACAO_TRABALHO_EXTRA, 121) >= '" & cFields(1) & "'"
       End If
    End If
    sSQL = sSQL & " ORDER BY ITS_USU_CHAPA"

    Case 38 ' EXISTENCIA DE MOVIMENTAÇÃO NO CADASTRO DE VT

     sSQL = "SELECT * FROM MovVale "
     sSQL = sSQL & _
            "WHERE CODFUNC = " & cFields(2) & _
            " AND COLIGADA = " & cFields(1) & _
            " AND CONVERT(NVARCHAR(10),Data, 121) = '" & cFields(3) & "'"

    Case 39 ' EXISTENCIA DO PFUNC NO CADASTRO

     sSQL = "SELECT * FROM PFUNC "
     sSQL = sSQL & _
            "WHERE CODCOLIGADA = " & cFields(1) & _
            " AND CHAPA = " & cFields(2) & _
            " AND CodSituacao <> 'D' AND CODHORARIO <> 'ISENTO'"

    Case 40 'Consultar cadastro de horarios no banco da RM
    
    sSQL = "Select * from Ahorario " & _
           "where codcoligada = 1 "
    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " AND codigo = '" & cFields(1) & "'"
       Else
          sSQL = sSQL & " ORDER BY Codigo"
       End If
    Else
       sSQL = sSQL & " ORDER BY Codigo"
    End If

    Case 41 'Consultar FERIADO, pela data, verificando por mes/ano e dia normal
    
    sSQL = "SELECT   CODIGO, " & _
                    "DESCRICAO " & _
                    "FROM TABFERIADO "
    
    sSQL = sSQL & " WHERE DATA = '" & Replace(cFields(1), "-", "") & "' AND TIPO = 0"

    sSQL = sSQL & " UNION "
    sSQL = sSQL & _
                    "SELECT   CODIGO, " & _
                    "DESCRICAO " & _
                    "FROM TABFERIADO "
    
    sSQL = sSQL & " WHERE MESDIA = '" & Mid$(cFields(1), 6, 2) & Mid$(cFields(1), 9, 2) & "' AND TIPO = 1"
    
    
    Case 42 'inserir novos registros na tabela de movimento de vale transporte
    sSQL = "INSERT INTO MovVale (" & _
           "Coligada, " & _
           "CodFunc, " & _
           "TipoMov, " & _
           "TipoExtra, " & _
           "Data, " & _
           "Tipo1, " & _
           "Qtd1, " & _
           "Tipo2, " & _
           "Qtd2, " & _
           "Recebimento, " & _
           "valpago, " & _
           "sequencial, " & _
           "Login, " & _
           "usu, " & _
           "visto, " & _
           "RecDev, " & _
           "DataDig"

    sSQL = sSQL & ") VALUES (" & _
                    cFields(1) & ", " & _
                    cFields(2) & ", " & _
                    cFields(3) & ", " & _
                    cFields(4) & ", " & _
                    cFields(5) & ", " & _
                    cFields(6) & ", " & _
                    cFields(7) & ", " & _
                    cFields(8) & ", " & _
                    cFields(9) & ", " & _
                    cFields(10) & ", " & _
                    cFields(11) & ", " & _
                    cFields(12) & ", " & _
                    cFields(13) & ", " & _
                    cFields(14) & ", " & _
                    cFields(15) & ", " & _
                    cFields(16) & ", " & _
                    cFields(17) & ")"
           
    Case 43 ' RECEBER REGISTRO DA TABELA DO CADASTRO DE VALE DOS FUNCIONARIOS NO cadvale
    
           sSQL = "SELECT * FROM CADVALE "
           
           sSQL = sSQL & _
                  "WHERE COLIGADA =" & cFields(1) & " " & _
                  "AND   CODFUNC =" & cFields(2) & _
                  "AND TipoVale1 = '*' AND TipoVale2 = '*' AND TipoVale3='*' AND TipoVale4 = '*' "
    
    Case 44 'Consultar TABELA DE CREDENCIAIS VERIFICANDO SE HA O REGISTRO CADASTRADO
           sSQL = "SELECT " & _
                  "E.CD_ESTRUTURA_ORGANIZACIONAL "
           sSQL = sSQL & _
                  "FROM PESSOA AS P " & _
                  "INNER JOIN ESTRUTURA_ORGANIZACIONAL AS E " & _
                  "ON P.CD_ESTRUTURA_ORGANIZACIONAL = E.CD_ESTRUTURA_ORGANIZACIONAL " & _
                  "WHERE P.NU_MATRICULA = " & cFields(1)
    
    Case 45 'Consultar CADASTRO DE SETORES DO RM PARA COMPARAR COM O DO GED E INCLRUIR OS QUE FALTAM
           sSQL = "SELECT " & _
                  "CODCOLIGADA, " & _
                  "CODINTERNO, " & _
                  "DESCRICAO "
           
           sSQL = sSQL & _
                  "FROM PCODINTGERFUNC "
                  
    Case 46 'Consultar CADASTRO DE SETORES DO GED P/INCLRUIR OS QUE FALTAM
           sSQL = "SELECT " & _
                  "TSE_EMP_CODIGO, " & _
                  "TSE_CODIGO, " & _
                  "TSE_NOME "
           
           sSQL = sSQL & _
                  "FROM GED_TAB_SETOR "
           
           sSQL = sSQL & _
                  " WHERE GED_TAB_SETOR.TSE_EMP_CODIGO =" & cFields(1) & _
                  " AND GED_TAB_SETOR.TSE_CODIGO = '" & cFields(2) & "'"
                  
    Case 47 'ATUALIZAR A TABELA DE SETORES VINDO DO RM
           sSQL = "INSERT INTO GED_TAB_SETOR (" & _
                  "TSE_EMP_CODIGO, " & _
                  "TSE_CODIGO, " & _
                  "TSE_NOME"
       
           sSQL = sSQL & ") VALUES (" & _
                           cFields(1) & ", '" & _
                           cFields(2) & "', '" & _
                           cFields(3) & "')"
    
    Case 48 'Consultar CADASTRO DE FUNCOES DO RM PARA COMPARAR COM O DO GED E INCLRUIR OS QUE FALTAM
           sSQL = "SELECT " & _
                  "CODCOLIGADA, " & _
                  "CODIGO, " & _
                  "NOME "
           
           sSQL = sSQL & _
                  "FROM PFUNCAO "
                  
    Case 49 'Consultar CADASTRO DE FUNCOES DO GED P/INCLRUIR OS QUE FALTAM
           sSQL = "SELECT " & _
                  "TFU_EMP_CODIGO, " & _
                  "TFU_CODIGO, " & _
                  "TFU_NOME "
           
           sSQL = sSQL & _
                  "FROM GED_TAB_FUNCAO "
           
           sSQL = sSQL & _
                  " WHERE GED_TAB_FUNCAO.TFU_EMP_CODIGO =" & cFields(1) & _
                  " AND GED_TAB_FUNCAO.TFU_CODIGO = '" & cFields(2) & "'"
                  
    Case 50 'ATUALIZAR A TABELA DE FUNCOES VINDO DO RM
           sSQL = "INSERT INTO GED_TAB_FUNCAO (" & _
                  "TFU_EMP_CODIGO, " & _
                  "TFU_CODIGO, " & _
                  "TFU_NOME "
       
           sSQL = sSQL & ") VALUES (" & _
                           cFields(1) & ",'" & _
                           cFields(2) & "', '" & _
                           cFields(3) & "')"
    
    Case 51 'SELECIONAR EXISTENCIA DO REGISTRO NA TABELA GED_TAB_ITEM_SOLICITACAO, DAS HORAS EXTRAS
    
           sSQL = "SELECT * "
           sSQL = sSQL & "FROM GED_TAB_SOLICITACAO AS SOL, GED_TAB_ITEM_SOLICITACAO AS ISO "
           sSQL = sSQL & _
                  "WHERE SOL.SOL_CODIGO = ISO.ITS_SOL_CODIGO " & _
                  "AND   SOL.SOL_SITUACAO = 'F' " & _
                  "AND   CONVERT(NVARCHAR(10),SOL_DATA_EMISSAO, 121) >='" & cFields(1) & "' " & _
                  "AND   SOL.SOL_EMP_CODIGO = " & cFields(2)
                  
'    Case 52 'SELECIONAR REGISTRO DAS SOLICITAÇÕES E COMPARAR COM OS HORARIOS, CASO TENHA DIVERGENCIA, ATUALIZAR.
'
'           sSQL = "SELECT * "
'           sSQL = sSQL & "FROM GED_TAB_SOLICITACAO AS SOL, GED_TAB_ITEM_SOLICITACAO AS ISO "
'           sSQL = sSQL & _
'                  "WHERE SOL.SOL_CODIGO = ISO.ITS_SOL_CODIGO " & _
'                  "AND   CONVERT(NVARCHAR(10),SOL.SOL_DATA_EMISSAO, 121) = '" & cFields(1) & "'"
                  
    Case 52 'SELECIONAR REGISTRO DA BATIDA DE UM FUNCIONARIO PARA COMPARAR COM O GED
           
           sSQL = "SELECT " & _
                  "AAFHTFUN.CHAPA, " & _
                  "AAFHTFUN.DATA, " & _
                  "AAFHTFUN.EXTRAAUTORIZADO AS NUMHORAS, "
           sSQL = sSQL & _
                  "(SELECT TOP 1 SUBSTRING(CODSECAO,10,3)  FROM PFHSTSEC " & _
                  " WHERE CODCOLIGADA = " & cFields(1) & " " & _
                  " AND   CHAPA = AAFHTFUN.CHAPA " & _
                  " AND   DTMUDANCA <=  AAFHTFUN.DATA " & _
                  " ORDER BY DTMUDANCA DESC) AS INTEGRGERENCIAL, "
           sSQL = sSQL & _
                  "(SELECT NOME FROM PFUNC WHERE PFUNC.CHAPA = AAFHTFUN.CHAPA AND PFUNC.CODCOLIGADA = '" & cFields(1) & "') " & " AS NOME_FUNCIONARIO, "
           sSQL = sSQL & _
                  "(SELECT PFUNCAO.NOME FROM PFUNCAO,PFUNC WHERE PFUNC.CODFUNCAO = PFUNCAO.CODIGO AND PFUNC.CODCOLIGADA = PFUNCAO.CODCOLIGADA AND PFUNC.CHAPA = AAFHTFUN.CHAPA AND PFUNCAO.CODCOLIGADA = '" & cFields(1) & "') " & " AS NOME_FUNCAO "
           sSQL = sSQL & _
                  "FROM AAFHTFUN "
           sSQL = sSQL & _
                  " WHERE AAFHTFUN.CODCOLIGADA = '" & cFields(1) & "' " & _
                  " AND   CONVERT(NVARCHAR(10),AAFHTFUN.DATA, 121) = '" & cFields(2) & "' " & _
                  " AND   AAFHTFUN.CHAPA = '" & cFields(3) & "'" & _
                  " AND   EXTRAAUTORIZADO <> 0 "
                  
     Case 53 ' atualizar os itens da solicitacao, com os dados da nova batida
           sSQL = "UPDATE GED_TAB_ITEM_SOLICITACAO " & _
                  "SET ITS_HORA_EXTRA_REALIZADA = " & cFields(2) & " "
           sSQL = sSQL & _
                  "WHERE ITS_CODIGO = " & cFields(1) & " "
     
     Case 54 ' Email de horas extras extrapoladas ou não autorizadas
           sSQL = "SELECT SOL.SOL_TSE_CODIGO,SOL.SOL_TMS_CODIGO, " & _
                  "(SELECT TMS_DESCRICAO FROM GED_TAB_MOTIVO_SOLICITACAO  " & _
                  " WHERE TMS_EMP_CODIGO = " & cFields(1) & _
                  " AND TMS_ID = SOL.SOL_TMS_CODIGO) AS DESC_MOTIVO, " & _
                  "(SELECT USU_NOME FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " (SELECT TAW_USU_NIVEL1 FROM GED_TAB_APROVACAO_WORKFLOW AS GTW " & _
                  " WHERE GTW.TAW_SETOR_CODIGO = SOL.SOL_TSE_CODIGO)) AS NOME_N1, " & _
                  "(SELECT USU_EMAIL FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " (SELECT TAW_USU_NIVEL1 FROM GED_TAB_APROVACAO_WORKFLOW AS GTW  " & _
                  " WHERE GTW.TAW_SETOR_CODIGO = SOL.SOL_TSE_CODIGO)) AS EMAIL_N1, " & _
                  "(SELECT USU_NOME FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " (SELECT TAW_USU_NIVEL2 FROM GED_TAB_APROVACAO_WORKFLOW AS GTW  " & _
                  " WHERE GTW.TAW_SETOR_CODIGO = SOL.SOL_TSE_CODIGO)) AS NOME_N2,  " & _
                  "(SELECT USU_EMAIL FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " (SELECT TAW_USU_NIVEL2 FROM GED_TAB_APROVACAO_WORKFLOW AS GTW  " & _
                  " WHERE GTW.TAW_SETOR_CODIGO = SOL.SOL_TSE_CODIGO)) AS EMAIL_N2, " & _
                  "ISO.ITS_USU_CHAPA, " & _
                  "USU_NOME AS NOME_USU, " & _
                  "(SELECT TFU_NOME FROM GED_TAB_FUNCAO WHERE TFU_CODIGO = USU.USU_FUNCAO_CODIGO) AS NOME_FUNCAO, "
          sSQL = sSQL & _
                  "CONVERT(NVARCHAR(10),ISO.ITS_HORA_EXTRA_ENTRADA, 121) AS DATA, " & _
                  "(RIGHT('00' + CONVERT(VARCHAR, DATEDIFF(HOUR, ISO.ITS_HORA_EXTRA_ENTRADA, ISO.ITS_HORA_EXTRA_SAIDA) % 24), 2) + ':'  " & _
                  "+ RIGHT('00' + CONVERT(VARCHAR, DATEDIFF(MINUTE, ISO.ITS_HORA_EXTRA_ENTRADA, ISO.ITS_HORA_EXTRA_SAIDA) % 60), 2)) AS TOTAL_SOLICITADO, "
          sSQL = sSQL & _
                  "ISO.ITS_HORA_EXTRA_REALIZADA AS DTA_R,CONVERT(NVARCHAR(10),ISO.ITS_HORA_EXTRA_REALIZADA, 108) AS TOTAL_REALIZADO "
          sSQL = sSQL & _
                  "FROM GED_TAB_SOLICITACAO AS SOL, " & _
                  "GED_TAB_ITEM_SOLICITACAO AS ISO, " & _
                  "GED_TAB_USUARIO As USU "
          sSQL = sSQL & _
                  "WHERE SOL.SOL_CODIGO = ISO.ITS_SOL_CODIGO " & _
                  " AND   ISO.ITS_USU_CHAPA = USU.USU_CHAPA " & _
                  " AND   SOL.SOL_EMP_CODIGO = " & cFields(1) & _
                  " AND   ISO.ITS_HORA_EXTRA_REALIZADA IS NOT NULL " & _
                  " AND   ( (SOL.SOL_SITUACAO = 'F' AND SOL.SOL_TMS_CODIGO <> 60) OR (SOL.SOL_TMS_CODIGO = 60) ) " & _
                  " AND   CONVERT(NVARCHAR(10),ISO.ITS_HORA_EXTRA_ENTRADA, 121) =" & cFields(2) & ""
          If cFields(3) = "1" Then
             sSQL = sSQL & " ORDER BY NOME_N1 "
          Else
             sSQL = sSQL & " ORDER BY NOME_N2 "
          End If
          
     Case 55 ' VERIFICAR FERIADO NA TABELA DO RM
           sSQL = "SELECT NOME FROM GFERIADO WHERE CODCALENDARIO = '01' " & _
                  "AND DIAFERIADO = " & cFields(1) & " "

    Case 56 'INSERIR NA TABELA DE GED_TAB_EMAIL
    
           
           sSQL = "INSERT INTO GED_TAB_EMAIL (" & _
                  "TE_ASSUNTO, " & _
                  "TE_ENVIADO, " & _
                  "TE_MENSAGEM, " & _
                  "TE_PARA, " & _
                  "TE_ANEXO, " & _
                  "TE_DATA_CRIACAO, " & _
                  "TE_DATA_ENVIO, " & _
                  "TE_ENVIADO_ST, TE_USUARIO "
           
           sSQL = sSQL & ") VALUES ('" & _
                           cFields(1) & "', " & _
                           cFields(2) & ", '" & _
                           cFields(3) & "', '" & _
                           cFields(4) & "', '" & _
                           cFields(5) & "', '" & _
                           cFields(6) & "', '" & _
                           cFields(7) & "', " & _
                           cFields(8) & " ,'SERVIDOR')"
Rem testar aqui
     Case 57 ' pegar o email dos setores 651 do nivel 1 2 nivel 2
           sSQL = "SELECT TAW_USU_NIVEL1, "
           sSQL = sSQL & _
                  " (SELECT USU_EMAIL FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " TAW_USU_NIVEL1) AS EMAIL_N1, "
           sSQL = sSQL & " TAW_USU_NIVEL2, "
           sSQL = sSQL & _
                  "(SELECT USU_EMAIL FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " TAW_USU_NIVEL2) AS EMAIL_N2 "
           sSQL = sSQL & "FROM GED_TAB_APROVACAO_WORKFLOW AS GTW " & _
                  " WHERE GTW.TAW_SETOR_CODIGO = '" & cFields(2) & "'"
    
     Case 58 ' Email de horas extras extrapoladas ou não autorizadas para envio ao rh de 15 a 16 do outro mes, ver sempre no dia 18 de cada mes.
           sSQL = "SELECT SOL.SOL_TSE_CODIGO,SOL.SOL_TMS_CODIGO, " & _
                  "(SELECT TMS_DESCRICAO FROM GED_TAB_MOTIVO_SOLICITACAO  " & _
                  " WHERE TMS_EMP_CODIGO = " & cFields(1) & _
                  " AND TMS_ID = SOL.SOL_TMS_CODIGO) AS DESC_MOTIVO, " & _
                  "(SELECT USU_NOME FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " (SELECT TAW_USU_NIVEL1 FROM GED_TAB_APROVACAO_WORKFLOW AS GTW " & _
                  " WHERE GTW.TAW_SETOR_CODIGO = SOL.SOL_TSE_CODIGO)) AS NOME_N1, " & _
                  "(SELECT USU_EMAIL FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " (SELECT TAW_USU_NIVEL1 FROM GED_TAB_APROVACAO_WORKFLOW AS GTW  " & _
                  " WHERE GTW.TAW_SETOR_CODIGO = SOL.SOL_TSE_CODIGO)) AS EMAIL_N1, " & _
                  "(SELECT USU_NOME FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " (SELECT TAW_USU_NIVEL2 FROM GED_TAB_APROVACAO_WORKFLOW AS GTW  " & _
                  " WHERE GTW.TAW_SETOR_CODIGO = SOL.SOL_TSE_CODIGO)) AS NOME_N2,  " & _
                  "(SELECT USU_EMAIL FROM GED_TAB_USUARIO WHERE USU_CODIGO =  " & _
                  " (SELECT TAW_USU_NIVEL2 FROM GED_TAB_APROVACAO_WORKFLOW AS GTW  " & _
                  " WHERE GTW.TAW_SETOR_CODIGO = SOL.SOL_TSE_CODIGO)) AS EMAIL_N2, " & _
                  "ISO.ITS_USU_CHAPA, " & _
                  "USU_NOME AS NOME_USU, " & _
                  "(SELECT TFU_NOME FROM GED_TAB_FUNCAO WHERE TFU_CODIGO = USU.USU_FUNCAO_CODIGO) AS NOME_FUNCAO, "
          sSQL = sSQL & _
                  "CONVERT(NVARCHAR(10),ISO.ITS_HORA_EXTRA_ENTRADA, 121) AS DATA, " & _
                  "(RIGHT('00' + CONVERT(VARCHAR, DATEDIFF(HOUR, ISO.ITS_HORA_EXTRA_ENTRADA, ISO.ITS_HORA_EXTRA_SAIDA) % 24), 2) + ':'  " & _
                  "+ RIGHT('00' + CONVERT(VARCHAR, DATEDIFF(MINUTE, ISO.ITS_HORA_EXTRA_ENTRADA, ISO.ITS_HORA_EXTRA_SAIDA) % 60), 2)) AS TOTAL_SOLICITADO, "
          sSQL = sSQL & _
                  "ISO.ITS_HORA_EXTRA_REALIZADA AS DTA_R,CONVERT(NVARCHAR(10),ISO.ITS_HORA_EXTRA_REALIZADA, 108) AS TOTAL_REALIZADO "
          sSQL = sSQL & _
                  "FROM GED_TAB_SOLICITACAO AS SOL, " & _
                  "GED_TAB_ITEM_SOLICITACAO AS ISO, " & _
                  "GED_TAB_USUARIO As USU "
          sSQL = sSQL & _
                  "WHERE SOL.SOL_CODIGO = ISO.ITS_SOL_CODIGO " & _
                  "AND   ISO.ITS_USU_CHAPA = USU.USU_CHAPA " & _
                  "AND   SOL.SOL_EMP_CODIGO = " & cFields(1) & _
                  "AND   ISO.ITS_HORA_EXTRA_REALIZADA IS NOT NULL " & _
                  "AND   ( (SOL.SOL_SITUACAO = 'F' AND SOL.SOL_TMS_CODIGO <> 60) OR (SOL.SOL_TMS_CODIGO = 60) ) " & _
                  "AND   CONVERT(NVARCHAR(10),ISO.ITS_HORA_EXTRA_ENTRADA, 121) >=" & cFields(2) & " " & _
                  "AND   CONVERT(NVARCHAR(10),ISO.ITS_HORA_EXTRA_ENTRADA, 121) <=" & cFields(3)
                  
          sSQL = sSQL & " ORDER BY SOL.SOL_TSE_CODIGO,NOME_USU,ISO.ITS_HORA_EXTRA_ENTRADA "
          
     Case 59 ' Selecionar as notas fiscais de uma determinada data e st = 16
           sSQL = "SELECT DT_ST_PROC,NR_DOCUMENTO_FISCAL "
           sSQL = sSQL & _
                  "FROM NFE_NF " & _
                  " WHERE DM_ST_PROC = '16' " & _
                  " AND   DT_ST_PROC >= '" & cFields(1) & "'"
     Case 60 ' ATUALIZAR as notas fiscais de uma determinada data e st = 16, PARA 99
           sSQL = "UPDATE NFE_NF SET "
           sSQL = sSQL & _
                  " DM_ST_PROC = '99' " & _
                  " WHERE NR_DOCUMENTO_FISCAL = '" & cFields(1) & "'"
    Case 61 'Consultar TABELA DE CREDENCIAL DA RODBEL VERIFICANDO SE HA O REGISTRO CADASTRADO
           sSQL = "SELECT " & _
                  "CD_CREDENCIAL, "
           sSQL = sSQL & "(SELECT CD_PESSOA FROM PESSOA WHERE NU_MATRICULA = " & cFields(1) & ") AS CD_PESSOA "
           sSQL = sSQL & _
                  "FROM CREDENCIAL "
       
           sSQL = sSQL & _
                  " WHERE NU_CREDENCIAL = " & cFields(1)
                  
    Case 62 'Consultar CADASTRO DE PESSOAL VERIFICANDO A situacao t,p,e,o
           sSQL = "SELECT " & _
                  "CONVERT(BIGINT,PFUNC.CHAPA) AS CHAPA, " & _
                  "PFUNC.NOME, " & _
                  "PFUNC.INTEGRGERENCIAL, " & _
                  "PPESSOA.CPF, " & _
                  "PPESSOA.CARTIDENTIDADE "
           
           sSQL = sSQL & _
                  "FROM PFUNC " & _
                  "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO"
       
           sSQL = sSQL & _
                  " WHERE PFUNC.CODCOLIGADA = 1" & _
                  " AND  (PFUNC.CODSITUACAO = 'T' " & _
                  " OR   PFUNC.CODSITUACAO = 'P' " & _
                  " OR   PFUNC.CODSITUACAO = 'O' " & _
                  " OR   PFUNC.CODSITUACAO = 'E') "
    
    End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

'If nAction = 22 Then Call Gerar_Sql_Log(sSQL)

rs.Open sSQL, cCONECCAO

Set RegBanco_dados = rs
Set rs = Nothing

Exit Function

Erro:

Set rs = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Rem *****************************************************************************
Rem *****************************************************************************
Rem FUNCAO PARA REGISTRAR EM ARQUIVO DE LOG DO TIPO TXT ONDE O EXECUTAVEL ESTIVER
Rem AS ATUALIZACOES E ERROS QUE HOUVER NA ATUALIZACAO DOS REGISTROS
Rem *****************************************************************************
Rem *****************************************************************************
Public Function Gerar_Situacao_Log(c_fields As Collection)

Dim Y As Double
Dim nx As Integer

Open App.Path & "\LogRmRodBel.TXT" For Random Access Read Write As #11 Len = Len(sTexto)

Y = LOF(11) / Len(sTexto)

For nx = 1 To c_fields.Count

    sTexto.Texto = c_fields(nx)

    sTexto.FFinal = Chr$(13) + Chr$(10)
    Y = Y + 1
    Put 11, Y, sTexto

Next

Close #11

End Function

Rem *****************************************************************************
Rem *****************************************************************************
Rem FUNCAO PARA REGISTRAR EM ARQUIVO DE LOG DO TIPO TXT ONDE O EXECUTAVEL ESTIVER
Rem AS ATUALIZACOES E ERROS QUE HOUVER NA ATUALIZACAO DOS REGISTROS
Rem *****************************************************************************
Rem *****************************************************************************
Public Function Gerar_Sql_Log(c_fields As String)

Dim Y As Double
Dim nx As Integer

Open App.Path & "\LogMsgSql.TXT" For Random Access Read Write As #12 Len = Len(ArqTextoSql)

Y = LOF(12) / Len(ArqTextoSql)

ArqTextoSql.TextoSql = c_fields

ArqTextoSql.FFinal = Chr$(13) + Chr$(10)
Y = Y + 1
Put 12, Y, ArqTextoSql

Close #12

End Function


