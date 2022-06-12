#Include 'Protheus.ch'
#Include 'fileio.ch'

/*
_____________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+----------+----------+-------+-----------------------+------+----------+¦¦
¦¦¦ Programa ¦ BSINDCAD ¦ Autor ¦ Lincoln Vasconcelos   ¦ Data ¦ 19/04/16 ¦¦¦
¦¦+----------+----------+-------+-----------------------+------+----------+¦¦
¦¦¦Descrição ¦ Rélatorio Base de indicadores, gera um excel com as        ¦¦¦
¦¦¦          ¦ seguintes planilhas (Admissoes, Desligamentos,             ¦¦¦
¦¦¦          ¦ Base Mensal, QLP, Absteismo, Turnover) com os dados dos    ¦¦¦
¦¦¦          ¦ funcionarios.                                              ¦¦¦   
¦¦+----------+------------------------------------------------------------+¦¦
¦¦¦ Uso      ¦                                                            ¦¦¦
¦¦+----------+------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
*/     

User Function BSINDCAD()

Local cPerg := 'RELBSIND' // Grupo de pergunta.
Local aMeses := {} // Meses de saida do relatorio.
Local cEmp // Recebe a tela com os codigos das empresa e filiais.
Local cTitulo := 'Relatorio Base Indicadores' // Titulo da janela de escolhe de parametros de empresas.
Local aEmp // Recebe empresas e filiais.
Local cArqTmpADM // Nome da tabela temporia na system de Admissões.
Local cTmpADM := 'TMPADM' // Nome da tabela temporaria de Admissões.
Local cArqTmpDEM // Nome da tabela temporaria na system de Demissões(Desligamentos).
Local cTmpDEM := 'TMPDEM' // Nome da tabela temporaria de Demissões(Desligamentos).
Local cArqTmpBMS // Nome da tabela temporaria na system de Base mensal.
Local cTmpBMS := 'TMPBMES' // Nome da tabela temporria de Base mensal.
Local cTrbQry := 'TRB' // Nome da tabela nas querys.
Local oExcel := FWMsExcelEx():New() // Instancia objeto que monta excel para impresão do relatorio.

CriaPerg(cPerg) // Cria grupo de perguntas.
If Empty(Pergunte(cPerg,.T.)) // Chama o grupo de perguntas.
	MsgAlert('Rélatorio Cancelado!')
	Return
Else
	aMeses := GrvMeses() // Grava todos os meses selecionados entre os periodos.
EndIf

cEmp := U_TelaEmp(aEmpresas, cTitulo) // Chama tela para seleção de empresas e filiais.
aEmp := U_EmpToArray(cEmp) // Destrincha o resultado da tela em um vetor com os codigos das empresas e filiais.      

/*Static Functions que cria as tabelas temporarias para receber
os dados das WorkSheets(Planilhas no Excel) no excel*/
cArqTmpADM := CriaTmpADM(cTmpADM) // Cria Tabela temporaria para area de Admissões.
cArqTmpDEM := CriaTmpDEM(cTmpDEM) // Cria Tabela temporaria para area de Desligamentos(Demissões).
cArqTmpBMS := CriaTmpBMS(cTmpBMS) // Cria Tabela temporaria para area de Base Mensal.

/*Static Functions que gravam os dados das consultas no banco nas tabelas 
temporarias para as WorkSheets(Planilhas no Excel) no excel*/
GrvTmpADM(aEmp, cTrbQry, cTmpADM, aMeses) // Carrega os dados do Banco e grava na tabela temporaria da area de Admissões.
GrvTmpDEM(aEmp, cTrbQry, cTmpDEM, aMeses) // Carrega os dados do Banco e grava na tabela temporaria de Demissões.
GrvTmpBMS(aEmp, cTrbQry, cTmpBMS, aMeses) // Carrega os dados do Banco e grava na tabela temporaria de Base Mensal.

/*Static Functions que criam, configuram, gravam dados nas WorkSheets(Planilhas no Excel) */
ConfWSAdm(oExcel, cTmpADM) // Configura Worksheet(Planilha) de Admissões.
ConfWSDem(oExcel, cTmpDEM) // Configura Worksheet(Planilha) de Demissões.
ConfWSBms(oExcel, cTmpBMS) // Configura Worksheet(Planilha) de Base Mensal.

oExcel:Activate() // Ativa o Excel.
oExcel:GetXMLFile(Alltrim(mv_par05)+'Base Indicadores(explicativa).xml') // Exportar o Excel para o diretorio especificado.

(cTmpADM)->(dbCloseArea())
(cTmpDEM)->(dbCloseArea())
(cTmpBMS)->(dbCloseArea())

Return 

/*Cria grupo de perguntas para o relatorio.*/
Static Function CriaPerg(cPerg)

Local 	aPergunta 	:= {}

// mv_par01 // De matricula.
// mv_par02 // Até matricula.
// mv_par03 // Da Periodo inicial.                     	
// mv_par04 // Até o Periodo final.            			       					   
// mv_par03 // Caminho do arquivo sem o nome.

Aadd(aPergunta, {"De Matricula","C",6,0,0,"G",{"","","","",""},"","",{"De matricula inicial."}})
Aadd(aPergunta, {"Até Matricula","C",6,0,0,"G",{"","","","",""},"","",{"Até matricula final."}})
Aadd(aPergunta, {"Da Periodo","D",8,0,0,"G",{"","","","",""},"","",{"Periodo inicial."}})
Aadd(aPergunta, {"Ate Periodo","D",8,0,0,"G",{"","","","",""},"","",{"Periodo final."}})
Aadd(aPergunta, {"Caminho do Arquivo xml","C",60,0,0,"G",{"","","","",""},"","",{"Diretorio do arquivo","pasta destino sem o nome.","exemplo: C:\Alura\"}})

U_CriaPerg(cPerg, aPergunta)

Return

/*Static Function que percorre todos os meses
e grava os meses no vetor.*/
Static Function GrvMeses()

Local dDataUM := CTOD("01/"+AllTrim(StrZero(Month(mv_par03),2))+"/"+AllTrim(Str(Year(mv_par03)))) // Primeiro dia do mes.
Local aMeses := {} // Meses de saida do relatorio.

While dDataUM <= mv_par04 // Enquanto a data atual for menor que a ultima data selecionada pelo usuario.
	/* 
	Layout aMeses, vetor usado para Base Mensal.
	aMeses[x][1] // Numero do Mes em caracter ex: '01' -> Janeiro
	aMeses[x][2] // Numero do Ano em caracter ex: '2016'
	aMeses[x][3] // Nome do Mes por extenso em caracter ex: JANEIRO
	aMeses[x][4] // Primeiro dia do mes corrente em data ex: 01/01/2016
	aMeses[x][5] // Ultimo dia do mes corrente em data ex: 31/01/2016
	*/	
	AADD(aMeses,{AllTrim(StrZero(Month(dDataUM),2)), AllTrim(Str(Year(dDataUM))), Upper(AllTrim(MesExtenso(Month(dDataUM)))), dDataUM, Lastday(dDataUM, 0)})
	dDataUM := Lastday(dDataUM, 0)+1 // Retorna o primeiro dia do outro mes
EndDo

Return aMeses

/*Static Function que cria tabela temporaria para a area de Admissões*/
Static Function CriaTmpADM(cTmpADM)

Local aEstrutura := {} // Array com os campos das tabela
Local cArqTmpADM // Nome da tabela 

/*Monta estrutura para area de Admissões */
AADD(aEstrutura,{"NMES"			,"C"						,2							,0							}) // Numero do mês. 
AADD(aEstrutura,{"CODEMP"		,"C"						,2							,0							}) // Codigo da empresa.
AADD(aEstrutura,{"CODFIL"		,"C"						,2							,0							}) // Codigo da filial.
AADD(aEstrutura,{"DESCMES"		,"C"						,10							,0							}) // Nome do mês.
AADD(aEstrutura,{"DESCFILI"		,"C"						,41							,0							}) // Descrião da filial da empresa do funcionario.
AADD(aEstrutura,{"MATRICULA"	,TamSX3("RA_MAT")[3]		,TamSX3("RA_MAT")[1] 	,TamSX3("RA_MAT")[2]		}) // Matricula do Funcionario.
AADD(aEstrutura,{"CPF"			,TamSX3("RA_CIC")[3]		,TamSX3("RA_CIC")[1]+3	,TamSX3("RA_CIC")[2]		}) // CPF do funcionario.
AADD(aEstrutura,{"CODCC"    	,TamSX3("RA_CC")[3]		,TamSX3("RA_CC")[1]		,TamSX3("RA_CC")[2]		}) // Codigo do Centro de Custo do Funcionario.
AADD(aEstrutura,{"DESCCC"		,TamSX3("CTT_DESC01")[3]	,TamSX3("CTT_DESC01")[1]	,TamSX3("CTT_DESC01")[2]	}) // Descrição do Centro de Custo do Funcionario.
AADD(aEstrutura,{"NOME"			,TamSX3("RA_NOME")[3]	,TamSX3("RA_NOME")[1]	,TamSX3("RA_NOME")[2]	}) // Nome do Funcionario.
AADD(aEstrutura,{"FUNCAO"		,TamSX3("RJ_DESC")[3]	,TamSX3("RJ_DESC")[1]	,TamSX3("RJ_DESC")[2]	}) // Descrição da funcão do funcionario.
AADD(aEstrutura,{"DTADMISAO"	,TamSX3("RA_ADMISSA")[3]	,TamSX3("RA_ADMISSA")[1]	,TamSX3("RA_ADMISSA")[2]	}) // Data de Admissão do funcionario.
AADD(aEstrutura,{"DTDEMISAO"	,TamSX3("RA_DEMISSA")[3]	,TamSX3("RA_DEMISSA")[1]	,TamSX3("RA_DEMISSA")[2]	}) // Data de Demissão do funcionario.
AADD(aEstrutura,{"SALARIO"		,TamSX3("RA_SALARIO")[3]	,TamSX3("RA_SALARIO")[1]	,TamSX3("RA_SALARIO")[2]	}) // Salario do Funcionario.

/*Se a tabela ja estiver aberta fecha*/
If Select(cTmpADM) > 0
	(cTmpADM)->(dbCloseArea())
EndIf

cArqTmpADM := CriaTrab(aEstrutura,.T.)
Use (cArqTmpADM) Alias (cTmpADM) New Exclusive  

Index On NMES+CODEMP+CODFIL+MATRICULA To (cArqTmpADM)

Return cArqTmpADM

/*Static Function que cria tabela temporaria para a area de Desligamentos(Demissões)*/
Static Function CriaTmpDEM(cTmpDEM)

Local aEstrutura := {} // Array com os campos das tabela
Local cArqTmpDEM // Nome da tabela 

/*Monta estrutura para area de Admissões */

AADD(aEstrutura,{"NMES"			,"C"						,2							,0							}) // Numero do mês. 
AADD(aEstrutura,{"CODEMP"		,"C"						,2							,0							}) // Codigo da empresa.
AADD(aEstrutura,{"CODFIL"		,"C"						,2							,0							}) // Codigo da filial.
AADD(aEstrutura,{"DESCMES"		,"C"						,10							,0							}) // Nome do mês.
AADD(aEstrutura,{"DESCFILI"		,"C"						,41							,0							}) // Descrião da filial da empresa do funcionario.
AADD(aEstrutura,{"MATRICULA"	,TamSX3("RA_MAT")[3]		,TamSX3("RA_MAT")[1]  	,TamSX3("RA_MAT")[2]		}) // Matricula do Funcionario.
AADD(aEstrutura,{"CPF"			,TamSX3("RA_CIC")[3]		,TamSX3("RA_CIC")[1]+3	,TamSX3("RA_CIC")[2]		}) // CPF do funcionario.
AADD(aEstrutura,{"CODCC"    	,TamSX3("RA_CC")[3]		,TamSX3("RA_CC")[1]		,TamSX3("RA_CC")[2]		}) // Codigo do Centro de Custo do Funcionario.
AADD(aEstrutura,{"DESCCC"		,TamSX3("CTT_DESC01")[3]	,TamSX3("CTT_DESC01")[1]	,TamSX3("CTT_DESC01")[2]	}) // Descrição do Centro de Custo do Funcionario.
AADD(aEstrutura,{"NOME"			,TamSX3("RA_NOME")[3]	,TamSX3("RA_NOME")[1]	,TamSX3("RA_NOME")[2]	}) // Nome do Funcionario.
AADD(aEstrutura,{"FUNCAO"		,TamSX3("RJ_DESC")[3]	,TamSX3("RJ_DESC")[1]	,TamSX3("RJ_DESC")[2]	}) // Descrição da funcão do funcionario.
AADD(aEstrutura,{"DTADMISAO"	,TamSX3("RA_ADMISSA")[3]	,TamSX3("RA_ADMISSA")[1]	,TamSX3("RA_ADMISSA")[2]	}) // Data de Admissão do funcionario.
AADD(aEstrutura,{"DTDEMISAO"	,TamSX3("RA_DEMISSA")[3]	,TamSX3("RA_DEMISSA")[1]	,TamSX3("RA_DEMISSA")[2]	}) // Data de Demissão do funcionario.
AADD(aEstrutura,{"SALARIO"		,TamSX3("RA_SALARIO")[3]	,TamSX3("RA_SALARIO")[1]	,TamSX3("RA_SALARIO")[2]	}) // Salario do Funcionario.

/*Se a tabela ja estiver aberta fecha*/
If Select(cTmpDEM) > 0
	(cTmpDEM)->(dbCloseArea())
EndIf

cArqTmpDEM := CriaTrab(aEstrutura,.T.)
Use (cArqTmpDEM) Alias (cTmpDEM) New Exclusive  

Index On NMES+CODEMP+CODFIL+MATRICULA To (cArqTmpDEM)

Return cArqTmpDEM

/*Static Function que cria tabela temporaria para a area de Base Mensal*/
Static Function CriaTmpBMS(cTmpBMS)

Local aEstrutura := {} // Array com os campos das tabela
Local cArqTmpBMS // Nome da tabela 

/*Monta estrutura para area de Admissões */
AADD(aEstrutura,{"NMES"			,"C"						,2							,0							}) // Numero do mês. 
AADD(aEstrutura,{"CODEMP"		,"C"						,2							,0							}) // Codigo da empresa.
AADD(aEstrutura,{"CODFIL"		,"C"						,2							,0							}) // Codigo da filial.
AADD(aEstrutura,{"DESCMES"		,"C"						,10							,0							}) // Nome do mês.
AADD(aEstrutura,{"DESCFILI"		,"C"						,41							,0							}) // Descrião da filial da empresa do funcionario.
AADD(aEstrutura,{"MATRICULA"	,TamSX3("RA_MAT")[3]		,TamSX3("RA_MAT")[1] 	,TamSX3("RA_MAT")[2]		}) // Matricula do Funcionario.
AADD(aEstrutura,{"CPF"			,TamSX3("RA_CIC")[3]		,TamSX3("RA_CIC")[1]+3	,TamSX3("RA_CIC")[2]		}) // CPF do funcionario.
AADD(aEstrutura,{"CODCC"    	,TamSX3("RA_CC")[3]		,TamSX3("RA_CC")[1]		,TamSX3("RA_CC")[2]		}) // Codigo do Centro de Custo do Funcionario.
AADD(aEstrutura,{"DESCCC"		,TamSX3("CTT_DESC01")[3]	,TamSX3("CTT_DESC01")[1]	,TamSX3("CTT_DESC01")[2]	}) // Descrição do Centro de Custo do Funcionario.
AADD(aEstrutura,{"NOME"			,TamSX3("RA_NOME")[3]	,TamSX3("RA_NOME")[1]	,TamSX3("RA_NOME")[2]	}) // Nome do Funcionario.
AADD(aEstrutura,{"FUNCAO"		,TamSX3("RJ_DESC")[3]	,TamSX3("RJ_DESC")[1]	,TamSX3("RJ_DESC")[2]	}) // Descrição da funcão do funcionario.
AADD(aEstrutura,{"DTADMISAO"	,TamSX3("RA_ADMISSA")[3]	,TamSX3("RA_ADMISSA")[1]	,TamSX3("RA_ADMISSA")[2]	}) // Data de Admissão do funcionario.
AADD(aEstrutura,{"DTDEMISAO"	,TamSX3("RA_DEMISSA")[3]	,TamSX3("RA_DEMISSA")[1]	,TamSX3("RA_DEMISSA")[2]	}) // Data de Demissão do funcionario.
AADD(aEstrutura,{"SITFOLHA"	 	,TamSX3("RA_SITFOLH")[3]	,TamSX3("RA_SITFOLH")[1]	,TamSX3("RA_SITFOLH")[2]	}) // Situação da folha do funcionario(' '->SITUACAO NORMAL, 'A'->AFASTADO TEMP., 'D'->DEMITIDO, 'F'->FERIAS, 'T'->TRANSFERIDO).	
AADD(aEstrutura,{"SALARIO"		,TamSX3("RA_SALARIO")[3]	,TamSX3("RA_SALARIO")[1]	,TamSX3("RA_SALARIO")[2]	}) // Salario do Funcionario.
AADD(aEstrutura,{"DIASUTEIS"	,"N"						,2							,0							}) // Dias Uteis.
AADD(aEstrutura,{"CH"			,"C"						,6							,0							}) // CH -> carga horaria.
AADD(aEstrutura,{"CHMESPREV"	,"C"						,6							,0							}) // CH Mensal prevista (Dias Uteis . CH)
AADD(aEstrutura,{"ABONOAUT"		,"C"						,6							,0							}) // Abono autorizado.
AADD(aEstrutura,{"ATESMED"		,"C"						,6							,0							}) // Atestado medico.
AADD(aEstrutura,{"CHNAOREA"		,"C"						,6							,0							}) // CH Não realizada, soma(Atestado - Horas negativas trimestre - abono)
AADD(aEstrutura,{"ABSTEISMO"	,"C"						,5							,0							}) // Absenteismo, Percentual(CH Mensal prevista / CH não realizada)
AADD(aEstrutura,{"HEPAGA"		,"C"						,6							,0							}) // não sei
AADD(aEstrutura,{"FADESC"		,"C"						,6							,0							}) // Faltas, atrasos descontados.
AADD(aEstrutura,{"HPOSTRIMES"	,"C"						,6							,0							}) // Hora + trimestre. -> não sei
AADD(aEstrutura,{"HNEGTRIMES"	,"C"						,6							,0							}) // Hora - trimestre. -> não sei

/*Se a tabela ja estiver aberta fecha*/
If Select(cTmpBMS) > 0
	(cTmpBMS)->(dbCloseArea())
EndIf

cArqTmpBMS := CriaTrab(aEstrutura,.T.)
Use (cArqTmpBMS) Alias (cTmpBMS) New Exclusive  

Index On NMES+CODEMP+CODFIL+MATRICULA To (cArqTmpBMS)

Return cArqTmpBMS

/*Static Function que carrega é grava os dados 
da area de admissão na tabela temporaria.*/
Static Function GrvTmpADM(aEmp, cTrbQry, cTmpADM, aMeses)

Local query // query
Local nCont // percorre as empresas escolhidas pelo usuario
Local nContFil // percorre as filiais escolhidas pelo usuario 
Local cFiliais := "" // salva as filiais escolhidas para uso em query

For nCont := 1 to Len(aEmp)

	// Adiciona as filiais no SQL	
	cFiliais := ""
	For nContFil := 1 to Len(aEmp[nCont][2])		
		If Empty(cFiliais)
			cFiliais := "('"  + aEmp[nCont][2][nContFil]
		Else
			cFiliais += "','" + aEmp[nCont][2][nContFil]
		EndIf
	Next nContFil
	cFiliais += "') "		

	For x := 1 To Len(aMeses)

		cQuery := " SELECT '"+aMeses[x][3]+"' MES, "
		cQuery += "        SRA.RA_FILIAL FILIAL, "
		cQuery += "        SRA.RA_MAT MATRICULA, "
		cQuery += "        SRA.RA_CIC CPF, "
		cQuery += "	     SRA.RA_CC CODCC, "
		cQuery += "	     CTT.CTT_DESC01 DESCCC, "
		cQuery += "	     SRA.RA_NOME NOME, "
		cQuery += "        SRJ.RJ_DESC FUNCAO, "
		cQuery += "        SRA.RA_ADMISSA DTADMISAO, "
		cQuery += "        SRA.RA_DEMISSA DTDEMISAO, "
		cQuery += "        SRA.RA_SALARIO SALARIO "
		cQuery += " FROM "+U_RetSqlEmp(aEmp[nCont][1],"SRA")+" SRA "
		cQuery += " INNER JOIN "+U_RetSqlEmp(aEmp[nCont][1],"CTT")+" CTT ON (SRA.RA_CC = CTT.CTT_CUSTO AND CTT.D_E_L_E_T_ = '') "
		cQuery += " INNER JOIN "+U_RetSqlEmp(aEmp[nCont][1],"SRJ")+" SRJ ON (SRA.RA_CODFUNC = SRJ.RJ_FUNCAO AND SRJ.D_E_L_E_T_ = '') "
		cQuery += " WHERE SRA.RA_FILIAL IN "+cFiliais
		cQuery += "   AND SRA.RA_ADMISSA BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
		cQuery += "   AND SRA.D_E_L_E_T_ = '' "

		// Se já existir a tabela da query entao fecha
		If Select(cTrbQry) > 0		
			(cTrbQry)->(dbCloseArea())
		EndIf
		
		dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)
		
		/*Grava o resultado da query na tabela temporaria*/	
		(cTrbQry)->(dbGoTop())	
		While !(cTrbQry)->(EOF())			
		
			RecLock(cTmpADM,.T.)
			(cTmpADM)->NMES := aMeses[x][1] 
			(cTmpADM)->CODEMP := aEmp[nCont][1]
			(cTmpADM)->CODFIL := (cTrbQry)->FILIAL			
			(cTmpADM)->DESCMES := (cTrbQry)->MES
			(cTmpADM)->DESCFILI  := FWFilialName(aEmp[nCont][1],(cTrbQry)->FILIAL)
			(cTmpADM)->MATRICULA := (cTrbQry)->MATRICULA		 	
			(cTmpADM)->CPF       := Transform(AllTrim((cTrbQry)->CPF), "@R 999.999.999-99")
			(cTmpADM)->CODCC     := (cTrbQry)->CODCC
			(cTmpADM)->DESCCC    := (cTrbQry)->DESCCC
			(cTmpADM)->NOME      := (cTrbQry)->NOME
			(cTmpADM)->FUNCAO    := (cTrbQry)->FUNCAO
			(cTmpADM)->DTADMISAO := STOD((cTrbQry)->DTADMISAO)
			(cTmpADM)->DTDEMISAO := STOD((cTrbQry)->DTDEMISAO)
			(cTmpADM)->SALARIO   := (cTrbQry)->SALARIO
			MsUnLock(cTmpADM)
			
			(cTrbQry)->(dbSkip())
		EndDo

		(cTrbQry)->(dbCloseArea())

	Next x	

Next nCont

Return

/*Static Function que carrega é grava os dados 
da area de Demissões na tabela temporaria.*/
Static Function GrvTmpDEM(aEmp, cTrbQry, cTmpDEM, aMeses)

Local query // query
Local nCont // percorre as empresas escolhidas pelo usuario
Local nContFil // percorre as filiais escolhidas pelo usuario 
Local cFiliais := "" // salva as filiais escolhidas para uso em query

For nCont := 1 to Len(aEmp)

	// Adiciona as filiais no SQL	
	cFiliais := ""
	For nContFil := 1 to Len(aEmp[nCont][2])		
		If Empty(cFiliais)
			cFiliais := "('"  + aEmp[nCont][2][nContFil]
		Else
			cFiliais += "','" + aEmp[nCont][2][nContFil]
		EndIf
	Next nContFil
	cFiliais += "') "		

	For x := 1 To Len(aMeses)	

		cQuery := " SELECT '"+aMeses[x][3]+"' MES, "
		cQuery += "        SRA.RA_FILIAL FILIAL, "
		cQuery += "        SRA.RA_MAT MATRICULA, "
		cQuery += "        SRA.RA_CIC CPF, "
		cQuery += "	     SRA.RA_CC CODCC, "
		cQuery += "	     CTT.CTT_DESC01 DESCCC, "
		cQuery += "	     SRA.RA_NOME NOME, "
		cQuery += "        SRJ.RJ_DESC FUNCAO, "
		cQuery += "        SRA.RA_ADMISSA DTADMISAO, "
		cQuery += "        SRA.RA_DEMISSA DTDEMISAO, "
		cQuery += "        SRA.RA_SALARIO SALARIO "	
		cQuery += " FROM "+U_RetSqlEmp(aEmp[nCont][1],"SRA")+" SRA "
		cQuery += " INNER JOIN "+U_RetSqlEmp(aEmp[nCont][1],"CTT")+" CTT ON (SRA.RA_CC = CTT.CTT_CUSTO AND CTT.D_E_L_E_T_ = '') "
		cQuery += " INNER JOIN "+U_RetSqlEmp(aEmp[nCont][1],"SRJ")+" SRJ ON (SRA.RA_CODFUNC = SRJ.RJ_FUNCAO AND SRJ.D_E_L_E_T_ = '') "
		cQuery += " WHERE SRA.RA_FILIAL IN "+cFiliais
		cQuery += "   AND SRA.RA_DEMISSA BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
		cQuery += "   AND SRA.D_E_L_E_T_ = '' "
	
		// Se já existir a tabela da query entao fecha
		If Select(cTrbQry) > 0		
			(cTrbQry)->(dbCloseArea())
		EndIf
			
		dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)
	
		/*Grava o resultado da query na tabela temporaria*/	
		(cTrbQry)->(dbGoTop())	
		While !(cTrbQry)->(EOF())			
		
			RecLock(cTmpDEM,.T.)
			(cTmpDEM)->NMES := aMeses[x][1] 
			(cTmpDEM)->CODEMP := aEmp[nCont][1]
			(cTmpDEM)->CODFIL := (cTrbQry)->FILIAL			
			(cTmpDEM)->DESCMES := (cTrbQry)->MES			
			(cTmpDEM)->DESCFILI   := FWFilialName(aEmp[nCont][1],(cTrbQry)->FILIAL)
			(cTmpDEM)->MATRICULA  := (cTrbQry)->MATRICULA		 	
			(cTmpDEM)->CPF        := Transform(AllTrim((cTrbQry)->CPF), "@R 999.999.999-99")
			(cTmpDEM)->CODCC      := (cTrbQry)->CODCC
			(cTmpDEM)->DESCCC     := (cTrbQry)->DESCCC
			(cTmpDEM)->NOME       := (cTrbQry)->NOME
			(cTmpDEM)->FUNCAO     := (cTrbQry)->FUNCAO
			(cTmpDEM)->DTADMISAO	 := STOD((cTrbQry)->DTADMISAO)
			(cTmpDEM)->DTDEMISAO  := STOD((cTrbQry)->DTDEMISAO)
			(cTmpDEM)->SALARIO    := (cTrbQry)->SALARIO
			MsUnLock(cTmpDEM)
			
			(cTrbQry)->(dbSkip())
		EndDo
	
		(cTrbQry)->(dbCloseArea())

	Next x

Next nCont

Return 

/*Static Function que carrega e grava os dados 
da area de Base Mensal na tabela temporaria.*/
Static Function GrvTmpBMS(aEmp, cTrbQry, cTmpBMS, aMeses)

Local nCont // percorre as empresas escolhidas pelo usuario
Local nContFil // percorre as filiais escolhidas pelo usuario 
Local cFiliais := "" // salva as filiais escolhidas para uso em query
Local cQuery // Consulta ao banco
Local aFerRCG := {} // Feriado na RCG, para calculo de dias uteis.
Local cFil // Recebe a filial do funcionario atualmente posicionado.
Local cMat // Recebe a matricula do funcionario atualmente posicionado.
Local cTurno // Recebe o turno do funcionario atualmente posicionado.
Local dDemissao // Data de demissão do funcionario.
Local aAfast := {} // Periodo de afastamento, para calculo de dias uteis.
Local aDiasUteis := {} // Quantidade de dias uteis, e quais são os dias uteis.
Local nCH // Recebe a carga horaria diaria do funcionario atualmente posicionado.
Local nCHMesPrev // Recebe a carga horaria mensal prevista pro funcionario posicionado.
Local nSomaAbono // Recebe o total de horas abonadas.
Local nSomaAtest // Recebe o total de horas de atestado em cima dos dias uteis.
Local nAbstsm // Recebe o % do Absenteismo.
Local nHEPaga // Hora extra paga.
Local nFTDesc // Hora Faltas e descontos.
Local nPosNegTri // Hora mais menos trimestre -> pega da regra do relatorio TOTVS -> PONTO ELETRONICO -> RELATORIOS -> BANCO DE HORAS -> RELATORIO DE HORAS(PONR100.PRX)
Local aTransTudo := {} // Recebe as variaveis transformadas em caracteres, para impresão do excel em horas.

For nCont := 1 to Len(aEmp)

	// Adiciona as filiais no SQL
	cFiliais := ""	
	For nContFil := 1 to Len(aEmp[nCont][2])		
		If Empty(cFiliais)
			cFiliais := "('"  + aEmp[nCont][2][nContFil]
		Else
			cFiliais += "','" + aEmp[nCont][2][nContFil]
		EndIf
	Next nContFil
	cFiliais += "') "		
	
	For x := 1 To Len(aMeses)
		
		cQuery := " SELECT DISTINCT "
		cQuery += " '"+aMeses[x][3]+"' MES, "
		cQuery += " SRA.RA_FILIAL FILIAL, " 
	   	cQuery += " SRA.RA_MAT MATRICULA, "
	   	cQuery += " SRA.RA_CIC CPF, "
	   	cQuery += " SRA.RA_CC CODCC, "
	   	cQuery += " CTT.CTT_DESC01 DESCCC, " 
	   	cQuery += " SRA.RA_NOME NOME, " 
       cQuery += " SRJ.RJ_DESC FUNCAO, "
       cQuery += " SRA.RA_ADMISSA DTADMISAO, "
       cQuery += " SRA.RA_DEMISSA DTDEMISAO, "
       cQuery += " SRA.RA_SALARIO SALARIO, "
       cQuery += " SRA.RA_TNOTRAB TURNO   		
 		cQuery += " FROM "+U_RetSqlEmp(aEmp[nCont][1],"SRA")+" SRA " 
 		cQuery += " INNER JOIN "+U_RetSqlEmp(aEmp[nCont][1],"CTT")+" CTT ON "
 		cQuery += " (SRA.RA_CC = CTT.CTT_CUSTO "
  		cQuery += "  AND CTT.D_E_L_E_T_ = '') " 
 		cQuery += " INNER JOIN "+U_RetSqlEmp(aEmp[nCont][1],"SRJ")+" SRJ ON "
 		cQuery += " (SRA.RA_CODFUNC = SRJ.RJ_FUNCAO " 
  		cQuery += "  AND SRJ.D_E_L_E_T_ = '') "  
 		cQuery += " WHERE SRA.RA_FILIAL IN "+cFiliais
 		cQuery += "   AND SRA.RA_MAT BETWEEN '"+mv_par01+"' AND '"+mv_par02+"' "
 		cQuery += "   AND (SRA.RA_DEMISSA >= '"+dTos(aMeses[x][4])+"' OR SRA.RA_DEMISSA = '') " 
   		cQuery += "   AND SRA.RA_CATFUNC NOT IN ('E','A','P') "
   		cQuery += "   AND SRA.D_E_L_E_T_ = '' "
 		cQuery += " ORDER BY SRA.RA_FILIAL, SRA.RA_MAT "						 					
		
		// Se já existir a tabela da query entao fecha
		If Select(cTrbQry) > 0		
			(cTrbQry)->(dbCloseArea())
		EndIf

		dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)			

		(cTrbQry)->(dbGoTop())	
		While !(cTrbQry)->(Eof())
			cFil := (cTrbQry)->FILIAL // Filial.
			cMat := (cTrbQry)->MATRICULA // Matricula.			
			cTurno := (cTrbQry)->TURNO // Turno trabalhado.
			dDemissao := STOD((cTrbQry)->DTDEMISAO) // Data de demissão do Funcionario, para calculo dos dias uteis.
			
			aFerRCG := VeriFerRCG(cFil, cMat, aEmp, nCont, x, aMeses, cTurno) /*Verifica os dias de feriado na tabela RCG por mês*/									
			aAfast := PerAfast(aEmp, nCont, cFil, cMat, aMeses, x) /*Periodo de Afastamento do funcionario, para calculo de dias uteis.*/
			aDiasUteis := CalcDUT(cFil, cMat, aMeses, aFerRCG, aAfast, dDemissao) /*Calcula a quantidade de dias uteis no mês do funcionario.*/			
			nCH := CalcCH(aEmp, nCont, cTurno) /*Calcula a carga horaria do funcionario.*/
			nCHMesPrev := CalcCHMes(aDiasUteis[1][1], nCH) /*Calcula a carga horaria mensal do funcionario.*/
			nSomaAbono := SomaAbono(aEmp, nCont, cFil, cMat, aMeses, x) /*Total de horas abonadas desse funcionario nesse mes.*/			
			nSomaAtest := SomaAtest(aEmp, nCont, cFil, cMat, aMeses, x, aDiasUteis[1][2], nCH) /*Soma os atestados medicos.*/
			nCHNaoReal := CHNaoReal(nSomaAbono, nSomaAtest) /*Horas não realizadas, Abonos + Atetados medicos*/ 
			nAbstsm := Abstsm(nCHMesPrev, nCHNaoReal) // Percentual do Absenteismo.											
			nHEPaga := HEPaga(aEmp, nCont, cFil, cMat, aMeses, x) // Hora extra paga ao funcionario.
			nFTDesc := HFTDesc(aEmp, nCont, cFil, cMat, aMeses, x) // Hora Falta e desconto.
			nPosNegTri := PosNegTri(cFil, cMat, aMeses, x) // Hora mais trimestre -> pega da regra do relatorio TOTVS -> PONTO ELETRONICO -> RELATORIOS -> BANCO DE HORAS -> RELATORIO DE HORAS(PONR100.PRX)
			
			aTransTudo := TransTudo(nCH, nCHMesPrev, nSomaAbono, nSomaAtest, nCHNaoReal, nAbstsm, nHEPaga, nFTDesc, nPosNegTri) // Transforma todas as horas pra impressão do excel.						
			
			RecLock(cTmpBMS,.T.)
			(cTmpBMS)->NMES := aMeses[x][1] 
			(cTmpBMS)->CODEMP := aEmp[nCont][1]
			(cTmpBMS)->CODFIL := cFil
			(cTmpBMS)->DESCMES := (cTrbQry)->MES	
			(cTmpBMS)->DESCFILI   := FWFilialName(aEmp[nCont][1],(cTrbQry)->FILIAL)
			(cTmpBMS)->MATRICULA  := cMat	 	
			(cTmpBMS)->CPF        := Transform(AllTrim((cTrbQry)->CPF), "@R 999.999.999-99")
			(cTmpBMS)->CODCC      := (cTrbQry)->CODCC
			(cTmpBMS)->DESCCC     := (cTrbQry)->DESCCC
			(cTmpBMS)->NOME       := (cTrbQry)->NOME
			(cTmpBMS)->FUNCAO     := (cTrbQry)->FUNCAO
			(cTmpBMS)->DTADMISAO	 := STOD((cTrbQry)->DTADMISAO)
			(cTmpBMS)->DTDEMISAO  := STOD((cTrbQry)->DTDEMISAO)
			(cTmpBMS)->SALARIO    := (cTrbQry)->SALARIO
			(cTmpBMS)->DIASUTEIS  := aDiasUteis[1][1]
			(cTmpBMS)->CH := aTransTudo[1]
			(cTmpBMS)->CHMESPREV := aTransTudo[2]
			(cTmpBMS)->ABONOAUT := aTransTudo[3]
			(cTmpBMS)->ATESMED := aTransTudo[4]
			(cTmpBMS)->CHNAOREA := aTransTudo[5]
			(cTmpBMS)->ABSTEISMO := aTransTudo[6]				
			(cTmpBMS)->HEPAGA := aTransTudo[7]
			(cTmpBMS)->FADESC := aTransTudo[8]			
			(cTmpBMS)->HPOSTRIMES := aTransTudo[9]
			(cTmpBMS)->HNEGTRIMES := aTransTudo[10]
			MsUnLock(cTmpBMS)
			
			(cTrbQry)->(dbSkip())
		EndDo						
 																		
	Next x
	
	(cTrbQry)->(dbCloseArea())

Next nCont

Return

/*Static Function que cria, configura, grava 
dados na WorkSheet(Planilha) de Admissões*/
Static Function ConfWSAdm(oExcel, cTmpADM)

Local cWorkSheet := 'Admissões' // Recebe o Nome da WorkSheet(Planilha).
Local cTable := 'ADMISSÕES - TAB' // Recebe o Nome da tabela.
Local aLinhaADM := {} // Recebe os dados da tabela temporaria de Admissão linha a linha. 

oExcel:AddWorkSheet(cWorkSheet) // Adiciona uma WorkSheet(Planilha) no Excel.
oExcel:AddTable(cWorkSheet, cTable) // Adiciona uma tabela dentro da WorkSheet(Planilha) selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'Mes'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Empresa'					,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Matricula'					,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'CPF'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'CC'							,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Desc. Centro de custo'	,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Nome'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Função'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Admissão'					,1,4) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Demissão'					,1,4) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Salário'					,1,2) // Adiciona uma coluna na tabela selecionada.

(cTmpADM)->(dbGoTop())
While !(cTmpADM)->(EOF())

	aLinhaADM := {(cTmpADM)->DESCMES,;
					(cTmpADM)->DESCFILI,; 
				   	(cTmpADM)->MATRICULA,; 
					(cTmpADM)->CPF,;
					(cTmpADM)->CODCC,;
					(cTmpADM)->DESCCC,;
					(cTmpADM)->NOME,;
					(cTmpADM)->FUNCAO,;
					(cTmpADM)->DTADMISAO,;					
					IIF(Empty((cTmpADM)->DTDEMISAO), " ",(cTmpADM)->DTDEMISAO),;
					(cTmpADM)->SALARIO}		

	oExcel:AddRow(cWorkSheet, cTable, aLinhaADM)
	
	(cTmpADM)->(dbSkip())

EndDo

Return

/*Static Function que cria, configura, grava 
dados na WorkSheet(Planilha) de Demissões*/
Static Function ConfWSDem(oExcel, cTmpDEM) 

Local cWorkSheet := 'Desligamentos' // Recebe o Nome da WorkSheet(Planilha).
Local cTable := 'DESLIGAMENTOS - TAB' // Recebe o Nome da tabela.
Local aLinhaDEM := {} // Recebe os dados da tabela temporaria de Demissões linha a linha. 

oExcel:AddWorkSheet(cWorkSheet) // Adiciona uma WorkSheet(Planilha) no Excel.
oExcel:AddTable(cWorkSheet, cTable) // Adiciona uma tabela dentro da WorkSheet(Planilha) selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'Mes'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Empresa'					,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Matricula'					,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'CPF'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'CC'							,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Desc. Centro de custo'	,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Nome'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Função'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Admissão'					,1,4) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Demissão'					,1,4) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Salário'					,1,2) // Adiciona uma coluna na tabela selecionada.

(cTmpDEM)->(dbGoTop())
While !(cTmpDEM)->(EOF())

	aLinhaDEM := {(cTmpDEM)->DESCMES,;
					(cTmpDEM)->DESCFILI,; 
				   	(cTmpDEM)->MATRICULA,; 
					(cTmpDEM)->CPF,;
					(cTmpDEM)->CODCC,;
					(cTmpDEM)->DESCCC,;
					(cTmpDEM)->NOME,;
					(cTmpDEM)->FUNCAO,;
					(cTmpDEM)->DTADMISAO,;
					(cTmpDEM)->DTDEMISAO,;
					(cTmpDEM)->SALARIO}		

	oExcel:AddRow(cWorkSheet, cTable, aLinhaDEM)
	
	(cTmpDEM)->(dbSkip())

EndDo

Return

/*Static Function que cria, configura, grava 
dados na WorkSheet(Planilha) de Base Mensal*/
Static Function ConfWSBms(oExcel, cTmpBMS) 

Local cWorkSheet := 'Base Mensal' // Recebe o Nome da WorkSheet(Planilha).
Local cTable := 'BASE MENSAL - TAB' // Recebe o Nome da tabela.
Local aLinhaBMS := {} // Recebe os dados da tabela temporaria de Demissões linha a linha. 

oExcel:AddWorkSheet(cWorkSheet) // Adiciona uma WorkSheet(Planilha) no Excel.
oExcel:AddTable(cWorkSheet, cTable) // Adiciona uma tabela dentro da WorkSheet(Planilha) selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Mes'							,1,1) // Adiciona uma coluna na tabela selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'Empresa'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Matricula'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'CPF'							,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'CC'								,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Desc. Centro de custo'		,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Nome'							,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Função'							,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Admissão'						,1,4) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Demissão'						,1,4) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Salário'						,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'Dias uteis'					,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'C.H'							,1,1) // Adiciona uma coluna na tabela selecionada.
oExcel:AddColumn(cWorkSheet, cTable, 'C.H Mensal prevista'			,1,1) // Adiciona uma coluna na tabela selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'Abono autorizado'        	,1,1) // Adiciona uma coluna na tabela selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'Atestado medico'         	,1,1) // Adiciona uma coluna na tabela selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'C.H. não realizada'      	,1,1) // Adiciona uma coluna na tabela selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'Absenteismo'             	,1,1) // Adiciona uma coluna na tabela selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'HE Paga'                 	,1,1) // Adiciona uma coluna na tabela selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'Faltas/Atrasos descontados'	,1,1) // Adiciona uma coluna na tabela selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'H ( + ) Trimestre'       	,1,1) // Adiciona uma coluna na tabela selecionada. 
oExcel:AddColumn(cWorkSheet, cTable, 'H ( - ) Trimestre'       	,1,1) // Adiciona uma coluna na tabela selecionada. 

(cTmpBMS)->(dbGoTop())
While !(cTmpBMS)->(EOF())

	aLinhaBMS := {(cTmpBMS)->DESCMES,;
					(cTmpBMS)->DESCFILI,; 
				   	(cTmpBMS)->MATRICULA,; 
					(cTmpBMS)->CPF,;
					(cTmpBMS)->CODCC,;
					(cTmpBMS)->DESCCC,;
					(cTmpBMS)->NOME,;
					(cTmpBMS)->FUNCAO,;
					(cTmpBMS)->DTADMISAO,;
					IIF(Empty((cTmpBMS)->DTDEMISAO), " ",(cTmpBMS)->DTDEMISAO),;
					(cTmpBMS)->SALARIO,;
					(cTmpBMS)->DIASUTEIS,;
					(cTmpBMS)->CH,;
					(cTmpBMS)->CHMESPREV,;
					(cTmpBMS)->ABONOAUT,;
					(cTmpBMS)->ATESMED,;
					(cTmpBMS)->CHNAOREA,;
					(cTmpBMS)->ABSTEISMO,;
					(cTmpBMS)->HEPAGA,; 					
					(cTmpBMS)->FADESC,;
					(cTmpBMS)->HPOSTRIMES,;
					(cTmpBMS)->HNEGTRIMES}		

	oExcel:AddRow(cWorkSheet, cTable, aLinhaBMS)
	
	(cTmpBMS)->(dbSkip())

EndDo

Return

/*------------------------------------Static Function Auxiliares------------------------------------*/

/*Verifica os dias de feriados, final de semana, dias não trabalhados na 
tabela RCG por periodo e turno, para calculo de dias uteis.*/
Static Function VeriFerRCG(cFil, cMat, aEmp, nCont, x, aMeses, cTurno)

Local dDataIni := aMeses[x][4] // recebe o primeiro dia do mês corrente.
Local cQuery // query
Local aFerRCG := {}
Local cTrbQry := "QFERRCG"

/*Percorre todos os dias para verificar se houve troca de turno
para o funcionario, e se houve troca de turno verifica se o dia
atual é um dia util para o funcionario.*/
While dDataIni <= aMeses[x][5]

	/*Se houve uma troca de turno pra esse funcionario,
	verifica se esse dia é util nesse novo turno.*/
	If SPF->(MsSeek(cFil+cMat+dTos(dDataIni))) // se houve troca de turno, considera o turno da tabela, SPF		
		
		cQuery := " SELECT RCG.RCG_DIAMES FROM "+U_RetSqlEmp(aEmp[nCont][1],"RCG")+" RCG " 
		cQuery += "  WHERE RCG.RCG_DIAMES = '"+dTos(dDataIni)+"' " 			
		cQuery += "    AND RCG.RCG_TNOTRA = '"+SPF->PF_TURNOPA+"' "
		cQuery += "    AND RCG.RCG_TIPDIA <> '1' "
		cQuery += "    AND RCG.D_E_L_E_T_ = '' "		
	
	Else // se não houve troca de turno, considera o turno do cadastro de funcionario.
	
		cQuery := " SELECT RCG.RCG_DIAMES FROM "+U_RetSqlEmp(aEmp[nCont][1],"RCG")+" RCG " 
		cQuery += "  WHERE RCG.RCG_DIAMES = '"+dTos(dDataIni)+"' "				
		cQuery += "    AND RCG.RCG_TNOTRA = '"+cTurno+"' "
		cQuery += "    AND RCG.RCG_TIPDIA <> '1' "
		cQuery += "    AND RCG.D_E_L_E_T_ = '' "
	
	EndIf

	// Se já existir a tabela da query entao fecha
	If Select(cTrbQry) > 0		
		(cTrbQry)->(dbCloseArea())
	EndIf
			
	dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)
	
	(cTrbQry)->(dbGoTop())
	While !(cTrbQry)->(Eof())
		AADD(aFerRCG, (cTrbQry)->(RCG_DIAMES)) // Grava feriados, finais de semanas, dias não trabalhados do mês -> Formato AAAAMMDD.
		(cTrbQry)->(dbSkip())
	EndDo
	
	(cTrbQry)->(dbCloseArea())

	dDataIni += 1 // proximo dia.

EndDo


Return aFerRCG

/*Periodo de Afastamento do funcionario, para calculo de dias uteis.*/
Static Function PerAfast(aEmp, nCont, cFil, cMat, aMeses, x)

Local cQuery // query
Local cTrbQry := "QAFA" 
Local aAfast := {}

/*Essa query tras o afastamento do funcionario posicionado nesse mes, 
a regra é a seguinte, se a data do inicio do afastamento ou do fim do
afastamento estiver entre o mes que esta sendo processado e o afastamento 
for diferente de atestado medico, tras o registro, ou se a data do inicio do
afastamento ou do fim do afastamento estiver entre o mes que esta sendo 
processado e o afastamento for igual a atestado medico e a quantidade de
dias uteis for maior igual a 15 tras o registro, ou então se o funcionario 
esta afastado antes do primeiro dia do primeiro mes ate depois do ultimo dia
do ultimo mes.*/
cQuery := " SELECT SR8.R8_TIPO TPAFAST, " 	
cQuery += "        SR8.R8_DATAINI INIAFAST, " 
cQuery += "        SR8.R8_DATAFIM FIMAFAST, "
cQuery += "        SR8.R8_DURACAO NUMDAFAST " 
cQuery += "   FROM "+U_RetSqlEmp(aEmp[nCont][1],"SR8")+" SR8 "	   	 	
cQuery += "  WHERE SR8.R8_FILIAL = '"+cFil+"' " 
cQuery += "    AND SR8.R8_MAT = '"+cMat+"' "
cQuery += "    AND (((SR8.R8_DATAINI BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
cQuery += "     OR SR8.R8_DATAFIM BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"') "
cQuery += "    AND SR8.R8_TIPO <> 'P') "  
cQuery += "     OR ((SR8.R8_DATAINI BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
cQuery += "     OR SR8.R8_DATAFIM BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"') "
cQuery += "    AND (SR8.R8_TIPO = 'P' AND SR8.R8_DURACAO >= 16)) " 
cQuery += "     OR (SR8.R8_DATAINI <= '"+dTos(aMeses[x][4])+"' AND SR8.R8_DATAFIM >= '"+dTos(aMeses[x][5])+"' )) "
cQuery += "    AND SR8.D_E_L_E_T_ = '' "

// Se já existir a tabela da query entao fecha
If Select(cTrbQry) > 0		
	(cTrbQry)->(dbCloseArea())
EndIf

dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)

(cTrbQry)->(dbGoTop())
While !(cTrbQry)->(Eof())
	AADD(aAfast, {(cTrbQry)->TPAFAST, (cTrbQry)->INIAFAST, (cTrbQry)->FIMAFAST, (cTrbQry)->NUMDAFAST})
	(cTrbQry)->(dbSkip())
EndDo		

(cTrbQry)->(dbCloseArea())

Return aAfast

/*Caclculo de dias uteis do funcionario*/
Static Function CalcDUT(cFil, cMat, aMeses, aFerRCG, aAfast, dDemissao) 

/* Layout dos vetores
Layout aMeses, vetor usado para Base Mensal.
aMeses[y][1] // Numero do Mes em caracter ex: '01' -> Janeiro
aMeses[y][2] // Numero do Ano em caracter ex: '2016'
aMeses[y][3] // Nome do Mes por extenso em caracter ex: JANEIRO
aMeses[y][4] // Primeiro dia do mes corrente em data ex: 01/01/2016
aMeses[y][5] // Ultimo dia do mes corrente em data ex: 31/01/2016	

aFerRCG[y] -> Grava feriados, finais de semanas, dias não trabalhados no mês dos funcionarios -> Formato AAAAMMDD.

aAfast[y][1] -> Tipo de Afastamento -> 'P'='Atestado médico'
aAfast[y][2] -> Dia inicial afastamento -> Formato -> 'AAAAMMDD'
aAfast[y][3] -> Dia final do afastamento -> Formato -> 'AAAAMMDD'
aAfast[y][4] -> Número de dias afastados tipo numerico
*/

Local dDataIni := aMeses[x][4] // recebe o primeiro dia do mês corrente.
Local lFlag := .T. // Controla o teste de se o dia é util em alguns dos testes.
Local nDiasUteis := 0 // quantidade de dias uteis desse funcionario nesse mês.
Local cDiasUteis := "" // data dos dias uteis.
Local aDiasUteis := {} // Vetor com os dias a quantidade de dias uteis e a data dos dias uteis.

SPF->(dbSetOrder(1)) // PF_FILIAL+PF_MAT+DTOS(PF_DATA)

/*Percorre todos os dias do mes do dia 1 ao ultimo dia 
e verifica se aquele dia é util para o funcionario.*/
While dDataIni <= aMeses[x][5] 
	
	/*enquanto lFlag for .T. o dia é util*/
	lFlag := .T. // a cada dia lFlag zera para teste novamente.		
	
	/*se a data atual for maior ou igual a 
	demissão do funcionario sai do laço.*/
	If (!Empty(dDemissao) .And. dDataIni >= dDemissao)
		Exit	
	EndIf	
		
	// testa pra ver se o dia é de feriado, DSR, não trabalhado na tabela RCG.		
	If aScan(aFerRCG, dTos(dDataIni)) <> 0 
		lFlag := .F.	
	EndIf
	
	If lFlag // se ainda não entrou em nenhuma das regras acima, testa na regra de afastamento.
	
		// Percorre todos os afastamentos desse funcionario nesse mes.
		For y := 1 To Len(aAfast)
		
			// se a data atual estiver entre o periodo de afastamento eo tipo de afastamento for diferente de atestado medico.
			If aAfast[y][1] <> 'P' .And. (dTos(dDataIni) >= aAfast[y][2] .And. dTos(dDataIni) <= aAfast[y][3]) 
				lFlag := .F. 		
			EndIf	
			
			// se a data atual estiver entre o periodo de afastamento eo tipo de afastamento for igual a atestado medico
			// conta a quantidade de dias uteis faltados de atestado nesse mes depois de 15 dias.
			If aAfast[y][1] == 'P' .And. ((dTos(dDataIni) >= dTos(cTod(aAfast[y][2])+16)) .And. (dTos(dDataIni) <= aAfast[y][3]))   			       				
				lFlag := .F.
			Endif						
		
		Next y
	   
	EndIf
	   	   
	If lFlag // se o dia for util soma.
		nDiasUteis += 1 // Conta dias uteis.
		cDiasUteis += dTos(dDataIni)+'-' // esse dia é util
	EndIf	   	   
	   	   	
	dDataIni += 1 // proximo dia.
		
EndDo

cDiasUteis := Left(cDiasUteis, Rat('-',cDiasUteis)-1)
AADD(aDiasUteis, {nDiasUteis, cDiasUteis})

Return aDiasUteis

/*Calcul a carga horaria do funcionario*/
Static Function CalcCH(aEmp, nCont, cTurno)

Local cQuery // query
Local cTrbQry := "QCH"
Local nCH := 0

cQuery := " SELECT TOP 1 SPJ.PJ_HRSTRAB+SPJ.PJ_HRSTRA2 CH "
cQuery += "   FROM "+U_RetSqlEmp(aEmp[nCont][1],"SPJ")+" SPJ "
cQuery += "  WHERE SPJ.PJ_TURNO = '"+cTurno+"' "
cQuery += "    AND SPJ.PJ_TPDIA = 'S' "
cQuery += "    AND SPJ.D_E_L_E_T_ = '' "	

// Se já existir a tabela da query entao fecha
If Select(cTrbQry) > 0		
	(cTrbQry)->(dbCloseArea())
EndIf

dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)
nCH := (cTrbQry)->CH
(cTrbQry)->(dbCloseArea())

Return nCH

/*Calcula a carga horaria mensal do funcionario.*/
Static Function CalcCHMes(nDiasUteis, nCH)

Local nHoras := 0 // Recebe as horas
Local nMinutos := 0 // Recebe os minutos
Local nDecMin := 0 // Recebe a decimal dos minutos
Local cCHMesPrev := "" // Hora e minuto em caracter
Local nCHMesPrev := 0 // Recebe a hora mensal

nHoras := Int(nCH) // ex: 8,45 -> 8 horas
nMinutos := (nCH - Int(nCH)) * 100 // ex: 8,45 - 8 -> 0.45 * 100 -> 45minutos.

nHoras := nDiasUteis*nHoras // 8*20 = 160 horas
nMinutos := nDiasUteis*nMinutos // 45 * 19 -> 855 minutos
nHoras := nHoras+(Int(nMinutos/60)) // 855/60 -> 14,25 -> 14 Horas + 160 horas -> 174 horas
If nMinutos >= 60 // Se minutos for maior ou igual a 60.
	nMinutos := Mod(nMinutos, 60) // Resto da divisão dos minutos por 60 -> 855/60 -> 15 Minutos.
EndIf
cCHMesPrev := AllTrim(Str(nHoras))+"."+AllTrim(Str(nMinutos)) // Concatena Horas + Minutos -> "174.15"
nCHMesPrev := Val(cCHMesPrev) // Transforma de Caracter pra numero -> 174.15

Return nCHMesPrev

/*Total de horas abonadas do funcionario posicionado no mês*/
Static Function SomaAbono(aEmp, nCont, cFil, cMat, aMeses, x)

Local cQuery // query
Local cTrbQry := "QABO" 
Local nSomaAbono := 0
Local nHoras := 0
Local nMinutos := 0

// SPC -> APONTAMENTOS -> CONTEM ABONO DO MES DE APONTAMENTO EM ABERTO.
// SPH -> 	ACUMULADOS DE APONTAMENTOS -> HISTORICO DOS APONTAMENTOS QUE JA FECHARAM.
cQuery := " SELECT SUM(SOMAABONO) SOMAABONO FROM ( "
cQuery += " SELECT ISNULL(SUM(SPC.PC_QTABONO),0) SOMAABONO FROM "+U_RetSqlEmp(aEmp[nCont][1],"SPC")+" SPC " 
cQuery += "  WHERE SPC.PC_FILIAL = '"+cFil+"' "
cQuery += "    AND SPC.PC_MAT = '"+cMat+"' "
cQuery += "    AND SPC.PC_DATA BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
cQuery += "    AND SPC.PC_ABONO <> '' "
cQuery += "    AND SPC.D_E_L_E_T_ = '' "
cQuery += "  UNION ALL "
cQuery += " SELECT ISNULL(SUM(SPH.PH_QTABONO),0) SOMAABONO FROM "+U_RetSqlEmp(aEmp[nCont][1],"SPH")+" SPH " 
cQuery += "  WHERE SPH.PH_FILIAL = '"+cFil+"' "
cQuery += "    AND SPH.PH_MAT = '"+cMat+"' " 
cQuery += "    AND SPH.PH_DATA BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
cQuery += "    AND SPH.PH_ABONO <> '' "
cQuery += "    AND SPH.D_E_L_E_T_ = '' "
cQuery += " ) TMP "

// Se já existir a tabela da query entao fecha
If Select(cTrbQry) > 0		
	(cTrbQry)->(dbCloseArea())
EndIf

dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)
nSomaAbono := (cTrbQry)->SOMAABONO
(cTrbQry)->(dbCloseArea())

/*Sequencia de condicoes para transformar minutos em horas*/
nHoras := Int(nSomaAbono) // ex: 16,9 -> 16Horas
nMinutos := (nSomaAbono - Int(nSomaAbono)) * 100 // ex: 16,9 - 16 -> 0.9 * 100 -> 90Minutos
nHoras := nHoras+(Int(nMinutos/60)) // ex: 16+((90/60) = 1,5) = 16 + 1 -> 17Horas
If nMinutos >= 60 // Se minutos for maior ou igual a 60.
	nMinutos := Mod(nMinutos, 60) // retorna o resto da divisão, para retirar os minutos corretos se os minutos for acima de 60 ex: 90m/60m resto -> 30min.
EndIf
cSomaAbono := AllTrim(Str(nHoras))+"."+AllTrim(Str(nMinutos)) // ex: "17"+"."+"30" -> "17.30"
nSomaAbono := Val(cSomaAbono) // ex: "17.30" -> 17.30

Return nSomaAbono

/*Soma o total de horas de atestados medicos apenas 
em cima dos dias uteis do funcionario.*/
Static Function SomaAtest(aEmp, nCont, cFil, cMat, aMeses, x, cDiasUteis, nCH) 

Local cQuery // query
Local cTrbQry := "QATES" // query
Local dDataIni := aMeses[x][4] // data inicial
Local nDiasUtAt := 0 // dias uteis em atestado.
Local nHoras := 0 // Soma das Horas.
Local nMinutos := 0 // Soma dos Minutos.
Local cHoras := "" // Horas em caracter.
Local nSomaAtest := 0 // soma dos atestados em horas.

cQuery := " SELECT SR8.R8_TIPO TPAFAST, "  	
cQuery += "        SR8.R8_DATAINI INIAFAST, "
cQuery += "        SR8.R8_DATAFIM FIMAFAST, " 
cQuery += "        SR8.R8_DURACAO NUMDAFAST " 
cQuery += "   FROM "+U_RetSqlEmp(aEmp[nCont][1],"SR8")+" SR8 "	   	 	
cQuery += "  WHERE SR8.R8_FILIAL = '"+cFil+"' "  
cQuery += "    AND SR8.R8_MAT = '"+cMat+"' "
cQuery += "    AND ((SR8.R8_DATAINI BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' " 
cQuery += "     OR SR8.R8_DATAFIM BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"') " 
cQuery += "    AND SR8.R8_TIPO = 'P') "

// Se já existir a tabela da query entao fecha
If Select(cTrbQry) > 0		
	(cTrbQry)->(dbCloseArea())
EndIf

dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)
    
(cTrbQry)->(dbGoTop())    
While !(cTrbQry)->(Eof())

	dDataIni := aMeses[x][4]

	/*Percorre todos os dias do mes para cada registro de 
	afastamento, para verificar se aquele dia do mes é util
	para o funcioanrio.*/
	While dDataIni <= aMeses[x][5] 
		
		/*se o dia que esta sendo lido estiver entre 
		as datas de afastamento por atestado e for
		um dia util para o funcionario, conta+1.*/
		If (dTos(dDataIni) >= (cTrbQry)->INIAFAST .And. dTos(dDataIni) <= (cTrbQry)->FIMAFAST) .And. dTos(dDataIni)$cDiasUteis	
			nDiasUtAt += 1			
		EndIf
		dDataIni += 1		    
	EndDo

	(cTrbQry)->(dbSkip())
EndDo
(cTrbQry)->(dbCloseArea())

nHoras := Int(nCH) // -> 8
nMinutos := (nCH - Int(nCH)) * 100 // (8.45 - 8) = 0.45 * 100 -> 45
nHoras := nHoras*nDiasUtAt // 8*7 -> 56
nMinutos := nMinutos*nDiasUtAt // 45 * 7 -> 315
nHoras := nHoras+(Int(nMinutos/60)) // 56 + 5 -> 61
If nMinutos >= 60 // Se minutos for maior ou igual a 60.
	nMinutos := Mod(nMinutos, 60) // 315/60 - resto 15 -> 15
EndIf
cHoras := AllTrim(Str(nHoras))+"."+AllTrim(Str(nMinutos)) // "61"+"."+"15" -> "61.15" -> horas minutos caracter
nSomaAtest := Val(cHoras) // 60.15, soma dos atestados em horas.
    
Return nSomaAtest  

/*Calcula a soma dos abonos com atestados*/
Static Function CHNaoReal(nSomaAbono, nSomaAtest)

Local nHorasAbo := 0 // Horas Abonos
Local nMinAbo := 0 // Minutos abonos
Local nHorasAte := 0 // Horas atestado
Local nMinAte := 0 // Minutos Atestado
Local nSomaHor := 0 // Soma das horas do abono e do atestado
Local nSomaMin := 0 // Soma os minutos do abono e do atestado
Local cHoras := "" // Horas.
Local nCHNaoReal := 0 // retorna o valor das somas

nHorasAbo := Int(nSomaAbono) // 8.45 -> 8
nMinAbo := (nSomaAbono - nHorasAbo) * 100 // -> 8.45 - 8 = 0.45*100 -> 45 
nHorasAte := Int(nSomaAtest) // 8.45 -> 8
nMinAte := (nSomaAtest - nHorasAte) * 100 // -> 8.45 - 8 = 0.45*100 -> 45
nSomaHor := nHorasAbo + nHorasAte // 8+8 -> 16
nSomaMin := nMinAbo + nMinAte // 45+45 -> 90 
nSomaHor := nSomaHor+(Int(nSomaMin/60)) // 16+(90/60) = 16+1 -> 17
If nSomaMin >= 60 // Se minutos for maior ou igual a 60.
	nSomaMin := Mod(nSomaMin/60) // Resto da divisão dos minutos por 60, para encontrar os minutos.
EndIf
cHoras := AllTrim(Str(nSomaHor))+"."+AllTrim(Str(nSomaMin)) // "17.30"
nCHNaoReal := Val(cHoras) // "17.30" -> 17.30

Return nCHNaoReal

/*Percentual da hora não realizado em cima da hora prevista.*/
Static Function Abstsm(nCHMesPrev, nCHNaoReal)

Local nHoraPrev := 0 // Horas previstas.
Local nMinPrev := 0 // minutos das horas previstas.
Local nHoraNReal := 0 // Horas não realizadas.
Local nMinNReal := 0 // Minutos das horas não realizadas.
Local nAbstsm := 0 // % do Absenteismo.

nHoraPrev := Int(nCHMesPrev) // 83.45 -> 83
nMinPrev := (nCHMesPrev - nHoraPrev) * 100 // 83.45 - 83.00 = 0.45 * 100 -> 45  
nHoraNReal := Int(nCHNaoReal) // 83.45 -> 83
nMinNReal := (nCHNaoReal - nHoraNReal) * 100 // 83.45 - 83.00 = 0.45 * 100 -> 45

nHoraPrev := nHoraPrev * 60 // transforma horas em minutos 83 * 60 -> 4980 
nHoraPrev := nHoraPrev + nMinPrev // 4980 + 45 -> 5025
nHoraNReal := nHoraNReal * 60 // transforma horas em minutos 83 * 60 -> 4980 
nHoraNReal := nHoraNReal + nMinNReal // 4980 + 45 -> 5025
nAbstsm := (nHoraNReal*100) / nHoraPrev // 43*100 / 2000 = 4300 / 2000 -> 2.15
nAbstsm := Round(nAbstsm,0) // Arredonda com duas casas ex 14.25

Return nAbstsm

/*Hora extra paga ao funcionario.*/
Static Function HEPaga(aEmp, nCont, cFil, cMat, aMeses, x) 

Local cQuery // query
Local cVerbasHE := SuperGetMv("MV_VBHEPAG", .F., "('105', '106','160','198','888','890','891')") // Verbas de Horas extras, busca no parametro MV_VBHEPAG
Local cTrbQry := "QHE"
Local nHEPaga := 0 // Hora extra paga.

/*Query para Somar as horas extras do funcionario.*/
cQuery := " SELECT ISNULL(SUM(HE),0) HE FROM ( "
cQuery += " SELECT SRC.RC_HORAS HE FROM "+U_RetSqlEmp(aEmp[nCont][1],"SRC")+" SRC "
cQuery += "  WHERE SRC.RC_FILIAL = '"+cFil+"' "
cQuery += "    AND SRC.RC_MAT = '"+cMat+"' "
cQuery += "    AND SRC.RC_PD IN "+cVerbasHE  
cQuery += "    AND SRC.RC_DATA BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
cQuery += "    AND SRC.D_E_L_E_T_ = '' "
cQuery += "  UNION ALL "
cQuery += " SELECT SRD.RD_HORAS HE FROM "+U_RetSqlEmp(aEmp[nCont][1],"SRD")+" SRD "
cQuery += "  WHERE SRD.RD_FILIAL = '"+cFil+"' "
cQuery += "    AND SRD.RD_MAT = '"+cMat+"' "
cQuery += "    AND SRD.RD_PD IN "+cVerbasHE
cQuery += "    AND SRD.RD_DATPGT BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
cQuery += "    AND SRD.D_E_L_E_T_ = '' "
cQuery += " ) TMP "
   
// Se já existir a tabela da query entao fecha
If Select(cTrbQry) > 0		
	(cTrbQry)->(dbCloseArea())
EndIf

dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)
nHEPaga :=  fConvHr((cTrbQry)->HE, "H") // fConvHr converte Numero decimal para horas, função interna da TOTVS, encontrada no arquivo RHLIBHRS.PRX
(cTrbQry)->(dbCloseArea())

Return nHEPaga

// Hora Falta e desconto.
Static Function HFTDesc(aEmp, nCont, cFil, cMat, aMeses, x)

Local cQuery // query
Local cVerbasFDs := SuperGetMv("MV_VBFDESC", .F., "('409', '412', '423', '509')") // Verbas de Faltas/Descontos, busca no parametro MV_VBFDESC.
Local cTrbQry := "QFD"
Local nFTDesc := 0// Hora Falta e desconto.

/*Query para Somar as faltas/descontos do funcionario.*/
cQuery := " SELECT ISNULL(SUM(FD),0) FD FROM ( "
cQuery += " SELECT SRC.RC_HORAS FD FROM "+U_RetSqlEmp(aEmp[nCont][1],"SRC")+" SRC " 
cQuery += " WHERE SRC.RC_FILIAL = '"+cFil+"' "
cQuery += "   AND SRC.RC_MAT = '"+cMat+"' "
cQuery += "   AND SRC.RC_PD IN "+cVerbasFDs 
cQuery += "   AND SRC.RC_DATA BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
cQuery += "   AND SRC.D_E_L_E_T_ = '' "
cQuery += " UNION ALL " 
cQuery += " SELECT SRD.RD_HORAS FD FROM "+U_RetSqlEmp(aEmp[nCont][1],"SRD")+" SRD "
cQuery += " WHERE SRD.RD_FILIAL = '"+cFil+"' "
cQuery += "   AND SRD.RD_MAT = '"+cMat+"' "
cQuery += "   AND SRD.RD_PD IN "+cVerbasFDs 
cQuery += "   AND SRD.RD_DATPGT BETWEEN '"+dTos(aMeses[x][4])+"' AND '"+dTos(aMeses[x][5])+"' "
cQuery += "   AND SRD.D_E_L_E_T_ = '' "
cQuery += " ) TMP "

// Se já existir a tabela da query entao fecha
If Select(cTrbQry) > 0		
	(cTrbQry)->(dbCloseArea())
EndIf

dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), cTrbQry, .F., .T.)
nFTDesc :=  fConvHr((cTrbQry)->FD, "H") // fConvHr converte Numero decimal para horas, função interna da TOTVS, encontrada no arquivo RHLIBHRS.PRX
(cTrbQry)->(dbCloseArea())

Return nFTDesc

/*Hora mais trimestre -> pega da regra do relatorio 
TOTVS -> PONTO ELETRONICO -> RELATORIOS -> BANCO DE HORAS -> RELATORIO DE HORAS(PONR100.PRX)*/
Static Function PosNegTri(cFil, cMat, aMeses, x)

Local nHoras := 1
Local nSaldo := 0
Local nSaldoAnt := 0
Local nHPosNegTri := 0 // Hora mais menos trimestre -> pega da regra do relatorio TOTVS -> PONTO ELETRONICO -> RELATORIOS -> BANCO DE HORAS -> RELATORIO DE HORAS(PONR100.PRX)
Private dDataAux  := Ctod('') 	//-- Variavel auxiliar para armazenar a ultima data considerada no calculo do Saldo Anterior

dDataAux  := CTOD(SPACE(8))		

SPI->(dbSetOrder(2)) // PI_FILIAL+PI_MAT+Dtos(PI_DATA)+PI_PD
SPI->(dbGoTop()) // Primeiro Registro
SPI->(dbSeek(cFil+cMat)) // posiciona no banco de horas

While SPI->(!Eof()) .And. SPI->(PI_FILIAL+PI_MAT) == cFil+cMat .And. SPI->PI_DATA <= aMeses[x][5]

	PosSP9(SPI->PI_PD,cFil,"P9_TIPOCOD") // não faco ideia do que faz é função TOTVS.
	
	If SP9->P9_TIPOCOD $  "1*3"				
		nValor:=If(SPI->PI_STATUS=="B",0,If(nHoras=1,SPI->PI_QUANT,SPI->PI_QUANTV))
		//-- Para valor nao nulo considera a Data para Referencia do Saldo
	    dDataAux:=If(Empty(nValor),dDataAux,SPI->PI_DATA)
		nSaldoAnt:=__TimeSum(nSaldoAnt,nValor)  			
	Else				
		nValor:=If(SPI->PI_STATUS=="B",0,If(nHoras=1,SPI->PI_QUANT,SPI->PI_QUANTV))
		//-- Para valor nao nulo considera a Data para Referencia do Saldo
		dDataAux:=If(Empty(nValor),dDataAux,SPI->PI_DATA)
		nSaldoAnt:=__TimeSub(nSaldoAnt,nValor)			
	Endif		

	nSaldo := nSaldoAnt

	SPI->(dbSkip())

EndDo

nHPosNegTri := nSaldo

Return nHPosNegTri			

/*Transforma tudo pra hora para impresão no excel.*/
Static Function TransTudo(nCH, nCHMesPrev, nSomaAbono, nSomaAtest, nCHNaoReal, nAbstsm, nHEPaga, nFTDesc, nPosNegTri)

Local nHCH := 0 // Hora do NCH.
Local nMCH := 0 // Minuto do NCH.
Local cCH := " " // Carga Horaria em caracter.

Local nHCHMesP := 0 // Hora Mes Previsto.
Local nMCHMesP := 0 // Minuto Mes Previsto.
Local cCHMesPrev := " " // Hora Mes Previsto.

Local nHSAbono := 0 // Hora abonado.
Local nMSAbono := 0 // Minutos abonado.
Local cSomaAbono := " " // Hora e Minuto abonado.

Local nHSAtest := 0 // Hora Atestado.
Local nMSAtest := 0 // Minutos Atestado.
Local cSomaAtest := " " // Hora + Minuto atestado.

Local nHCHNRea := 0 // Hora Não realizada.
Local nMCHNRea := 0 // Minutos não realizado.
Local cCHNaoReal := " " // Hora + Minuto atestado.

Local cAbstsm := " " // Percentual do Absenteismo.

Local nHHEPaga := 0 // Hora extra paga.
Local nMHEPaga := 0 // Minutos das horas extras pagas.
Local cHEPaga := " " // Hora + Minutos extras

Local nHFTDesc := 0 // Hora Falta e desconto
Local nMFTDesc := 0 // Minuto falta e desconto
Local cFTDesc := " " // Hora + Minutos faltas e descontos

Local nHPosNeg := 0 // Hora +- Trimestre
Local nMPosNeg := 0 // Minutos +- Trimestre
Local cPosTri := " " // Hora + Minutos positivos trimestre
Local cNegTri := " " // Hora + Minutos negativo trimestre

Local aTransTudo := {} // Recebe tudo transformado em caracter.

nHCH := Int(nCH) // 8.45 -> 8
nMCH := (nCH - nHCH) * 100 // 8.45 - 8 = 0.45 * 100 -> 45
cCH := AllTrim(Str(nHCH))+':'+AllTrim(StrZero(nMCH,2)) // '8:45'

nHCHMesP := Int(nCHMesPrev) // 44.45 -> 44
nMCHMesP := (nCHMesPrev - nHCHMesP) * 100 // 44.45 - 44 = 0.45 * 100 -> 45
cCHMesPrev := AllTrim(Str(nHCHMesP))+':'+ AllTrim(StrZero(nMCHMesP,2)) // '44:45'

nHSAbono := Int(nSomaAbono) // Hora abonado.
nMSAbono := (nSomaAbono - nHSAbono) * 100 // Minutos abonado.
cSomaAbono := AllTrim(Str(nHSAbono))+':'+AllTrim(StrZero(nMSAbono,2)) // Hora + Minuto Abonado.

nHSAtest := Int(nSomaAtest) // Hora Atestado.
nMSAtest := (nSomaAtest - nHSAtest) * 100 // Minutos Atestado.
cSomaAtest := AllTrim(Str(nHSAtest))+':'+AllTrim(StrZero(nMSAtest,2)) // Hora + Minuto atestado.

nHCHNRea := Int(nCHNaoReal) // Hora Não realizada.
nMCHNRea := (nCHNaoReal - nHCHNRea) * 100 // Minutos não realizado.
cCHNaoReal := AllTrim(Str(nHCHNRea))+':'+AllTrim(StrZero(nMCHNRea,2)) // Hora + Minuto atestado.

cAbstsm := AllTrim(Str(nAbstsm))+'%' // Percentual do Absenteismo -> '25%'

nHHEPaga := Int(nHEPaga) // Hora extra paga.
nMHEPaga := (nHEPaga-nHHEPaga) * 100 // Minutos das horas extras pagas.
cHEPaga := AllTrim(Str(nHHEPaga))+':'+AllTrim(Str(nMHEPaga)) // Hora + Minutos extras

nHFTDesc := Int(nFTDesc) // Hora Falta e desconto
nMFTDesc := (nFTDesc-nHFTDesc) * 100 // Minuto falta e desconto
cFTDesc := AllTrim(Str(nHFTDesc))+':'+AllTrim(Str(nMFTDesc)) // Hora + Minutos faltas e descontos

If nPosNegTri >= 0
	nHPosNeg := Int(nPosNegTri) // Hora + Trimestre
	nMPosNeg := (nPosNegTri-nHPosNeg) * 100 // Minutos + Trimestre
	cPosTri := AllTrim(Str(nHPosNeg))+':'+AllTrim(Str(nMPosNeg)) // Hora + Minutos positivos trimestre
	cNegTri := "0:0" 
Else
	nPosNegTri := nPosNegTri * -1 // Transforma de negativo pra positivo
	nHPosNeg := Int(nPosNegTri) // Hora + Trimestre
	nMPosNeg := (nPosNegTri-nHPosNeg) * 100 // Minutos + Trimestre
	cNegTri := AllTrim(Str(nHPosNeg))+':'+AllTrim(Str(nMPosNeg)) // Hora + Minutos faltas e descontos
	cPosTri := "0:0" // Hora + Minutos negativo trimestre
EndIf

aTransTudo := {cCH, cCHMesPrev, cSomaAbono, cSomaAtest, cCHNaoReal, cAbstsm, cHEPaga, cFTDesc, cPosTri, cNegTri} // Vetor com as horas transformadas.

Return aTransTudo
