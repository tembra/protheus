#include 'rwmake.ch'
#include 'protheus.ch'
///////////////////////////////////////////////////////////////////////////////
User Function MyRDIEF2()
///////////////////////////////////////////////////////////////////////////////
// Data : 22/09/2014
// User : Thieres Tembra
// Desc : Relat�rio referente ao Anexo II - Substitui��o Tribut�ria da DIEF PA
///////////////////////////////////////////////////////////////////////////////
Local cTitulo := 'Relat�rio Anexo II - DIEF'
Local cPerg := '#MYRDIEF2'

Private _cNoLock := ''

CriaSX1(cPerg)

If !Pergunte(cPerg, .T., cTitulo)
	Return Nil
End If

If MV_PAR01 == Nil .or. MV_PAR01 == CTOD('') .or. MV_PAR02 == Nil .or. MV_PAR02 == CTOD('')
	Alert('Informe as datas para gera��o do relat�rio.')
	Return Nil
ElseIf MV_PAR01 > MV_PAR02
	Alert('A data final deve ser maior que a data inicial.')
	Return Nil
EndIf

//ativa NOLOCK nas queries SQL caso seja referente a um ano anterior do corrente
If Year(MV_PAR02) < Year(Date()) .and. 'MSSQL' $ TCGetDB()
	_cNoLock := 'WITH (NOLOCK)'
EndIf

Processa({|| Executa(cTitulo) },cTitulo,'Aguarde...')

Return Nil

/* ------------------- */

Static Function Executa(cTitulo)

Local aExcel := {}
Local cArq := 'MYRDIEF2'
Local cAux, cRet, cQry
Local cNat, cTip, cOpe, cMun, cCod
Local nTotal
Local nI, nJ, nMaxI, nMaxJ

Private _aDIEF := {}
Private _aAnal := {}

aAdd(aExcel, {cTitulo})
aAdd(aExcel, {'Relat�rio emitido em '+DTOC(Date())+' �s '+Time()+' por '+AllTrim(cUsername)})
aAdd(aExcel, {'Per�odo: '+DTOC(MV_PAR01)+' at� '+DTOC(MV_PAR02)})
aAdd(aExcel, {''})
aAdd(aExcel, {'Empresa: '+cEmpAnt+'/'+cFilAnt+'-'+AllTrim(SM0->M0_NOME)+' / '+AllTrim(SM0->M0_FILIAL)})
aAdd(aExcel, {''})

//09 opera��es no total
ProcRegua(9)

//05 de entrada
Dados('1', '1')
Dados('1', '2')
Dados('1', '3')
Dados('1', '4')
Dados('1', '5')
//04 de sa�da
Dados('2', '1')
Dados('2', '2')
Dados('2', '3')
Dados('2', '4')

//ordenando e adicionando tudo no array do Excel
aSort(_aDIEF,,,{|x,y| x[2]+x[3]+x[4]+x[1]+x[5]+x[6] < y[2]+y[3]+y[4]+y[1]+y[5]+y[6] })
aSort(_aAnal,,,{|x,y| x[2]+x[3]+x[4]+x[1]+x[5]+x[6] < y[2]+y[3]+y[4]+y[1]+y[5]+y[6] })
aAdd(aExcel, {'CPF/CNPJ', 'Natureza', 'Tipo', 'Opera��o', 'UF', 'Munic�pio', 'C�digo DIEF', 'Valor Cont�bil', 'ICMS ST'})
nMaxI := Len(_aDIEF)
For nI := 1 to nMaxI
	aAdd(aExcel, aClone(_aDIEF[nI]))
	
	If MV_PAR03 == 2
		aAdd(aExcel, {'', 'Nota', 'S�rie', 'Nome', 'C�digo', 'Loja', 'Valor Cont�bil', 'ICMS ST'})
		nMaxJ := Len(_aAnal)
		For nJ := 1 to nMaxJ
			If _aDIEF[nI][1] == _aAnal[nJ][1] .and.;
			Left(_aDIEF[nI][2],1) == _aAnal[nJ][2] .and.;
			Left(_aDIEF[nI][3],1) == _aAnal[nJ][3] .and.;
			Left(_aDIEF[nI][4],1) == _aAnal[nJ][4]
				aAdd(aExcel, {'', _aAnal[nJ][5], _aAnal[nJ][6], _aAnal[nJ][7], _aAnal[nJ][8], _aAnal[nJ][9], _aAnal[nJ][10], _aAnal[nJ][11]})
			EndIf
		Next nJ
		aAdd(aExcel, {''})
	EndIf
Next nI

IncProc()

cAux := AllTrim(cGetFile('CSV (*.csv)|*.csv', 'Selecione o diret�rio onde ser� salvo o relat�rio', 1, 'C:\', .T., nOR( GETF_LOCALHARD, GETF_LOCALFLOPPY, GETF_NETWORKDRIVE, GETF_RETDIRECTORY ), .F., .T.))
If cAux <> ''
	cAux := SubStr(cAux, 1, RAt('\', cAux)) + cArq
	cAux := cAux + '-' + DTOS(Date()) + '-' + StrTran(Time(), ':', '') + '.csv'
	
	cRet := U_MyArrCsv(aExcel, cAux, Nil, cTitulo)
	If cRet <> ''
		Alert(cRet)
	EndIf
Else
	Alert('A gera��o do relat�rio foi cancelada!')
EndIf

Return Nil

/* ------------------- */

Static Function Query(cNat, cOpe)

Local cQry := ""
Local cAux1 := ""
Local cAux2 := ""

If cNat == '1'
	cQry += CRLF + " SELECT"
	cQry += CRLF + "    A2_CGC AS CPFCNPJ"
	cQry += CRLF + "   ,A2_NOME AS NOME"
	cQry += CRLF + "   ,FT_ESTADO AS ESTADO"
	cQry += CRLF + "   ,FT_NFISCAL AS NOTA"
	cQry += CRLF + "   ,FT_SERIE AS SERIE"
	cQry += CRLF + "   ,FT_CLIEFOR AS CODIGO"
	cQry += CRLF + "   ,FT_LOJA AS LOJA"
	cQry += CRLF + "   ,LEFT(FT_CFOP,1) AS CFOP"
	cQry += CRLF + "   ,SUM(FT_VALCONT) AS VALOR"
	cQry += CRLF + "   ,SUM(FT_ICMSRET) AS ICMSRET"
	cQry += CRLF + " FROM " + RetSqlName('SFT') + " SFT " + _cNoLock
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SA2') + " SA2 " + _cNoLock
	cQry += CRLF + " ON  SA2.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND A2_FILIAL = '" + xFilial('SA2') + "'"
	cQry += CRLF + " AND A2_COD = FT_CLIEFOR"
	cQry += CRLF + " AND A2_LOJA = FT_LOJA"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SD1') + " SD1 " + _cNoLock
	cQry += CRLF + " ON  SD1.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND D1_FILIAL = '" + xFilial('SD1') + "'"
	cQry += CRLF + " AND D1_DOC = FT_NFISCAL"
	cQry += CRLF + " AND D1_SERIE = FT_SERIE"
	cQry += CRLF + " AND D1_FORNECE = FT_CLIEFOR"
	cQry += CRLF + " AND D1_LOJA = FT_LOJA"
	cQry += CRLF + " AND D1_COD = FT_PRODUTO"
	cQry += CRLF + " AND D1_ITEM = FT_ITEM"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SF4') + " SF4 " + _cNoLock
	cQry += CRLF + " ON  SF4.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND F4_FILIAL = '" + xFilial('SF4') + "'"
	cQry += CRLF + " AND F4_CODIGO = D1_TES"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SB1') + " SB1 " + _cNoLock
	cQry += CRLF + " ON  SB1.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND B1_FILIAL = '" + xFilial('SB1') + "'"
	cQry += CRLF + " AND B1_COD = FT_PRODUTO"
	cQry += CRLF + " WHERE SFT.D_E_L_E_T_ <> '*'"
	cQry += CRLF + "   AND FT_FILIAL = '" + xFilial('SFT') + "'"
	cQry += CRLF + "   AND LEFT(FT_CFOP,1) IN ('1','2')"
	cQry += CRLF + "   AND FT_TIPO NOT IN ('D','B')"
	cQry += CRLF + "   AND FT_ENTRADA BETWEEN '" + DTOS(MV_PAR01) + "' AND '" + DTOS(MV_PAR02) + "'"
	cQry += CRLF + " %cAux1%"
	cQry += CRLF + " GROUP BY"
	cQry += CRLF + "    A2_CGC"
	cQry += CRLF + "   ,A2_NOME"
	cQry += CRLF + "   ,FT_ESTADO"
	cQry += CRLF + "   ,FT_NFISCAL"
	cQry += CRLF + "   ,FT_SERIE"
	cQry += CRLF + "   ,FT_CLIEFOR"
	cQry += CRLF + "   ,FT_LOJA"
	cQry += CRLF + "   ,LEFT(FT_CFOP,1)"
	
	cQry += CRLF + " UNION ALL"
	
	cQry += CRLF + " SELECT"
	cQry += CRLF + "    A1_CGC AS CPFCNPJ"
	cQry += CRLF + "   ,A1_NOME AS NOME"
	cQry += CRLF + "   ,FT_ESTADO AS ESTADO"
	cQry += CRLF + "   ,FT_NFISCAL AS NOTA"
	cQry += CRLF + "   ,FT_SERIE AS SERIE"
	cQry += CRLF + "   ,FT_CLIEFOR AS CODIGO"
	cQry += CRLF + "   ,FT_LOJA AS LOJA"
	cQry += CRLF + "   ,LEFT(FT_CFOP,1) AS CFOP"
	cQry += CRLF + "   ,SUM(FT_VALCONT) AS VALOR"
	cQry += CRLF + "   ,SUM(FT_ICMSRET) AS ICMSRET"
	cQry += CRLF + " FROM " + RetSqlName('SFT') + " SFT " + _cNoLock
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SA1') + " SA1 " + _cNoLock
	cQry += CRLF + " ON  SA1.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND A1_FILIAL = '" + xFilial('SA1') + "'"
	cQry += CRLF + " AND A1_COD = FT_CLIEFOR"
	cQry += CRLF + " AND A1_LOJA = FT_LOJA"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SD1') + " SD1 " + _cNoLock
	cQry += CRLF + " ON  SD1.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND D1_FILIAL = '" + xFilial('SD1') + "'"
	cQry += CRLF + " AND D1_DOC = FT_NFISCAL"
	cQry += CRLF + " AND D1_SERIE = FT_SERIE"
	cQry += CRLF + " AND D1_FORNECE = FT_CLIEFOR"
	cQry += CRLF + " AND D1_LOJA = FT_LOJA"
	cQry += CRLF + " AND D1_COD = FT_PRODUTO"
	cQry += CRLF + " AND D1_ITEM = FT_ITEM"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SF4') + " SF4 " + _cNoLock
	cQry += CRLF + " ON  SF4.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND F4_FILIAL = '" + xFilial('SF4') + "'"
	cQry += CRLF + " AND F4_CODIGO = D1_TES"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SB1') + " SB1 " + _cNoLock
	cQry += CRLF + " ON  SB1.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND B1_FILIAL = '" + xFilial('SB1') + "'"
	cQry += CRLF + " AND B1_COD = FT_PRODUTO"
	cQry += CRLF + " WHERE SFT.D_E_L_E_T_ <> '*'"
	cQry += CRLF + "   AND FT_FILIAL = '" + xFilial('SFT') + "'"
	cQry += CRLF + "   AND LEFT(FT_CFOP,1) IN ('1','2')"
	cQry += CRLF + "   AND FT_TIPO IN ('D','B')"
	cQry += CRLF + "   AND FT_ENTRADA BETWEEN '" + DTOS(MV_PAR01) + "' AND '" + DTOS(MV_PAR02) + "'"
	cQry += CRLF + " %cAux2%"
	cQry += CRLF + " GROUP BY"
	cQry += CRLF + "    A1_CGC"
	cQry += CRLF + "   ,A1_NOME"
	cQry += CRLF + "   ,FT_ESTADO"
	cQry += CRLF + "   ,FT_NFISCAL"
	cQry += CRLF + "   ,FT_SERIE"
	cQry += CRLF + "   ,FT_CLIEFOR"
	cQry += CRLF + "   ,FT_LOJA"
	cQry += CRLF + "   ,LEFT(FT_CFOP,1)"
	
	If cOpe == '1'
		//Entradas Internas/Interestaduais
		//Opera��o 1: Com ICMS Destacado na NF do Substituto Tribut�rio
		//S�o consideradas as notas cujo CST do ICMS for 10 e o campo do Valor do ICMS Retido estiver preenchido.
		cAux1 += CRLF + "   AND F4_SITTRIB = '10'"
		cAux1 += CRLF + "   AND FT_ICMSRET > 0"
	ElseIf cOpe == '2'
		//Entradas Internas/Interestaduais
		//Opera��o 2: Com Antecipa��o na Entrada feita pelo Declarante
		//S�o consideradas as notas cujo campo F4_ANTICMS (TES) esteja como 1-Sim e o
		//campo Valor do ICMS Antecipado estiverem preenchidos.
		cAux1 += CRLF + "   AND F4_ANTICMS = '1'"
		cAux1 += CRLF + "   AND FT_ANTICMS > 0"
	ElseIf cOpe == '3'
		//Entradas Internas/Interestaduais
		//Opera��o 3: Sem Tributa��o em Decorr�ncia de Substitui��o Tribut�ria ou Antecipa��o na Opera��o Antecedente
		//S�o consideradas as notas cujo CST do ICMS for 60.
		cAux1 += CRLF + "   AND F4_SITTRIB = '60'"
	ElseIf cOpe == '4'
		//Entradas Internas/Interestaduais
		//Opera��o 4: Transfer�ncia Interna ou Interestadual n�o Sujeita ao Regime de ST
		//S�o consideradas as notas cujo CFOP for um dos listados abaixo:
		//1151,1152,1153,1154,1408,1409,1552,1557,1658,1659,2151,2152,2153,2154,2408,2409,2552,2557,2658,2659
		cAux1 += CRLF + "   AND FT_CFOP IN ('1151','1152','1153','1154','1408','1409','1552','1557','1658','1659','2151','2152','2153','2154','2408','2409','2552','2557','2658','2659')"
	ElseIf cOpe == '5'
		//Entradas Internas/Interestaduais
		//Opera��o 5: Outras Entradas sem Aplica��o do Regime ST (Liminares e Outras)
		//S�o consideradas as notas cujo campo Cadastro de Produtos B1_PICMENT estiver preenchido,
		//o CST do ICMS for diferente de 10 e 60 e CFOP diferente do citado acima, na lista de transfer�ncia.
		cAux1 += CRLF + "   AND B1_PICMENT > 0"
		cAux1 += CRLF + "   AND F4_SITTRIB NOT IN ('10','60')"
		cAux1 += CRLF + "   AND FT_CFOP NOT IN ('1151','1152','1153','1154','1408','1409','1552','1557','1658','1659','2151','2152','2153','2154','2408','2409','2552','2557','2658','2659')"
	EndIf
	
	If cAux2 == ''
		cAux2 := cAux1
	EndIf
	cQry := StrTran(cQry, '%cAux1%', cAux1)
	cQry := StrTran(cQry, '%cAux2%', cAux2)
ElseIf cNat == '2'
	cQry += CRLF + " SELECT"
	cQry += CRLF + "    A1_CGC AS CPFCNPJ"
	cQry += CRLF + "   ,A1_NOME AS NOME"
	cQry += CRLF + "   ,FT_ESTADO AS ESTADO"
	cQry += CRLF + "   ,FT_NFISCAL AS NOTA"
	cQry += CRLF + "   ,FT_SERIE AS SERIE"
	cQry += CRLF + "   ,FT_CLIEFOR AS CODIGO"
	cQry += CRLF + "   ,FT_LOJA AS LOJA"
	cQry += CRLF + "   ,ZZF_CODIGO AS CODMUN"
	cQry += CRLF + "   ,ZZF_DESC AS MUNDESC"
	cQry += CRLF + "   ,SUM(FT_VALCONT) AS VALOR"
	cQry += CRLF + "   ,SUM(FT_ICMSRET) AS ICMSRET"
	cQry += CRLF + " FROM " + RetSqlName('SFT') + " SFT " + _cNoLock
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SA1') + " SA1 " + _cNoLock
	cQry += CRLF + " ON  SA1.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND A1_FILIAL = '" + xFilial('SA1') + "'"
	cQry += CRLF + " AND A1_COD = FT_CLIEFOR"
	cQry += CRLF + " AND A1_LOJA = FT_LOJA"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SD2') + " SD2 " + _cNoLock
	cQry += CRLF + " ON  SD2.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND D2_FILIAL = '" + xFilial('SD2') + "'"
	cQry += CRLF + " AND D2_DOC = FT_NFISCAL"
	cQry += CRLF + " AND D2_SERIE = FT_SERIE"
	cQry += CRLF + " AND D2_CLIENTE = FT_CLIEFOR"
	cQry += CRLF + " AND D2_LOJA = FT_LOJA"
	cQry += CRLF + " AND D2_COD = FT_PRODUTO"
	cQry += CRLF + " AND D2_ITEM = FT_ITEM"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SF4') + " SF4 " + _cNoLock
	cQry += CRLF + " ON  SF4.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND F4_FILIAL = '" + xFilial('SF4') + "'"
	cQry += CRLF + " AND F4_CODIGO = D2_TES"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SB1') + " SB1 " + _cNoLock
	cQry += CRLF + " ON  SB1.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND B1_FILIAL = '" + xFilial('SB1') + "'"
	cQry += CRLF + " AND B1_COD = FT_PRODUTO"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('ZZF') + " ZZF " + _cNoLock
	cQry += CRLF + " ON  ZZF.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND ZZF_FILIAL = '" + xFilial('ZZF') + "'"
	cQry += CRLF + " AND ZZF_UF = A1_EST"
	cQry += CRLF + " AND ZZF_TIPO = '1'"
	cQry += CRLF + " AND ZZF_CODIGO = A1_DIEF"
	cQry += CRLF + " WHERE SFT.D_E_L_E_T_ <> '*'"
	cQry += CRLF + "   AND FT_FILIAL = '" + xFilial('SFT') + "'"
	cQry += CRLF + "   AND LEFT(FT_CFOP,1) IN ('5')"
	cQry += CRLF + "   AND FT_TIPO NOT IN ('D','B')"
	cQry += CRLF + "   AND FT_ENTRADA BETWEEN '" + DTOS(MV_PAR01) + "' AND '" + DTOS(MV_PAR02) + "'"
	cQry += CRLF + " %cAux1%"
	cQry += CRLF + " GROUP BY"
	cQry += CRLF + "    A1_CGC"
	cQry += CRLF + "   ,A1_NOME"
	cQry += CRLF + "   ,FT_ESTADO"
	cQry += CRLF + "   ,FT_NFISCAL"
	cQry += CRLF + "   ,FT_SERIE"
	cQry += CRLF + "   ,FT_CLIEFOR"
	cQry += CRLF + "   ,FT_LOJA"
	cQry += CRLF + "   ,ZZF_CODIGO"
	cQry += CRLF + "   ,ZZF_DESC"
	
	cQry += CRLF + " UNION ALL"
	
	cQry += CRLF + " SELECT"
	cQry += CRLF + "    A2_CGC AS CPFCNPJ"
	cQry += CRLF + "   ,A2_NOME AS NOME"
	cQry += CRLF + "   ,FT_ESTADO AS ESTADO"
	cQry += CRLF + "   ,FT_NFISCAL AS NOTA"
	cQry += CRLF + "   ,FT_SERIE AS SERIE"
	cQry += CRLF + "   ,FT_CLIEFOR AS CODIGO"
	cQry += CRLF + "   ,FT_LOJA AS LOJA"
	cQry += CRLF + "   ,ZZF_CODIGO AS CODMUN"
	cQry += CRLF + "   ,ZZF_DESC AS MUNDESC"
	cQry += CRLF + "   ,SUM(FT_VALCONT) AS VALOR"
	cQry += CRLF + "   ,SUM(FT_ICMSRET) AS ICMSRET"
	cQry += CRLF + " FROM " + RetSqlName('SFT') + " SFT " + _cNoLock
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SA2') + " SA2 " + _cNoLock
	cQry += CRLF + " ON  SA2.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND A2_FILIAL = '" + xFilial('SA2') + "'"
	cQry += CRLF + " AND A2_COD = FT_CLIEFOR"
	cQry += CRLF + " AND A2_LOJA = FT_LOJA"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SD2') + " SD2 " + _cNoLock
	cQry += CRLF + " ON  SD2.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND D2_FILIAL = '" + xFilial('SD2') + "'"
	cQry += CRLF + " AND D2_DOC = FT_NFISCAL"
	cQry += CRLF + " AND D2_SERIE = FT_SERIE"
	cQry += CRLF + " AND D2_CLIENTE = FT_CLIEFOR"
	cQry += CRLF + " AND D2_LOJA = FT_LOJA"
	cQry += CRLF + " AND D2_COD = FT_PRODUTO"
	cQry += CRLF + " AND D2_ITEM = FT_ITEM"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SF4') + " SF4 " + _cNoLock
	cQry += CRLF + " ON  SF4.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND F4_FILIAL = '" + xFilial('SF4') + "'"
	cQry += CRLF + " AND F4_CODIGO = D2_TES"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('SB1') + " SB1 " + _cNoLock
	cQry += CRLF + " ON  SB1.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND B1_FILIAL = '" + xFilial('SB1') + "'"
	cQry += CRLF + " AND B1_COD = FT_PRODUTO"
	cQry += CRLF + " LEFT JOIN " + RetSqlName('ZZF') + " ZZF " + _cNoLock
	cQry += CRLF + " ON  ZZF.D_E_L_E_T_ <> '*'"
	cQry += CRLF + " AND ZZF_FILIAL = '" + xFilial('ZZF') + "'"
	cQry += CRLF + " AND ZZF_UF = A2_EST"
	cQry += CRLF + " AND ZZF_TIPO = '1'"
	cQry += CRLF + " AND ZZF_CODIGO = A2_DIEF"
	cQry += CRLF + " WHERE SFT.D_E_L_E_T_ <> '*'"
	cQry += CRLF + "   AND FT_FILIAL = '" + xFilial('SFT') + "'"
	cQry += CRLF + "   AND LEFT(FT_CFOP,1) IN ('5')"
	cQry += CRLF + "   AND FT_TIPO IN ('D','B')"
	cQry += CRLF + "   AND FT_ENTRADA BETWEEN '" + DTOS(MV_PAR01) + "' AND '" + DTOS(MV_PAR02) + "'"
	cQry += CRLF + " %cAux2%"
	cQry += CRLF + " GROUP BY"
	cQry += CRLF + "    A2_CGC"
	cQry += CRLF + "   ,A2_NOME"
	cQry += CRLF + "   ,FT_ESTADO"
	cQry += CRLF + "   ,FT_NFISCAL"
	cQry += CRLF + "   ,FT_SERIE"
	cQry += CRLF + "   ,FT_CLIEFOR"
	cQry += CRLF + "   ,FT_LOJA"
	cQry += CRLF + "   ,ZZF_CODIGO"
	cQry += CRLF + "   ,ZZF_DESC"
	
	If cOpe == '1'
		//Sa�das Internas
		//Opera��o 1: Com ICMS Destacado na NF do Substituto Tribut�rio
		//S�o consideradas as notas cujo CST do ICMS for 10 e o campo do Valor do ICMS Retido estiver preenchido.
		cAux1 += CRLF + "   AND F4_SITTRIB = '10'"
		cAux1 += CRLF + "   AND FT_ICMSRET > 0"
	ElseIf cOpe == '2'
		//Sa�das Internas
		//Opera��o 2: Com ICMS n�o Destacado na Nota de Sa�da, em decorr�ncia de Substitui��o Tributa��o ou Antecipa��o na Opera��o Antecedente
		//S�o consideradas as notas cujo CST do ICMS for 60.
		cAux1 += CRLF + "   AND F4_SITTRIB = '60'"
	ElseIf cOpe == '3'
		//Sa�das Internas
		//Opera��o 3: Transfer�ncia Interna sem Aplica��o da ST
		//S�o consideradas as notas cujo CFOP for um dos listados abaixo:
		//5151,5152,5153,5155,5156,5408,5409,5552,5557,5601,5602,5605,5658,5659
		cAux1 += CRLF + "   AND FT_CFOP IN ('5151','5152','5153','5155','5156','5408','5409','5552','5557','5601','5602','5605','5658','5659')"
	ElseIf cOpe == '4'
		//Sa�das Internas
		//Opera��o 4: Outras Sa�das sem Aplica��o do Regime ST, Liminares e Outras
		//S�o consideradas as notas cujo campo Cadastro de Produtos B1_PICMRET estiver preenchido,
		//o CST do ICMS for diferente de 10 e 60 e CFOP diferente do citado acima, na lista de transfer�ncia.
		cAux1 += CRLF + "   AND B1_PICMRET > 0"
		cAux1 += CRLF + "   AND F4_SITTRIB NOT IN ('10','60')"
		cAux1 += CRLF + "   AND FT_CFOP NOT IN ('5151','5152','5153','5155','5156','5408','5409','5552','5557','5601','5602','5605','5658','5659')"
	EndIf
	
	If cAux2 == ''
		cAux2 := cAux1
	EndIf
	cQry := StrTran(cQry, '%cAux1%', cAux1)
	cQry := StrTran(cQry, '%cAux2%', cAux2)
EndIf

Return cQry

/* ------------------- */

Static Function Dados(cNat, cOpe)

Local cQry
Local cTip, cMun, cCod
Local cUlt
Local nPos, nVlr, nICMSRet

cQry := Query(cNat,cOpe)
If _cNoLock == ''
	cQry := ChangeQuery(cQry)
EndIf
dbUseArea(.T.,'TOPCONN',TCGenQry(,,cQry),'MQRY',.T.)

cUlt := ''
nPos := 0

While !MQRY->(Eof())
	If cNat == '1'
		cTip := MQRY->CFOP //1=Internas / 2=Interestaduais
		cMun := ''
		cCod := ''
	ElseIf cNat == '2'
		cTip := '1' //1=Interna
		cMun := AllTrim(MQRY->MUNDESC)
		cCod := MQRY->CODMUN
	EndIf

	If cUlt <> MQRY->CPFCNPJ
		If nPos > 0
			_aDIEF[nPos][8] := nVlr
			_aDIEF[nPos][9] := nICMSRet
		EndIf
		
		aAdd(_aDIEF, {MQRY->CPFCNPJ, RetNat(cNat), RetTip(cTip), RetOpe(cNat,cOpe), MQRY->ESTADO, cMun, cCod, 0, 0})
		nPos := Len(_aDIEF)
		
		nVlr := 0
		nICMSRet := 0
		cUlt := MQRY->CPFCNPJ
	EndIf
	
	nVlr += MQRY->VALOR
	nICMSRet += MQRY->ICMSRET
	
	If MV_PAR03 == 2
		//anal�tico
		aAdd(_aAnal, {MQRY->CPFCNPJ, cNat, cTip, cOpe, MQRY->NOTA, MQRY->SERIE, MQRY->NOME, MQRY->CODIGO, MQRY->LOJA, MQRY->VALOR, MQRY->ICMSRET})
	EndIf
	
	MQRY->(dbSkip())
EndDo

If nPos > 0
	_aDIEF[nPos][8] := nVlr
	_aDIEF[nPos][9] := nICMSRet
EndIf
		
MQRY->(dbCloseArea())

IncProc()

Return Nil

/* ------------------- */

Static Function RetNat(cNat)

Local cRet := ''

If cNat == '1'
	cRet := '1 - Entradas'
ElseIf cNat == '2'
	cRet := '2 - Sa�das'
EndIf

Return cRet

/* ------------------- */

Static Function RetTip(cTip)

Local cRet := ''

If cTip == '1'
	cRet := '1 - Internas'
ElseIf cTip == '2'
	cRet := '2 - Interestaduais'
EndIf

Return cRet

/* ------------------- */

Static Function RetOpe(cNat,cOpe)

Local cRet := ''

If cNat == '1'
	If cOpe == '1'
		cRet := '1 - Com ICMS Destacado na NF do Substituto Tribut�rio'
	ElseIf cOpe == '2'
		cRet := '2 - Com Antecipa��o na Entrada feita pelo Declarante'
	ElseIf cOpe == '3'
		cRet := '3 - Sem Tributa��o em Decorr�ncia de Substitui��o Tribut�ria ou Antecipa��o na Opera��o Antecedente'
	ElseIf cOpe == '4'
		cRet := '4 - Transfer�ncia Interna ou Interestadual n�o Sujeita ao Regime de ST'
	ElseIf cOpe == '5'
		cRet := '5 - Outras Entradas sem Aplica��o do Regime ST (Liminares e Outras)'
	EndIf
ElseIf cNat == '2'
	If cOpe == '1'
		cRet := '1 - Com ICMS Destacado na NF do Substituto Tribut�rio'
	ElseIf cOpe == '2'
		cRet := '2 - Com ICMS n�o Destacado na Nota de Sa�da, em decorr�ncia de Substitui��o Tributa��o ou Antecipa��o na Opera��o Antecedente'
	ElseIf cOpe == '3'
		cRet := '3 - Transfer�ncia Interna sem Aplica��o da ST'
	ElseIf cOpe == '4'
		cRet := '4 - Outras Sa�das sem Aplica��o do Regime ST, Liminares e Outras'
	EndIf
EndIf

Return cRet

/* ------------------- */

Static Function CriaSX1(cPerg,cCFPad)

Local nTamGrp := Len(SX1->X1_GRUPO)
Local aHelpPor := {}, aHelpEng := {}, aHelpSpa := {}
Local cNome

aHelpPor := {}
aAdd(aHelpPor, 'Informe a data inicial/final para    ')
aAdd(aHelpPor, 'gera��o do relat�rio.                ')
cNome := 'Data inicial'
PutSx1(PadR(cPerg,nTamGrp), '01', cNome, cNome, cNome,;
'MV_CH1', 'D', 8, 0, 0, 'G', '', '', '', '', 'MV_PAR01',;
'', '', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
aClone(aHelpPor), aClone(aHelpEng), aClone(aHelpSpa))

cNome := 'Data final'
PutSx1(PadR(cPerg,nTamGrp), '02', cNome, cNome, cNome,;
'MV_CH2', 'D', 8, 0, 0, 'G', '', '', '', '', 'MV_PAR02',;
'', '', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
aClone(aHelpPor), aClone(aHelpEng), aClone(aHelpSpa))

aHelpPor := {}
aAdd(aHelpPor, 'Selecione se deseja emitir o         ')
aAdd(aHelpPor, 'relat�rio Sint�tico ou Anal�tico.    ')
cNome := 'Tipo'
PutSx1(PadR(cPerg,nTamGrp), '03', cNome, cNome, cNome,;
'MV_CH3', 'N', 1, 0, 1, 'C', '', '', '', '', 'MV_PAR03',;
'Sint�tico', 'Sint�tico', 'Sint�tico', '',;
'Anal�tico', 'Anal�tico', 'Anal�tico',;
'', '', '',;
'', '', '',;
'', '', '',;
aClone(aHelpPor), aClone(aHelpEng), aClone(aHelpSpa))

Return Nil