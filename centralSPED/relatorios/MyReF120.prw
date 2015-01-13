#include 'rwmake.ch'
#include 'protheus.ch'
///////////////////////////////////////////////////////////////////////////////
User Function MyReF120()
///////////////////////////////////////////////////////////////////////////////
// Data : 16/10/2014
// User : Thieres Tembra
// Desc : Relat�rio dos registros gerados para o bloco F120 atrav�s
//        da fun��o DeprecAtivo padr�o Protheus
///////////////////////////////////////////////////////////////////////////////
Local cTitulo := 'Relat�rio do registro F120'
Local cPerg := '#MYREF120'

CriaSX1(cPerg)

If !Pergunte(cPerg, .T., cTitulo)
	Return Nil
End If

If AllTrim(MV_PAR01) == '' .or. AllTrim(MV_PAR02) == ''
	Alert('Informe as datas para gera��o do relat�rio.')
	Return Nil
ElseIf MV_PAR01 > MV_PAR02
	Alert('O ano m�s final deve ser igual ou maior que o ano m�s inicial.')
	Return Nil
EndIf

Processa({|| Executa(cTitulo) },cTitulo,'Aguarde...')

Return Nil

/* ------------------- */

Static Function Executa(cTitulo)

Local aExcel   := {}
Local cArq     := 'MYREF120'
Local aResult
Local cAliasF120
Local cDescCta, cDescCC, cNomeFor, dDtAquis, nVlAquis
Local aCST := {}
Local nPos, nI, nMax, nMaxF, nTot
Local aMes := {}
Local aProcItem := {}
Local cMes
Local nMes, nAno
Local cCSTCRED	:= '50/51/52/53/54/55/56/60/61/62/63/64/65/66/67'

aAdd(aExcel, {cTitulo})
aAdd(aExcel, {'Relat�rio emitido em '+DTOC(Date())+' �s '+Time()+' por '+AllTrim(cUsername)})
aAdd(aExcel, {'Per�odo: '+MV_PAR01+' at� '+MV_PAR02})
aAdd(aExcel, {''})
aAdd(aExcel, {'Empresa: '+cEmpAnt+'/'+cFilAnt+'-'+AllTrim(SM0->M0_NOME)+' / '+AllTrim(SM0->M0_FILIAL)})
aAdd(aExcel, {''})

cMes := MV_PAR01
While cMes <= MV_PAR02
	aAdd(aMes, cMes)
	If Right(cMes,2) == '12'
		cMes := StrZero(Val(Left(cMes,4))+1,4) + '01'
	Else
		cMes := Left(cMes,4) + StrZero(Val(Right(cMes,2))+1,2)
	EndIf
EndDo

nMaxF := 0
nMax := Len(aMes)

ProcRegua(nMax)

aAdd(aExcel, {;
				'M�s Refer�ncia'							,;
				'CodBCC'										,;
				'Identifica��o'							,;
				'Origem'										,;
				'Utiliza��o'								,;
				'Valor de Deprecia��o/Amortiza��o'						,;
				'Parcela a Excluir do Valor de Deprecia��o/Amortiza��o'		,;
				'CST PIS'									,;
				'Base C�lculo PIS'						,;
				'Al�q. PIS'									,;
				'Valor PIS'									,;
				'CST COFINS'								,;
				'Base C�lculo COFINS'					,;
				'Al�q. COFINS'								,;
				'Valor COFINS'								,;
				'Conta Anal�tica'							,;
				'Nome da Conta'							,;
				'Centro de Custo'							,;
				'Descri��o CC'								,;
				'Descri��o Complementar'				,;
				'Processo Referenciado'				,;
				'Indicador de Origem do Processo'				,;
				'Bem'											,;
				'Item'										,;
				'Data Aquisi��o'							,;
				'Valor Aquisi��o'							,;
				'Nota Fiscal'								,;
				'S�rie'										,;
				'Fornecedor'								,;
				'Loja'										,;
				'Nome'										 ;
})
		
For nI := 1 to nMax
	IncProc('Analisando m�s ' + aMes[nI] + '..')
	
	nMes := Val(Right(aMes[nI],2))
	nAno := Val(Left(aMes[nI],4))
	
	cAliasF120 := GetNextAlias()
	aProcItem := {'          ','ZZZZZZZZZZ',CTOD(''),U_MyDia(2,nMes,nAno),cAliasF120}
	aResult := _DeprecAtivo(U_MyDia(1,nMes,nAno),aProcItem[4],.T.,.F.,aProcItem,,.F.,'09/11',SM0->M0_CODFIL,.F.)
	
	(cAliasF120)->(dbSetOrder(1))
	(cAliasF120)->(dbGoTop())
	
	If !(cAliasF120)->(Eof()) .and. (cAliasF120)->(FieldPos('NATBCCRED')) > 0 .and. Len(aResult) > 0
		cAliasF120 := aResult[1,2]
				
		nTot := 0
		(cAliasF120)->(dbEval({ || nTot++ }))
		(cAliasF120)->(dbGoTop())
		nMaxF += nTot
		
		While !(cAliasF120)->(Eof())
			
			If (cAliasF120)->CSTPIS $ cCSTCRED .or. (cAliasF120)->CSTCOFINS $ cCSTCRED			
				cDescCta := Posicione('CT1',1,xFilial('CT1') + (cAliasF120)->CODCONTA,'CT1_DESC01')
				cDescCC  := Posicione('CTT',1,xFilial('CTT') + (cAliasF120)->CODCCUSTO,'CTT_DESC01')
				cNomeFor := Posicione('SA2',1,xFilial('SA2') + (cAliasF120)->FORNECEDOR + (cAliasF120)->LOJA,'A2_NOME')
				dDtAquis := Posicione('SN1',1,xFilial('SN1') + (cAliasF120)->BEM + (cAliasF120)->ITEM,'N1_AQUISIC')
				nVlAquis := SN1->N1_VLAQUIS
				If nVlAquis == 0
					//Tipo 01=Aquisi��o
					//Ocorr�ncia 05=Implanta��o
					nVlAquis := Posicione('SN4',4,xFilial('SN4') + (cAliasF120)->BEM + (cAliasF120)->ITEM + '01' + '05','N4_VLROC1')
					If nVlAquis == 0
						nVlAquis := SN4->N4_VLROC2
					EndIf
					If nVlAquis == 0
						nVlAquis := SN4->N4_VLROC3
					EndIf
					If nVlAquis == 0
						nVlAquis := SN4->N4_VLROC4
					EndIf
					If nVlAquis == 0
						nVlAquis := SN4->N4_VLROC5
					EndIf
				EndIf
				
				aAdd(aExcel, {;
								aMes[nI]                         ,;
								(cAliasF120)->NATBCCRED          ,;
								IndIde((cAliasF120)->INDBEMIMOB) ,;
								IndOri((cAliasF120)->INDORIGCRD) ,;
								IndUti((cAliasF120)->INDUTILBEM) ,;
								(cAliasF120)->VRET               ,;
								0                                ,;
								(cAliasF120)->CSTPIS             ,;
								(cAliasF120)->VLRBCPIS           ,;
								(cAliasF120)->ALIQPIS            ,;
								(cAliasF120)->VLRPIS             ,;
								(cAliasF120)->CSTCOFINS          ,;
								(cAliasF120)->VLRBCCOFIN         ,;
								(cAliasF120)->ALIQCOFINS         ,;
								(cAliasF120)->VLRCOFINS          ,;
								(cAliasF120)->CODCONTA           ,;
								cDescCta                         ,;
								(cAliasF120)->CODCCUSTO          ,;
								cDescCC                          ,;
								(cAliasF120)->DESCBEMIMO         ,;
								(cAliasF120)->NUMPRO             ,;
								IndPro((cAliasF120)->INDPRO)     ,;
								(cAliasF120)->BEM                ,;
								(cAliasF120)->ITEM               ,;
								dDtAquis                         ,;
								nVlAquis                         ,;
								(cAliasF120)->NOTAFISCAL         ,;
								(cAliasF120)->SERIE              ,;
								(cAliasF120)->FORNECEDOR         ,;
								(cAliasF120)->LOJA               ,;
								cNomeFor                          ;
				})
				
				nPos := aScan(aCST, {|x| x[1] == (cAliasF120)->CSTPIS })
				If nPos <> 0
					aCST[nPos][2] += (cAliasF120)->VLRBCPIS
					aCST[nPos][3] += (cAliasF120)->VLRBCCOFIN
					aCST[nPos][4] += (cAliasF120)->VLRPIS
					aCST[nPos][5] += (cAliasF120)->VLRCOFINS
				Else
					aAdd(aCST, {(cAliasF120)->CSTPIS, (cAliasF120)->VLRBCPIS, (cAliasF120)->VLRBCCOFIN, (cAliasF120)->VLRPIS, (cAliasF120)->VLRCOFINS})
				EndIf
			EndIf
			
			(cAliasF120)->(dbSkip())
		EndDo
		(cAliasF120)->(dbCloseArea())
	EndIf
Next nI

If nMaxF > 0
	aAdd(aExcel, {''})
	aAdd(aExcel, {'Totais por CST'})
	aAdd(aExcel, {'CST','Base PIS','Base COFINS','Valor PIS','Valor COFINS'})
	aSort(aCST,,,{|x,y| x[1] < y[1] })
	nMax := Len(aCST)
	For nI := 1 to nMax
		aAdd(aExcel, {aCST[nI][1], aCST[nI][2], aCST[nI][3], aCST[nI][4], aCST[nI][5]})
	Next nI
Else
	aAdd(aExcel, {'N�o existem registros a serem apresentados.'})
EndIf

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

Static Function IndIde(cInd)

Local cRet := ''
Local aInd := {;
	{'01', '01 - Edifica��es e Benfeitorias em Im�veis Pr�prios'},;
	{'02', '02 - Edifica��es e Benfeitorias em Im�veis de Terceiros'},;
	{'03', '03 - Instala��es'},;
	{'04', '04 - M�quinas'},;
	{'05', '05 - Equipamentos'},;
	{'06', '06 - Ve�culos'},;
	{'99', '99 - Outros'};
}
Local nPos

nPos := aScan(aInd, {|x| x[1] == cInd })
If nPos > 0
	cRet := aInd[nPos][2]
EndIf

Return cRet

/* ------------------- */

Static Function IndOri(cInd)

Local cRet := ''
Local aInd := {;
	{'0', '0 - Aquisi��o no Mercado Interno'},;
	{'1', '1 - Aquisi��o no Mercado Externo (Importa��o)'};
}
Local nPos

nPos := aScan(aInd, {|x| x[1] == cInd })
If nPos > 0
	cRet := aInd[nPos][2]
EndIf

Return cRet

/* ------------------- */

Static Function IndUti(cInd)

Local cRet := ''
Local aInd := {;
	{'1', '1 - Produ��o de Bens Destinados a Venda'},;
	{'2', '2 - Presta��o de Servi�os'},;
	{'3', '3 - Loca��o a Terceiros'},;
	{'9', '9 - Outros'};
}
Local nPos

nPos := aScan(aInd, {|x| x[1] == cInd })
If nPos > 0
	cRet := aInd[nPos][2]
EndIf

Return cRet

/* ------------------- */

Static Function IndPro(cInd)

Local cRet := ''
Local aInd := {;
	{'0', '0 - SEFAZ'},;
	{'1', '1 - Justi�a Federal'},;
	{'2', '2 - Justi�a Estadual'},;
	{'3', '3 - Secretaria da Receita Federal do Brasil'},;
	{'9', '9 - Outros'};
}
Local nPos

nPos := aScan(aInd, {|x| x[1] == cInd })
If nPos > 0
	cRet := aInd[nPos][2]
EndIf

Return cRet

/* ------------------- */

Static Function CriaSX1(cPerg,cCFPad)

Local nTamGrp := Len(SX1->X1_GRUPO)
Local aHelpPor := {}, aHelpEng := {}, aHelpSpa := {}
Local cNome

aHelpPor := {}
aAdd(aHelpPor, 'Informe o ano/m�s inicial para       ')
aAdd(aHelpPor, 'gera��o do relat�rio. Deve ser       ')
aAdd(aHelpPor, 'informado a data de aquisi��o.       ')
aAdd(aHelpPor, 'Ex: 201201                           ')
cNome := 'Do Ano Mes (Ex: 201201)'
PutSx1(PadR(cPerg,nTamGrp), '01', cNome, cNome, cNome,;
'MV_CH1', 'C', 6, 0, 0, 'G', '', '', '', '', 'MV_PAR01',;
'', '', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
aClone(aHelpPor), aClone(aHelpEng), aClone(aHelpSpa))

aHelpPor := {}
aAdd(aHelpPor, 'Informe o ano/m�s final para         ')
aAdd(aHelpPor, 'gera��o do relat�rio. Deve ser       ')
aAdd(aHelpPor, 'informado a data de aquisi��o.       ')
aAdd(aHelpPor, 'Ex: 201203                           ')
cNome := 'Ate Ano Mes (Ex: 201203)'
PutSx1(PadR(cPerg,nTamGrp), '02', cNome, cNome, cNome,;
'MV_CH2', 'C', 6, 0, 0, 'G', '', '', '', '', 'MV_PAR02',;
'', '', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
aClone(aHelpPor), aClone(aHelpEng), aClone(aHelpSpa))

Return Nil