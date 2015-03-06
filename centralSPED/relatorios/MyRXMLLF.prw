#include 'rwmake.ch'
#include 'protheus.ch'

//defines para o vetor retornado pela função DIRECTORY()
#define F_NAME        1
#define F_SIZE        2
#define F_DATE        3
#define F_TIME        4
#define F_ATTR        5
//tamanho do vetor
#define F_LEN         5

//defines para o vetor _aXML
#define XML_TIPO      01
#define XML_MODELO    02
#define XML_SERIE     03
#define XML_NUMERO    04
#define XML_EMISSAO   05
#define XML_CLIEFOR   06
#define XML_NOME      07
#define XML_VALOR     08
#define XML_CHAVE     09
#define XML_CANCELADO 10
#define XML_PARTFRETE 11
#define XML_TOMAFRETE 12
#define XML_LIVRO     13
#define XML_FINAL     14

//defines para o vetor aVerifica
#define VER_TAG          01
#define VER_ERRO_TAG     02
#define VER_COMPARA      03
#define VER_ERRO_COMPARA 04
#define VER_SALVA        05
#define VER_FINAL        06

//defines para o vetor _aLivro
#define LIVRO_TIPO      01
#define LIVRO_ESPECIE   02
#define LIVRO_SERIE     03
#define LIVRO_NUMERO    04
#define LIVRO_EMISSAO   05
#define LIVRO_ENTRADA   06
#define LIVRO_CLIEFOR   07
#define LIVRO_LOJA      08
#define LIVRO_CNPJ      09
#define LIVRO_NOME      10
#define LIVRO_VALOR     11
#define LIVRO_CHAVE     12
#define LIVRO_CANCELADO 13
#define LIVRO_XML       14
#define LIVRO_FINAL     15
///////////////////////////////////////////////////////////////////////////////
User Function MyRXMLLF()
///////////////////////////////////////////////////////////////////////////////
// Data : 08/01/2015
// User : Thieres Tembra
// Desc : Emite o relatório de cruzamento entre os XMLs e o Livro Fiscal
// Ação : A rotina emite o relatório de cruzamento dos documentos fiscais
//        eletrônicos (NF-e ou CT-e) de entrada ou saída através dos XMLs
//        disponibilizados no servidor mensalmente contra os documentos
//        lançados no sistema (Livro Fiscal).
///////////////////////////////////////////////////////////////////////////////

Local cTitulo := 'Relatório de XML x Livro Fiscal'
Local cPerg := '#MYRXMLLF'
Local aArea := GetArea()
Local aAreaSM0 := SM0->(GetArea())
Private _cPathXML := '\XML-SEFA\'
Private _aXML := {}
Private _aLivro := {}
Private _aDados := {}
Private _nSemRelac := 0

CriaSX1(cPerg)

If !Pergunte(cPerg, .T., cTitulo)
	Return Nil
End If

If MV_PAR01 == Nil .or. MV_PAR01 > 12 .or. MV_PAR01 < 1
	Alert('Informe um mês válido para geração do relatório.')
	Return Nil
ElseIf MV_PAR02 == Nil .or. MV_PAR02 < 1000
	Alert('Informe um ano válido para geração do relatório.')
	Return Nil
EndIf

Processa({|| Executa(cTitulo) },cTitulo,'Aguarde...')

SM0->(RestArea(aAreaSM0))
RestArea(aArea)

Return Nil

/* -------------- */

Static Function Executa(cTitulo)

Local lErro := .F.

ProcRegua(3)

TelaXML(cTitulo, @lErro)
IncProc()
If lErro
	Return Nil
EndIf

Processa({|| Livro(cTitulo, @lErro) },cTitulo,'Analisando Livro Fiscal...')
IncProc()
If lErro
	Return Nil
EndIf

Processa({|| Dados(cTitulo, @lErro) },cTitulo,'Relacionando documentos...')
IncProc()
If lErro
	Return Nil
EndIf

Return Nil

/* -------------- */

Static Function TelaXML(cTitulo, lErro)

Local nI
Local lEnt, lSai, lNFe, lCTe
Local oPnl1, oPnl2, oPnl3
Private _oDlg, _oGrid, _oEdt
Private _oBtn1, _oBtn2, _oBtn3
Private _lContinua := .T.
Private _cMsgErro := ''
Private _cMsgLoad := '<br><br><center><font size="4"><b>Aguarde...</b></font></center>'
Private _cBtnDis := 'QPushButton{ background-image: none; background-repeat: none; padding-left: 3px; padding-top: 1px; background-origin: content; background-clip: margin; }'
Private _aGridData
Private _nGridMax := 5
Private _nGridIni := 1

lEnt := MV_PAR03 == 1 .or. MV_PAR03 == 3
lSai := MV_PAR03 == 2 .or. MV_PAR03 == 3
lNFe := MV_PAR04 == 1 .or. MV_PAR04 == 3
lCTe := MV_PAR04 == 2 .or. MV_PAR04 == 3

//unchecked, checked, nochecked, next_pq, updwarning
_aGridData := {}
aAdd(_aGridData, {'', 'Verificando caminho dos XMLs...', .T.})
//Entrada
aAdd(_aGridData, {'', 'Entrada - Verificando diretório...', lEnt})
aAdd(_aGridData, {'', 'Entrada - Verificando empresa...', lEnt})
aAdd(_aGridData, {'', 'Entrada - Verificando filial...', lEnt})
aAdd(_aGridData, {'', 'Entrada - Verificando ano...', lEnt})
aAdd(_aGridData, {'', 'Entrada - Verificando mês...', lEnt})
aAdd(_aGridData, {'', 'Entrada - Verificando NF-e/CT-e...', lEnt})
aAdd(_aGridData, {'', 'Entrada - Lendo arquivos NF-e...', lEnt .and. lNFe})
aAdd(_aGridData, {'', 'Entrada - Lendo arquivos CT-e...', lEnt .and. lCTe})
//Saída
aAdd(_aGridData, {'', 'Saída - Verificando diretório...', lSai})
aAdd(_aGridData, {'', 'Saída - Verificando empresa...', lSai})
aAdd(_aGridData, {'', 'Saída - Verificando filial...', lSai})
aAdd(_aGridData, {'', 'Saída - Verificando ano...', lSai})
aAdd(_aGridData, {'', 'Saída - Verificando mês...', lSai})
aAdd(_aGridData, {'', 'Saída - Verificando NF-e/CT-e...', lSai})
aAdd(_aGridData, {'', 'Saída - Lendo arquivos NF-e...', lSai .and. lNFe})
aAdd(_aGridData, {'', 'Saída - Lendo arquivos CT-e...', lSai .and. lCTe})

_oDlg := TDialog():New(0,0,145,600,cTitulo,,,,,CLR_BLACK,CLR_WHITE,,,.T.) 

oPnl1 := TPanelCSS():New(0,0,'',_oDlg,,.F.,,,,_oDlg:nClientWidth/4,(_oDlg:nClientHeight/2)+20)
oPnl1:Align := CONTROL_ALIGN_LEFT

_oGrid := TGrid():New(oPnl1,0,0,oPnl1:nClientWidth/2,(oPnl1:nClientHeight/2)+20)
_oGrid:addColumn(1, '', 30, 0)
_oGrid:addColumn(2, 'Analisando XMLs...', 250, CONTROL_ALIGN_LEFT)
For nI := 1 to _nGridMax
	_oGrid:setRowData(nI, {|o| {'', ''} })
Next nI
_oGrid:setSelectedRow(_nGridMax+1)

oPnl2 := TPanelCSS():New(0,0,'',_oDlg,,.F.,,,,_oDlg:nClientWidth/4,(_oDlg:nClientHeight/2)+20)
oPnl2:Align := CONTROL_ALIGN_RIGHT

_oEdt := TSimpleEditor():New(0,0,oPnl2,oPnl2:nClientWidth/2,(oPnl2:nClientHeight/2)-18,'',.F.,{|u| Iif(PCount()>0, _cMsgErro := u, _cMsgErro ) },,.T.,{|| .T. },{|| .T. })
_oEdt:Load(_cMsgLoad)
_oEdt:Align := CONTROL_ALIGN_TOP

oPnl3 := TPanelCSS():New(0,0,'',oPnl2,,.F.,,,,oPnl2:nClientWidth/2,18)
oPnl3:Align := CONTROL_ALIGN_BOTTOM

/*
_oBtn1 := TButton():New(0,0,'      + Ações',oPnl3,{|| .T. },51,18,,,.F.,.T.,.F.,,.F.,,,.F. )
_oBtn1:SetCss('QPushButton{ background-image: url(rpo:parametros.png); background-repeat: none; padding-left: 3px; padding-top: 1px; background-origin: content; background-clip: margin; }')
_oBtn1:Align := CONTROL_ALIGN_LEFT
_oBtn1:Disable()
*/
_oBtn1 := TButton():New(0,0,'...',oPnl3,{|| .T. },51,18,,,.F.,.T.,.F.,,.F.,,,.F. )
_oBtn1:SetCss(_cBtnDis)
_oBtn1:Align := CONTROL_ALIGN_LEFT
_oBtn1:Disable()

_oBtn2 := TButton():New(0,0,'...',oPnl3,{|| .T. },51,18,,,.F.,.T.,.F.,,.F.,,,.F. )
_oBtn2:SetCss(_cBtnDis)
_oBtn2:Align := CONTROL_ALIGN_ALLCLIENT
_oBtn2:Disable()

_oBtn3 := TButton():New(0,0,'...',oPnl3,{|| Continua() },51,18,,,.F.,.T.,.F.,,.F.,,,.F. )
_oBtn3:SetCss(_cBtnDis)
_oBtn3:Align := CONTROL_ALIGN_RIGHT
_oBtn3:Disable()

_oDlg:Activate(,,,.T.,{|| .T. },,{|| AnaXML(cTitulo, @lErro), _oDlg:End() })

Return Nil

/* -------------- */

Static Function Continua()

_oEdt:Load(_cMsgLoad)
_oBtn3:SetCss(_cBtnDis)
_oBtn3:SetText('...')
_oBtn3:Disable()
_lContinua := .T.

Return Nil

/* -------------- */

Static Function GridData(cES, nRow, cImage, lScroll, cMsg)

Local cAux
Local nPos, nI
Default lScroll := .F.
Default cMsg := ''

If cES == 'S'
	//se for saída pula as 8 linhas da entrada
	nRow += 8
EndIf

If lScroll .and. nRow > _nGridMax
	nPos := 0
	For nI := nRow to 1 step -1
		If _aGridData[nI][3]
			nPos += 1
		EndIf
		If nPos == _nGridMax
			_nGridIni := nI
			Exit
		EndIf
	Next nI
EndIf

_aGridData[nRow][1] := 'RPO_IMAGE=' + cImage
If cMsg <> ''
	_aGridData[nRow][2] := cMsg
EndIf

nPos := 1
_oGrid:clearRows()
For nI := _nGridIni to nRow
	If _aGridData[nI][3]
		cAux := '{|o| {_aGridData[' + cValToChar(nI) + '][1], _aGridData[' + cValToChar(nI) + '][2]} }'
		_oGrid:setRowData(nPos, &(cAux))
		nPos += 1
	EndIf
Next nI

While nPos <= _nGridMax
	_oGrid:setRowData(nPos, {|o| {'', ''} })
	nPos += 1
EndDo

//evita travamento e atualiza tela
ProcessMessages()

Return Nil

/* -------------- */

Static Function AnaXML(cTitulo, lErro)

Local cPath := _cPathXML
Local cErroPath
Local nRow

nRow := 1
GridData('', nRow, 'NEXT_PQ', .T.)

If Right(cPath, 1) <> '\'
	cPath := cPath + '\'
EndIf

cErroPath := 'caminho (' + cPath + ') dos XMLs'

If !ExistDir(cPath)
	GridData('', nRow, 'NOCHECKED')
	lErro := .T.
	Alert('Não foi possível acessar o ' + cErroPath + '.')
	Return Nil
EndIf

GridData('', nRow, 'CHECKED')

//Entrada
If MV_PAR03 == 1 .or. MV_PAR03 == 3
	XML('E', cPath, cErroPath, @lErro)
	If lErro
		Return Nil
	EndIf
EndIf
//Saída
If MV_PAR03 == 2 .or. MV_PAR03 == 3
	XML('S', cPath, cErroPath, @lErro)
	If lErro
		Return Nil
	EndIf
EndIf

Return Nil

/* -------------- */

Static Function XML(cES, cPath, cErroPath, lErro)

Local cName, cAttr
Local cEntSai
Local lNFe, lCTe, lErrNFeCTe
Local aFiles
Local nI
Local nMaxI
Local nRow
Local cPathEmp, cPathFil, cPathAno, cPathMes
Local cRet := ''

//buscando diretório de Entrada ou Saída
nRow := 2
GridData(cES, nRow, 'NEXT_PQ', .T.)
cEntSai := ''
aFiles := Directory(cPath + '*.*', 'D', Nil, .T.)
nMaxI := Len(aFiles)
For nI := 1 to nMaxI
	cName := aFiles[nI][F_NAME]
	cAttr := aFiles[nI][F_ATTR]
	//se for Diretório
	If cAttr == 'D'
		//se estiver processando Entrada
		If cES == 'E'
			//se primeiro caractere for E de Entrada
			//e parâmetro Livro estiver configurado como Entrada ou Ambos
			//e caminho de entrada ainda não estiver definido
			If Left(cName, 1) == 'E' .and. (MV_PAR03 == 1 .or. MV_PAR03 == 3) .and. cEntSai == ''
				cEntSai := cPath + cName + '\'
			EndIf
		ElseIf cES == 'S'
			//se primeiro caractere for S de Saída
			//e parâmetro Livro estiver configurado como Saída ou Ambos
			//e caminho de saída ainda não estiver definido
			If Left(cName, 1) == 'S' .and. (MV_PAR03 == 2 .or. MV_PAR03 == 3) .and. cEntSai == ''
				cEntSai := cPath + cName + '\'
			EndIf
		EndIf
	EndIf
Next nI

If cEntSai == '' .or. !ExistDir(cEntSai)
	GridData(cES, nRow, 'NOCHECKED')
	lErro := .T.
	ErroXML(cES, cErroPath)
	Return Nil
EndIf

GridData(cES, nRow, 'CHECKED')

//buscando diretório da empresa
nRow := 3
GridData(cES, nRow, 'NEXT_PQ', .T.)
cPathEmp := ''
aFiles := Directory(cEntSai + '*.*', 'D', Nil, .T.)
nMaxI := Len(aFiles)
For nI := 1 to nMaxI
	cName := aFiles[nI][F_NAME]
	cAttr := aFiles[nI][F_ATTR]
	//se for Diretório
	//e os dois primeiros caracteres forem igual ao Código da Empresa
	//e caminho da empresa ainda não estiver definido
	If cAttr == 'D' .and. Left(cName, 2) == cEmpAnt .and. cPathEmp == ''
		cPathEmp := cName
		cEntSai += cPathEmp + '\'
	EndIf
Next nI

If cPathEmp == '' .or. !ExistDir(cEntSai)
	GridData(cES, nRow, 'NOCHECKED')
	lErro := .T.
	ErroXML(cES, cErroPath, '<b>Empresa:</b> ' + cEmpAnt)
	Return Nil
EndIf

GridData(cES, nRow, 'CHECKED')

//buscando diretório da filial
nRow := 4
GridData(cES, nRow, 'NEXT_PQ', .T.)
cPathFil := ''
aFiles := Directory(cEntSai + '*.*', 'D', Nil, .T.)
nMaxI := Len(aFiles)
For nI := 1 to nMaxI
	cName := aFiles[nI][F_NAME]
	cAttr := aFiles[nI][F_ATTR]
	//se for Diretório
	//e os dois primeiros caracteres forem igual ao Código da Filial
	//e caminho da filial ainda não estiver definido
	If cAttr == 'D' .and. Left(cName, 2) == cFilAnt .and. cPathFil == ''
		cPathFil := cName
		cEntSai += cPathFil + '\'
	EndIf
Next nI

If cPathFil == '' .or. !ExistDir(cEntSai)
	GridData(cES, nRow, 'NOCHECKED')
	lErro := .T.
	ErroXML(cES, cErroPath, '<b>Empresa:</b> ' + cEmpAnt + ' - <b>Filial:</b> ' + cFilAnt)
	Return Nil
EndIf

GridData(cES, nRow, 'CHECKED')

//verificando se existe diretório do ano
nRow := 5
GridData(cES, nRow, 'NEXT_PQ', .T.)
cPathAno := StrZero(MV_PAR02, 4)
cEntSai += cPathAno + '\'
If !ExistDir(cEntSai)
	GridData(cES, nRow, 'NOCHECKED')
	lErro := .T.
	ErroXML(cES, cErroPath, '<b>Empresa:</b> ' + cEmpAnt + ' - <b>Filial:</b> ' + cFilAnt + ' - <b>Ano:</b> ' + StrZero(MV_PAR02, 4))
	Return Nil
EndIf

GridData(cES, nRow, 'CHECKED')

//buscando diretório do mês
nRow := 6
GridData(cES, nRow, 'NEXT_PQ', .T.)
cPathMes := ''
aFiles := Directory(cEntSai + '*.*', 'D', Nil, .T.)
nMaxI := Len(aFiles)
For nI := 1 to nMaxI
	cName := aFiles[nI][F_NAME]
	cAttr := aFiles[nI][F_ATTR]
	//se for Diretório
	//e os dois primeiros caracteres forem igual ao Código da Filial
	//e caminho da filial ainda não estiver definido
	If cAttr == 'D' .and. Left(cName, 2) == StrZero(MV_PAR01, 2) .and. cPathMes == ''
		cPathMes := cName
		cEntSai += cPathMes + '\'
	EndIf
Next nI

If cPathMes == '' .or. !ExistDir(cEntSai)
	GridData(cES, nRow, 'NOCHECKED')
	lErro := .T.
	ErroXML(cES, cErroPath, '<b>Empresa:</b> ' + cEmpAnt + ' - <b>Filial:</b> ' + cFilAnt + ' - <b>Ano:</b> ' + StrZero(MV_PAR02, 4) + ' - <b>Mês:</b> ' + StrZero(MV_PAR01, 2))
	Return Nil
EndIf

GridData(cES, nRow, 'CHECKED')

//buscando diretórios de NF-e e CT-e
nRow := 7
GridData(cES, nRow, 'NEXT_PQ', .T.)
lNFe := .F.
lCTe := .F.
lErrNFeCTe := .F.
aFiles := Directory(cEntSai + '*.*', 'D', Nil, .T.)
nMaxI := Len(aFiles)
For nI := 1 to nMaxI
	cName := aFiles[nI][F_NAME]
	cAttr := aFiles[nI][F_ATTR]
	//se for Diretório
	If cAttr == 'D'
		//se o nome do diretório for NFe
		If cName == 'NFE'
			lNFe := .T.
		//se o nome do diretório for CTe
		ElseIf cName == 'CTE'
			lCTe := .T.
		EndIf
	EndIf
Next nI

//NF-e
If MV_PAR04 == 1 .or. MV_PAR04 == 3
	If !lNFe .or. !ExistDir(cEntSai + 'NFE\')
		GridData(cES, nRow, 'UPDWARNING')
		lNFe := .F.
		lErrNFeCTe := .T.
		ErroXML(cES, cErroPath, '<b>Empresa:</b> ' + cEmpAnt + ' - <b>Filial:</b> ' + cFilAnt + ' - <b>Ano:</b> ' + StrZero(MV_PAR02, 4) + ' - <b>Mês:</b> ' + StrZero(MV_PAR01, 2) + CRLF + '<b>Espécie:</b> NF-e', .T.)
	EndIf
Else
	lNFe := .F.
EndIf
//CT-e
If MV_PAR04 == 2 .or. MV_PAR04 == 3
	If !lCTe .or. !ExistDir(cEntSai + 'CTE\')
		GridData(cES, nRow, 'UPDWARNING')
		lCTe := .F.
		lErrNFeCTe := .T.
		ErroXML(cES, cErroPath, '<b>Empresa:</b> ' + cEmpAnt + ' - <b>Filial:</b> ' + cFilAnt + ' - <b>Ano:</b> ' + StrZero(MV_PAR02, 4) + ' - <b>Mês:</b> ' + StrZero(MV_PAR01, 2) + CRLF + '<b>Espécie:</b> CT-e', .T.)
	EndIf
Else
	lCTe := .F.
EndIf

If !lErrNFeCTe
	GridData(cES, nRow, 'CHECKED')
EndIf

//lendo arquivos de NF-e
If lNFe
	nRow := 8
	GridData(cES, nRow, 'NEXT_PQ', .T.)
	aFiles := Directory(cEntSai + 'NFE\*.xml', '', Nil, .T.)
	nMaxI := Len(aFiles)
	For nI := 1 to nMaxI
		GridData(cES, nRow, 'NEXT_PQ', .F., Iif(cES=='E','Entrada','Saída') + ' - Lendo arquivos NF-e... ' + cValToChar(nI) + '/' + cValToChar(nMaxI))
		cName := aFiles[nI][F_NAME]
		cAttr := aFiles[nI][F_ATTR]
		//analisa XML
		If !LeXML(cES, nRow, '55', cEntSai + 'NFE\', cName)
			lErro := .T.
			Return Nil
		EndIf
	Next nI
	GridData(cES, nRow, 'CHECKED')
Else
	If cES == 'E'
		_aGridData[8][3] := .F.
	ElseIf cES == 'S'
		_aGridData[16][3] := .F.
	EndIf
EndIf

If lCTe
	nRow := 9
	GridData(cES, nRow, 'NEXT_PQ', .T.)
	aFiles := Directory(cEntSai + 'CTE\*.xml', '', Nil, .T.)
	nMaxI := Len(aFiles)
	For nI := 1 to nMaxI
		GridData(cES, nRow, 'NEXT_PQ', .F., Iif(cES=='E','Entrada','Saída') + ' - Lendo arquivos CT-e... ' + cValToChar(nI) + '/' + cValToChar(nMaxI))
		cName := aFiles[nI][F_NAME]
		cAttr := aFiles[nI][F_ATTR]
		//analisa XML
		If !LeXML(cES, nRow, '57', cEntSai + 'CTE\', cName)
			lErro := .T.
			Return Nil
		EndIf
	Next nI
	GridData(cES, nRow, 'CHECKED')
Else
	If cES == 'E'
		_aGridData[9][3] := .F.
	ElseIf cES == 'S'
		_aGridData[17][3] := .F.
	EndIf
EndIf

Return Nil

/* -------------- */

Static Function ErroXML(cES, cErroPath, cAux, lAviso)

Local cEntSai := ''
Local cMsg := ''
Default cAux := ''
Default lAviso := .F.

If cES == 'E'
	cEntSai := 'Entrada'
ElseIf cES == 'S'
	cEntSai := 'Saída'
EndIf

If cAux <> ''
	cAux := CRLF + cAux
EndIf

cMsg := 'Não foi possível encontrar o diretório que contém ' +;
'as informações abaixo dentro do ' + cErroPath + ':' + CRLF +;
'<b>Documento:</b> ' + cEntSai + cAux

cMsg := StrTran(cMsg, CRLF, '<br />')

MsgErro(cMsg, lAviso)

Return Nil

/* -------------- */

Static Function LeXML(cES, nRow, cModelo, cPath, cFile)

Local lRet := .T.
Local cError := ''
Local cWarning := ''
Local lNFe200 := .T.
Local lNFe310 := .T.
Local lCTe := .T.
Local lCanc := .T.
Local lEvento := .T.
Local cXML
Local nI, nMaxI, nPos
Local aVerifica, aAux
Local nTomador
Local cAux, cTpNF, cTagToma, cNomeToma, cTomador
Local lToma03, lToma4, lCNPJ, lCPF, lEntToma
Local aCTePart := {;
	{'_REM'       , 'remetente'   },;
	{'_EXPED'     , 'expedidor'   },;
	{'_RECEB'     , 'recebedor'   },;
	{'_DEST'      , 'destinatário'},;
	{'_IDE:_TOMA4', 'tomador'     } ;
}

Private _oXML := Nil

_oXML := XmlParserFile(cPath + cFile, '_', @cError, @cWarning)

If cError <> ''
	lRet := .F.
	GridData(cES, nRow, 'NOCHECKED')
	LeErro(cES, cFile, cError)
	Return lRet
EndIf

If cWarning <> ''
	GridData(cES, nRow, 'UPDWARNING')
	LeErro(cES, cFile, cWarning, .T.)
EndIf

If _oXML == Nil
	lRet := .F.
	GridData(cES, nRow, 'NOCHECKED')
	LeErro(cES, cFile, 'Não foi possível carregar o XML.')
	Return lRet
EndIf

/* ///////////////
 *
 * NF-e
 *
 * ///////////////
 */
If cModelo == '55'
	If Type('_oXML:_NFEPROC:_NFE:_INFNFE:_IDE:_DEMI') == 'U'
		//não existe tag do XML da NF-e layout 2.00
		lNFe200 := .F.
	EndIf
	If Type('_oXML:_NFEPROC:_NFE:_INFNFE:_IDE:_DHEMI') == 'U'
		//não existe tag do XML da NF-e layout 3.10
		lNFe310 := .F.
	EndIf
	If Type('_oXML:_PROCCANCNFE:_CANCNFE:_INFCANC') == 'U'
		//não existe tag do XML do Cancelamento da NF-e layout 2.00
		lCanc := .F.
	EndIf
	If Type('_oXML:_PROCEVENTONFE:_EVENTO:_INFEVENTO:_DETEVENTO:_DESCEVENTO') == 'U'
		//não existe tag do XML do Cancelamento por Evento layout 1.00
		lEvento := .F.
	EndIf
	
	If !lNFe200 .and. !lNFe310 .and. !lCanc .and. !lEvento
		lRet := .F.
		GridData(cES, nRow, 'NOCHECKED')
		LeErro(cES, cFile, 'O XML não está em um formato reconhecido pela análise.')
		Return lRet
	EndIf
	
	If lNFe200 .or. lNFe310
		//verificando tipo NF (0=Entrada e 1=Saída)
		If Type('_oXML:_NFEPROC:_NFE:_INFNFE:_IDE:_TPNF') == 'U'
			lRet := .F.
			GridData(cES, nRow, 'NOCHECKED')
			LeErro(cES, cFile, 'Impossível verificar o tipo do documento.')
			Return lRet
		EndIf
		cTpNF := AllTrim(_oXML:_NFEPROC:_NFE:_INFNFE:_IDE:_TPNF:TEXT)
		
		aVerifica := {}
		
		If (cTpNF == '1' .and. cES == 'E') .or.;
		(cTpNF == '0' .and. cES == 'S')
			//documento de saída normal emitido pelo nosso fornecedor/cliente
			//documento de entrada (possivelmente devolução) emitido pelo nosso fornecedor/cliente
			
			//destinatário
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_DEST:_CNPJ'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o CNPJ do destinatário do documento.'
			aAux[VER_COMPARA]      := SM0->M0_CGC
			aAux[VER_ERRO_COMPARA] := 'O destinatário do documento é diferente da empresa que você está logado.'
			aAux[VER_SALVA]        := Nil
			aAdd(aVerifica, aClone(aAux))
			
			//emitente (fornecedor/cliente)
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_EMIT:_CNPJ'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o CNPJ do emitente do documento.'
			aAux[VER_COMPARA]      := Nil
			aAux[VER_ERRO_COMPARA] := Nil
			aAux[VER_SALVA]        := XML_CLIEFOR
			aAdd(aVerifica, aClone(aAux))
			
			//nome do emitente (fornecedor/cliente)
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_EMIT:_XNOME'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o nome do emitente do documento.'
			aAux[VER_COMPARA]      := Nil
			aAux[VER_ERRO_COMPARA] := Nil
			aAux[VER_SALVA]        := XML_NOME
			aAdd(aVerifica, aClone(aAux))
		ElseIf (cTpNF == '0' .and. cES == 'E') .or.;
		(cTpNF == '1' .and. cES == 'S')
			//documento de entrada (possivelmente devolução) emitido pela própria empresa
			//documento de saída emitido pela própria empresa
			
			//emitente
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_EMIT:_CNPJ'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o CNPJ do emitente do documento.'
			aAux[VER_COMPARA]      := SM0->M0_CGC
			aAux[VER_ERRO_COMPARA] := 'O emitente do documento (devolução) é diferente da empresa que você está logado.'
			aAux[VER_SALVA]        := Nil
			aAdd(aVerifica, aClone(aAux))
			
			//destinatário (fornecedor/cliente)
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_DEST:_CNPJ'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o CNPJ do destinatário do documento (devolução).'
			aAux[VER_COMPARA]      := Nil
			aAux[VER_ERRO_COMPARA] := Nil
			aAux[VER_SALVA]        := XML_CLIEFOR
			aAdd(aVerifica, aClone(aAux))
			
			//nome do destinatário (fornecedor/cliente)
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_DEST:_XNOME'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o nome do destinatário do documento (devolução).'
			aAux[VER_COMPARA]      := Nil
			aAux[VER_ERRO_COMPARA] := Nil
			aAux[VER_SALVA]        := XML_NOME
			aAdd(aVerifica, aClone(aAux))
		EndIf
		
		//modelo
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_IDE:_MOD'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar modelo do documento.'
		aAux[VER_COMPARA]      := cModelo
		aAux[VER_ERRO_COMPARA] := 'O modelo do documento é diferente de NF-e.'
		aAux[VER_SALVA]        := XML_MODELO
		aAdd(aVerifica, aClone(aAux))
		
		//série
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_IDE:_SERIE'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a série do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_SERIE
		aAdd(aVerifica, aClone(aAux))
		
		//número
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_IDE:_NNF'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar o número do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_NUMERO
		aAdd(aVerifica, aClone(aAux))
		
		//emissão
		If lNFe200
			cAux := '_DEMI'
		ElseIf lNFe310
			cAux := '_DHEMI'
		EndIf
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_IDE:' + cAux
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a data de emissão do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_EMISSAO
		aAdd(aVerifica, aClone(aAux))
		
		//valor
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_NFEPROC:_NFE:_INFNFE:_TOTAL:_ICMSTOT:_VNF'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar o valor do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_VALOR
		aAdd(aVerifica, aClone(aAux))
		
		//chave
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_NFEPROC:_PROTNFE:_INFPROT:_CHNFE'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a chave do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_CHAVE
		aAdd(aVerifica, aClone(aAux))
	ElseIf lCanc
		aVerifica := {}
		
		//serviço
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_PROCCANCNFE:_CANCNFE:_INFCANC:_XSERV'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar serviço requisitado.'
		aAux[VER_COMPARA]      := 'CANCELAR'
		aAux[VER_ERRO_COMPARA] := 'O serviço requisitado é diferente de cancelamento.'
		aAux[VER_SALVA]        := Nil
		aAdd(aVerifica, aClone(aAux))
		
		//chave
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_PROCCANCNFE:_RETCANCNFE:_INFCANC:_CHNFE'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a chave do documento cancelado.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_CHAVE
		aAdd(aVerifica, aClone(aAux))
		
		//data do cancelamento
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_PROCCANCNFE:_RETCANCNFE:_INFCANC:_DHRECBTO'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a data do cancelamento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_CANCELADO
		aAdd(aVerifica, aClone(aAux))
	ElseIf lEvento
		aVerifica := {}
		
		//descrição do evento
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_PROCEVENTONFE:_EVENTO:_INFEVENTO:_DETEVENTO:_DESCEVENTO'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar descrição do evento.'
		aAux[VER_COMPARA]      := 'Cancelamento'
		aAux[VER_ERRO_COMPARA] := 'O evento é diferente de cancelamento.'
		aAux[VER_SALVA]        := Nil
		aAdd(aVerifica, aClone(aAux))
		
		//chave
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_PROCEVENTONFE:_RETEVENTO:_INFEVENTO:_CHNFE'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a chave do documento do evento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_CHAVE
		aAdd(aVerifica, aClone(aAux))
		
		//data do evento de cancelamento
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_PROCEVENTONFE:_RETEVENTO:_INFEVENTO:_DHREGEVENTO'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a data do evento de cancelamento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_CANCELADO
		aAdd(aVerifica, aClone(aAux))
	EndIf
	
	aAux := Array(XML_FINAL - 1)
	nMaxI := Len(aVerifica)
	For nI := 1 to nMaxI
		If Type(aVerifica[nI][VER_TAG]) == 'U'
			lRet := .F.
			GridData(cES, nRow, 'NOCHECKED')
			LeErro(cES, cFile, aVerifica[nI][VER_ERRO_TAG])
			Return lRet
		EndIf
		
		cAux := AllTrim(&(aVerifica[nI][VER_TAG] + ':TEXT'))
		
		If aVerifica[nI][VER_COMPARA] <> Nil
			If cAux <> aVerifica[nI][VER_COMPARA]
				lRet := .F.
				GridData(cES, nRow, 'NOCHECKED')
				LeErro(cES, cFile, aVerifica[nI][VER_ERRO_COMPARA])
				Return lRet
			EndIf
		EndIf
		
		If aVerifica[nI][VER_SALVA] <> Nil
			aAux[aVerifica[nI][VER_SALVA]] := cAux
		EndIf
	Next nI
	
	If lNFe200 .or. lNFe310
		/*
		If cES == 'E'
			If cTpNF == '1'
				//documento de saída do Emitente (Fornecedor)
				//Entrada para a empresa
				aAux[XML_TIPO] := 'Entrada'
			ElseIf cTpNF == '0'
				//documento de entrada do Emitente (Própria Empresa)
				//Entrada-Devolução para a empresa
				aAux[XML_TIPO] := 'Entrada-Devolução'
			EndIf
		ElseIf cES == 'S'
			If cTpNF == '0'
				//documento de entrada do Emitente (Fornecedor/Cliente)
				//Saída para a empresa
				aAux[XML_TIPO] := 'Saída'
			ElseIf cTpNF == '1'
				//documento de saída do Emitente (Própria Empresa)
				//Saída-Devolução para a empresa
				aAux[XML_TIPO] := 'Saída-Devolução'
			EndIf
		EndIf
		*/
		If cTpNF == '0'
			aAux[XML_TIPO] := 'Entrada'
		Elseif cTpNF == '1'
			aAux[XML_TIPO] := 'Saída'
		EndIf
		
		nPos := aScan(_aXML, {|x| x[XML_CHAVE] == aAux[XML_CHAVE] })
		If nPos > 0
			//já existe no vetor (é Cancelamento)
			//coloca data do cancelamento no vetor auxiliar
			//substitui posição do vetor existente pelo auxiliar
			aAux[XML_CANCELADO] := _aXML[nPos][XML_CANCELADO]
			_aXML[nPos] := aClone(aAux)
		Else
			aAdd(_aXML, aClone(aAux))
		EndIf
	ElseIf lCanc .or. lEvento
		nPos := aScan(_aXML, {|x| x[XML_CHAVE] == aAux[XML_CHAVE] })
		If nPos > 0
			//já existe no vetor (é NF-e)
			//coloca data do cancelamento ou do evento do cancelamento na posição do vetor existente
			_aXML[nPos][XML_CANCELADO] := aAux[XML_CANCELADO]
		Else
			aAdd(_aXML, aClone(aAux))
		EndIf
	EndIf
	
/* ///////////////
 *
 * CT-e
 *
 * ///////////////
 */
ElseIf cModelo == '57'
	If Type('_oXML:_CTEPROC:_CTE:_INFCTE:_IDE') == 'U'
		//não existe tag do XML do CT-e layout 2.00
		lCTe := .F.
	EndIf
	
	If !lCTe
		lRet := .F.
		GridData(cES, nRow, 'NOCHECKED')
		LeErro(cES, cFile, 'O XML não está em um formato reconhecido pela análise.')
		Return lRet
	EndIf
	
	If lCTe
		//verificando tomador
		lToma03 := .T.
		lToma4 := .T.
		
		If Type('_oXML:_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA03:_TOMA') == 'U'
			lToma03 := .F.
		EndIf
		
		If Type('_oXML:_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA4:_TOMA') == 'U'
			lToma4 := .F.
		EndIf
		
		If !lToma03 .and. !lToma4
			lRet := .F.
			GridData(cES, nRow, 'NOCHECKED')
			LeErro(cES, cFile, 'Impossível verificar o tomador do documento.')
			Return lRet
		EndIf
		
		nTomador := -1
		If lToma03
			nTomador := Val(AllTrim(_oXML:_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA03:_TOMA:TEXT))
		ElseIf lToma4
			nTomador := Val(AllTrim(_oXML:_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA4:_TOMA:TEXT))
		EndIf
			
		aVerifica := {}
		
		If cES == 'E'
			//documento de entrada emitido pelo nosso fornecedor
			
			//verifica se CNPJ da empresa logada faz parte do CT-e
			lCNPJ := .F.
			lEntToma := .F.
			cTomador := ''
			nMaxI := Len(aCTePart)
			For nI := 1 to nMaxI
				cTagToma := aCTePart[nI][1]
				cNomeToma := aCTePart[nI][2]
				cAux := '_oXML:_CTEPROC:_CTE:_INFCTE:' + cTagToma + ':_CNPJ'
				If Type(cAux) <> 'U'
					cAux := AllTrim(&(cAux + ':TEXT'))
					If cAux == SM0->M0_CGC
						lCNPJ := .T.
						cTomador += Upper(Left(cNomeToma, 1)) + SubStr(cNomeToma, 2) + ', '
						If (nTomador + 1) == nI
							lEntToma := .T.
						EndIf
					EndIf
				EndIf
			Next nI
			If !lCNPJ
				lRet := .F.
				GridData(cES, nRow, 'NOCHECKED')
				LeErro(cES, cFile, 'A empresa que você está logado não está entre os participantes do documento.')
				Return lRet
			EndIf
			cTomador := SubStr(cTomador, 1, Len(cTomador) - 2)
			
			If nTomador <> -1
				cTagToma := aCTePart[nTomador + 1][1]
				cNomeToma := aCTePart[nTomador + 1][2]
			Else
				cTagToma := ''
				cNomeToma := ''
			EndIf
			
			//emitente (fornecedor)
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:_EMIT:_CNPJ'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o CNPJ do emitente do documento.'
			aAux[VER_COMPARA]      := Nil
			aAux[VER_ERRO_COMPARA] := Nil
			aAux[VER_SALVA]        := XML_CLIEFOR
			aAdd(aVerifica, aClone(aAux))
			
			//nome do emitente (fornecedor)
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:_EMIT:_XNOME'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o nome do emitente do documento.'
			aAux[VER_COMPARA]      := Nil
			aAux[VER_ERRO_COMPARA] := Nil
			aAux[VER_SALVA]        := XML_NOME
			aAdd(aVerifica, aClone(aAux))
		ElseIf cES == 'S'
			//documento de saída emitido pela própria empresa
			
			//emitente
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:_EMIT:_CNPJ'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o CNPJ do emitente do documento.'
			aAux[VER_COMPARA]      := SM0->M0_CGC
			aAux[VER_ERRO_COMPARA] := 'O emitente do documento é diferente da empresa que você está logado.'
			aAux[VER_SALVA]        := Nil
			aAdd(aVerifica, aClone(aAux))
			
			If nTomador <> -1
				cTagToma := aCTePart[nTomador + 1][1]
				cNomeToma := aCTePart[nTomador + 1][2]
			Else
				cTagToma := ''
				cNomeToma := ''
			EndIf
			
			lCNPJ := .F.
			lCPF := .F.
			
			If Type('_oXML:_CTEPROC:_CTE:_INFCTE:' + cTagToma + ':_CNPJ') <> 'U'
				lCNPJ := .T.
			EndIf
			
			If Type('_oXML:_CTEPROC:_CTE:_INFCTE:' + cTagToma + ':_CPF') <> 'U'
				lCPF := .T.
			EndIf

			If !lCNPJ .and. !lCPF
				lRet := .F.
				GridData(cES, nRow, 'NOCHECKED')
				LeErro(cES, cFile, 'Impossível verificar o CNPJ/CPF do ' + cNomeToma + Iif(nTomador<>4,' (tomador)','') + ' do documento.')
				Return lRet
			EndIf
			
			If lCNPJ
				aAux := Array(VER_FINAL - 1)
				aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:' + cTagToma + ':_CNPJ'
				aAux[VER_ERRO_TAG]     := 'Impossível verificar o CNPJ do ' + cNomeToma + Iif(nTomador<>4,' (tomador)','') + ' do documento.'
				aAux[VER_COMPARA]      := Nil
				aAux[VER_ERRO_COMPARA] := Nil
				aAux[VER_SALVA]        := XML_CLIEFOR
				aAdd(aVerifica, aClone(aAux))
			EndIf
			If lCPF
				aAux := Array(VER_FINAL - 1)
				aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:' + cTagToma + ':_CPF'
				aAux[VER_ERRO_TAG]     := 'Impossível verificar o CPF do ' + cNomeToma + Iif(nTomador<>4,' (tomador)','') + ' do documento.'
				aAux[VER_COMPARA]      := Nil
				aAux[VER_ERRO_COMPARA] := Nil
				aAux[VER_SALVA]        := XML_CLIEFOR
				aAdd(aVerifica, aClone(aAux))
			EndIf
			
			aAux := Array(VER_FINAL - 1)
			aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:' + cTagToma + ':_XNOME'
			aAux[VER_ERRO_TAG]     := 'Impossível verificar o nome do ' + cNomeToma + Iif(nTomador<>4,' (tomador)','') + ' do documento.'
			aAux[VER_COMPARA]      := Nil
			aAux[VER_ERRO_COMPARA] := Nil
			aAux[VER_SALVA]        := XML_NOME
			aAdd(aVerifica, aClone(aAux))
		EndIf
		
		//modelo
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:_IDE:_MOD'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar modelo do documento.'
		aAux[VER_COMPARA]      := cModelo
		aAux[VER_ERRO_COMPARA] := 'O modelo do documento é diferente de CT-e.'
		aAux[VER_SALVA]        := XML_MODELO
		aAdd(aVerifica, aClone(aAux))
		
		//série
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:_IDE:_SERIE'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a série do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_SERIE
		aAdd(aVerifica, aClone(aAux))
		
		//número
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:_IDE:_NCT'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar o número do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_NUMERO
		aAdd(aVerifica, aClone(aAux))
		
		//emissão
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:_IDE:_DHEMI'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a data de emissão do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_EMISSAO
		aAdd(aVerifica, aClone(aAux))
		
		//valor
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_CTEPROC:_CTE:_INFCTE:_VPREST:_VTPREST'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar o valor do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_VALOR
		aAdd(aVerifica, aClone(aAux))
		
		//chave
		aAux := Array(VER_FINAL - 1)
		aAux[VER_TAG]          := '_oXML:_CTEPROC:_PROTCTE:_INFPROT:_CHCTE'
		aAux[VER_ERRO_TAG]     := 'Impossível verificar a chave do documento.'
		aAux[VER_COMPARA]      := Nil
		aAux[VER_ERRO_COMPARA] := Nil
		aAux[VER_SALVA]        := XML_CHAVE
		aAdd(aVerifica, aClone(aAux))
	EndIf
	
	aAux := Array(XML_FINAL - 1)
	nMaxI := Len(aVerifica)
	For nI := 1 to nMaxI
		If Type(aVerifica[nI][VER_TAG]) == 'U'
			lRet := .F.
			GridData(cES, nRow, 'NOCHECKED')
			LeErro(cES, cFile, aVerifica[nI][VER_ERRO_TAG])
			Return lRet
		EndIf
		
		cAux := AllTrim(&(aVerifica[nI][VER_TAG] + ':TEXT'))
		
		If aVerifica[nI][VER_COMPARA] <> Nil
			If cAux <> aVerifica[nI][VER_COMPARA]
				lRet := .F.
				GridData(cES, nRow, 'NOCHECKED')
				LeErro(cES, cFile, aVerifica[nI][VER_ERRO_COMPARA])
				Return lRet
			EndIf
		EndIf
		
		If aVerifica[nI][VER_SALVA] <> Nil
			aAux[aVerifica[nI][VER_SALVA]] := cAux
		EndIf
	Next nI
	
	If lCTe
		If cES == 'E'
			aAux[XML_TOMAFRETE] := lEntToma
			aAux[XML_PARTFRETE] := cTomador
		ElseIf cES == 'S'
			aAux[XML_TOMAFRETE] := .F.
			aAux[XML_PARTFRETE] := 'Emitente'
		EndIf
		
		nPos := aScan(_aXML, {|x| x[XML_CHAVE] == aAux[XML_CHAVE] })
		If nPos > 0
			//já existe no vetor (é Cancelamento)
			//coloca data do cancelamento no vetor auxiliar
			//substitui posição do vetor existente pelo auxiliar
			aAux[XML_CANCELADO] := _aXML[nPos][XML_CANCELADO]
			_aXML[nPos] := aClone(aAux)
		Else
			aAdd(_aXML, aClone(aAux))
		EndIf
	EndIf
EndIf

Return lRet

/* -------------- */

Static Function LeErro(cES, cFile, cErro, lAviso)

Local cMsgErro := ' - <b>Espécie:</b> NF-e' + CRLF + '<b>Empresa:</b> ' + cEmpAnt + ' - <b>Filial:</b> ' + cFilAnt + ' - <b>Ano:</b> ' + StrZero(MV_PAR02, 4) + ' - <b>Mês:</b> ' + StrZero(MV_PAR01, 2)
Local cEntSai := ''
Local cMsg := ''
Default lAviso := .F.

If cES == 'E'
	cEntSai := 'Entrada'
ElseIf cES == 'S'
	cEntSai := 'Saída'
EndIf

cMsg := cErro + CRLF +;
'<b>Arquivo:</b> ' + cFile + CRLF +;
'<b>Documento:</b> ' + cEntSai + cMsgErro

cMsg := StrTran(cMsg, CRLF, '<br />')

MsgErro(cMsg, lAviso)

Return Nil

/* -------------- */

Static Function MsgErro(cMsg, lAviso)

Local cImage := ''
Local cTexto := ''
Default cMsg := 'Erro!'
Default lAviso := .F.

If lAviso
	cImage := 'updwarning.png'
	cTexto := '     Continuar'
Else
	cImage := 'upderror.png'
	cTexto := '     Finalizar'
EndIf

_lContinua := .F.
_oEdt:Load(cMsg)
_oBtn3:SetCss('QPushButton{ background-image: url(rpo:' + cImage + '); background-repeat: none; padding-left: 3px; padding-top: 1px; background-origin: content; background-clip: margin; }')
_oBtn3:SetText(cTexto)
_oBtn3:Enable()

While !_lContinua
	ProcessMessages()
EndDo

Return Nil

/* -------------- */

Static Function Livro(cTitulo, lErro)

Local cAnoMes := StrZero(MV_PAR02,4) + StrZero(MV_PAR01,2)
Local cQry
Local nQtd, nCnt, nPos
Local aAux, cAux

ProcRegua(0)
ProcessMessages()

cQry := ""
cQry += CRLF + " SELECT "
cQry += CRLF + "    FT_TIPOMOV"
cQry += CRLF + "   ,FT_TIPO"
cQry += CRLF + "   ,FT_ESPECIE"
cQry += CRLF + "   ,FT_SERIE"
cQry += CRLF + "   ,FT_NFISCAL"
cQry += CRLF + "   ,FT_EMISSAO"
cQry += CRLF + "   ,FT_ENTRADA"
cQry += CRLF + "   ,FT_CLIEFOR"
cQry += CRLF + "   ,FT_LOJA"
cQry += CRLF + "   ,A1_CGC"
cQry += CRLF + "   ,A1_NOME"
cQry += CRLF + "   ,A2_CGC"
cQry += CRLF + "   ,A2_NOME"
cQry += CRLF + "   ,SUM(FT_VALCONT) VALCONT"
cQry += CRLF + "   ,FT_CHVNFE"
cQry += CRLF + "   ,FT_DTCANC"
cQry += CRLF + " FROM " + RetSqlName('SFT') + " SFT"
cQry += CRLF + " LEFT JOIN " + RetSqlName('SA2') + " SA2"
cQry += CRLF + " ON  SA2.D_E_L_E_T_ <> '*'"
cQry += CRLF + " AND A2_FILIAL = '" + xFilial('SA2') + "'"
cQry += CRLF + " AND A2_COD = FT_CLIEFOR"
cQry += CRLF + " AND A2_LOJA = FT_LOJA"
cQry += CRLF + " AND ("
cQry += CRLF + "   ("
cQry += CRLF + "         FT_TIPOMOV = 'E'"
cQry += CRLF + "     AND FT_TIPO NOT IN ('D','B')"
cQry += CRLF + "   ) OR ("
cQry += CRLF + "         FT_TIPOMOV = 'S'"
cQry += CRLF + "     AND FT_TIPO IN ('D','B')"
cQry += CRLF + "   )"
cQry += CRLF + " )"
cQry += CRLF + " LEFT JOIN " + RetSqlName('SA1') + " SA1"
cQry += CRLF + " ON  SA1.D_E_L_E_T_ <> '*'"
cQry += CRLF + " AND A1_FILIAL = '" + xFilial('SA1') + "'"
cQry += CRLF + " AND A1_COD = FT_CLIEFOR"
cQry += CRLF + " AND A1_LOJA = FT_LOJA"
cQry += CRLF + " AND ("
cQry += CRLF + "   ("
cQry += CRLF + "         FT_TIPOMOV = 'S'"
cQry += CRLF + "     AND FT_TIPO NOT IN ('D','B')"
cQry += CRLF + "   ) OR ("
cQry += CRLF + "         FT_TIPOMOV = 'E'"
cQry += CRLF + "     AND FT_TIPO IN ('D','B')"
cQry += CRLF + "   )"
cQry += CRLF + " )"
cQry += CRLF + " WHERE SFT.D_E_L_E_T_ <> '*'"
cQry += CRLF + "   AND FT_FILIAL = '" + xFilial('SFT') + "'"
cQry += CRLF + "   AND LEFT(FT_EMISSAO,6) = '" + cAnoMes + "'"
If MV_PAR03 == 1
	cQry += CRLF + "   AND FT_TIPOMOV = 'E'"
ElseIf MV_PAR03 == 2
	cQry += CRLF + "   AND FT_TIPOMOV = 'S'"
EndIf
If MV_PAR04 == 1
	cQry += CRLF + "   AND FT_ESPECIE = 'SPED'"
ElseIf MV_PAR04 == 2
	cQry += CRLF + "   AND FT_ESPECIE = 'CTE'"
ElseIf MV_PAR04 == 3
	cQry += CRLF + "   AND FT_ESPECIE IN ('SPED','CTE')"
EndIf
cQry += CRLF + " GROUP BY"
cQry += CRLF + "    FT_TIPOMOV"
cQry += CRLF + "   ,FT_TIPO"
cQry += CRLF + "   ,FT_ESPECIE"
cQry += CRLF + "   ,FT_SERIE"
cQry += CRLF + "   ,FT_NFISCAL"
cQry += CRLF + "   ,FT_EMISSAO"
cQry += CRLF + "   ,FT_ENTRADA"
cQry += CRLF + "   ,FT_CLIEFOR"
cQry += CRLF + "   ,FT_LOJA"
cQry += CRLF + "   ,A1_CGC"
cQry += CRLF + "   ,A1_NOME"
cQry += CRLF + "   ,A2_CGC"
cQry += CRLF + "   ,A2_NOME"
cQry += CRLF + "   ,FT_CHVNFE"
cQry += CRLF + "   ,FT_DTCANC"

dbUseArea(.T.,'TOPCONN',TCGenQry(,,cQry),'MQRY',.T.)
nQtd := 0
nCnt := 0
MQRY->(dbEval({|| nQtd++ }))
MQRY->(dbGoTop())
ProcRegua(nQtd)
While !MQRY->(Eof())
	nCnt += 1
	
	IncProc('Analisando Livro Fiscal ' + cValToChar(nCnt) + '/' + cValToChar(nQtd) + '...')
	
	aAux := Array(LIVRO_FINAL - 1)
	aAux[LIVRO_TIPO]      := ''
	If AllTrim(MQRY->FT_TIPOMOV) == 'E'
		aAux[LIVRO_TIPO]      += 'Entrada'
	ElseIf AllTrim(MQRY->FT_TIPOMOV) == 'S'
		aAux[LIVRO_TIPO]      += 'Saída'
	EndIf
	aAux[LIVRO_TIPO]      += TipoNota(AllTrim(MQRY->FT_TIPO))
	aAux[LIVRO_ESPECIE]   := AllTrim(MQRY->FT_ESPECIE)
	aAux[LIVRO_SERIE]     := AllTrim(MQRY->FT_SERIE)
	aAux[LIVRO_NUMERO]    := AllTrim(MQRY->FT_NFISCAL)
	aAux[LIVRO_EMISSAO]   := AllTrim(MQRY->FT_EMISSAO)
	aAux[LIVRO_ENTRADA]   := AllTrim(MQRY->FT_ENTRADA)
	aAux[LIVRO_CLIEFOR]   := AllTrim(MQRY->FT_CLIEFOR)
	aAux[LIVRO_LOJA]      := AllTrim(MQRY->FT_LOJA)
	If AllTrim(MQRY->A1_CGC) <> ''
		aAux[LIVRO_CNPJ]      := AllTrim(MQRY->A1_CGC)
		aAux[LIVRO_NOME]      := AllTrim(MQRY->A1_NOME)
	ElseIf AllTrim(MQRY->A2_CGC) <> ''
		aAux[LIVRO_CNPJ]      := AllTrim(MQRY->A2_CGC)
		aAux[LIVRO_NOME]      := AllTrim(MQRY->A2_NOME)
	EndIf
	aAux[LIVRO_VALOR]     := MQRY->VALCONT
	aAux[LIVRO_CHAVE]     := AllTrim(MQRY->FT_CHVNFE)
	aAux[LIVRO_CANCELADO] := AllTrim(MQRY->FT_DTCANC)
	
	nPos := aScan(_aXML, {|x| x[XML_CHAVE] == aAux[LIVRO_CHAVE] })
	If nPos > 0
		aAux[LIVRO_XML] := nPos
	Else
		cAux := ''
		If aAux[LIVRO_ESPECIE] == 'SPED'
			cAux += '55'
		ElseIf aAux[LIVRO_ESPECIE] == 'CTE'
			cAux += '57'
		EndIf
		cAux += StrZero(Val(aAux[LIVRO_SERIE]),3) + StrZero(Val(aAux[LIVRO_NUMERO]),9) + aAux[LIVRO_CNPJ]
		nPos := aScan(_aXML, {|x| x[XML_MODELO] + StrZero(Val(x[XML_SERIE]),3) + StrZero(Val(x[XML_NUMERO]),9) + x[XML_CLIEFOR] == cAux })
		If nPos > 0
			aAux[LIVRO_XML] := nPos
		EndIf
	EndIf
	If aAux[LIVRO_XML] == Nil
		_nSemRelac += 1
	EndIf
	
	aAdd(_aLivro, aClone(aAux))
	
	MQRY->(dbSkip())
EndDo
MQRY->(dbCloseArea())

Return Nil

/* -------------- */

Static Function TipoNota(cTipo)

Local aTipo := {;
	{'D','Devolução'},;
	{'B','Beneficiamento'},;
	{'I','Compl.ICMS'},;
	{'I','Compl.Preço/Frete'},;
	{'L','Nota em Lote'},;
	{'S','Nota com ISS'} ;
}
Local cRet := ''
Local nPos := 0

If cTipo <> ''
	nPos := aScan(aTipo, {|x| x[1] == cTipo })
	If nPos > 0
		cRet := '-' + aTipo[nPos][2]
	EndIf
EndIf

Return cRet

/* -------------- */

Static Function Dados(cTitulo, lErro)

Local cRet, cAux
Local cArq := 'MYRXMLLF'
Local aExcel := {}
Local aAux
Local nI, nJ, nMaxI, nMaxJ
Local aLivro := {'Entrada','Saída','Ambos'}
Local aEspDoc := {'NF-e','CT-e','Ambos'}
Local uAux

ProcRegua(0)
ProcessMessages()

aAdd(aExcel, {cTitulo})
aAdd(aExcel, {'Relatório emitido em '+DTOC(Date())+' às '+Time()+' por '+AllTrim(cUsername)})
aAdd(aExcel, {'Empresa/Filial: '+cEmpAnt+'/'+cFilAnt+' - '+AllTrim(SM0->M0_NOME)+' / '+AllTrim(SM0->M0_FILIAL)})
aAdd(aExcel, {'Mês: '+StrZero(MV_PAR01,2)+' - Ano: '+StrZero(MV_PAR02,4)})
aAdd(aExcel, {'Livro: '+aLivro[MV_PAR03]+' - NF-e/CT-e: '+aEspDoc[MV_PAR04]})
aAdd(aExcel, {''})

aAdd(aExcel, {'Documentos com XML e lançados no Livro Fiscal'})
aAdd(aExcel, {;
	'XML', 'XML' , 'XML', 'XML', 'XML', 'XML', 'XML', 'XML', 'XML', 'XML', 'XML', 'XML', ;
	'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal', 'L.Fiscal' ;
})
aAdd(aExcel, {;
	'Tipo', 'Modelo' , 'Série', 'Número', 'Emissão', 'CNPJ', 'Nome', 'Valor', 'Chave', 'Cancelado', 'Participante', 'Tomador',; //XML
	'Tipo', 'Espécie', 'Série', 'Número', 'Emissão', 'Entrada', 'Cli/For', 'Loja', 'CNPJ', 'Nome', 'Valor Contábil', 'Chave', 'Cancelado' ; //L.Fiscal
})

nMaxI := Len(_aXML)

ProcRegua(nMaxI + _nSemRelac)

For nI := 1 to nMaxI
	IncProc()
	
	nPos := aScan(_aLivro, {|x| x[LIVRO_XML] == nI })
	If nPos > 0
		aAux := {}
		ConvXML(@aAux, nI)
		ConvLivro(@aAux, nPos)
		aAdd(aExcel, aClone(aAux))
		
		_aXML[nI][XML_LIVRO] := nPos
	EndIf
Next nI

aAdd(aExcel, {''})
aAdd(aExcel, {'Documentos XML que não estão lançados no Livro Fiscal'})
aAdd(aExcel, {;
	'Tipo', 'Modelo' , 'Série', 'Número', 'Emissão', 'CNPJ', 'Nome', 'Valor', 'Chave', 'Cancelado', 'Participante', 'Tomador', 'Observação' ; //XML
})
//adiciona XML sem relacionamento
nMaxI := Len(_aXML)
For nI := 1 to nMaxI
	If _aXML[nI][XML_LIVRO] == Nil
		IncProc()
		
		aAux := {}
		ConvXML(@aAux, nI)
		If _aXML[nI][XML_MODELO] == '57' .and. !_aXML[nI][XML_TOMAFRETE]
			aAdd(aAux, 'A empresa é participante apenas como parte da negociação e não como prestador ou tomador do serviço.')
		EndIf
		aAdd(aExcel, aClone(aAux))
	EndIf
Next nI

aAdd(aExcel, {''})
aAdd(aExcel, {'Documentos lançados no Livro Fiscal que não estão na pasta dos XMLs'})
aAdd(aExcel, {;
	'Tipo', 'Espécie', 'Série', 'Número', 'Emissão', 'Entrada', 'Cli/For', 'Loja', 'CNPJ', 'Nome', 'Valor Contábil', 'Chave', 'Cancelado' ; //L.Fiscal
})
//adiciona Livro sem relacionamento
nMaxI := Len(_aLivro)
For nI := 1 to nMaxI
	If _aLivro[nI][LIVRO_XML] == Nil
		IncProc()
		
		aAux := {}
		ConvLivro(@aAux, nI)
		aAdd(aExcel, aClone(aAux))
	EndIf
Next nI

IncProc()

cAux := AllTrim(cGetFile('CSV (*.csv)|*.csv', 'Selecione o diretório onde será salvo o relatório', 1, 'C:\', .T., nOR( GETF_LOCALHARD, GETF_LOCALFLOPPY, GETF_NETWORKDRIVE, GETF_RETDIRECTORY ), .F., .T.))
If cAux <> ''
	cAux := SubStr(cAux, 1, RAt('\', cAux)) + cArq
	cAux := cAux + '-' + DTOS(Date()) + '-' + StrTran(Time(), ':', '') + '.csv'
	
	cRet := U_MyArrCsv(aExcel, cAux, Nil, cTitulo)
	If cRet <> ''
		Alert(cRet)
	EndIf
Else
	Alert('A geração do relatório foi cancelada!')
EndIf
			
Return Nil

/* -------------- */

Static Function ConvXML(aAux, nPos)

Local nJ, nMaxJ
Local uAux

nMaxJ := XML_FINAL - 2 //o XML_LIVRO não entra
For nJ := 1 to nMaxJ
	uAux := _aXML[nPos][nJ]
	If uAux <> Nil
		If nJ == XML_SERIE
			uAux := StrZero(Val(uAux), 3)
		ElseIf nJ == XML_NUMERO
			uAux := StrZero(Val(uAux), 9)
		ElseIf nJ == XML_EMISSAO .or. nJ == XML_CANCELADO
			If uAux <> ''
				uAux := STOD(StrTran(Left(uAux, 10), '-', ''))
			EndIf
		ElseIf nJ == XML_VALOR
			uAux := Val(uAux)
		ElseIf nJ == XML_TOMAFRETE
			If uAux
				uAux := 'Sim'
			Else
				uAux := 'Não'
			EndIf
		EndIf
	Else
		uAux := ''
	EndIf
	aAdd(aAux, uAux)
Next nJ
	
Return Nil

/* -------------- */

Static Function ConvLivro(aAux, nPos)

Local nJ, nMaxJ
Local uAux

nMaxJ := LIVRO_FINAL - 2 //o LIVRO_XML não entra
For nJ := 1 to nMaxJ
	uAux := _aLivro[nPos][nJ]
	If uAux <> Nil
		If nJ == LIVRO_SERIE
			uAux := StrZero(Val(uAux), 3)
		ElseIf nJ == LIVRO_NUMERO
			uAux := StrZero(Val(uAux), 9)
		ElseIf nJ == LIVRO_EMISSAO .or. nJ == LIVRO_ENTRADA .or. nJ == LIVRO_CANCELADO
			If uAux <> ''
				uAux := STOD(uAux)
			EndIf
		EndIf
	Else
		uAux := ''
	EndIf
	aAdd(aAux, uAux)
Next nJ

Return Nil

/* -------------- */

Static Function CriaSX1(cPerg)

Local nTamGrp := Len(SX1->X1_GRUPO)
Local aHelpPor := {}, aHelpEng := {}, aHelpSpa := {}
Local cNome

aHelpPor := {}
aAdd(aHelpPor, 'Informe o mês para geração do        ')
aAdd(aHelpPor, 'relatório.                           ')
cNome := 'Mês'
PutSx1(PadR(cPerg,nTamGrp), '01', cNome, cNome, cNome,;
'MV_CH1', 'N', 2, 0, 0, 'G', '', '', '', '', 'MV_PAR01',;
'', '', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
aClone(aHelpPor), aClone(aHelpEng), aClone(aHelpSpa))

aHelpPor := {}
aAdd(aHelpPor, 'Informe o ano para geração do        ')
aAdd(aHelpPor, 'relatório.                           ')
cNome := 'Ano'
PutSx1(PadR(cPerg,nTamGrp), '02', cNome, cNome, cNome,;
'MV_CH2', 'N', 4, 0, 0, 'G', '', '', '', '', 'MV_PAR02',;
'', '', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
'', '', '',;
aClone(aHelpPor), aClone(aHelpEng), aClone(aHelpSpa))

aHelpPor := {}
aAdd(aHelpPor, 'Informe de qual livro os documentos  ')
aAdd(aHelpPor, 'devem ser listados:                  ')
aAdd(aHelpPor, '1-Entrada                            ')
aAdd(aHelpPor, '2-Saída                              ')
aAdd(aHelpPor, '3-Ambos                              ')
cNome := 'Livro'
PutSx1(PadR(cPerg,nTamGrp), '03', cNome, cNome, cNome,;
'MV_CH3', 'N', 1, 0, 3, 'C', '', '', '', '', 'MV_PAR03',;
'Entrada', 'Entrada', 'Entrada', '',;
'Saída', 'Saída', 'Saída',;
'Ambos', 'Ambos', 'Ambos',;
'', '', '',;
'', '', '',;
aClone(aHelpPor), aClone(aHelpEng), aClone(aHelpSpa))

aHelpPor := {}
aAdd(aHelpPor, 'Informe qual espécie de documentos   ')
aAdd(aHelpPor, 'devem ser listados:                  ')
aAdd(aHelpPor, '1-NF-e                               ')
aAdd(aHelpPor, '2-CT-e                               ')
aAdd(aHelpPor, '3-Ambos                              ')
cNome := 'Espécie'
PutSx1(PadR(cPerg,nTamGrp), '04', cNome, cNome, cNome,;
'MV_CH4', 'N', 1, 0, 3, 'C', '', '', '', '', 'MV_PAR04',;
'NF-e', 'NF-e', 'NF-e', '',;
'CT-e', 'CT-e', 'CT-e',;
'Ambos', 'Ambos', 'Ambos',;
'', '', '',;
'', '', '',;
aClone(aHelpPor), aClone(aHelpEng), aClone(aHelpSpa))

Return Nil