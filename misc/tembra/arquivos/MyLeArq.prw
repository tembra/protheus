#Include 'rwmake.ch'
#Include 'protheus.ch'
#Include 'fileio.ch'
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
User Function MyLeArq(cArquivo, lString)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DATA       : 02/12/2010
// USER       : THIERES TEMBRA
// ACAO       : LE UM ARQUIVO PARA UM ARRAY
// RETORNO    : Array / String
// PARÂMETROS : cArquivo - Caminho do arquivo
//            : lString  - Identifica se o retorno será em string
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	Local uRet := Nil
	Local aRetorno := {}
	Local cRetorno := ''
	Local cLinha
	Default lString := .F.
	
	If !File(cArquivo)
		If lString
			uRet := cRetorno
		Else
			uRet := aClone(aRetorno)
		EndIf
		Return uRet
	EndIf
	
	fT_fUse(cArquivo)
	fT_fGoTop()
	
	While (!fT_fEof())
		cLinha := fT_fReadLn()
		If lString
			cRetorno += cLinha
		Else
			aAdd(aRetorno, cLinha)
		EndIf
		fT_fSkip()
	EndDo
	
	fT_fUse()
	
	If lString
		uRet := cRetorno
	Else
		uRet := aClone(aRetorno)
	EndIf
Return uRet