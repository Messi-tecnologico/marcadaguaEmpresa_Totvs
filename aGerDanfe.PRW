//Bibliotecas
#Include "Protheus.ch"
#Include "TBIConn.ch"
#Include "Colors.ch"
#Include "RPTDef.ch"
#Include "FWPrintSetup.ch"

/*/{Protheus.doc} zGerDanfe
Função que gera a danfe e o xml de uma nota em uma pasta passada por parâmetro
@author Atilio
@since 10/02/2019
@version 1.0
@param cNota, characters, Nota que será buscada
@param cSerie, characters, Série da Nota
@param cPasta, characters, Pasta que terá o XML e o PDF salvos
@type function
@example u_aGerDanfe("000123ABC", "1", "C:\TOTVS\NF")
@obs Para o correto funcionamento dessa rotina, é necessário:
    1. Ter baixado e compilado o rdmake danfeii.prw
    2. Ter baixado e compilado o zSpedXML.prw - https://terminaldeinformacao.com/2017/12/05/funcao-retorna-xml-de-uma-nota-em-advpl/
/*/
User Function aGerDanfe(cNota, cSerie, cPasta)
	Local aArea     := GetArea()
	Local cIdent    := ""
	Local cArquivo  := ""
	Local oDanfe    := Nil
	Local lEnd      := .F.
	Local nTamNota  := TamSX3('F2_DOC')[1]
	Local nTamSerie := TamSX3('F2_SERIE')[1]
	Local dDataDe   := sToD("20040101")
	Local dDataAt   := Date()
	local cmsgalert := ""
	Private PixelX
	Private PixelY
	Private nConsNeg
	Private nConsTex
	Private oRetNF
	Private nColAux
	Default cNota   := ""
	Default cSerie  := ""
	Default cCpf    := ""
	Default APasta  := GetTempPath()



	If !Empty(cNota)
		//Pega o IDENT da empresa
		cIdent := RetIdEnti()

		//Se o último caracter da pasta não for barra, será barra para integridade
		If SubStr(APasta, Len(APasta), 1) != "\"
			APasta += "\"
		EndIf

		//Gera o XML da Nota--cNota
		cArquivo := "nf" + "_" +dToS(Date()) + "_" + Substr(StrTran(Time(), ":", "-"),1,5)
		//u_zSpedXML(cNota, cSerie, cPasta + cArquivo  + ".xml", .F.)///gerar XML

		//Define as perguntas da DANFE
		Pergunte("NFSIGW",.F.)
		MV_PAR01 := PadR(cNota,  nTamNota)     //Nota Inicial
		MV_PAR02 := PadR(cNota,  nTamNota)     //Nota Final
		MV_PAR03 := PadR(cSerie, nTamSerie)    //Série da Nota
		MV_PAR04 := 2                          //NF de Saida
		MV_PAR05 := 1                          //Frente e Verso = Sim
		MV_PAR06 := 2                          //DANFE simplificado = Nao
		MV_PAR07 := dDataDe                    //Data De
		MV_PAR08 := dDataAt                    //Data Até

		//Cria a Danfe
		oDanfe := FWMSPrinter():New(cArquivo, IMP_PDF, .F., , .T.)

		//Propriedades da DANFE
		oDanfe:SetResolution(78)
		oDanfe:SetPortrait()
		oDanfe:SetPaperSize(DMPAPER_A4)
		oDanfe:SetMargin(60, 60, 60, 60)

		//Força a impressão em PDF
		oDanfe:nDevice  := 6
		oDanfe:cPathPDF := APasta
		oDanfe:lServer  := .F.
		oDanfe:lViewPDF := .F.

		//Variáveis obrigatórias da DANFE (pode colocar outras abaixo)
		PixelX    := oDanfe:nLogPixelX()
		PixelY    := oDanfe:nLogPixelY()
		nConsNeg  := 0.4
		nConsTex  := 0.5
		oRetNF    := Nil
		nColAux   := 0

		//Chamando a impressão da danfe no RDMAKE
		RptStatus({|lEnd| DanfeProc(@oDanfe, @lEnd, cIdent, , , .T.)}, "Imprimindo Danfe...")
		//Tentando abrir o objeto
		ShellExecute("open", "printdanfe.bat", "", "C:/SMARTCLIENT/", 1)

		oDanfe:Print()
	EndIf

	cmsgalert := "         "
	cmsgalert += CRLF
	cmsgalert += CRLF
	cmsgalert += CRLF
	cmsgalert += "     DANFE GERADA COM SUCESSO..      "  //JRE-8u401
	cmsgalert += CRLF
	cmsgalert += CRLF
	cmsgalert += "    "
	Aviso( "  2 VIA DANFE   ", cmsgalert, {"OK"}, 3 )


	RestArea(aArea)
Return
