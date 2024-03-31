#INCLUDE "PROTHEUS.CH"
#INCLUDE "rwmake.ch"
#INCLUDE "PARMTYPE.CH"
#INCLUDE "FWMVCDEF.ch"
#INCLUDE "RESTFUL.CH"

/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Função    ¦ Final      ¦ Autor ¦                      ¦ Data ¦ 27/02/2024 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦                                                               ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/

User Function aGer2VIA()
	PRIVATE cIndexName := ''
	PRIVATE cIndexKey  := ''
	PRIVATE cFilter    := ''
	Private cCpf       := Space(14)
	Private cNome      := Space(40)
	Private cCodigo    := Space(6)
	Private cLoja      := Space(2)
	Private cCodUser   := RetCodUsr(SubStr(cUsuario,7,15)) //Retorna o Codigo do Usuario 
	Private cNamUser   := UsrRetName( cCodUser )
	Private cCaixa     := Posicione("SA6",2,XFILIAL("SA6")+cNamUser,"A6_NREDUZ")
	Private cOperado   := Posicione("SA6",2,XFILIAL("SA6")+cNamUser,"A6_COD")
	Private cPedido    := ""
	Private cDoc       := ""
	Private cSerie     := ""
	Private oJanela

	@ 304,453 To 477,913 Dialog oJanela Title OemToAnsi("Solicitação 2ª Via da Nota Fiscal")

	@ 15,10 Say OemToAnsi("Código.: (F3)") Size 46,8 Pixel of oJanela
	@ 15,147 Say OemToAnsi("Loja.:") Size 16,8  Pixel of oJanela
	@ 27,10 Say OemToAnsi("CPF.:") Size 46,8 Pixel of oJanela
	@ 40,10 Say OemToAnsi("Nome do Cliente.:") Size 46,8 Pixel of oJanela

	@ 15,65  MSGET oCodigo VAR cCodigo Picture "@!" SIZE 76,10 of oJanela PIXEL F3 "SA1"  VALID  GetCPF() COLOR CLR_HBLUE
	@ 15,165 MsGet oLoja  var cLoja Size 25,10 Pixel of oJanela When .T.
	@ 27,65  MsGet oCpf    var cCpf  Size 76,10 Pixel of oJanela Picture "@R 999.999.999-99" WHEN .F. //Altera .Or. Inclui VALID If(Empty(cCpf),.T.,ChkCPF(cCpf)) .AND. GetCPF()
	@ 40,65  MsGet oNome   var cNome Picture "@!" Size 149,10 Pixel of oJanela When .F.
	@ 64,129 Button OemToAnsi("&Confirmar") Size 36,16 OF oJanela Action PesqCupom() Pixel
	@ 64,177 Button OemToAnsi("&Fechar") Size 36,16 OF oJanela Action oJanela:End() Pixel

	Activate Dialog oJanela Centered
Return


Static Function PesqCupom()
	Local aCpoBrw := {}
	Local aStruct := {}
	Local cAlias  := Alias()
	Local aReg    := {}
	Local aCupom  := {}
	Local x

	
	oJanela:End()
	
	cMarca  := GetMark()
   // Cria tabela temporária para marcação dos itens
	AAdd( aStruct , { "TMP_FIL"    , "C", 02, 0} )  // Filial
	AAdd( aStruct , { "TMP_ORC"    , "C", 06, 0} )  // Orçamento
	AAdd( aStruct , { "TMP_DC"     , "C", 09, 0} )  // Cupom
	AAdd( aStruct , { "TMP_SER"    , "C", 03, 0} )  // Serie
	AAdd( aStruct , { "TMP_DATA"   , "D", 08, 0} )  // Data Emissao
	AAdd( aStruct , { "TMP_TOTAL"  , "N", 16, 4} )  // Valor da Compra
	AAdd( aStruct , { "TMP_CLI"    , "C", 06, 0} )  // Cliente
	AAdd( aStruct , { "TMP_LOJA"   , "C", 02, 0} )  // Loja do Cliente
	AAdd( aStruct , { "TMP_VEND"   , "C", 06, 0} )  // Vendedor
	AAdd( aStruct , { "TMP_HORA"   , "C", 05, 0} )  // Valor da Compra

	cArq := CriaTrab(aStruct,.T.)
	cInd := Criatrab(Nil,.F.)
	Use &cArq Alias TMP New Exclusive
	IndRegua("TMP", cInd, "TMP_DATA",,, "Aguarde selecionando registros....")
	
	cQuery := " SELECT D1_FILIAL L1_FILIAL, '' L1_NUM,D1_FORNECE L1_CLIENTE, D1_LOJA L1_LOJA,D1_EMISSAO L1_EMISNF,D1_TOTAL*-1 L1_VLRTOT,D1_DOC L1_DOC,D1_SERIE L1_SERIE, D1_NFORI L1_DOCPED, D1_SERIORI L1_SERPED, '' L1_HORA, '' L1_VEND "
	cQuery += " FROM "+RetSqlName("SD1")+" SD1, "+RetSqlName("SD2")+" SD2 "
	cQuery += " WHERE SD1.D_E_L_E_T_ = '' AND SD2.D_E_L_E_T_ = '' "
	cQuery += " AND D1_FILIAL = D2_FILIAL "
	cQuery += " AND D1_NFORI = D2_DOC "
	cQuery += " AND D1_SERIORI = D2_SERIE "
	cQuery += " AND D1_COD = D2_COD "
	cQuery += " AND D1_FORNECE = D2_CLIENTE "
	cQuery += " AND D1_LOJA = D2_LOJA "
	cQuery += " AND D1_FORNECE = '"+cCodigo+"'"
	cQuery += " AND D1_LOJA    = '"+cLoja+"'"
	dbUseArea( .T., "TOPCONN", TcGenQry(,,CHANGEQUERY(cQuery)), "TBB", .T., .F. )
	
	dbSelectArea("TBB")
	dbGoTop()
	While !Eof()
		_nPos := aScan( aReg, {|x| x[1] == L1_FILIAL .and. x[3] == L1_CLIENTE  .and. x[9] == L1_DOCPED .and. x[10] == L1_SERPED})
		If _nPos == 0
			aAdd( aReg, { L1_FILIAL, L1_NUM, L1_CLIENTE, L1_LOJA,L1_EMISNF, L1_VLRTOT,L1_DOC, L1_SERIE, L1_DOCPED, L1_SERPED, L1_HORA, L1_VEND } )
		Endif
		dbSkip()
	End
	dbSelectArea("TBB")
	dbCloseArea()
	
	cQuery := " SELECT L1_FILIAL , L1_NUM , L1_CLIENTE,L1_LOJA,L1_EMISNF, L1_VLRTOT ,L1_DOC , L1_SERIE, L1_DOCPED, L1_SERPED,L1_HORA,L1_VEND "
	cQuery += " FROM "+RetSqlName("SL1")+" A, "+RetSqlName("SL2")+" B, "+RetSqlName("SB1")+" C"
	cQuery += " WHERE A.D_E_L_E_T_ = ''  AND B.D_E_L_E_T_ = '' AND C.D_E_L_E_T_ = ''"
	cQuery += " AND L1_FILIAL = L2_FILIAL "
	cQuery += " AND L1_NUM = L2_NUM "
	cQuery += " AND L2_PRODUTO = B1_COD "
	cQuery += " AND B1_GRUPO NOT IN ('R104','R105','R102','R101','R108','H407') "
	cQuery += " AND L1_TIPO IN ('V') "
	cQuery += " AND L1_CLIENTE = '"+cCodigo+"'"
	cQuery += " AND L1_LOJA = '"+cLoja+"'"
	cQuery += " AND L1_FILIAL NOT IN ('60','00') "
	cQuery += " UNION ALL "
	cQuery += " SELECT (CASE C5_PEDSITE WHEN 0 THEN D2_FILIAL ELSE '00' END) L1_FILIAL,C5_NUM L1_NUM,C5_CLIENTE L1_CLIENTE,C5_LOJACLI L1_LOJA, C5_EMISSAO L1_EMISNF,SUM(D2_TOTAL+D2_VALFRE+D2_DESPESA) L1_VLRTOT,D2_DOC L1_DOC, D2_SERIE L1_SERIE, ''L1_DOCPED ,'' L1_SERPED,''L1_HORA,C5_VEND1 L1_VEND "
	cQuery += " FROM "+RetSqlName("SD2")+" SD2 "
	cQuery += " INNER JOIN " + RetSQLName("SB0") + " SB0 ON B0_COD = D2_COD AND SB0.D_E_L_E_T_ = ' '"
	cQuery += " INNER JOIN " + RetSQLName("SB1") + " SB1 ON B1_COD = D2_COD AND SB1.D_E_L_E_T_ = ' '"
	cQuery += " INNER JOIN " + RetSQLName("SF4") + " SF4 ON F4_CODIGO = D2_TES AND SF4.D_E_L_E_T_ = ' '"
	cQuery += " INNER JOIN " + RetSQLName("SC5") + " SC5 ON C5_FILIAL = D2_FILIAL AND C5_NUM = D2_PEDIDO AND SC5.D_E_L_E_T_ = ' '"
	cQuery += " INNER JOIN " + RetSQLName("SF2") + " SF2 ON F2_FILIAL = D2_FILIAL AND F2_DOC = D2_DOC AND F2_SERIE = D2_SERIE "
	cQuery += " WHERE SD2.D_E_L_E_T_ = '' AND D2_FILIAL = '00' AND SUBSTRING(D2_CF,2,3) IN ('102','108','117','403','404','405') "
	cQuery += " AND C5_CLIENTE = '"+cCodigo+"'"
	cQuery += " AND C5_LOJACLI = '"+cLoja+"'"
	cQuery += " GROUP BY (CASE C5_PEDSITE WHEN 0 THEN D2_FILIAL ELSE '00' END),C5_NUM,C5_CLIENTE,C5_LOJACLI,C5_VEND1,D2_DOC, D2_SERIE,D2_LOCAL,C5_EMISSAO "
	cQuery += " ORDER BY L1_FILIAL "

	dbUseArea( .T., "TOPCONN", TcGenQry(,,CHANGEQUERY(cQuery)), "TBB", .T., .F. )
	
	dbSelectArea("TBB")
	dbGoTop()
	While !Eof()
		_nPos := aScan( aReg, {|x| x[1] == L1_FILIAL .and. x[3] == L1_CLIENTE  .and. x[9] == L1_DOC .and. x[10] == L1_SERIE})
		If _nPos == 0
			aAdd( aCupom, { L1_FILIAL, L1_NUM, L1_CLIENTE, L1_LOJA,L1_EMISNF, L1_VLRTOT,L1_DOC, L1_SERIE, L1_DOCPED, L1_SERPED, L1_HORA, L1_VEND } )
		Endif
		dbSkip()
	End
	dbSelectArea("TBB")
	dbCloseArea()
  // Gravação da tabela temporária para marcação
	For x:=1 to Len(aCupom)
		RecLock("TMP",.T.)
		TMP->TMP_FIL     := aCupom[x,1] 
		TMP->TMP_ORC     := aCupom[x,2] 
		TMP->TMP_DC      := iif(!Empty(aCupom[x,7]),aCupom[x,7],aCupom[x,9])
		TMP->TMP_SER     := iif(!Empty(aCupom[x,8]),aCupom[x,8],aCupom[x,10])
		TMP->TMP_DATA    := STOD(aCupom[x,5])
		TMP->TMP_TOTAL   := aCupom[x,6] 
		TMP->TMP_CLI     := aCupom[x,3] 
		TMP->TMP_LOJA    := aCupom[x,4]
		TMP->TMP_VEND    := aCupom[x,12] 
		TMP->TMP_HORA    := aCupom[x,11] 
		MsunLock() // libera os registros bloqueados pela função RecLock()
	Next
	
	DbSelectArea("TMP")
	dbGoTop()
	
	If !(TMP->(Bof()) .And. TMP->(Eof()))
		//Adiciona os campos a serem exibidos no Browsed de Seleção
		aAdd( aCpoBrw, { "TMP_FIL"     ,, "FILIAL"      , "@!" 			 	     } )
		aAdd( aCpoBrw, { "TMP_ORC"     ,, "ORÇAMENTO"   , "@!" 			 	     } )
		aAdd( aCpoBrw, { "TMP_CLI"     ,, "CLIENTE"     , "@!" 			 	     } )
		aAdd( aCpoBrw, { "TMP_DC"      ,, "NOTA"        , "@!" 			 	     } )
		aAdd( aCpoBrw, { "TMP_SER"     ,, "SERIE"       , "@!" 			 	     } )
		aAdd( aCpoBrw, { "TMP_DATA"    ,, "EMISSÃO"     , "@!" 			 	     } )
		aAdd( aCpoBrw, { "TMP_TOTAL"   ,, "VALOR TOTAL" , "@E 99,999,999,999.99"  } )
		
		DEFINE MSDIALOG oDlg TITLE "Histórico de Compras do Cliente" FROM 00,00 TO 400,700 PIXEL
		
		oMark := MsSelect():New( "TMP", ,,aCpoBrw,, cMarca, { 001, 001, 170, 350 } ,,, )
		oMark:oBrowse:Refresh()
		oMark:oBrowse:lHasMark    := .T.
		oMark:oBrowse:lCanAllMark := .F.
		
		@ 175,001 Say OemToAnsi("**** Selecione 1 (um) orçamento acima para imprimir") Size 250,8 Pixel of oDlg COLOR CLR_HRED
		
		@ 180,310 BUTTON  oBut1 PROMPT "&Avançar" SIZE 30,12 OF oDlg PIXEL Action ConCupom()
		@ 180,250 BUTTON oBut2 PROMPT "&Cancelar" SIZE 30,12 OF oDlg PIXEL Action Final(oDlg)
		ACTIVATE MSDIALOG oDlg CENTERED
	Else
		MsgStop("Não foram encontrados dados para seleção!")
	Endif
	DbSelectArea(cAlias)
Return Nil
/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Função    ¦ Final      ¦ Autor ¦                      ¦ Data ¦ 27/02/2024 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦                                                               ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/
Static Function Final(oDlg)
    TMP->(dbCloseArea())
    FErase(cArq+GetDBExtension())
    FErase(cInd+OrdBagExt())        
    oDlg:End()
Return

/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Função    ¦    Marcar  ¦ Autor ¦                      ¦ Data ¦ 27/02/2024 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦                                                               ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/
Static Function Marcar(cMarca)
	RecLock("TMP",.F.)
	TMP->TMP_OK := If( Empty(TMP->TMP_OK) , cMarca, Space(Len(TMP->TMP_OK)))
	MsUnLock()  // libera os registros bloqueados pela função RecLock()
Return
/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Função    ¦    Marcar  ¦ Autor ¦                      ¦ Data ¦ 27/02/2024 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦                                                               ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/
Static Function ConCupom()
    Local aBrw    := {}   
    Local aTabela := {}
    Local cQuery  := ""  
    Local cAlias  := Alias()
    Private oQuadro
    _cFilial      := ""
    _cNota        := ""
    _cSerie       := ""
    _cProduto     := ""
    _cCliente     := ""
    _cEmail       := ""
    _cChaveNf     := ""
	 
    // Cria tabela temporária para marcação dos itens
    AAdd( aTabela , { "TMD_FIL"   , "C", 02, 0} )  // Filial
    AAdd( aTabela , { "TMD_ITEM"  , "C", 02, 0} )  // Item
    AAdd( aTabela , { "TMD_PROD"  , "C", 15, 0} )  // Produto
    AAdd( aTabela , { "TMD_DESC"  , "C", 30, 0} )  // Descrição do Produto
    AAdd( aTabela , { "TMD_QDT"   , "N", 09, 2} )  // Quantidade Comprada
    AAdd( aTabela , { "TMD_VLRU"  , "N", 10, 4} )  // Valor Unitario 
    AAdd( aTabela , { "TMD_VLRT"  , "N", 11, 4} )  // Valor Total
    AAdd( aTabela , { "TMD_NFCE"  , "C", 45, 0} )  //    
                              
    cArq1 := CriaTrab(aTabela,.T.)
    cInd1 := Criatrab(Nil,.F.)
    Use &cArq1 Alias TMD New Exclusive
    IndRegua("TMD", cInd1, "TMD_ITEM",,, "Aguarde selecionando registros....")  

   cQuery := " SELECT L2_FILIAL,L2_ITEM,L2_PRODUTO,L2_DESCRI,L2_QUANT,L2_VRUNIT,L2_VLRITEM,L1_DOC,L1_SERIE,LEFT(L2_DOCSER,6) L2_DOCSER,CASE F2_CHVNFE WHEN '' THEN 'NAO EXISTE CHAVE NFCE' ELSE F2_CHVNFE END AS L1_KEYNFCE,A1_NOME,A1_EMAIL  "  
   cQuery += " FROM "+RetSQLName("SB1")+" SB1, "+RetSQLName("SL2")+" SL2 "   
   cQuery += " LEFT JOIN "+RetSQLName("SL1")+" SL1 ON SL1.D_E_L_E_T_ = ' ' AND SL1.L1_FILIAL = SL2.L2_FILIAL AND SL1.L1_NUM = SL2.L2_NUM "
   cQuery += " INNER JOIN "+RetSQLName("SF2")+" F2 ON F2.D_E_L_E_T_ = ' '  AND F2.F2_DOC=SL2.L2_DOC AND F2.F2_SERIE=SL2.L2_SERIE AND F2.F2_FILIAL=SL2.L2_FILIAL AND F2.F2_DOC=SL1.L1_DOC AND F2.F2_SERIE=SL1.L1_SERIE AND F2.F2_FILIAL=SL1.L1_FILIAL "
   cQuery += " LEFT JOIN "+RetSQLName("SA1")+" SA1 ON SA1.D_E_L_E_T_ = ' ' AND SA1.A1_COD = SL1.L1_CLIENTE AND SA1.A1_LOJA = SL1.L1_LOJA  "
	cQuery += " WHERE SL2.D_E_L_E_T_ = ' ' AND SB1.D_E_L_E_T_= ' '  "
	cQuery += " AND L2_PRODUTO = B1_COD AND SL1.L1_TIPO='V' "
	cQuery += " AND B1_GRUPO NOT IN ('R104','R105','R102','R101','R108','H407') "
	cQuery += " AND L2_NUM = '"+TMP->TMP_ORC+"'"
   cQuery += " AND L2_FILIAL ='"+TMP->TMP_FIL+"'"  
   cQuery += " UNION ALL  " 
	cQuery += " SELECT (CASE C5_PEDSITE WHEN 0 THEN D2_FILIAL ELSE '00' END) L2_FILIAL,D2_ITEM L2_ITEM,D2_COD L2_PRODUTO,B1_DESC L2_DESCRI,D2_QUANT L2_QUANT,D2_PRCVEN L2_VRUNIT,SUM(D2_TOTAL+D2_VALFRE+D2_DESPESA)L2_VLRITEM,D2_DOC L1_DOC,D2_SERIE L1_SERIE,LEFT(D2_DOCFIN,6) L2_DOCSER,CASE F2_CHVNFE WHEN '' THEN 'NAO EXISTE CHAVE NFCE' ELSE F2_CHVNFE END AS L1_KEYNFCE,''A1_NOME,''A1_EMAIL "
	cQuery += " FROM "+RetSQLName("SD2")+" SD2 "
   cQuery += " INNER JOIN " + RetSQLName("SB0") + " SB0 ON B0_COD = D2_COD AND SB0.D_E_L_E_T_ = ' '"
   cQuery += " INNER JOIN " + RetSQLName("SB1") + " SB1 ON B1_COD = D2_COD AND SB1.D_E_L_E_T_ = ' '"
   cQuery += " INNER JOIN " + RetSQLName("SF4") + " SF4 ON F4_CODIGO = D2_TES AND SF4.D_E_L_E_T_ = ' '"
   cQuery += " INNER JOIN " + RetSQLName("SC5") + " SC5 ON C5_FILIAL = D2_FILIAL AND C5_NUM = D2_PEDIDO AND SC5.D_E_L_E_T_ = ' '"
   cQuery += " INNER JOIN " + RetSQLName("SF2") + " SF2 ON F2_FILIAL = D2_FILIAL AND F2_DOC = D2_DOC AND F2_SERIE = D2_SERIE "
   cQuery += " LEFT JOIN  " + RetSQLName("SA1") + " SA1 ON SA1.D_E_L_E_T_ = ' ' AND SA1.A1_COD = SD2.D2_CLIENTE AND SA1.A1_LOJA = SD2.D2_LOJA "
   cQuery += " WHERE SD2.D_E_L_E_T_ = '' AND D2_FILIAL = '00' AND SUBSTRING(D2_CF,2,3) IN ('102','108','117','403','404','405') " 
   cQuery += " AND C5_NUM = '"+TMP->TMP_ORC+"'"
   cQuery += " AND C5_FILIAL = '"+TMP->TMP_FIL+"'"
   cQuery += " GROUP BY (CASE C5_PEDSITE WHEN 0 THEN D2_FILIAL ELSE '00' END),C5_NUM,C5_CLIENTE,C5_LOJACLI,C5_VEND1,D2_DOC, D2_SERIE,D2_LOCAL,C5_EMISSAO,D2_ITEM,D2_COD,B1_DESC,D2_QUANT,D2_PRCVEN,D2_DOCFIN,F2_CHVNFE,A1_NOME,A1_EMAIL  "
	cQuery += " UNION ALL  " 
	cQuery += " SELECT SL1.L1_FILIAL L2_FILIAL,SL2.L2_ITEM,SL2.L2_PRODUTO,SL2.L2_DESCRI,SL2.L2_QUANT,SL2.L2_VRUNIT,SL2.L2_VLRITEM, "
	cQuery += " RES.D2_DOC L1_DOC,RES.D2_SERIE L1_SERIE,''L2_DOCSER,RES.F2_CHVNFE L1_KEYNFCE,''A1_NOME,''A1_EMAIL "	
	cQuery += " FROM "+RetSQLName("SL1")+" SL1 "
	cQuery += " INNER JOIN "+RetSQLName("SL2")+" SL2 ON SL2.D_E_L_E_T_ = ' ' AND SL1.L1_FILIAL = SL2.L2_FILIAL AND SL1.L1_NUM = SL2.L2_NUM "
	cQuery += " INNER JOIN "+RetSQLName("SA1")+" SA1 ON SA1.D_E_L_E_T_ = ' ' AND SA1.A1_COD = SL1.L1_CLIENTE AND SA1.A1_LOJA = SL1.L1_LOJA "
	cQuery += " LEFT  JOIN "+RetSQLName("SC6")+" SC6 ON SC6.D_E_L_E_T_ = ' ' AND SL1.L1_FILIAL = SC6.C6_FILRES AND SL1.L1_NUM = SC6.C6_ORCRES "  
	cQuery += " AND SC6.C6_FILIAL = '00' AND SL2.L2_PRODUTO = SC6.C6_PRODUTO "
	cQuery += " LEFT OUTER JOIN "
	cQuery += " (SELECT SD2.D2_FILIAL, SD2.D2_PEDIDO, SD2.D2_ITEMPV, SD2.D2_COD, SD2.D2_DOCFIN, SD2.D2_SERFIN, SD2.D2_EMISSAO, "
	cQuery += " SUM(SD2.D2_QUANT) D2_QUANT "
	cQuery += " FROM "+RetSQLName("SD2")+" SD2 "
	cQuery += " WHERE SD2.D_E_L_E_T_ = ' ' AND SD2.D2_FILIAL = '00' "
	cQuery += " GROUP BY SD2.D2_FILIAL, SD2.D2_PEDIDO, SD2.D2_ITEMPV, SD2.D2_COD, SD2.D2_DOCFIN, SD2.D2_SERFIN, SD2.D2_EMISSAO "
	cQuery += " ) SD2 ON SD2.D2_FILIAL = SC6.C6_FILIAL AND SD2.D2_PEDIDO = SC6.C6_NUM AND SD2.D2_ITEMPV = SC6.C6_ITEM "
	cQuery += " AND SD2.D2_COD = SC6.C6_PRODUTO "
	cQuery += " LEFT OUTER JOIN "
	cQuery += " (	
	cQuery += " SELECT SD2.D2_FILIAL, SD2.D2_DOC, SD2.D2_SERIE, SD2.D2_COD, SD2.D2_EMISSAO, SUM(SD2.D2_QUANT) D2_QUANT, SUM(SD2.D2_TOTAL) D2_TOTAL ,F2.F2_CHVNFE "
	cQuery += " FROM "+RetSQLName("SD2")+" SD2 "
	cQuery += " INNER JOIN "+RetSQLName("SF2")+" F2 ON F2.D_E_L_E_T_ = ' ' AND F2.F2_DOC=SD2.D2_DOC AND F2.F2_SERIE=SD2.D2_SERIE AND F2.F2_FILIAL=SD2.D2_FILIAL "  
	cQuery += " WHERE SD2.D_E_L_E_T_ = ' ' "
	cQuery += " GROUP BY SD2.D2_FILIAL, SD2.D2_DOC, SD2.D2_SERIE, SD2.D2_COD, SD2.D2_EMISSAO,F2.F2_CHVNFE "
	cQuery += " ) RES ON RES.D2_FILIAL = SL1.L1_FILIAL AND RES.D2_COD = SD2.D2_COD AND RES.D2_QUANT = SD2.D2_QUANT "
	cQuery += " AND RES.D2_EMISSAO = SD2.D2_EMISSAO AND RES.D2_DOC = SD2.D2_DOCFIN AND RES.D2_SERIE = SD2.D2_SERFIN "
	cQuery += " WHERE SL1.D_E_L_E_T_ = ' ' "
	cQuery += " AND SL1.L1_TIPO = 'P' AND SUBSTRING(L2_CF,2,3) IN ('102','108','403','404','405') "
	cQuery += " AND SL1.L1_DOCPED!='' AND SL1.L1_SERPED!='' AND SL2.L2_ENTREGA='1' "
   cQuery += " AND SL1.L1_NUM = '"+TMP->TMP_ORC+"'"
   cQuery += " AND SL1.L1_FILIAL ='"+TMP->TMP_FIL+"'"
	cQuery += "ORDER BY L2_ITEM "
    
    dbUseArea( .T., "TOPCONN", TcGenQry(,,CHANGEQUERY(cQuery)), "TMB", .T., .F. )
          
    DbSelectArea("TMB")
    dbGoTop()
    While !TMB->(EOF()) 
          RecLock("TMD",.T.)
          TMD->TMD_FIL    := TMB->L2_FILIAL
          TMD->TMD_ITEM   := TMB->L2_ITEM
          TMD->TMD_PROD   := TMB->L2_PRODUTO
          TMD->TMD_DESC   := TMB->L2_DESCRI
          TMD->TMD_QDT    := TMB->L2_QUANT
          TMD->TMD_VLRU   := TMB->L2_VRUNIT
          TMD->TMD_VLRT   := TMB->L2_VLRITEM
          TMD->TMD_NFCE   := TMB->L1_KEYNFCE    
          _cFilial        := TMB->L2_FILIAL
          _cNota          := TMB->L1_DOC
          _cSerie	        := TMB->L1_SERIE   
          _cProduto       := Alltrim(TMB->L2_DESCRI)
          _cCliente       := Alltrim(TMB->A1_NOME)
          _cEmail         := Alltrim(TMB->A1_EMAIL)
          _cChaveNf       := TMB->L1_KEYNFCE
          MsunLock() // libera os registros bloqueados pela função RecLock()
          if !EMPTY(TMB->L2_DOCSER)
             cPedido := TMB->L2_DOCSER
          endif
          TMB->(dbSkip())
    Enddo       
    TMB->(dbCloseArea())
    
    if !EMPTY(cPedido)
       cQuery := " SELECT D2_DOCFIN,D2_SERFIN "
       cQuery += " FROM "+RetSQLName("SD2")+" SD2 "
       cQuery += " WHERE SD2.D_E_L_E_T_ = ' ' "
       cQuery += " AND D2_PEDIDO = '"+cPedido+"'"
       cQuery += " AND D2_FILIAL = '00' "
       cQuery += " AND D2_GRUPO NOT IN ('R104','R105','R102','R101','R108','H407') "
       
       dbUseArea( .T., "TOPCONN", TcGenQry(,,CHANGEQUERY(cQuery)), "TD2", .T., .F. )
            
       DbSelectArea("TD2")
       dbGoTop()
       While !TD2->(EOF()) 
           cDoc   :=  TD2->D2_DOCFIN
           cSerie :=  TD2->D2_SERFIN
           TD2->(dbSkip())       
       Enddo       
       TD2->(dbCloseArea())
       cPedido := ""
    endif
     
    DbSelectArea("TMD")
    dbGoTop()
    
    If !(TMD->(Bof()) .And. TMD->(Eof()))
       //Adiciona os campos a serem exibidos no Browsed de Seleção
       aAdd( aBrw, { "TMD_FIL"    ,, "FILIAL"         , "@!" 			 	     } )    
       aAdd( aBrw, { "TMD_ITEM"   ,, "ITEM"           , "@!" 			 	     } )
       aAdd( aBrw, { "TMD_PROD"   ,, "PRODUTO"        , "@!" 			 	     } )
       aAdd( aBrw, { "TMD_DESC"   ,, "DESCRIÇAO"      , "@!" 			 	     } )
       aAdd( aBrw, { "TMD_QDT"    ,, "QUANTIDADE"     , "@!" 			 	     } )
       aAdd( aBrw, { "TMD_VLRU"   ,, "VALOR UNITÁRIO" , "@E 99,999,999,999.99"    } )
       aAdd( aBrw, { "TMD_VLRT"   ,, "VALOR TOTAL"    , "@E 99,999,999,999.99"    } )
       aAdd( aBrw, { "TMD_NFCE"   ,, "CHAVE NOTA"     , "@!" 			 	     } )

       DEFINE MSDIALOG oQuadro TITLE "Confirmação dos dados da Compra" FROM 00,00 TO 400,850 PIXEL
       
       oPanel := MsSelect():New( "TMD",,,aBrw,,, { 001, 001, 170, 420 } ,,, )
       oPanel:oBrowse:Refresh()
      
       @ 180,350 BUTTON oBut2 PROMPT "&Retornar"  SIZE 30,12 OF oQuadro PIXEL Action  Finalizar(oQuadro)	
       @ 180,390 BUTTON oBut1 PROMPT "&Confirmar" SIZE 30,12 OF oQuadro PIXEL Action Pag2Via(_cNota,_cSerie,_cCliente,_cProduto,_cChaveNf,_cEmail,_cFilial)
       ACTIVATE MSDIALOG oQuadro CENTERED
    Else 
      MsgStop("Não foram encontrados dados para seleção!")
      TMD->(dbCloseArea())
    Endif   
    DbSelectArea(cAlias)
Return Nil  

/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Função    ¦ Finalizar  ¦ Autor ¦                      ¦ Data ¦ 27/02/2024 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦                                                               ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/
Static Function Finalizar(oQuadro)
    TMD->(dbCloseArea())
    FErase(cArq1+GetDBExtension())
    FErase(cInd1+OrdBagExt())        
	oQuadro:End()
Return
/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Função    ¦ Pag2Via    ¦ Autor ¦                      ¦ Data ¦ 27/02/2024 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦                                                               ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/
Static Function Pag2Via(_cNota,_cSerie,_cCliente,_cProduto,_cChaveNf,_cEmail,_cFilial)
Local cmsgalert := ""
cmsgalert := "                  ***ATENÇÃO ***                    " 
cmsgalert += CRLF
cmsgalert += CRLF
cmsgalert += "                  Para imprimir a 2ª VIA (DANFE)             "
cmsgalert += CRLF
cmsgalert += CRLF
cmsgalert += "     Você precisa esta na mesma loja onde foi finalizado a venda     "
cmsgalert += CRLF
cmsgalert += CRLF
cmsgalert += "                    (Usar Caixa Personalizado)                                "
cmsgalert += "                                                    "

If Alltrim(_cSerie) == "1" .Or. Alltrim(_cSerie) == "2" .Or. Alltrim(_cSerie) == "3"
    
     If _cFilial == cFilAnt

if MsgBox (" Imprimir 2ª Via ( Danfe ) "," 2ª Via ","YESNO")
   Processa({|| aXMLGER1(_cNota,_cSerie,_cEmail)},"Imprimindo a 2ª. Via, Aguarde...")   
 
else
   MsgStop("Cancelado")
EndIf
      Else
        Aviso( " 2ª VIA DANFE   ", cmsgalert, {"OK"}, 3 )
      EndIf
Else

if MsgBox (" E-mail cupom fiscal ."," 2ª Via ","YESNO")
   Processa({|| aXMLGER2(_cNota,_cSerie,_cCliente,_cProduto,_cChaveNf,_cEmail)},"Imprimindo a 2ª. Via, Aguarde...")   

else
   MsgStop("Cancelado")
 EndIf
EndIf

Finalizar(oQuadro)
Return 

Static Function aXMLGER1(_cNota,_cSerie,_cEmail)
local cmsgalert := ""
cmsgalert := "                  ***ATENÇÃO ***                  " 
cmsgalert += CRLF
cmsgalert += CRLF
cmsgalert += "  AGUARDE A DANFE SER IMPRESSA EM SUA TELA        " 
cmsgalert += CRLF
cmsgalert += CRLF
cmsgalert += " OBSERVAÇÃO VERIFIQUE SE O SOFTWARE  -- JAVA --   "  //JRE-8u401
cmsgalert += CRLF
cmsgalert += CRLF            
cmsgalert += " ESTA INSTALADO EM SEU COMPUTADOR                  "

  Aviso( " GERAR NF - 2ª VIA (DANFE)   ", cmsgalert, {"OK"}, 3 )
	
     u_aGerDanfe(_cNota,_cSerie)
  
Return
Static Function aXMLGER2(_cNota,_cSerie,_cCliente,_cProduto,_cChaveNf,_cEmail)
local cmsgalert := ""
Default APasta  := GetTempPath()
cmsgalert := "           ***ATENÇÃO ***                    " 
cmsgalert += CRLF
cmsgalert += CRLF
cmsgalert += CRLF
cmsgalert += "E-MAIL ENVIADO COM LINK DO CUPOM FISCAL PARA O CLIENTE "
cmsgalert += CRLF
cmsgalert += CRLF

  Aviso(" GERAR NF - 2ª VIA    ", cmsgalert, {"OK"}, 3 )
	
        u_aRsEmail(_cNota,_cSerie,_cCliente,_cProduto,_cChaveNf,_cEmail)
      
Return   



****************************
Static Function GetCPF()
****************************
Local aArea  := GetArea()
Local nH 
Default APasta  := GetTempPath()
DBSelectArea("SA1")
DBSetOrder(1)
DBSeek(xFilial("SA1")+cCodigo)
IF !Empty(cCodigo) 
  cLoja := SA1->A1_LOJA
  If !DBSeek(xFilial("SA1")+cCodigo+cLoja)
     MsgStop("Cliente não cadastrado !!!")
     Return (.F.)
  Endif
Endif
cCpf    := SA1->A1_CGC
cNome   := SA1->A1_NOME
_cEmail := SA1->A1_EMAIL

///criar arquivo para e-mail cliente
nH := fCreate(APasta+"email.txt")
If nH == -1
//"Falha ao criar arquivo - erro 
   Return
Endif
// Escreve o texto mais a quebra de linha CRLF
fWrite(nH,_cEmail+chr(13)+chr(10) )
fClose(nH)
//Arquivo criado 


RestArea(aArea)
Return(.T.) 

     



 