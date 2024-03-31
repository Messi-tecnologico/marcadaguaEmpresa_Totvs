#Include "Protheus.ch"
#Include "Tbiconn.ch"
#include "ap5mail.CH"
/*/{Protheus.doc} Mail
Função para disparo do e-mail utilizando TMailMessage e tMailManager com anexos.
@version 1.0
@type function
@obs Deve-se configurar os parâmetros:
    * MV_RELACNT - Conta de login do e-Mail    - Ex.: email@servidor.com.br
    * MV_RELPSW  - Senha de login do e-Mail    - Ex.: senha
    * MV_RELSERV - Servidor SMTP do e-Mail     - Ex.: smtp.servidor.com.br:587
    * MV_RELTIME - TimeOut do e-Mail           - Ex.: 120
/*/

/*/
+-------------------------------------------------------------------------------+
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Programa  ¦aXMLemail   ¦ Autor ¦MESSI PESSOA          ¦ Data ¦ 27/02/2024 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ aXMLemail ¦ Empresa ABC GERAR NF - 2 VIA CUPOM                            ¦¦¦
¦¦¦           ¦                                       			              ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
+-------------------------------------------------------------------------------+
/*/

User Function aRsEmail(_cNota,_cSerie,_cCliente,_cProduto,_cChaveNf,_cEmail,cArquivo)
	Local cHtml		:= ""
	Local cServer	:= GetMv("MV_YSERVIDO")		// IP do servidor de e-mails.
	Local cAccount	:= GetMv("MV_YCONTA")		// Conta de envio de relatórios por e-mail.
	Local cPass		:= GetMv("MV_YPASSW")	    // Senha da conta de e-mails.
	Local cEmail	:= AllTrim(_cEmail)		   // Relação de e-mails que irão receber o e-mail sobre a diferença de preço da garantia.
	Local cBCC		:= ""
	Local lResult	:= .F.
	local cmsgalert := ""
	Default cPara      := ""
	Default cAssunto   := ""
	Default cCorpo     := ""
	Default aAnexos    := {}
	Default lMostraLog := .F.
	Default lUsaTLS    := .T.

////novo

	cHtml := "<html>"
	cHtml += "<head>"
	cHtml += "<title> EMPRESA ABC 2 VIA CUPOM </title>"
	cHtml += "</head>"
	cHtml += "<body>"
	cHtml += "<br>"
	cHtml += "Olá: " + ALLTrim(_cCliente) + "<br><br>"
	cHtml += "<br>"
	cHtml += "Produto que você comprou na Empresa ABC : "+_cProduto+"<br><br>"
	cHtml += "abaixo está o link para consultar no site da Sefaz seu cupom fiscal <br>"
	cHtml += "<hr>"
	cHtml += "<p>Link Sefaz: <a href='https://sistemas.sefaz.XXXX.gov.br/nfceweb/consultarNFCe.jsp?p="+Alltrim(_cChaveNf)+"' target='_blank'>CUPOM FISCAL</a></p> <br><br>"
	cHtml += "<br><br>"
	cHtml += "    <center>Empresa ABC SUA MENSAGEM </center>"
	cHtml += "<br>"
	cHtml += "</body>"
	cHtml += "</html>"

	//+---------------------------------------------------------------------+
	//| Conecta no servidor de e-mail                                       |
	//+---------------------------------------------------------------------+
	CONNECT SMTP SERVER cServer ACCOUNT cAccount PASSWORD cPass RESULT lResult
	MailAuth(cAccount,cPass)

	//+---------------------------------------------------------------------+
	//| Se a conexão com o SMPT foi estabelecida, envia o e-mail            |
	//+---------------------------------------------------------------------+
	If lResult
		If !Empty(cArquivo) //envia Danfe
			SEND MAIL FROM cAccount TO cEmail BCC cBCC SUBJECT "Empresa ABC - 2VIA NOTA FISCAL" BODY cHtml ATTACHMENT cArquivo RESULT lResult
		Else                //envia cupom
			SEND MAIL FROM cAccount TO cEmail BCC cBCC SUBJECT "Empresa ABC - 2VIA NOTA FISCAL" BODY cHtml RESULT lResult
			cmsgalert := "         "
			cmsgalert += CRLF
			cmsgalert += CRLF
			cmsgalert += CRLF
			cmsgalert += "     EMAIL ENVIADO COM SUCESSO...    "  //JRE-8u401
			cmsgalert += CRLF
			cmsgalert += CRLF
			cmsgalert += "    "
			Aviso( "   NFCE - CUPOM FISCAL   ", cmsgalert, {"OK"}, 3 )
		EndIf

		DISCONNECT SMTP SERVER

	Else
		GET MAIL ERROR CERROR
	Endif


Return
