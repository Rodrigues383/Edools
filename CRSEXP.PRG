********************************************************************************
*  NOME PROGRAMA : CRSEX100                                                    *
*  AUTOR         : (MARCOS RODRIGUES)VS SISTEMAS DE INFORMAÇÃO LTDA            *
*  FUNÇÃO        : EXPROTAR DADOS DO CONTAS A RECEBER                          *
********************************************************************************

#pragma Oem2Ansi(OFF)

#include "Gra.ch"
#include "Xbp.ch"
#include "Common.ch"
#include "Appevent.ch"
#include "Font.ch"
#include "Class.ch"
#include "Dll.ch" 
#include "Bap.ch"
#include "Nls.ch"
#include "Directry.ch"
#include "Fileio.ch"
#include "odbcdbe.ch"
#include "dbstruct.ch"

#include "S:\SOURCE\comum\recursos\wvssc_res.ch"
#include "S:\SOURCE\_main\vss.ch"

******************************************************************************

procedure AppSys()

	set EXCLUSIVE OFF
	set INTENSITY ON
	set SCOREBOARD OFF
	set CONFIRM ON
	set DATE BRITISH
	set CENTURY ON
	set DELETED ON
	set WRAP ON
	set EXACT OFF
	set EPOCH to 1970
	set NULLVALUE OFF

	set(_SET_DECIMALS, 10)
	
	//set OPTIMIZE OFF
	//set RUSHMORE OFF
	//set SMARTFILTER OFF
	
	SET COLLATION to ASCII // para garantir a interoperabilidade com o CLIPPER
	SET CHARSET TO ansi

return

******************************************************************************

procedure dbeSys()
return

******************************************************************************

//================================================================================================================================

procedure Main()

public CRLF := chr(13)+chr(10)
PUBLIC __SystemCollation := set(_SET_COLLATION) // ver AppSys em WVSSC.PRG
PUBLIC __SystemCharSet := set(_SET_CHARSET) // ver AppSys em WVSSC.PRG

private fHndExpLog := ""
	CriaLogExpCrs()
	EscreverNoCrsExpLog(CabLog("LOG DA EXPORTAÇÃO DE TÍTULOS DO CONTAS A RECEBER", .f.), .f.)

private cTipExp := "", cArqDSN := "", cTabCRS := "", cTabCRI := "", cTabEMP := "", cDelimit := ""
private dbDestino
private cTipoBDDestino := "", cDelimNomeTabelaIni := "", cDelimNomeTabelaFim := "", cDelimNomeCampoIni := "", cDelimNomeCampoFim := "", cSepCmd := ";"
private aEstruTabExpCRS := {}, aEstruTabExpCRI := {}, aEstruTabExpEMP := {}
private XDirWrk := "ODBC(XDirWrk)"
private cArqDSN := PegaPathDoArqDSN()
private aContasConsumo := {}
private aDefArqExp := {}

private cTIPPES  := ""
private cNOMLOJ  := ""
private cENDLOJ  := ""
private cCIDEST  := ""
private cBAILOJ  := ""
private cCEPLOJ  := ""
private cINSEST  := ""
private cIDEPES  := ""
private cTELFAX  := ""
private cESTCIV  := ""
private cPROFIS  := ""

private cENDER   := ""
private cNUMEND  := ""
private cCOMEND  := ""
private cCIDADE  := ""
private cESTADO  := ""
private cDDD     := ""
private cTEL     := ""
private cRG      := ""
private cORGEXP  := ""
private cEMAIL   := ""

private cCodUF   := ""
private cCodIBGE := ""

private cCODCTAREC := ""
private cDSCCTAREC := ""
private cCODBCO    := ""
private cCODAGC    := ""
private cDVCODAGC  := ""
private cNOMAGC    := ""
private cCODCTA    := ""
private cDVCODCTA  := ""
private cCODCNV    := ""
private cCODCAR    := ""
private cNOSNUM    := ""
private cDVNOSNUM  := ""
private cRAZSOCTIT := ""
private cCNPJTIT   := ""
private cENDTIT    := ""
private cLOGOBCO   := ""
private cCMPLIV    := ""

private aContasCorrentes := {}
private aContasRec := {{"", "", "", "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, ""}}
private aTotRec    :=  {"", "", "", "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, ""}
private aLayoutCRS := {}
private aLayoutCRI := {}
private aLayoutEMP := {}

private fHndCRI := 0
private fHndCRS := 0
private fHndEMP := 0

private cCodMov  := ""
private cDataIni := ""
private cDataFin := ""
private aEmpExp  := {}
private cDlmtTexto := ""
private aNomeExp := {}
private aTipoNomeHnd := {}
private aTmpIdeLoj := {}
private cStrDoLog
private dbVSSC
private DataExp
private Sequencial
private cSERVER := "", cDATABASE := ""

	aInfDoExecutavel := Directory(OnlyPath(AppName(.t.))+"\EXPCRS.EXE")
	cNumVersao := "1.0"
	cDataCompilacao := DtoC(aInfDoExecutavel[1,3]) + " - " + aInfDoExecutavel[1,4]
	cDesVersao := "[versão " + cNumVersao + " - " + cDataCompilacao + "]"
	EscreverNoCrsExpLog(cDesVersao)

	aConfExp := GetConfExpCrs()
	I_COMMON()
	LoadDBEs()

	EscreverNoCrsExpLog("Conectando ao banco de dados de destino...")
	if !ConectarBD(XDirWrk, cArqDSN)
		return
	endif

	ApagarLogsExportacao()
	dbVSSC := getODBCObj(XDirWrk)

	cPathTXT   := AllTrim(GetParm("VSS_PRM", "Integração com outros sistemas - Local_TXT", ""))
	cDlmtTexto := AllTrim(GetParm("VSS_PRM", "Integração com outros sistemas - DelimitadorDeTexto", ""))
	DataExp    := AllTrim(GetParm("VSS_PRM", "CRSPE100_DataExp", ""))
	DataExp    := StrTran(DataExp, "/", "")
	Sequencial := AllTrim(GetParm("VSS_PRM", "CRSPE100_CodSave", ""))

	if DataExp < DtoS(date())
		SetParm("VSS_PRM", "CRSPE100_CodSave", StrZero(0,2))
	endif

	// aTotRec[ 1] = aContasRec[n,  1] = Código do Grupo
	// aTotRec[ 2] = aContasRec[n,  2] = Descrição do Grupo
	// aTotRec[ 3] = aContasRec[n,  3] = Código da Conta
	// aTotRec[ 4] = aContasRec[n,  4] = Descrição da Conta
	// aTotRec[ 5] = aContasRec[n,  5] = Principal Devido (valor)
	// aTotRec[ 6] = aContasRec[n,  6] = Principal Devido (texto)
	// aTotRec[ 7] = aContasRec[n,  7] = Total Devido (valor)
	// aTotRec[ 8] = aContasRec[n,  8] = Total Devido (texto)
	// aTotRec[ 9] = aContasRec[n,  9] = Principal Corrigido (valor)
	// aTotRec[10] = aContasRec[n, 10] = Principal Corrigido (texto)
	// aTotRec[11] = aContasRec[n, 11] = Correção Monetária Devida (valor)
	// aTotRec[12] = aContasRec[n, 12] = Correção Monetária Devida (texto)
	// aTotRec[13] = aContasRec[n, 13] = Juros Devidos (valor)
	// aTotRec[14] = aContasRec[n, 14] = Juros Devidos (texto)
	// aTotRec[15] = aContasRec[n, 15] = Multa Devida (valor)
	// aTotRec[16] = aContasRec[n, 16] = Multa Devida (texto)
	// aTotRec[17] = aContasRec[n, 17] = Total Pago (valor)
	// aTotRec[18] = aContasRec[n, 18] = Total Pago (texto)
	// aTotRec[19] = aContasRec[n, 19] = Principal Pago (valor)
	// aTotRec[20] = aContasRec[n, 20] = Principal Pago (texto)
	// aTotRec[21] = aContasRec[n, 21] = Correção Monetária Paga (valor)
	// aTotRec[22] = aContasRec[n, 22] = Correção Monetária Paga (texto)
	// aTotRec[23] = aContasRec[n, 23] = Juros Pagos (valor)
	// aTotRec[24] = aContasRec[n, 24] = Juros Pagos (texto)
	// aTotRec[25] = aContasRec[n, 25] = Multa Paga (valor)
	// aTotRec[26] = aContasRec[n, 26] = Multa Paga (texto)

	//--------------------------------------------------------------------------
	// Definição do lay-out do arquivo de exportação.
	//--------------------------------------------------------------------------

	EscreverNoCrsExpLog(CabLog("Carregando lay-out do arquivo de exportação..."), .f.)
	CarregarLayoutArqExp()
	ObterLayoutExp()
	
	if empty(aDefArqExp)
		cStrDoLog := ""
		cStrDoLog := "Não foi possível carregar as definições do lay-out do arquivo de exportação.;;" + ;
		        "A tabela CRS_DAE deve estar vazia.;;" + ;
		        "Ela é carregada automaticamente ao entrar no contas a receber, a partir da importação do arquivo " + ;
		        "'TabVS_CRS_CRB_DAE.txt', que deve estar localizado na pasta definida como padrão para " + ;
		        "importação de arquivos (geralmente a pasta IMPORTA, localizada abaixo da pasta onde se " + ;
		        "encontra o executável WVSSC).;"
		EscreverNoCrsExpLog(cStrDoLog)
		return
	endif
	//--------------------------------------------------------------------------
	EscreverNoCrsExpLog(CabLog("Carregando dados de contas de consumo..."), .f.)
	GetContasConsumo()

	EscreverNoCrsExpLog(CabLog("Carregando dados de contas correntes..."), .f.)
	CarregarContasCorrentes()

	EscreverNoCrsExpLog(CabLog("Recuperando parâmetros de exportação..."), .f.)
	RecuperarParamExportacao()

	EscreverNoCrsExpLog(CabLog("Recuperando dados da última exportação..."), .f.)
	RecuperarDadosUltExportacao()

	if Left(cTipExp,3) <> "TXT"

		EscreverNoCrsExpLog("Obtendo estrutura das tabelas de destino...")
		if !ObterEstruTabsDestino()
			return
		endif

		EscreverNoCrsExpLog("Exportar para " + cDATABASE + " em " + cSERVER + ".")

	endif

	Exportar()

	ODBC_Disconnect("ODBC(XDirWrk)")

	EscreverNoCrsExpLog(CabLog("Exportação concluida!"), .f.)

	unlock all
	DbCommitAll()
	dbCloseAll()
	
	Quit

return

//================================================================================================================================

function CriaLogExpCrs()
local pFile, fHnd

	pFile := OnlyPath(AppName(.t.))

	if file(pFile)
		if fErase(pFile) <> 0
			return .f.
		endif
	endif

	fHnd := fcreate(pFile+"\"+"LOGEXP"+dtos(date())+strtran(time(),":","")+".TXT")

	if fHnd = -1
		return .f.
	endif

	fHndExpLog := fHnd

return .t.

//================================================================================================================================

function CabLog(cString, lPadraso)
local cRetorno

	default lPadraso to .t.

	if lPadraso
		cRetorno := "["+DtoC(date())+" - "+time()+"]  "+replicate("=", 115) + crlf
		cRetorno += "["+DtoC(date())+" - "+time()+"]  "+cString + crlf
		cRetorno += "["+DtoC(date())+" - "+time()+"]  "+replicate("=", 115) + crlf

	else

		cRetorno := replicate("=", 140) + crlf
		cRetorno += padc(cString, 140," ") + crlf
		cRetorno += replicate("=", 140) + crlf

	endif

return cRetorno

//================================================================================================================================

function EscreverNoCrsExpLog(cTxt, lPadraso)

	default lPadraso to .t.

	if cTxt = NIL
		return .f.
	endif

	if ValType(cTxt) <> "C"
		return .f.
	endif

	if fHndExpLog <= 0
		return .f.
	endif

	FSeek(fHndExpLog, 0, FS_END)

	if lPadraso
		FWrite(fHndExpLog, "["+DtoC(date())+" - "+time()+"]  "+cTxt+CRLF)
	else 
		FWrite(fHndExpLog, cTxt)
	endif

return .t.

//================================================================================================================================

procedure ApagarLogsExportacao()
local cPath := OnlyPath(AppName(.t.))

#define NDIASDELLOG 2

	aFiles := Directory(cPath+"\LOGEXP*.TXT")
	for i := 1 to len(aFiles)
		if SubStr(aFiles[i, F_NAME], 7) < DtoS(date()-NDIASDELLOG)+"999999.TXT"
			delete file (cPath+"\"+aFiles[i, F_NAME])
			EscreverNoCrsExpLog("Excluiu o Log de exportação: " + cPath+"\"+aFiles[i, F_NAME])
		endif
	next

return

//================================================================================================================================

function PegaPathDoArqDSN()
local cPath, cFileDSN, nPosIni, nPosFin, nTam, cRet	

	cPath := ONLYPATH(APPNAME(.T.))+"\"+"expcrs.ini"
	cFileDSN := LoadIni(cPath)

	if right(cFileDSN, 2) <> CRLF
		cFileDSN += CRLF
	endif

	cRet := GetIni(cFileDSN,"PATHS","DADOS")

	nIni := at("FILEDSN=", cRet)
	if nIni <> 0
		nIni += 8
	else
		return ""
	endif
	nFim := at(";", cRet, nIni)
	if nFim <> 0
		cRet := SubStr(cRet, nIni, nFim-nini)
	else
		return ""
	endif

return cRet
//================================================================================================================================

static function ObterEstruTabsDestino()
local cSql

	if !empty(cTabCRS)

		cSql := " SELECT * FROM " + cTabCRS + "; "

		vsCloseArea("TABEXPCRS")
		dbDestino:Use(cSql, "TABEXPCRS") 
		if NetErr()
			cStrDoLog := ""
			cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
			       "Rotina: "+ProcName()+";"+ ;
			       "Erro: "+var2char(dbDestino:getLastError())+" - "+dbDestino:getLastMessage()
			EscreverNoCrsExpLog(cStrDoLog)
			return .f.
		endif

		aEstruTabExpCRS := TABEXPCRS->(DbStruct())
		vsCloseArea("TABEXPCRS")

	endif

	if !empty(cTabCRI)

		cSql := " SELECT * FROM " + cTabCRI + "; "

		vsCloseArea("TABEXPCRI")
		dbDestino:Use(cSql, "TABEXPCRI")
		if NetErr()
			cStrDoLog := ""
			cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
			       "Rotina: "+ProcName()+";"+ ;
			       "Erro: "+var2char(dbDestino:getLastError())+" - "+dbDestino:getLastMessage()
			EscreverNoCrsExpLog(cStrDoLog)
			return .f.
		endif

		aEstruTabExpCRI := TABEXPCRI->(DbStruct())
		vsCloseArea("TABEXPCRI")

	endif

	if !empty(cTabEMP)

		cSql := " SELECT * FROM " + cTabEMP + "; "

		vsCloseArea("TABEXPEMP")
		dbDestino:Use(cSql, "TABEXPEMP")
		if NetErr()
			cStrDoLog := ""
			cStrDoLog :="Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
			       "Rotina: "+ProcName()+";"+ ;
			       "Erro: "+var2char(dbDestino:getLastError())+" - "+dbDestino:getLastMessage()
			EscreverNoCrsExpLog(cStrDoLog)
			return .f.
		endif

		aEstruTabExpEMP := TABEXPEMP->(DbStruct())
		vsCloseArea("TABEXPEMP")

	endif

return .t.

//================================================================================================================================

static function RecuperarParamExportacao()

	cTipExp := AllTrim(GetParm("VSS_PRM", "Integração com outros sistemas - Exportar para", "TXT"))
	cArqDSN := AllTrim(GetParm("VSS_PRM", "Integração com outros sistemas - DSN", ""))
	cTabCRS := AllTrim(GetParm("VSS_PRM", "Integração com outros sistemas - Tipo Reg. = CRS", ""))
	cTabCRI := AllTrim(GetParm("VSS_PRM", "Integração com outros sistemas - Tipo Reg. = CRI", ""))
	cTabEMP := AllTrim(GetParm("VSS_PRM", "Integração com outros sistemas - Tipo Reg. = EMP", ""))
	cDelimit := AllTrim(GetParm("VSS_PRM", "Integração com outros sistemas - Delimitador", ";")) 

	if cTipExp == "TXT" .or. cTipExp == "TXTDLMT"
		return .t.
	endif

	if empty(cArqDSN)
		cStrDoLog := ""
		cStrDoLog := "A integração está configurada para exportar os dados para tabelas de outro banco de dados," + ;
		        "porém não foi informado o arquivo DSN correspondente.;;" + ;
		        "Corrija este problema em Configurações -> Integração com outros sistemas.;"
		EscreverNoCrsExpLog(cStrDoLog)
		return .f.
	endif

	if !file(cArqDSN)
		cStrDoLog := ""
		cStrDoLog := "A integração está configurada para exportar os dados para tabelas de outro banco de dados," + ;
		        "porém o arquivo DSN informado [" + cArqDSN + "] não foi encontrado.;" + ;
		        "Corrija este problema em Configurações -> Integração com outros sistemas.;"
		EscreverNoCrsExpLog(cStrDoLog)
		return .f.
	endif

return .t.

//================================================================================================================================

static function ConectarBD(pConn, pArqDSN)
local cConnStrBDDestino


	ODBC_Disconnect(pConn)

	EscreverNoCrsExpLog(CabLog("Verificando estrutura do arquivo DSN"), .f.)

	if empty(AllTrim(pArqDSN))
		EscreverNoCrsExpLog("Arquivo DSN vazio")
		return .f.
	endif

	if !VerifEstruDSN(pArqDSN)
		return .f.
	endif

	cConnStrBDDestino := "ODBC(DBE=ODBCDBE;FILEDSN="+pArqDSN+";)"

	if ODBC_Connect(pConn, cConnStrBDDestino) = nil
		EscreverNoCrsExpLog("Não foi possível se conectar ao BD " + cDATABASE + " do servidor " + cSERVER + ".;")

		return .f.
	endif

	if NetErr()
		EscreverNoCrsExpLog("Ocorreu um erro inesperado ao se conectar ao BD " + cDATABASE + " do servidor " + cSERVER + ".;")
		return .f.
	endif

	dbDestino := getODBCObj(pConn)

	EscreverNoCrsExpLog("Conexão ao BD " + cDATABASE + " do servidor " + cSERVER + " realizada com sucesso.;")

return .t.

//================================================================================================================================

static function VerifEstruDSN(pArqDSN)
local cArqDSNIn, cArqDSNOut, nPos := 0
local cConnStrBDDestino
local nPosSvr, nPosDrv, nPosDB, cDriver
local cMsg1, cMsg2, cMsg3
local nTamSrv := 0, nTamDB := 0

	cArqDSNIn := LoadArqDSN(pArqDSN)
	cArqDSNOut := ""

	if right(cArqDSNIn, 2) <> CRLF
		cArqDSNIn += CRLF
	endif

	// Elimina as linhas de comentário.
	do while len(cArqDSNIn) > 0
		nPos := at(CRLF, cArqDSNIn)
		cLin := AllTrim(left(cArqDSNIn, nPos+1))
		if left(cLin, 1) <> "*"
			cArqDSNOut += cLin
		endif
		cArqDSNIn := SubStr(cArqDSNIn, nPos+2)
	enddo

	cMsg1 := "A integração está configurada para exportar os dados para tabelas de outro banco de dados," + ;
		     "porém o arquivo DSN informado [" + pArqDSN + "] está com problemas.;;"

	cMsg3 := "Corrija este problema em Configurações -> Integração com outros sistemas.;"

	if (nPosDrv := at("DRIVER=", upper(cArqDSNOut))) = 0
		cMsg2 := "Falta a cláusula 'DRIVER' no arquivo DSN.;;"
		EscreverNoCrsExpLog(cMsg1 + cMsg2 + cMsg3)
		return .f.
	endif
	cDriver := SubStr(cArqDSNOut, nPosDrv+7)
	cDriver := left(cDriver, at(CRLF, cDriver)-1)
	do while at(" ", cDriver) > 0
		cDriver := StrTran(cDriver, " ", "")
	enddo

	if at("ORACLE", upper(cDriver)) > 0
		cTipoBDDestino := "Oracle"
		cDelimNomeTabelaIni := '"'
		cDelimNomeTabelaFim := '"'
		cDelimNomeCampoIni := '"'
		cDelimNomeCampoFim := '"'
		cSepCmd := ";"
	elseif at("SQLSERVER", upper(cDriver)) > 0
		cTipoBDDestino := "SQLSERVER"
		cDelimNomeTabelaIni := "["
		cDelimNomeTabelaFim := "]"
		cDelimNomeCampoIni := "["
		cDelimNomeCampoFim := "]"
		cSepCmd := ";"
	elseif at("MYSQL", upper(cDriver)) > 0
		cTipoBDDestino := "MySql"
		cDelimNomeTabelaIni := "`"
		cDelimNomeTabelaFim := "`"
		cDelimNomeCampoIni := "`"
		cDelimNomeCampoFim := "`"
		cSepCmd := ";"
	else
		cTipoBDDestino := ""
		cDelimNomeTabelaIni := ""
		cDelimNomeTabelaFim := ""
		cDelimNomeCampoIni := ""
		cDelimNomeCampoFim := ""
		cSepCmd := ";"
	endif

	if at("UID=", upper(cArqDSNOut)) = 0
		cMsg2 := "Falta a cláusula 'UID' no arquivo DSN.;;"
		EscreverNoCrsExpLog(cMsg1 + cMsg2 + cMsg3)
		return .f.
	endif

	if at("PWD=", upper(cArqDSNOut)) = 0
		cMsg2 := "Falta a cláusula 'PSW' no arquivo DSN.;;"
		EscreverNoCrsExpLog(cMsg1 + cMsg2 + cMsg3)
		return .f.
	endif

	if (nPosSvr := at("SERVER=", upper(cArqDSNOut))) = 0
		cMsg2 := "Falta a cláusula 'SERVER' no arquivo DSN.;;"
		EscreverNoCrsExpLog(cMsg1 + cMsg2 + cMsg3)
		return .f.
	endif

	if (nPosDB := at("DATABASE=", upper(cArqDSNOut))) = 0
		cMsg2 := "Falta a cláusula 'DATABASE' no arquivo DSN.;;"
		EscreverNoCrsExpLog(cMsg1 + cMsg2 + cMsg3)
		return .f.
	endif

	cSERVER := AllTrim(SubStr(cArqDSNOut, nPosSvr+7))
	nTamSrv := at(CRLF, AllTrim(cSERVER)) - 1
	cSERVER := AllTrim(Left(cSERVER, nTamSrv))

	cDATABASE := AllTrim(SubStr(cArqDSNOut, nPosDB+9))
	nTamDB := at(CRLF, AllTrim(cDATABASE)) - 1
	cDATABASE := AllTrim(Left(cDATABASE, nTamDB))

	EscreverNoCrsExpLog("Arquivo DSN OK")

return .t.

//================================================================================================================================

function LoadArqDSN(pArqDSN)
local LocalBuff := "", BuffRes := "", nIniLin, nFimLin, cLinha

	if file(pArqDSN)
	
		LocalBuff := MemoRead(pArqDSN)

		// Garante que haja um CRLF no final do arquivo.

		if right(LocalBuff, 2) <> CRLF
			LocalBuff += CRLF
		endif

		// Garante que qualquer item é precedido por um CRLF.
		
		do while at(CRLF+" ", LocalBuff) > 0
			LocalBuff := StrTran(LocalBuff, CRLF+" ", CRLF)
		enddo
		
		if left(LocalBuff, 2) <> CRLF
			LocalBuff := CRLF+LocalBuff
		endif
		
		// Garante que não haverá duas linhas em branco seguidas ou itens separados por linhas em branco.
		
		do while at(CRLF+CRLF, LocalBuff) > 0
			LocalBuff := StrTran(LocalBuff, CRLF+CRLF, CRLF)
		enddo
		
		// Garante que qualquer item é imediatamente seguido de '=', sem espaços.
		
		nIniLin := 1
		do while nIniLin <= len(LocalBuff)
			
			nFimLin := at(CRLF, SubStr(LocalBuff, nIniLin))
			if nFimLin = 0
				exit
			endif
			
			cLinha := SubStr(LocalBuff, nIniLin, nFimLin+1)
			nIniLin += nFimLin+1
			
			if (posAt := at("=", cLinha)) = 0
				BuffRes += cLinha
				loop
			endif
			
			if posAt = len(cLinha)-2
				BuffRes += AllTrim(left(cLinha, posAt-1)) + "=" + CRLF
			else
				BuffRes += AllTrim(left(cLinha, posAt-1)) + "=" + AllTrim(SubStr(cLinha, posAt+1))
			endif
			
		enddo
		
	endif

return BuffRes

//================================================================================================================================

static procedure RecuperarDadosUltExportacao()
local cUsuarioExp, cDataExp, cHoraExp, cDataInicial, cDataFinal, aDtAntigaERecente, cPathArq, cNomeArq
local aArqExp, nArqNumSeq

	cUsuarioExp  := AllTrim(GetParm("VSS_PRM", "CRSPE100_UsuarioExp", ""))
	cDataExp     := AllTrim(GetParm("VSS_PRM", "CRSPE100_DataExp", ""))
	cDataExp     := StrTran(cDataExp, "/", "")
	cHoraExp     := AllTrim(GetParm("VSS_PRM", "CRSPE100_HoraExp", ""))
	cDataInicial := AllTrim(GetParm("VSS_PRM", "CRSPE100_DataInicial", ""))
	cDataFinal   := AllTrim(GetParm("VSS_PRM", "CRSPE100_DataFinal", ""))

	cPathArq := AllTrim(GetParm("VSS_PRM", "CRSPE100_Arquivo", ""))
	cPathArq := iif(cPathArq <> ".", cPathArq, "")
	cPathArq := iif(!empty(cPathArq), cPathArq, EXPORTA)
	cPathArq := iif(right(cPathArq, 1) == "\", left(cPathArq, len(cPathArq)-1), cPathArq)

	nArqNumSeq := 1
	aArqExp := Directory(cPathArq+"\"+"EXP"+DtoS(date())+"*.txt")
	if len(aArqExp) > 0
		aSort(aArqExp,,, {|a1,a2| a1[F_NAME] < a2[F_NAME]})
		nArqNumSeq := val(right(OnlyName(aArqExp[len(aArqExp), F_NAME]), 2)) + 1
	endif
	cNomeArq := cPathArq+"\"

	aDtAntigaERecente := PegarDataMaisAntigaMaisRecente()

return

//================================================================================================================================

static function PegarDataMaisAntigaMaisRecente()
local cSql, aDtAntigaERecente := {}

	cSql := " SELECT "
	cSql += "    MIN(DATA) AS ANTIGA, "
	cSql += "    MAX(DATA) AS RECENTE "
	cSql += " FROM "
	cSql += "    VSS_MCT "
	cSql += " WHERE "
	cSql += " SISTEMA = 'CRB' "
	cSql += " ; "

	vsCloseArea("QRY")
	dbVSSC:Use(cSql, "QRY")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog)
		return {CtoD(""), CtoD("")}
	endif

	aAdd(aDtAntigaERecente, DtoC(QRY->ANTIGA))
	aAdd(aDtAntigaERecente, DtoC(QRY->RECENTE))
	
	vsCloseArea("QRY")	
	
return aDtAntigaERecente


//================================================================================================================================

static procedure Exportar()
local aChaveExp
private aIdeLoj := {}

	aIdeLoj := aTmpIdeLoj

	if cTipExp == "TXT" .or. cTipExp == "TXTDLMT"
		if !CritArqDestino()
			return
		endif
	endif

	aChaveExp := ObterTitulosMovimentadosNoPeriodo()

	//--------------------------------------------
	_Exportar(aChaveExp)
	//--------------------------------------------

	fClose(fHndCRI)
	fClose(fHndCRS)
	fClose(fHndEMP)
	SalvarDadosUltExportacao()
	PostAppEvent(xbeP_Close, nil, nil) 

return

//================================================================================================================================

static function CritArqDestino()
local cPath, cName, cExt, cFile, aNome_prm, aReg, lExisteNome := .f., nCodHnd := 0, aHnd := {}

	EscreverNoCrsExpLog(CabLog("Criação do(s) arquivo(s) de exportação"), .f.)

	if empty(cPathTXT)
		EscreverNoCrsExpLog("Não foi informada a pasta do arquivo a ser gerado.;")
		return .f.
	endif
	
	if Right(AllTrim(cPathTXT),1) <> "\"
		cPathTXT += "\"
	endif
	
	cPath := OnlyPath(AllTrim(cPathTXT))
	cPath := iif(cPath <> ".", cPath, "")
	cPath := iif(!empty(cPath), cPath, EXPORTA)
	cPath := iif(right(cPath, 1) == "\", left(cPath, len(cPath)-1), cPath)

	if !empty(cPath)
		if empty(Directory(cPath + "\", "D"))
			EscreverNoCrsExpLog("A pasta especificada [" + cPath + "] não existe.;")
			return .f.
		endif
	endif

	aNome_prm := RetornaTipoNome()

	cExt  := OnlyExt(AllTrim(cPath))
	if empty(cExt)
		cExt := "txt"
	endif


	for i := 1 to len(aNome_prm)

		if Empty(aNome_prm[i,2])
			loop
		endif

		try

			lExisteNome := IncrementaSequencial(aNome_prm[i,2])

		
			cValCpo := &(aNome_prm[i,2])
			if ValType(cValCpo) == "N"
				cValCpo := AllTrim(str(cValCpo))
			elseif ValType(cValCpo) == "B"
				cValCpo := eval(cValCpo)
				if ValType(cValCpo) == "N"
					cValCpo := AllTrim(str(cValCpo))
				endif
			endif
		catch_all(e)

			xExp := aNome_prm[i,2]

			cStrDoLog := ""
			cStrDoLog := "Ocorreu um erro na interpretação da expressão:;;" + ;
			           "[" + xExp + "];;" + ;;
			           "Nconfiguração da Interação com Outros Sistemas:;;" + ;
			           "Erro: " + var2char(e:description) + ";;" + ;
			EscreverNoCrsExpLog(cStrDoLog)
			return .f.

			loop
		end try
	
		
		nPos := Rat("=",aNome_prm[i,1])
		cTipo := substr(aNome_prm[i,1],nPos+1)
		
		aReg := {}
		aAdd(aReg, AllTrim(cTipo)) 
		aAdd(aReg, AllTrim(cValCpo))	
			
		if !lExisteNome 
			nCodHnd := CriaArquivo(cPath, iif(!empty(cValCpo), cValCpo ,""), iif(!empty(cExt), "."+cExt, ""))
			aHndAux := {}
			aAdd(aHndAux, AllTrim(Var2Char(cValCpo)))
			aAdd(aHndAux, nCodHnd)
			aAdd(aReg, nCodHnd)
			aAdd(aHnd, aHndAux) 
		else
			nAux := aScan(aHnd,cValCpo)
			aAdd(aReg, aHnd[nAux,2])
		endif
		
		aAdd(aTipoNomeHnd, aReg)
		
		
	next	

return .t.


//================================================================================================================================

function IncrementaSequencial(pNome)
local nSeq := 0, nTmpSeq := 0, lExiste := .t.

	n := aScan(aNomeExp,  pNome)
	if n = 0
		aAdd(aNomeExp, pNome)
		lExiste := .f.
	endif

	if lExiste = .f.
		if at("SEQUENCIAL", Upper(pNome)) <> 0
			nTmpSeq := AllTrim(GetParm("VSS_PRM", "CRSPE100_CodSave", ""))
			nSeq := Val(nTmpSeq)
			nSeq ++
			SetParm("VSS_PRM", "CRSPE100_CodSave", StrZero(nSeq,2))
			Sequencial := AllTrim(GetParm("VSS_PRM", "CRSPE100_CodSave", "00"))	
		endif
		return .f.
	else 
		Sequencial := AllTrim(GetParm("VSS_PRM", "CRSPE100_CodSave", "00"))	
	endif
	
return .t.

//================================================================================================================================

function RetornaTipoNome()
local cSql, aReg := {}, aTipoNome := {}

	cSql := " SELECT "
	cSql += "    NOME, "
	cSql += "    VALOR "
	cSql += " FROM "
	cSql += "    VSS_PRM "
	cSql += " WHERE "
	cSql += "       NOME LIKE 'Integração com outros sistemas - Tipo TXT. = %' "
	cSql += " ORDER BY "
	cSql += "    NOME; "

	vsCloseArea("QRY")
	dbVSSC:Use(cSql, "QRY")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog)
		break
	endif

	do while !QRY->(EOF())
		aReg := {}
		aAdd(aReg, AllTrim(QRY->NOME)) 
		aAdd(aReg, AllTrim(QRY->VALOR)) 
		aAdd(aTipoNome, aReg) 
		QRY->(dbSkip())
	enddo

return aTipoNome

//================================================================================================================================

function CriaArquivo(cPath, cNome, cExt)
local cFile := cPath + "\" + cNome + cExt
local nArqNumSeq
local aArqExp := {}
local nTamNome := 0
local cNewName

	if file(cFile)
		nArqNumSeq := 1
		nTamNome := Len(cNome)
		aArqExp := Directory(cPath + "\" + cNome + "*" + cExt)
		
		if len(aArqExp) > 0
			aSort(aArqExp,,, {|a1,a2| a1[F_NAME] < a2[F_NAME]})
			nArqNumSeq := val(right(onlyname(aArqExp[len(aArqExp), F_NAME]), 4)) + 1
		endif

		cNewName := cPath + "\" + cNome + StrZero(nArqNumSeq, 4) + cExt
		EscreverNoCrsExpLog("O arquivo [" + cFile + "] já existe e e sera renomeado para [" + cNewName + "].")

		if FRename(cFile, cNewName) == -1
			EscreverNoCrsExpLog("Não foi possível renomear o arquivo [" + cFile + "] para [" + cNewName + "].")
			return .f.
		endif

	endif

	fHnd := fCreate(cFile)
	if fHnd = -1
		EscreverNoCrsExpLog("Ocorreu um erro ao criar o arquivo [" + cFile + "].;")
		return .f.
	endif

return fHnd

//================================================================================================================================

static procedure SalvarDadosUltExportacao()
local cDataExp, cHoraExp, cDataInicial, cDataFinal, cUsuarioExp, cArquivo, nCodSave := 0, cPath

	cUsuarioExp :=__VSSLoged[1] + "-" + __VSSLoged[3]
	cDataExp := DtoS(date())
	cDataExp := StrTran(cDataExp, "/", "")
	cHoraExp := time()

	cPath := OnlyPath(AllTrim(cPathTXT))
	cPath := iif(cPath <> ".", cPath, "")
	cPath := iif(!empty(cPath), cPath, EXPORTA)
	cPath := iif(right(cPath, 1) == "\", left(cPath, len(cPath)-1), cPath)

	SetParm("VSS_PRM", "CRSPE100_UsuarioExp", cUsuarioExp)
	SetParm("VSS_PRM", "CRSPE100_DataExp", cDataExp)
	SetParm("VSS_PRM", "CRSPE100_HoraExp", cHoraExp)
	SetParm("VSS_PRM", "CRSPE100_Arquivo", cPath)

return

//================================================================================================================================

static function ObterTitulosMovimentadosNoPeriodo()
local cSql, aTit := {}, nCont := 0
local cCodMvt := ""

	EscreverNoCrsExpLog(CabLog("Obter últimos títulos movimentados"), .f.)

	cSql := " SELECT DISTINCT "
	cSql += "    MCT.CHAVEACC AS CHAVE, "
	cSql += "    MAX(MCT.CODMOV) AS CODMOV "
	cSql += " FROM "
	cSql += "    VSS_MCT MCT "
	cSql += " WHERE "
//	cSql += "        " + GET_SUBQRY(453, {}, "CONVERT(CHAR, DATA, 112)") + " BETWEEN '" + DtoS(oFrm:txtDtInicial:getData()) + "' AND '" + DtoS(oFrm:txtDtFinal:getData()) + "' "
	cSql += "     SISTEMA = 'CRB' AND "
	cSql += "     CODMOV >=   " + aConfExp[1] + " "
	cSql += "    AND EXISTS "
	cSql += "           (SELECT * "
	cSql += "            FROM "
	cSql += "               CRS_CRS CRS "
	cSql += "            WHERE "
	cSql += "                   CRS.NCODORG = SUBSTRING(MCT.CHAVEACC,  1,4) "
	cSql += "               AND CRS.ACRS    = SUBSTRING(MCT.CHAVEACC,  5,4) "
	cSql += "               AND CRS.MCRS    = SUBSTRING(MCT.CHAVEACC,  9,2) "
	cSql += "               AND CRS.NRECCRS = SUBSTRING(MCT.CHAVEACC, 11,5) "
	cSql += "               AND CRS.NSEQCRS = SUBSTRING(MCT.CHAVEACC, 16,2) "
	cSql += "               AND CRS.NPARCRS = SUBSTRING(MCT.CHAVEACC, 18,2) "
	cSql += "           ) "
	cSql += " GROUP "
	cSql += "    BY MCT.CHAVEACC "
	cSql += " ORDER "
	cSql += "    BY CODMOV "
	cSql += " ; "

	vsCloseArea("QRY")
	dbVSSC:Use(cSql, "QRY")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog)
		break
	endif

	do while !QRY->(EOF())
		nCont ++
		EscreverNoCrsExpLog("Lendo movimento " + AllTrim(transform(nCont, '@E ###,###')) + " de " + AllTrim(transform(QRY->(RecCount()), '@E ###,###')) + "   [" + AllTrim(QRY->CHAVE) + "].")
		aAdd(aTit, AllTrim(QRY->CHAVE)) 
		cCodMvt := Var2Char(QRY->CODMOV)
		QRY->(dbSkip())
	enddo

	if !Empty(cCodMvt)
		Atualizacodmov(cCodMvt)
	else
		EscreverNoCrsExpLog("Não existem títulos a serem exportados")
	endif

	vsCloseArea("QRY")

return aClone(aTit)

//================================================================================================================================

procedure ObterLayoutExp()
local cSql, aLayout := {}

	EscreverNoCrsExpLog(CabLog("Obtendo Layot de exportação da tabela CRS_DAE"), .f.)

	cSql := " SELECT DISTINCT "
	cSql += " 	NTIPLOTDAE "
	cSql += " FROM "
	cSql += " 	CRS_DAE "
	cSql += " WHERE "
	cSql += " 	XTIPREGDAE = 'CRS' "
	cSql += " ; "
	
	vsCloseArea("QRY_CRS")
	dbVSSC:Use(cSql, "QRY_CRS")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog)
		break
	endif
	
	do while !QRY_CRS->(EOF())
		aAdd(aLayoutCRS, AllTrim(QRY_CRS->NTIPLOTDAE)) 
		QRY_CRS->(dbSkip())
	enddo
	
	cSql := " SELECT DISTINCT "
	cSql += " 	NTIPLOTDAE "
	cSql += " FROM "
	cSql += " 	CRS_DAE "
	cSql += " WHERE "
	cSql += " 	XTIPREGDAE = 'CRI' "
	cSql += " ; "
	
	vsCloseArea("QRY_CRI")
	dbVSSC:Use(cSql, "QRY_CRI")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog)
		break
	endif

	do while !QRY_CRI->(EOF())
		aAdd(aLayoutCRI, AllTrim(QRY_CRI->NTIPLOTDAE)) 
		QRY_CRI->(dbSkip())
	enddo
	
	cSql := " SELECT DISTINCT "
	cSql += " 	NTIPLOTDAE "
	cSql += " FROM "
	cSql += " 	CRS_DAE "
	cSql += " WHERE "
	cSql += " 	XTIPREGDAE = 'EMP' " 
	cSql += " ; "
	
	vsCloseArea("QRY_EMP")
	dbVSSC:Use(cSql, "QRY_EMP")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog)
		break
	endif

	do while !QRY_EMP->(EOF())
		aAdd(aLayoutEMP, AllTrim(QRY_EMP->NTIPLOTDAE)) 
		QRY_EMP->(dbSkip())
	enddo
	
return 

//================================================================================================================================

static procedure _Exportar(aChaveExp)
local cSql, cCmd := "", i := 0, j := 0, nConta := 0, nTit := 0, cRecibo
private cNCODORG, cACRS, cMCRS, cNRECCRS, cNSEQCRS, cNPARCRS
private cCodGrupo, cDscGrupo, cCodConta, cDscConta
private nTOTDev, cTOTDev, nPRIDev, cPRIDev, nCRMDev, cCRMDev, nJRSDev, cJRSDev, nMLTDev, cMLTDev
private nTOTPag, cTOTPag, nPRIPag, cPRIPag, nCRMPag, cCRMPag, nJRSPag, cJRSPag, nMLTPag, cMLTPag
private nPRICor, cPRICor
private nTotTitTOTDev, nTotTitPRIDev, nTotTitCRMDev, nTotTitJRSDev, nTotTitMLTDev
private nTotTitTOTPag, nTotTitPRIPag, nTotTitCRMPag, nTotTitJRSPag, nTotTitMLTPag
private nTotTitPRICor, cCtaConsumo, fPos

private cTabela := ""

	for i := 1 to len(aChaveExp)

		cNCODORG := SubStr(aChaveExp[i],  1, 4)
		cACRS    := SubStr(aChaveExp[i],  5, 4)
		cMCRS    := SubStr(aChaveExp[i],  9, 2)
		cNRECCRS := SubStr(aChaveExp[i], 11, 5)
		cNSEQCRS := SubStr(aChaveExp[i], 16, 2)
		cNPARCRS := SubStr(aChaveExp[i], 18, 2)

		cSql := " SELECT * "
		cSql += " FROM "
		cSql += "    CRS_CRS "
		cSql += " WHERE "
		cSql += "        NCODORG = '" + cNCODORG + "' "
		cSql += "    AND ACRS    = '" + cACRS + "' "
		cSql += "    AND MCRS    = '" + cMCRS + "' "
		cSql += "    AND NRECCRS = '" + cNRECCRS + "' "
		cSql += "    AND NSEQCRS = '" + cNSEQCRS + "' "
		cSql += "    AND NPARCRS = '" + cNPARCRS + "' "
		cSql += " ; "

		vsCloseArea("CRS_CRS")
		dbVSSC:Use(cSql, "CRS_CRS")
		if NetErr()
			cStrDoLog := ""
			cStrDoLog :="Ocorreu um erro inesperado ao efetuar uma consulta à base de dados .;;"+ ;
			       "Rotina: "+ProcName()+";"+ ;
			       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog)
			break
		endif

		if CRS_CRS->(EOF())
			EscreverNoCrsExpLog("Não achou -> " + var2char(aChaveExp[i]) + ".;")
			EscreverNoCrsExpLog("Processo interrompido...")
		endif

		StoreDataRec()

		CarregarDadosPessoa(AllTrim(NIDELOJWRK))
		CarregarValores(NCODORGWRK, ACRBWRK, MCRBWRK, NRECCRBWRK, NSEQCRBWRK, NPARCRBWRK)
		CarregarDadosContaCorrente(NCODBCOWRK, XCODAGCWRK, XCODCTAWRK)

		//----------------------------------------
		if aScan(aIdeLoj, AllTrim(NIDELOJWRK)) = 0

			// Só exporta os dados da empresa se já não foram exportados antes.
			for j := 1 to len(aLayoutEMP)
				if cTipExp == "TXT" .or. cTipExp == "TXTDLMT"

					cReg := MontarRegistro("EMP", aLayoutEMP[j], cTipExp)

					if AllTrim(cReg) == "ERRO"
						return
					endif
					if !empty(AllTrim(cReg))
						
						n := aScan(aTipoNomeHnd, {|x| x[1] == "EMP"})
						if n <> 0
							fHndEMP := aTipoNomeHnd[n, 3]
						endif

						fSeek(fHndEMP, 0, FS_END)
						fWrite(fHndEMP, cReg+CRLF)
						AtualizaEmpresas(AllTrim(NIDELOJWRK), cNOMLOJ, cTIPPES)
					endif
				else

					cCmd := MontarInsert("EMP", aLayoutEMP[j])

					if AllTrim(cCmd) == "ERRO"
						return
					endif
					if !empty(AllTrim(cCmd))
						dbDestino:setProperty(ODBCSSN_CLEAR_ERROR, .t.)
						if !dbDestino:cmd(cCmd, .t.)
							cStrDoLog := ""
							cStrDoLog := "Ocorreu um erro inesperado ao inserir na tabela " + cTabela + ".;;"+ ;
							       "Rotina: "+ProcName()+";"+ ;
							       "Erro: "+var2char(dbDestino:getLastError())+" - "+dbDestino:getLastMessage()
							EscreverNoCrsExpLog(cStrDoLog)
							return
						endif
					endif
					AtualizaEmpresas(AllTrim(NIDELOJWRK), cNOMLOJ, cTIPPES)
				endif
			next
			aAdd(aIdeLoj, AllTrim(NIDELOJWRK))

		endif

		//----------------------------------------

		for j := 1 to len(aLayoutCRS)
			if cTipExp == "TXT" .or. cTipExp == "TXTDLMT"

				cReg := MontarRegistro("CRS", aLayoutCRS[j], cTipExp)

				if AllTrim(cReg) == "ERRO"
					return
				endif
				if !empty(AllTrim(cReg))
					n := aScan(aTipoNomeHnd, {|x| x[1] == "CRS"})
					if n <> 0
						fHndCRS := aTipoNomeHnd[n, 3]
					endif					
					fSeek(fHndCRS, 0, FS_END)
					fWrite(fHndCRS, cReg+CRLF)
				endif
			else

				cCmd := MontarInsert("CRS", aLayoutCRS[j])

				if AllTrim(cCmd) == "ERRO"
					return
				endif
				if !empty(AllTrim(cCmd))
					dbDestino:setProperty(ODBCSSN_CLEAR_ERROR, .t.)
					if !dbDestino:cmd(cCmd, .t.)
						cStrDoLog := ""
						cStrDoLog := "Ocorreu um erro inesperado ao inserir na tabela " + cTabela + ".;;"+ ;
						       "Rotina: "+ProcName()+";"+ ;
						       "Erro: "+var2char(dbDestino:getLastError())+" - "+dbDestino:getLastMessage()
						EscreverNoCrsExpLog(cStrDoLog)
						return
					endif
				endif
			endif 
		next

		//----------------------------------------

		for j := 1 to len(aLayoutCRI)
			for nConta := 1 to len(aContasRec)

				cCodGrupo := aContasRec[nConta,  1] // Código do Grupo
				cDscGrupo := aContasRec[nConta,  2] // Descrição do Grupo
				cCodConta := aContasRec[nConta,  3] // Código da Conta
				cDscConta := aContasRec[nConta,  4] // Descrição da Conta
				nPRIDev   := aContasRec[nConta,  5] // Principal Devido (valor)
				cPRIDev   := aContasRec[nConta,  6] // Principal Devido (texto)
				nTOTDev   := aContasRec[nConta,  7] // Total Devido (valor)
				cTOTDev   := aContasRec[nConta,  8] // Total Devido (texto)
				nPRICor   := aContasRec[nConta,  9] // Principal Corrigido (valor)
				cPRICor   := aContasRec[nConta, 10] // Principal Corrigido (texto)
				nCRMDev   := aContasRec[nConta, 11] // Correção Monetária Devida (valor)
				cCRMDev   := aContasRec[nConta, 12] // Correção Monetária Devida (texto)
				nJRSDev   := aContasRec[nConta, 13] // Juros Devidos (valor)
				cJRSDev   := aContasRec[nConta, 14] // Juros Devidos (texto)
				nMLTDev   := aContasRec[nConta, 15] // Multa Devida (valor)
				cMLTDev   := aContasRec[nConta, 16] // Multa Devida (texto)
				nTOTPag   := aContasRec[nConta, 17] // Total Pago (valor)
				cTOTPag   := aContasRec[nConta, 18] // Total Pago (texto)
				nPRIPag   := aContasRec[nConta, 19] // Principal Pago (valor)
				cPRIPag   := aContasRec[nConta, 20] // Principal Pago (texto)
				nCRMPag   := aContasRec[nConta, 21] // Correção Monetária Paga (valor)
				cCRMPag   := aContasRec[nConta, 22] // Correção Monetária Paga (texto)
				nJRSPag   := aContasRec[nConta, 23] // Juros Pagos (valor)
				cJRSPag   := aContasRec[nConta, 24] // Juros Pagos (texto)
				nMLTPag   := aContasRec[nConta, 25] // Multa Paga (valor)
				cMLTPag   := aContasRec[nConta, 26] // Multa Paga (texto)

				//----------------------------------------

				cCtaConsumo   := RetornaConsumo(cCodConta, )

				if cTipExp == "TXT" .or. cTipExp == "TXTDLMT"

					cReg := MontarRegistro("CRI", Alayoutcri[j], cTipExp)

					if AllTrim(cReg) == "ERRO"
						return
					endif
					if !empty(AllTrim(cReg))
						n := aScan(aTipoNomeHnd, {|x| x[1] == "CRI"})
						if n <> 0
							fHndCRI := aTipoNomeHnd[n, 3]
						endif									
						fSeek(fHndCRI, 0, FS_END)
						fWrite(fHndCRI, cReg+CRLF)
					endif
				else
					cCmd := MontarInsert("CRI", aLayoutCRI[j])

					if AllTrim(cCmd) == "ERRO"
						return
					endif
					if !empty(AllTrim(cCmd))
						dbDestino:setProperty(ODBCSSN_CLEAR_ERROR, .t.)
						if !dbDestino:cmd(cCmd, .t.)
							cStrDoLog := ""
							cStrDoLog := "Ocorreu um erro inesperado ao inserir na tabela " + cTabela + ".;;"+ ;
							       "Rotina: "+ProcName()+";"+ ;
							       "Erro: "+var2char(dbDestino:getLastError())+" - "+dbDestino:getLastMessage()
							EscreverNoCrsExpLog(cStrDoLog)
							return
						endif
					endif
				endif
				//----------------------------------------

			next
		next

		nTit ++
		cRecibo := cNCODORG + " " + cMCRS + "/" + cACRS + " " + cNRECCRS + "-" + cNSEQCRS + "-" + cNPARCRS
		EscreverNoCrsExpLog("Exportou título " + AllTrim(transform(nTit, "@E ###,###")) + " de " + AllTrim(transform(len(aChaveExp), "@E ###,###")) + cRecibo + "    " + Alltrim(NLUCWRK) + " - " + AllTrim(XFANWRK))
	next
	vsCloseArea("CRS_CRS")

return

//================================================================================================================================

static function MontarRegistro(cTipReg, pLayout, cTipExp)
local cReg

	if cTipExp == 'TXT'
		cReg := RegistroTXT(cTipReg, pLayout)
	else
		cReg := RegistroDLMT(cTipReg, pLayout)
	endif

return cReg

//================================================================================================================================

static function RegistroTXT(cTipReg, pLayout)
local i, cReg, nTamReg, xExp, cValCpo := ""
local nIniCpo, nFimCpo, nTamCpo
local aDefReg := {}

	if empty(cTipReg)
		return ""
	endif

	aDefReg := ObterLayOutRegistro(cTipReg, pLayout)

	if empty(aDefReg)
		return "ERRO"
	endif

	nTamReg := val(aTail(aDefReg)[4]) - val(aDefReg [3,1]) + 1

	cReg := space(nTamReg)

	for i := 1 to len(aDefReg)

		nIniCpo := val(aDefReg[i,3])
		nFimCpo := val(aDefReg[i,4])

		nTamCpo := nFimCpo - nIniCpo + 1

		if nTamCpo < 0
			nTamCpo := 0
		endif
		xExp := AllTrim(aDefReg[i,5])

		if empty(xExp)

			cValCpo := space(nTamCpo)

		else
			if len(xExp) >= 5 .and. upper(left(xExp, 5)) == upper("dbFn{")
				cValCpo := xExp
			else
				try
					cValCpo := &(xExp)
					if ValType(cValCpo) == "C"
						cValCpo := padr(cValCpo, nTamCpo)
					elseif ValType(cValCpo) == "N"
						cValCpo := padl(StrTran(StrTran(AllTrim(str(cValCpo)), "."), ","), nTamCpo)
					elseif ValType(cValCpo) == "B"
						cValCpo := eval(cValCpo)
						if ValType(cValCpo) == "C"
							cValCpo := padr(cValCpo, nTamCpo)
						elseif ValType(cValCpo) == "N"
							cValCpo := padl(StrTran(StrTran(AllTrim(str(cValCpo)), "."), ","), nTamCpo)
						endif
					else
						cValCpo := space(nTamCpo)
					endif
				catch_all(e)
					cStrDoLog := ""
					cStrDoLog := "Ocorreu um erro na conversão da expressão:;;" + ;
					           "[" + xExp + "];;" + ;;
					           "Erro: " + var2char(e:description) + ";;" + ;
					           "Tipo = " + aDefReg[i,1] + ";" + ;
					           "Campo = " + aDefReg[i,2] + ";" + ;
					           "Posição (início) = " + aDefReg[i,3] + ";" + ;
					           "Posição (final)  = " + aDefReg[i,4] + ";"
					EscreverNoCrsExpLog(cStrDoLog)
				end try
			endif
		endif

		cReg := stuff(cReg, (nIniCpo - val(aDefReg[1,3]) + 1), nTamCpo, cValCpo)

	next
	
return cReg

//================================================================================================================================

static function RegistroDLMT(cTipReg, pLayout)
local i, cReg, nTamReg, xExp, cValCpo := ""
local nIniCpo, nFimCpo, nTamCpo
local aDefReg := {}

	if empty(cTipReg)
		return ""
	endif

	aDefReg := ObterLayOutRegistro(cTipReg, pLayout)

	if empty(aDefReg)
		return "ERRO"
	endif

	nTamReg := val(aTail(aDefReg)[4]) - val(aDefReg [3,1]) + 1

	cReg := ""

	for i := 1 to len(aDefReg)

		nIniCpo := val(aDefReg[i,3])
		nFimCpo := val(aDefReg[i,4])

		nTamCpo := nFimCpo - nIniCpo + 1

		if nTamCpo < 0
			nTamCpo := 0
		endif
		xExp := AllTrim(aDefReg[i,5])

		if empty(xExp)

			cValCpo := space(nTamCpo)

		else
			if len(xExp) >= 5 .and. upper(left(xExp, 5)) == upper("dbFn{")
				cValCpo := xExp
			else
				try
					cValCpo := &(xExp)
					if ValType(cValCpo) == "C"
						cValCpo := (cValCpo)
					elseif ValType(cValCpo) == "N"
						cValCpo :=StrTran(StrTran(AllTrim(str(cValCpo)), "."), ",")
					elseif ValType(cValCpo) == "B"
						cValCpo := eval(cValCpo)
						if ValType(cValCpo) == "C"
							cValCpo := cValCpo
						elseif ValType(cValCpo) == "N"
							cValCpo := StrTran(StrTran(AllTrim(str(cValCpo)), "."), ",")
						endif
					else
						cValCpo := space(nTamCpo)
					endif
				catch_all(e)
					cStrDoLog := ""
					cStrDoLog := "Ocorreu um erro na conversão da expressão:;;" + ;
					           "[" + xExp + "];;" + ;;
					           "Erro: " + var2char(e:description) + ";;" + ;
					           "Tipo = " + aDefReg[i,1] + ";" + ;
					           "Campo = " + aDefReg[i,2] + ";" + ;
					           "Posição (início) = " + aDefReg[i,3] + ";" + ;
					           "Posição (final)  = " + aDefReg[i,4] + ";"
					EscreverNoCrsExpLog(cStrDoLog)
				end try
			endif
		endif

		if ValType(cValCpo) == "N"
			cReg += StrTran(cValCpo, ".", ",") + cDelimit
		elseif ValType(cValCpo) == "C"
			cReg += cDlmtTexto + AllTrim(cValCpo) + cDlmtTexto + cDelimit
		else
			cReg += AllTrim(Var2Char(cValCpo)) + cDelimit
		endif
	
	next
	
	if cTipExp == "TXTDLMT"
			cReg := left(cReg, len(cReg)-1)
	endif		
	
return cReg


//==============================================================================================================================================================

static function MontarInsert(cTipReg, pLayout)
local i, n, xExp, nTipCpo, nTamCpo, cValCpo := ""
local aDefReg := {}, aEstruTabExp := {}
local cFields := "", cValues := "", cInsert := ""

	if empty(cTipReg)
		return ""
	endif

	if !(cTipReg $ "CRS,CRI,EMP")
		return ""
	endif

	if upper(AllTrim(cTipReg)) == "CRS"
		cTabela := cTabCRS
		aEstruTabExp := aClone(aEstruTabExpCRS)
	elseif upper(AllTrim(cTipReg)) == "CRI"
		cTabela := cTabCRI
		aEstruTabExp := aClone(aEstruTabExpCRI)
	elseif upper(AllTrim(cTipReg)) == "EMP"
		cTabela := cTabEMP
		aEstruTabExp := aClone(aEstruTabExpEMP)
	else
		cTabela := ""
		aEstruTabExp := {}
	endif

	if empty(cTabela)
		return ""
	endif

	aDefReg := ObterLayOutRegistro(cTipReg, pLayout)

	if empty(aDefReg)
		return ""
	endif

	for i := 1 to len(aDefReg)

		if (n := aScan(aEstruTabExp, {|x| upper(AllTrim(x[DBS_NAME])) == upper(AllTrim(aDefReg[i,2]))})) = 0
			// O campo da definição do registro não faz parte da tabela.
			loop
		endif
		nTipCpo := aEstruTabExp[n, DBS_TYPE]
		nTamCpo := aEstruTabExp[n, DBS_LEN]

		xExp := AllTrim(aDefReg[i,5])

		if empty(xExp)

			cValCpo := ""

		else

			if len(xExp) >= 5 .and. upper(left(xExp, 5)) == upper("dbFn{")
				cValCpo := xExp
			else
				// Só interpreta a expressão se esta não for uma função de banco de dados (identificada por "dbFn{Expressão}").
				try
					cValCpo := &(xExp)
					if ValType(cValCpo) == "N"
						cValCpo := AllTrim(str(cValCpo))
					elseif ValType(cValCpo) == "B"
						cValCpo := eval(cValCpo)
						if ValType(cValCpo) == "N"
							cValCpo := AllTrim(str(cValCpo))
						endif
					endif
				catch_all(e)
				cStrDoLog := ""
					cStrDoLog :="Ocorreu um erro na conversão da expressão:;;" + ;
					           "[" + xExp + "];;" + ;;
					           "Erro: " + var2char(e:description) + ";;" + ;
					           "Tipo = " + aDefReg[i,1] + ";" + ;
					           "Campo = " + aDefReg[i,2] + ";" + ;
					           "Posição (início) = " + aDefReg[i,3] + ";" + ;
					           "Posição (final)  = " + aDefReg[i,4] + ";"
					EscreverNoCrsExpLog(cStrDoLog)
					return "ERRO"
				end try
			endif

		endif

		if len(cValCpo) >= 5 .and. upper(left(cValCpo, 5)) == upper("dbFn{")
			// xExp = dbFn{Expressão} ou xExp = dbFn{Expressão.
			// CValCpo é o valor da Expressão (funcão do banco de dados).
			cValCpo := SubStr(cValCpo, 6, iif(right(cValCpo, 1) = "}", len(cValCpo)-6, len(cValCpo)-5))
		elseif len(cValCpo) >= 5 .and. upper(SubStr(cValCpo, 2, 5)) == upper("dbFn{")
			// xExp = "dbFn{Expressão}" ou xExp = "dbFn{Expressão".
			// CValCpo é o valor da Expressão (funcão do banco de dados).
			cValCpo := SubStr(cValCpo, 7, iif(right(cValCpo, 1) = "}", len(cValCpo)-7, len(cValCpo)-6))
		else
			if nTipCpo $ "C, W, M"
				// O tipo de dado da coluna é caracter.
				cValCpo := "'" + cValCpo + "'"
			endif
		endif

		if !empty(cValCpo)
			cFields += cDelimNomeCampoIni + AllTrim(aDefReg[i,2]) + cDelimNomeCampoFim + ", "
			cValues += AllTrim(cValCpo) + ", "
		endif

	next

	if !empty(cFields)
		cFields := left(cFields, len(cFields) - 2)
		cValues := left(cValues, len(cValues) - 2)
	else
		EscreverNoCrsExpLog("O layout informado não é compatível com a estrutura da tabela ["+AllTrim(cTabela)+"].")
		return "ERRO"
	endif

	cInsert := " INSERT INTO " + cDelimNomeTabelaIni + cTabela + cDelimNomeTabelaFim
	cInsert += " (" + cFields + ") VALUES (" + cValues + "); "



return cInsert



//==============================================================================================================================================================

static function ObterLayOutRegistro(cTipReg, pLayout)
local i, aDefReg := {}

	EscreverNoCrsExpLog(CabLog("Obtendo lay-out dos registros de exportação..."), .f.)


	for i := 1 to len(aDefArqExp)
		if upper(AllTrim(aDefArqExp[i, 1])) == upper(AllTrim(cTipReg)) .and. upper(AllTrim(aDefArqExp[i, 7])) == upper(AllTrim(pLayout))
			if cTipExp == "TXT"
				if empty(aDefArqExp[i, 3])
					cStrDoLog := ""
					cStrDoLog := "Não foi informada a posição de início do campo abaixo,;" + ;
					       "na definição do lay-out do tipo de registro [" + cTipReg + "].;;" + ;
					       "Campo = " + aDefArqExp[i, 2] + ";" + ;
					       "Início = " + aDefArqExp[i, 3] + ";" + ;
					       "Fim = " + aDefArqExp[i, 4] + ";" + ;
					       "Expressão = " + left(aDefArqExp[i, 5], 100)
					EscreverNoCrsExpLog(cStrDoLog)
					return {}
				elseif empty(aDefArqExp[i, 4])
					cStrDoLog := ""
					cStrDoLog := "Não foi informada a posição de fim do campo abaixo,;" + ;
					       "na definição do lay-out do tipo de registro [" + cTipReg + "].;;" + ;
					       "Campo = " + aDefArqExp[i, 2] + ";" + ;
					       "Início = " + aDefArqExp[i, 3] + ";" + ;
					       "Fim = " + aDefArqExp[i, 4] + ";" + ;
					       "Expressão = " + left(aDefArqExp[i, 5], 100)
					EscreverNoCrsExpLog(cStrDoLog)
					return {}
				elseif empty(aDefArqExp[i, 5])
					cStrDoLog := ""
					cStrDoLog := "Não foi informada a expressão para o campo abaixo,;" + ;
					       "na definição do lay-out do tipo de registro [" + cTipReg + "].;;" + ;
					       "Campo = " + aDefArqExp[i, 2] + ";" + ;
					       "Início = " + aDefArqExp[i, 3] + ";" + ;
					       "Fim = " + aDefArqExp[i, 4] + ";" + ;
					       "Expressão = " + left(aDefArqExp[i, 5], 100)
					EscreverNoCrsExpLog(cStrDoLog)
					return {}
				elseif !TestarParenteses(aDefArqExp[i, 5])[1]
					cStrDoLog := ""
					cStrDoLog := "Existe inconsistência na abertura e fechamento de parênteses na expressão do campo abaixo,;" + ;
					       "na definição do lay-out do tipo de registro [" + cTipReg + "].;;" + ;
					       "Campo = " + aDefArqExp[i, 2] + ";" + ;
					       "Início = " + aDefArqExp[i, 3] + ";" + ;
					       "Fim = " + aDefArqExp[i, 4] + ";" + ;
					       "Expressão = " + left(aDefArqExp[i, 5], 100)
					EscreverNoCrsExpLog(cStrDoLog)
					return {}
				endif
			else
				if empty(aDefArqExp[i, 2])
					cStrDoLog := ""
					cStrDoLog := "Não foi informado o nome do campo para o campo abaixo,;" + ;
					       "na definição do lay-out do tipo de registro [" + cTipReg + "].;;" + ;
					       "Campo = " + aDefArqExp[i, 2] + ";" + ;
					       "Início = " + aDefArqExp[i, 3] + ";" + ;
					       "Fim = " + aDefArqExp[i, 4] + ";" + ;
					       "Expressão = " + left(aDefArqExp[i, 5], 100)
					EscreverNoCrsExpLog(cStrDoLog)
					return {}
				elseif empty(aDefArqExp[i, 5])
					cStrDoLog := ""
					cStrDoLog := "Não foi informada a expressão para o campo abaixo,;" + ;
					       "na definição do lay-out do tipo de registro [" + cTipReg + "].;;" + ;
					       "Campo = " + aDefArqExp[i, 2] + ";" + ;
					       "Início = " + aDefArqExp[i, 3] + ";" + ;
					       "Fim = " + aDefArqExp[i, 4] + ";" + ;
					       "Expressão = " + left(aDefArqExp[i, 5], 100)
					EscreverNoCrsExpLog(cStrDoLog)
					return {}
				elseif !TestarParenteses(aDefArqExp[i, 5])[1]
					cStrDoLog := ""
					cStrDoLog := "Existe inconsistência na abertura e fechamento de parênteses na expressão do campo abaixo,;" + ;
					       "na definição do lay-out do tipo de registro [" + cTipReg + "].;;" + ;
					       "Campo = " + aDefArqExp[i, 2] + ";" + ;
					       "Início = " + aDefArqExp[i, 3] + ";" + ;
					       "Fim = " + aDefArqExp[i, 4] + ";" + ;
					       "Expressão = " + left(aDefArqExp[i, 5], 100)
					EscreverNoCrsExpLog(cStrDoLog)
					return {}
				endif
			endif
			aAdd(aDefReg, aDefArqExp[i])
		endif
	next
	if cTipExp == "TXT"
		aSort(aDefReg,,, {|x,y| x[3] < y[3]})
	endif

return aClone(aDefReg)




//==============================================================================================================================================================

static procedure CarregarDadosPessoa(pNIDELOJ)
local cSql
local cNIDELOJ

	cTIPPES  := ""
	cNOMLOJ  := ""
	cENDLOJ  := ""
	cCIDEST  := ""
	cBAILOJ  := ""
	cCEPLOJ  := ""
	cINSEST  := ""
	cIDEPES  := ""
	cTELFAX  := ""
	cESTCIV  := ""
	cPROFIS  := ""

	cENDER   := ""
	cNUMEND  := ""
	cCOMEND  := ""
	cCIDADE  := ""
	cESTADO  := ""
	cDDD     := ""
	cTEL     := ""
	cRG      := ""
	cORGEXP  := ""
	cEMAIL   := ""

	cCodUF   := ""
	cCodIBGE := ""

	if empty(NIDELOJWRK)
		return
	endif

	cNIDELOJ := RemoveEdicaoCNPJouCPF(AllTrim(pNIDELOJ))

	if len(cNIDELOJ) >= 11 .and. len(cNIDELOJ) <= 12

		// É um CPF.

		cSql := " SELECT * "
		cSql += " FROM "
		cSql += "    CON_FIS "
		cSql += " WHERE "
//	cSql += "        " + GET_SUBQRY(454, {}, "LEFT(NCPFFIS, 11)") + " = '" + cNIDELOJ + "' "
		cSql += "         LEFT(NCPFFIS, 11)  = '" + cNIDELOJ + "' "
		if !Empty(aConfExp[2])
			cSql += " AND LEFT(NCPFFIS, 11) NOT IN ('" + aConfExp[2] + "') "
		endif
		cSql += " ; "

		vsCloseArea("QRY")
		dbVSSC:Use(cSql, "QRY")
		if NetErr()
			cStrDoLog := ""
			cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados .;;"+ ;
			       "Rotina: "+ProcName()+";"+ ;
			       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
			EscreverNoCrsExpLog(cStrDoLog)
			break
		endif

		if !Empty(QRY->EEMIIDEFIS)
			if ValType(QRY->EEMIIDEFIS) == "D"
				cEEMIIDEFIS := DtoS(QRY->EEMIIDEFIS)
			else
				cEEMIIDEFIS := StrTran(left(QRY->EEMIIDEFIS,10),"-","")
			endif
		else
			cEEMIIDEFIS := ""
		endif
		cTIPPES := "F"
		cNOMLOJ := left(QRY->XNOMPESFIS, 60)
		cENDLOJ := AllTrim(left(QRY->XENDPESFIS, 50))+space(02)+left(QRY->NNUMENDFIS, 4)+space(01)+AllTrim(left(QRY->XCOMENDFIS, 7))
		cCIDEST := AllTrim(left(QRY->XCIDPESFIS, 20))+space(02)+AllTrim(QRY->XESTPESFIS)
		cBAILOJ := left(QRY->XBAIPESFIS, 20)
		cCEPLOJ := AllTrim(QRY->NCEPPESFIS)+AllTrim(QRY->NCODCEP)
		cIDEPES := QRY->NIDEPESFIS+" ("+QRY->XORGEXPFIS+" - "+QRY->XESTIDEFIS+") "+cEEMIIDEFIS
		cTELFAX := iif(empty(QRY->NDDDPESFIS), "", "("+AllTrim(QRY->NDDDPESFIS)+") ")+AllTrim(QRY->NTELPESFIS)+iif(empty(QRY->NRAMPESFIS), "", " RAMAL "+AllTrim(QRY->NRAMPESFIS))
		cESTCIV := QRY->XESTCIVFIS
		cPROFIS := QRY->XPROPESFIS

		cENDER  := AllTrim(left(QRY->XENDPESFIS, 50))
		cNUMEND := AllTrim(left(QRY->NNUMENDFIS, 4))
		cCOMEND := AllTrim(left(QRY->XCOMENDFIS, 7))
		cCIDADE := AllTrim(left(QRY->XCIDPESFIS, 20))
		cESTADO := AllTrim(QRY->XESTPESFIS)
		cDDD    := AllTrim(QRY->NDDDPESFIS)
		cTEL    := AllTrim(QRY->NTELPESFIS)
		cRG     := AllTrim(QRY->NIDEPESFIS)
		cORGEXP := AllTrim(QRY->XORGEXPFIS)+" - "+AllTrim(QRY->XESTIDEFIS)+" - "+cEEMIIDEFIS
		cEMAIL  := AllTrim(QRY->XEMLPESFIS)

	elseif len(cNIDELOJ) = 14

		// É um CNPJ.

		cSql := " SELECT * "
		cSql += " FROM "
		cSql += "    CON_JUR "
		cSql += " WHERE "
		cSql += "        LEFT(NCGCJUR, 14) = '" + cNIDELOJ + "' "
		if !Empty(aConfExp[2])
			cSql += " AND LEFT(NCGCJUR, 14) NOT IN ('" + aConfExp[2] + "') "
		endif
		cSql += " ; "

		vsCloseArea("QRY")
		dbVSSC:Use(cSql, "QRY")
		if NetErr() 
			cStrDoLog := ""
			cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
			       "Rotina: "+ProcName()+";"+ ;
			       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
			EscreverNoCrsExpLog(cStrDoLog)
			break
		endif

		cTIPPES := "J"
		cNOMLOJ := left(QRY->XRAZSOCJUR, 60)
		cENDLOJ := AllTrim(left(QRY->XENDPESJUR, 50))+space(02)+left(QRY->NNUMENDJUR, 4)+space(01)+AllTrim(left(QRY->XCOMENDJUR, 7))
		cCIDEST := AllTrim(left(QRY->XCIDPESJUR, 20))+space(02)+AllTrim(QRY->XESTPESJUR)
		cBAILOJ := left(QRY->XBAIPESJUR, 20)
		cCEPLOJ := AllTrim(QRY->NCEPPESJUR)+AllTrim(QRY->NCODCEP)
		cINSEST := QRY->NINSESTJUR
		cTELFAX := iif(empty(QRY->NDDDPESJUR), "", "("+AllTrim(QRY->NDDDPESJUR)+") ")+AllTrim(QRY->NTELPESJUR)+iif(empty(QRY->NFAXPESJUR), "", " FAX "+AllTrim(QRY->NFAXPESJUR))
		cESTCIV := ""
		cPROFIS := ""

		cENDER  := AllTrim(left(QRY->XENDPESJUR, 50))
		cNUMEND := AllTrim(left(QRY->NNUMENDJUR, 4))
		cCOMEND := AllTrim(left(QRY->XCOMENDJUR, 7))
		cCIDADE := AllTrim(left(QRY->XCIDPESJUR, 20))
		cESTADO := AllTrim(QRY->XESTPESJUR)
		cDDD    := AllTrim(QRY->NDDDPESJUR)
		cTEL    := AllTrim(QRY->NTELPESJUR)
		cRG     := AllTrim("")
		cORGEXP := AllTrim("")
		cEMAIL  := AllTrim(QRY->XENDTLXJUR)

	endif

	aRet := ObterCodIBGE(cESTADO, cCIDADE)
	cCodUF   := aRet[1]
	cCodIBGE := aRet[2]

return

//==============================================================================================================================================================

static procedure CarregarContasCorrentes()
local cSql, aAux := {}

	cSql := " SELECT * FROM REC_CTA; "

	vsCloseArea("QRY")
	dbVSSC:Use(cSql, "QRY")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog)
		return
	endif

	do while !QRY->(EOF())

		aAux := {}
		aAdd(aAux, QRY->NCODCTAREC)
		aAdd(aAux, QRY->XDSCCTAREC)
		aAdd(aAux, QRY->NCODBCO)
		aAdd(aAux, QRY->XCODAGC)
		aAdd(aAux, QRY->XDVCODAGC)
		aAdd(aAux, QRY->XNOMAGC)
		aAdd(aAux, QRY->XCODCTA)
		aAdd(aAux, QRY->XDVCODCTA)
		aAdd(aAux, QRY->XCODCNV)
		aAdd(aAux, QRY->XCODCAR)
		aAdd(aAux, QRY->XNOSNUM)
		aAdd(aAux, QRY->XDVNOSNUM)
		aAdd(aAux, QRY->XRAZSOCTIT)
		aAdd(aAux, QRY->XCNPJTIT)
		aAdd(aAux, QRY->XENDTIT)
		aAdd(aAux, QRY->XLOGOBCO)
		aAdd(aAux, QRY->XCMPLIV)

		aAdd(aContasCorrentes, aClone(aAux))

		QRY->(dbSkip())

	enddo

return

//==============================================================================================================================================================

static procedure CarregarDadosContaCorrente(pCODBCO, pCODAGC, pCODCTA)
local n

	n := aScan(aContasCorrentes, {|x| x[3]+x[4]+x[7] = pCODBCO+pCODAGC+pCODCTA})

	if n <> 0
		cCODCTAREC := aContasCorrentes[n,  1]
		cDSCCTAREC := aContasCorrentes[n,  2]
		cCODBCO    := aContasCorrentes[n,  3]
		cCODAGC    := aContasCorrentes[n,  4]
		cDVCODAGC  := aContasCorrentes[n,  5]
		cNOMAGC    := aContasCorrentes[n,  6]
		cCODCTA    := aContasCorrentes[n,  7]
		cDVCODCTA  := aContasCorrentes[n,  8]
		cCODCNV    := aContasCorrentes[n,  9]
		cCODCAR    := aContasCorrentes[n, 10]
		cNOSNUM    := aContasCorrentes[n, 11]
		cDVNOSNUM  := aContasCorrentes[n, 12]
		cRAZSOCTIT := aContasCorrentes[n, 13]
		cCNPJTIT   := aContasCorrentes[n, 14]
		cENDTIT    := aContasCorrentes[n, 15]
		cLOGOBCO   := aContasCorrentes[n, 16]
		cCMPLIV    := aContasCorrentes[n, 17]
	endif

return

//==============================================================================================================================================================

static procedure CarregarValores(pNCODORG, pACRS, pMCRS, pNRECCRS, pNSEQCRS, pNPARCRS)
local i

////	aContasRec := LoadContasRec(pNCODORG, pACRS, pMCRS, pNRECCRS, pNSEQCRS, pNPARCRS)

	aTotRec := {"", "", "", "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, "", 0, ""}
	for i := 1 to len(aContasRec)
		nTotTitPRIDev := aTotRec[ 5] += aContasRec[i,  5]   // Principal Devido
		nTotTitTOTDev := aTotRec[ 7] += aContasRec[i,  7]   // Total Devido
		nTotTitPRICor := aTotRec[ 9] += aContasRec[i,  9]   // Principal Corrigido
		nTotTitCRMDev := aTotRec[11] += aContasRec[i, 11]   // Correção Monetária Devida
		nTotTitJRSDev := aTotRec[13] += aContasRec[i, 13]   // Juros Devidos
		nTotTitMLTDev := aTotRec[15] += aContasRec[i, 15]   // Multa Devida
		nTotTitTOTPag := aTotRec[17] += aContasRec[i, 17]   // Total Pago
		nTotTitPRIPag := aTotRec[19] += aContasRec[i, 19]   // Principal Pago
		nTotTitCRMPag := aTotRec[21] += aContasRec[i, 21]   // Correção Monetária Paga
		nTotTitJRSPag := aTotRec[23] += aContasRec[i, 23]   // Juros Pagos
		nTotTitMLTPag := aTotRec[25] += aContasRec[i, 25]   // Multa Paga
	next

return

//==============================================================================================================================================================

static procedure CarregarLayoutArqExp()
local cSql, aAux := {}

	aDefArqExp := {}

	cSql := " SELECT * FROM CRS_DAE ORDER BY XTIPREGDAE, NINICPODAE; "

	vsCloseArea("QRY")
	dbVSSC:Use(cSql, "QRY")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog)
		return
	endif

	do while !QRY->(EOF())

		aAux := {}
		aAdd(aAux, QRY->XTIPREGDAE)
		aAdd(aAux, QRY->XNOMCPODAE)
		aAdd(aAux, QRY->NINICPODAE)
		aAdd(aAux, QRY->NFIMCPODAE)
		aAdd(aAux, QRY->XEXPCNVDAE)
		aAdd(aAux, QRY->XDSCCPODAE)
		aAdd(aAux, QRY->NTIPLOTDAE)

		aAdd(aDefArqExp, aClone(aAux))

		QRY->(dbSkip())

	enddo

return

//==============================================================================================================================================================

function DePara(pTipo, pValorDe)

	if empty(pTipo)
		return ""
	endif

	if empty(pValorDe)
		return ""
	endif

	n := aScan(aDePara, {|x| upper(AllTrim(x[1]))+upper(AllTrim(x[2])) == upper(AllTrim(pTipo))+upper(AllTrim(pValorDe))})

	if n = 0
		return ""
	endif

return aDePara[n, 3]

//================================================================================================================================

function ObterCodIBGE(pUF, pMunicipio)
local cSql, dbVSSC := getODBCObj(XDIRWRK)
local aRet := {"", ""}

	if empty(pUF)
		return {"", ""}
	endif

	if empty(pMunicipio)
		return {"", ""}
	endif

	cSql := " SELECT "
	cSql += "    CODUF, "
	cSql += "    CODIBGE "
	cSql += " FROM "
	cSql += "    VSS_MUN "
	cSql += " WHERE "
	cSql += "        NOMMUNICIP COLLATE SQL_Latin1_General_CP1_CI_AI = '" + RemoverAcentos(pMunicipio) + "' "
	cSql += "    AND SIGLAUF    COLLATE SQL_Latin1_General_CP1_CI_AI = '" + RemoverAcentos(pUF) + "' "
	cSql += " ; "

	vsCloseArea("QRY")
	dbVSSC:Use(cSql, "QRY")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog) 
		return {"", ""}
	endif

	if !(QRY->(EOF()) .and. QRY->(BOF()))
		aRet := {QRY->CODUF, QRY->CODIBGE}
	endif

return aClone(aRet)

//================================================================================================================================

function TestarParenteses(cExp)
local i, aStack := {}, aMatch := {}, lMatch := .t.

	for i := 1 to len(cExp)

		if SubStr(cExp, i, 1) == "("
			aAdd(aStack, {i, 0})
		elseif SubStr(cExp, i, 1) == ")"
			if len(aStack) > 0
				aTail(aStack)[2] := i
				aAdd(aMatch, aTail(aStack))
				aDel(aStack, len(aStack))
				aSize(aStack, len(aStack)-1)
			else
				// ")" sem "(" correspondente.
				aAdd(aMatch, {0, i})
				lMatch := .f.
			end
		endif

	next

	for i := 1 to len(aStack)
		// "(" sem ")" correspondente.
		aAdd(aMatch, aStack[i])
		lMatch := .t.
	next

	aSort(aMatch,,, {|x,y| x[1] < y[1]})

return {lMatch, aMatch}

//================================================================================================================================

function De_Para(pTipo, pDe)
local cSql, dbVSSC := getODBCObj(XDIRWRK)

	if empty(pTipo)
		return ""
	endif

	if empty(pDe)
		return ""
	endif

	cSql := " SELECT "
	cSql += "    PARA "
	cSql += " FROM "
	cSql += "    VSS_DE_PARA "
	cSql += " WHERE "
	cSql += "        TIPO = '" + pTipo + "' "
	cSql += "    AND DE   = '" + pDe + "' "
	cSql += " ; "
	
	vsCloseArea("QRY")
	dbVSSC:Use(cSql, "QRY")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog := "Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		EscreverNoCrsExpLog(cStrDoLog) 
		return ""
	endif

	if QRY->(EOF()) .and. QRY->(BOF())
		cStrDoLog := ""
		cStrDoLog :="Consulta à tabela VSS_DE_PARA não retornou nenhum valor.;;" + ;
		       "TIPO = [" + pTipo + "];" + ;
		       "DE = [" + pDe + "];"
		EscreverNoCrsExpLog(cStrDoLog)
		return ""
	endif

	if QRY->(RecCount()) > 1
		cStrDoLog := ""
		cStrDoLog :="Consulta à tabela VSS_DE_PARA retornou mais de um valor.;;" + ;
		       "TIPO = [" + pTipo + "];" + ;
		       "DE = [" + pDe + "];"
		EscreverNoCrsExpLog(cStrDoLog)
		return QRY->PARA
	endif

return QRY->PARA

//=====================================================================================================================

function RetornaConsumo(pConta, pImprimeConsumo)
local cSql, dbVSSC := getODBCObj(XDirWrk)
local n, cConsumo

	default pImprimeConsumo to .t.

	if empty(pConta)
		return ""
	endif

	if !pImprimeConsumo
		return ""
	endif

	n := aScan(aContasConsumo, {|x| upper(AllTrim(x[2])) == upper(AllTrim(pConta))})

	if n = 0
		return ""
	endif

	if empty(aContasConsumo[n, 1])
		return ""
	endif

	cSql := " SELECT * "
	cSql += " FROM "
	cSql += "    CTA_INC "
	cSql += " WHERE "
	cSql += "        NCON = '" + NCONWRK + "' "
	cSql += "    AND NCODGRC = '" + XGRCENCWRK + "' "
	cSql += "    AND NCONREDPLC = '" + aContasConsumo[n, 1] + "' "
	cSql += "    AND AINC = '" + APRORECWRK + "' "
	cSql += "    AND MINC = '" + MPRORECWRK + "' "
	cSql += " ; "

	vsCloseArea("QRY")
	dbVSSC:use(cSql, "QRY")
	if NetErr()
		cStrDoLog := ""
		cStrDoLog :="Ocorreu um erro inesperado ao efetuar uma consulta à base de dados.;;"+ ;
		       "Rotina: "+ProcName()+";"+ ;
		       "Erro: "+var2char(dbVSSC:getLastError())+" - "+dbVSSC:getLastMessage()
		vsCloseArea("QRY")
		EscreverNoCrsExpLog(cStrDoLog)
		break
	endif

	if QRY->(EOF()) .and. QRY->(BOF())
		return ""
	endif

	cConsumo := " - " + AllTrim(transform(QRY->VENCINC, "@E ###,###,###.##")) + " " + AllTrim(aContasConsumo[n, 3])

return cConsumo

//================================================================================================================================

function GetConfExpCrs()
local cPath := OnlyPath(AppName(.t.)) + "\CONFEXPCRS.TXT"
local cCodMov 
local cEmpresas
local aConfExp := {}

	cCodMov := GetIni(MemoRead(cPath),"MOVIMENTO","CODMOV")
	
	cEmpresas := RetornaEmpresas(MemoRead(cPath))

	do while Left(cCodMov, 2) = CRLF
		cCodMov := Right(cCodMov, Len(cCodMov) - 2)
	enddo

	do while Right(cCodMov, 2) = CRLF
		cCodMov := Left(cCodMov, Len(cCodMov) - 2)
	enddo

	aAdd(aConfExp, cCodMov)
	aAdd(aConfExp, cEmpresas)

return  aConfExp

//================================================================================================================================

procedure AtualizaCodMov(pNewCodMov)
local cString := MemoRead(OnlyPath(AppName(.t.)) + "\CONFEXPCRS.TXT")
local cTmpString
local cStringFinal

	nIniString := at("CODMOV", cString) + 6
	nFinString := at(CRLF, cString, nIniString)
	cStringFinal := Right(cString, Len(cString) - nFinString)

	if Left(cStringFinal, 1) = CHR(10)
		cStringFinal := Right(cStringFinal, len(cStringFinal) - 1)
		if Left(cStringFinal, 2) <> CRLF
			cStringFinal := CRLF + cStringFinal
		endif
	elseif Left(cStringFinal, 2) <> CRLF
		cStringFinal := CRLF + cStringFinal
	endif

	cTmpString := Left(cString, nIniString) + pNewCodMov + cStringFinal
	MemoWrit(OnlyPath(AppName(.t.)) + "\CONFEXPCRS.TXT", cTmpString)

return

//================================================================================================================================

procedure AtualizaEmpresas(cNIDELOJ, cNOMLOJ, cTIPPES)
local cString := MemoRead(OnlyPath(AppName(.t.)) + "\CONFEXPCRS.TXT")
local cTmpString, nFinalDoArquivo, cCod, cDesc

	if Right(AllTrim(cString), 2) <> CRLF
		cString += CRLF
	endif

	nFinalDoArquivo := Len(cString)

	if cTIPPES == "J"
		cCod := PadR(Transform(cNIDELOJ,"@R ##.###.###/####-##"), 20)
	else
		cCod := PadR(Transform(cNIDELOJ,"@R ###.###.###-##"), 20)
	endif

	cDesc := PadR(cNOMLOJ, 100)

	cString += cCod + cDesc + CRLF

	MemoWrit(OnlyPath(AppName(.t.)) + "\CONFEXPCRS.TXT", cString)

return

//================================================================================================================================
function RetornaEmpresas(pString)
local cString, nIniString

	nIniString := at("[EMPRESAS]", pString)
	cString := SubStr(pString, nIniString + 10)

	if Len(cString) > 0

		do while Left(cString, 2) = CRLF
			cString := Right(cString, Len(cString) - 2)
		enddo

		do while Right(cString, 2) = CRLF
			cString := Left(cString, Len(cString) - 2)
		enddo

		cString := RetornaEmpresasFormatadas(cString)
	else 
		cString := ""
	endif

return cString

//================================================================================================================================

function RetornaEmpresasFormatadas(pString)
local cIdEmpresa, clinha, nTotLinhas, cString := ""


	nTotLinhas := MlCount(pString)
	for i := 1 to nTotLinhas
		clinha := MemoLine(pString,,i,4) 
		if !Empty(clinha)
			cIdEmpresa := AllTrim(Left(clinha,20))
			cString += cIdEmpresa + "','"
			CarregaArrayDeEmpressas(cIdEmpresa)
		endif
	next

	cString := Left(cString, Len(cString) - 3)

	cString := StrTran(cString, ".", "")

	cString := StrTran(cString, "/", "")

	cString := StrTran(cString, "-", "")

return cString

//================================================================================================================================

procedure CarregaArrayDeEmpressas(pIdEmpresa)

	pIdEmpresa := StrTran(pIdEmpresa, ".", "")

	pIdEmpresa := StrTran(pIdEmpresa, "/", "")

	pIdEmpresa := StrTran(pIdEmpresa, "-", "")

	aAdd(aTmpideloj, AllTrim(pIdEmpresa))

return
 

