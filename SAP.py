import time
import sys
import win32com.client
import pandas as pd
import traceback
import glob
import os
import datetime
import xlwings as xw
pd.set_option('future.no_silent_downcasting', True)

class ProgramarSAP:
    def __init__(self):
        self.path_ec = r'\\srvfile01\DADOS_PBI\Compartilhado_BI\DPCP\2. Programação\EC'
        self.path_ec_script = r'\\srvfile01\DADOS_PBI\Compartilhado_BI\DPCP\2. Programação\EC\Script EC'
        self.file_ec = '2. Controle_Etapa.xlsx'
        self.file_ec_script = 'Controle EC Script.xlsx'
        self.file_macro = 'Histórico cabeçalho.xlsx'
        self.file_pep = 'Elementos PEP.txt'
        self.file_itens  = 'Histórico componentes.xlsx'
        #self.path_ait = r'\\srvfile01\DADOS_PBI\Compartilhado_BI\DPCP\2. Programação\EC\Script EC\EC AIT'
        self.df_etapas = pd.read_excel(self.path_ec + "\\" + self.file_ec)
        self.df_etapas_script = pd.read_excel(self.path_ec + "\\" + self.file_ec_script)
        self.excel = win32com.client.Dispatch("Excel.Application")
        
    def verificando_planilhas_abertas(self):
        for arquivo in self.excel.Workbooks:
            if any(f in arquivo.Name for f in [self.file_macro, self.file_itens, self.file_pep, self.file_ec_script]):
                print("Um dos arquivos Histórico cabeçalho, Histórico componentes, Controle EC Script, Elementos PEP está aberto.")
                print("Feche os arquivos antes de executar o script.")
                print("O programa será encerrado.")
                time.sleep(8)
                sys.exit()
    
    def ler_input(self):
        self.df_input_ec = pd.read_excel(self.path_ec_script + "\\" + self.file_macro)
        self.df_programacao_script = self.df_input_ec[(self.df_input_ec['Status Sistema'] == 'PROCESSADA') & (self.df_input_ec['Status Programação SAP'] == 'PENDENTE') & (self.df_input_ec['Status ECS'] == 'OK')]
        print("Etapas a serem programadas: ", self.df_programacao_script['EC'].to_list())
        self.df_input_itens = pd.read_excel(self.path_ec_script + "\\" + self.file_itens)
    
    def verificar_incoerencias(self):
        ec_faltante = self.df_programacao_script.merge(self.df_input_itens[['EC']], on='EC', how='left', indicator=True)
        ec_faltante = ec_faltante[ec_faltante['_merge'] == 'left_only']['EC']

        if ec_faltante.to_list():
             print("As tabelas input EC e input materiais estão incoerentes, verifique se todas as etapas pendentes na Planilha1 estão na Planilha 2")
             print("As EC faltantes na segunda aba são:", ec_faltante.to_list())
             print('\n')
             time.sleep(5)
             sys.exit()

        if not self.df_programacao_script['EC'].is_unique:
            print("Atenção, há ECs duplicadas na planilha Histórico Cabeçalho. Analise a planilha e remova as linhas duplicadas")
            print("Etapas com linhas duplicadas:")
            print(self.df_programacao_script[self.df_programacao_script['EC'].duplicated()]['EC'])
            print('\n')
            time.sleep(5)
            sys.exit()
        
        if any(self.df_input_itens[self.df_input_itens['EC'].isin(self.df_programacao_script['EC'])]['Volume'].isnull()):
            print("Dentre as ECs a serem programadas, há itens com a coluna 'Volume' em nulo. Por questões de segurança, o script será finalizado")
            time.sleep(5)
            sys.exit()

    def tratar_dados(self):
        self.df_input_ec = pd.read_excel(self.path_ec_script + "\\" + self.file_macro)
        self.df_input_itens = pd.read_excel(self.path_ec_script + "\\" + self.file_itens)
        self.df_input_itens['Quantidade'] = self.df_input_itens['Quantidade'].fillna(0).astype(int)

    #Junta a tabela Input com a tabela Script EC (esta é uma tabela com o mesmo formato da planilha Controle de Etapa, mas destinada as EC programadas pelo script. O objetivo é informar a tabela Controle de Etapa as informações do script)
    def mesclar_tabelas(self):

        self.df_programacao_rename = self.df_programacao_script.rename(columns={
            'Data Planejada': 'DATA',
            'Obra': 'OBRA',
            'Filial': 'FILIAL',
            'Ordem de Venda': 'OV',
            'Origem': 'ORIGEM',
            'Elemento PEP': 'PEP',
            'Status Programação SAP': 'STATUS',
            'ENVIADO E-MAIL': 'ENVIADO EMAIL - ITENS DEPM' 
        })

        self.df_programacao_rename = self.df_programacao_rename[['EC', 'DATA', 'OBRA', 'FILIAL', 'OV', 'ORIGEM', 'PEP', 'STATUS', 'ENVIADO EMAIL - ITENS DEPM']]
        self.df_etapas_script = self.df_etapas_script[['EC', 'DATA', 'OBRA', 'FILIAL', 'OV', 'ORIGEM', 'PEP', 'STATUS', 'ENVIADO EMAIL - ITENS DEPM']]
        self.df_etapas_script = pd.concat([self.df_etapas_script, self.df_programacao_rename], ignore_index=True)

    def conectar_sap(self):

        """Pega a aplicação COM do SAP e a primeira sessão para uso no script
           Testa se o SAP está aberto e se há um usuário logado
           
           """
        try: #Garante que o SAP está aberto, caso contrário encerra o programa
        #Pega a aplicação COM do SAP e a primeira sessão para uso no script
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            self.application = self.SapGuiAuto.GetScriptingEngine
            self.connection = self.application.Children(0)
            self.sessions = self.connection.Children
            self.session = self.sessions[0]
            self.usuario = self.session.Info.User
        except:
            print("SAP não está aberto, o programa será finalizado")
            time.sleep(2)
            sys.exit(0)
        #Testa se há um usuário logado no SAP, caso não haja, encerra o programa
        self.usuario = self.session.Info.User
        if self.usuario == '':
            print("SAP não está logado, o programa será finalizado")
            time.sleep(2)
            sys.exit(0)
    
    def va02_cn33_cj20n_md51(self):
        def va02():
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = self.df_input_ec.loc[index,'Ordem de Venda']
            self.session.findById("wnd[0]").sendVKey(0)
            try:
                self.session.findById(r'wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[5,0]').text = round(float(self.df_programacao_script.loc[index,'Custo Total'].replace('.', '').replace(',', '.')), 2) #Usado para transformar possíveis strings com mais de duas casas decimais, em valores flutuantes com no máximo duas casas decimais
            except:
                self.session.findById(r'wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[5,0]').text = round(float(self.df_programacao_script.loc[index,'Custo Total']), 2)
                pass

            self.session.findById("wnd[0]").sendVKey(0)

            #Reporta a planilha input
            self.df_programacao_script.loc[index, 'Elemento PEP'] = self.session.findById(r'wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PS_PSP_PNR[12,0]').text
            self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Elemento PEP'] = self.session.findById(r'wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PS_PSP_PNR[12,0]').text
            self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Status Programação SAP'] = "Montante na VA02 modificado"

            #Preenche a planilha Controle de Etapa
            self.df_etapas_script.loc[self.df_etapas_script['EC'] == row['EC'], 'OV'] = row['Ordem de Venda']
            self.df_etapas_script.loc[self.df_etapas_script['EC'] == row['EC'], 'PEP'] = self.session.findById(r'wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-PS_PSP_PNR[12,0]').text
            self.df_etapas_script.loc[self.df_etapas_script['EC'] == row['EC'], 'OBSERVAÇÃO'] = "Programação via script: Montante preenchido"

            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            
        def cn33():
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/ncn33"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtRCNIP01-POSID_LOW").text = self.df_programacao_script.loc[index, 'Elemento PEP']
            self.session.findById("wnd[0]/usr/ctxtRCNIP01-PROFILE").text = "Z001"
            self.session.findById("wnd[0]/usr/ctxtRCNIP01-MATNR").text = "KIT_ETAPA"
            self.session.findById("wnd[0]/usr/ctxtRCNIP01-WERKS").text = "5001"
            self.session.findById("wnd[0]/usr/ctxtRCNIP01-STLAN").text = "1"
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()

            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except:
                pass
            time.sleep(1)

            self.session.findById("wnd[0]/usr/btnSEL_ALL").press() #Botão de selecionar tudo
            self.session.findById("wnd[0]/usr/subACTIVITIES:SAPLCN10:2010/tblSAPLCN10TABCNTR_2010").getAbsoluteRow(0).selected = True
            self.session.findById("wnd[0]/tbar[1]/btn[5]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()

        def cj20n():
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/ncj20n"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton("OPEN")
            self.session.findById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PRPS_EXT").text = self.df_programacao_script.loc[index, 'Elemento PEP']
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton("EBLM")
            self.alteracao_cj20n = False

            self.num_alteracoes = 0
            j = 4
            while True:
                i = str(j).zfill(2)
                try:
                    self.session.findById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = f'0000{i}'
                    time.sleep(1)
                    try:
                        if any(kit in self.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOMD:2801/ctxtRESBD-MATNR").text for kit in self.volumes):
                            try:
                                self.session.findById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subVIEW_AREA:SAPLCOMD:2800/tabsTABSTRIP_2700/tabpMKAG/ssubSUBSCR_2700:SAPLCOMD:2701/ctxtRESBD-BDTER").text = self.df_input_ec.loc[index, 'Data Planejada'].strftime("%d.%m.%Y")
                                j += 1
                                self.num_alteracoes += 1
                                self.alteracao_cj20n = True
                            except:
                                pass 
                        else:
                            j += 1
                    except:
                        j += 1  
                except:
                    break
            time.sleep(1)
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()

        def md51():
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NMD51"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/chkRM61X-PLALL").selected = True   
            self.session.findById("wnd[0]/usr/ctxtRM61X-PSPEL").text = self.df_programacao_script.loc[index, 'Elemento PEP']
            self.session.findById("wnd[0]/usr/ctxtRM61X-BANER").text = "3"
            self.session.findById("wnd[0]/usr/ctxtRM61X-PLMOD").text = "3"
            self.session.findById("wnd[0]/usr/ctxtRM61X-LIFKZ").text = "1"
            self.session.findById("wnd[0]/usr/ctxtRM61X-TRMPL").text = "2"
            self.session.findById("wnd[0]").sendVKey(0)
            try:
                self.session.findById("wnd[0]").sendVKey(0)
            except:
                pass

        self.peps_programados = []
        self.erros_pep = [] #Cria a lista que irá guardar os peps que possuirem erro ao longo da iteração
        #Para cada EC, roda a VA02, CN33, coloca a data na CJ20N e roda o MRP na MD51

        for index, row in self.df_programacao_script.iterrows():

            print("\nIniciando iteração da EC", row['EC'])
            try:
                va02()
                print("EC", row['EC'], " -> Etapa VA02 concluída com sucesso.")
                self.peps_programados.append(self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Elemento PEP'].values[0])
            except:
                self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Status Programação SAP'] = 'ERRO NA VA02'
                self.df_etapas_script.loc[self.df_etapas_script['EC'] == row['EC'], 'OBSERVAÇÃO'] = 'ERRO NA VA02'
                print("EC", row['EC'], " -> Erro na Etapa VA02, indo para a próxima EC.")
                continue
            
            try:
                cn33()
                print("EC", row['EC'], " -> Etapa CN33 concluída com sucesso.")
            except:
                self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Status Programação SAP'] = 'ERRO NA CN33'
                self.erros_pep.append(self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Elemento PEP'].values[0])
                self.df_etapas_script.loc[self.df_etapas_script['EC'] == row['EC'], 'OBSERVAÇÃO'] = 'ERRO NA CN33'
                print("EC", row['EC'], " -> Erro na Etapa CN33, indo para a próxima EC.")
                continue
        
            try:    
                self.volumes = self.df_input_itens[self.df_input_itens['EC'] == row['EC']]['Volume'] #Guarda os volumes da EC que está sendo analizada. É utilizado na CJ20N
                cj20n()
                print("EC", row['EC'], " -> Etapa CJ20N concluída com sucesso.")
            except:
                self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Status Programação SAP'] = 'ERRO NA CJ20N'
                self.erros_pep.append(self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Elemento PEP'].values[0])
                self.df_etapas_script.loc[self.df_etapas_script['EC'] == row['EC'], 'OBSERVAÇÃO'] = 'ERRO NA CJ20N'
                print("EC", row['EC'], " -> Erro na Etapa CJ20N, indo para a próxima EC.")
                continue

            if self.num_alteracoes != self.volumes.drop_duplicates().count():
                print("Falta alterações na CJ2ON, EC não será levada adiante")
                self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Status Programação SAP'] = 'ERRO NA CJ20N'
                self.df_etapas_script.loc[self.df_etapas_script['EC'] == row['EC'], 'OBSERVAÇÃO'] = 'ERRO NA CJ20N'
                self.erros_pep.append(self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Elemento PEP'].values[0])
                print("EC", row['EC'], " -> Erro na Etapa CJ2ON, nem todos os KITs foram alterados, indo para a próxima EC")
                continue

            try:
                md51()
                print("EC", row['EC'], " -> Etapa MD51 concluída com sucesso.")
            except:
                self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Status Programação SAP'] = 'ERRO NA MD51'
                self.erros_pep.append(self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'Elemento PEP'].values[0])
                self.df_etapas_script.loc[self.df_etapas_script['EC'] == row['EC'], 'OBSERVAÇÃO'] = 'ERRO NA MD51'
                print("EC", row['EC'], " -> Erro na Etapa MD51, indo para a próxima EC")
                continue

        self.peps_cohv = [pep for pep in self.peps_programados if pep not in self.erros_pep]
        self.peps_cohv = pd.Series(self.peps_cohv)

        if self.peps_cohv.empty:
            print("Todas as ECs tiveram erro SAP na geração da necessidade.")
            print("O script será encerrado.")
            with pd.ExcelWriter(self.path_ec_script + "\\" + self.file_macro, mode='a', if_sheet_exists='replace') as writer:
                self.df_input_ec.to_excel(writer, index=False)

            with pd.ExcelWriter(self.path_ec + "\\" + self.file_ec_script, mode='a', if_sheet_exists='replace') as writer:
                self.df_etapas_script.to_excel(writer, index=False)
            sys.exit()

        self.peps_cohv.to_csv(self.path_ec_script + "\\" + self.file_pep, index=False, header=False)

    def converter_ordens(self):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NCOHV"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/chkPPIO_ENTRY_SC1100-SELECT_PLANNEDORDS").selected = True
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/chkPPIO_ENTRY_SC1100-SELECT_PRODORDS").selected = False
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/ROBÔ_EC" #Definir como atributo da classe
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_FEVOR-LOW").text = "010"
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectColumn("PROJN")

        try:
            self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")
        except:
            pass

        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&MB_FILTER")
        self.session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN003_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[2]/tbar[0]/btn[23]").press()
        self.session.findById("wnd[3]/usr/ctxtDY_PATH").text = self.path_ec_script
        self.session.findById("wnd[3]/usr/ctxtDY_FILENAME").text = self.file_pep
        self.session.findById("wnd[3]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[2]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

        #Lógica criada para selecionar todas as linhas da COHV.
        #Não foi usado o botão "selecionar tudo" pois apresentava algum bug em algumas ocasiões 
        i = 0
        while True:
            try:
                self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectedRows = f'0-{i}'
            except:
                break
            i += 1

        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectedRows = f'0-{(i-1)}'
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("COWB_HVOM")
        self.session.findById("wnd[1]/usr/subFUNCTION_SETUP:SAPLCOWORK:0200/cmbCOWORK_FCT_SETUP-FUNCT").key = "210"
        self.session.findById("wnd[1]/usr/subFUNCTION_SETUP:SAPLCOWORK:0200/subFUNCTION_PARAM:SAPLCOWORK210:0100/ctxtCOWORK210_SETUP-AUART").text = "ZETE"
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()

    def liberar_ordens(self):

        #COHV LIBERAR OPLAs
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/ncohv"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/ROBÔ_EC"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/btn%_S_PROJN_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/tbar[0]/btn[23]").press()
        self.session.findById("wnd[2]/usr/ctxtDY_PATH").text = self.path_ec_script
        self.session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = self.file_pep
        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()

        #Lógica criada para selecionar todas as linhas da COHV.
        #Não foi usado o botão "selecionar tudo" pois apresentava algum bug em algumas ocasiões 
        i = 0
        while True:
            try:
                self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectedRows = f'0-{i}'
            except:
                break
            i += 1

        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectedRows = f'0-{(i-1)}'

        try:
            self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")
        except:
            pass

        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton("COWB_HVOM")
        self.session.findById("wnd[1]/usr/subFUNCTION_SETUP:SAPLCOWORK:0200/cmbCOWORK_FCT_SETUP-FUNCT").key = "130"
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()

        j = i - 1
        for tentativa in range(j):
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
            except:
                continue

        #Informa o usuário que programou a EC no SAP
        self.df_input_ec.loc[self.df_input_ec['Elemento PEP'].isin(self.peps_cohv), 'Usuário SAP'] = self.usuario
        self.df_input_ec.loc[self.df_input_ec['Elemento PEP'].isin(self.peps_cohv), 'Status Programação SAP'] = "PROGRAMADO"
        self.df_input_ec.loc[self.df_input_ec['Elemento PEP'].isin(self.peps_cohv), 'ENVIADO E-MAIL'] = "NÃO"

        self.df_etapas_script.loc[self.df_etapas_script['PEP'].isin(self.peps_cohv), 'OBSERVAÇÃO'] = "PROGRAMADO VIA SCRIPT PYTHON"
        self.df_etapas_script.loc[self.df_etapas_script['PEP'].isin(self.peps_cohv), 'STATUS'] = "PROGRAMADO"

    def gravar_relatorio(self):
        with pd.ExcelWriter(self.path_ec_script + "\\" + self.file_macro, mode='a', if_sheet_exists='replace') as writer:
            self.df_input_ec.to_excel(writer, index=False)

        with pd.ExcelWriter(self.path_ec + "\\" + self.file_ec_script, mode='a', if_sheet_exists='replace') as writer:
            self.df_etapas_script.to_excel(writer, index=False)
        
        print("Relatórios atualizados com sucesso!")


        
class EnviarEmail:
    def __init__(self, path_ec, path_ec_script, session, df_input_itens, df_etapas_script, file_macro, file_ec_script):
        self.outlook = win32com.client.Dispatch('outlook.application')
        self.excel = win32com.client.Dispatch('Excel.Application')
        self.session = session
        self.file_dados_programadores = '3. Dados  programadores MRP.xlsx'
        self.path_ec = path_ec
        self.path_ec_script = path_ec_script
        self.itens_ec = 'itens.txt'
        self.pasta_arquivos_zp058 = r'\\srvfile01\DADOS_PBI\Compartilhado_BI\DPCP\2. Programação\EC\Script EC\Dados ZP058'
        self.arquivos_zp058 = glob.glob(self.pasta_arquivos_zp058 + "\\" + '*.xlsx')
        self.email_cc = 'gustavo.ferraz@tkelevator.com;rodrigo.ramiro@tkelevator.com;gustavo.schmidt@tkelevator.com'
        self.df_input_itens = df_input_itens
        self.df_etapas_script = df_etapas_script
        self.file_macro = file_macro
        self.file_ec_script = file_ec_script
    
    def ler_input(self):
        self.df_input_ec = pd.read_excel(self.path_ec_script + "\\" + self.file_macro)
        self.df_etapas_script = pd.read_excel(self.path_ec + "\\" + self.file_ec_script)

    def filtra_df(self):
        self.df_input_ec_email = self.df_input_ec[(self.df_input_ec['Status Sistema'] == 'PROCESSADA') & (self.df_input_ec['Status Programação SAP'] == 'PROGRAMADO') & (self.df_input_ec['ENVIADO E-MAIL'] == 'NÃO')]
        print("Etapas que serão enviadas e-mail: ", self.df_input_ec_email['EC'])

    def le_dados_programadores(self):
        self.df_programadores = pd.read_excel(self.path_ec + "\\" + self.file_dados_programadores)
        self.df_planejador_mrp = pd.read_excel(self.path_ec + "\\" + self.file_dados_programadores, sheet_name='Programadores')
        self.df_email_programadores = pd.read_excel(self.path_ec + "\\" + self.file_dados_programadores, sheet_name='Emails')
    
    def compila_dados_programadores(self):
        self.df_programadores = self.df_planejador_mrp.merge(self.df_email_programadores, how='left', on=['Programador'])
    
    def extrair_zp058(self):
        def criar_txt_materiais():
            self.df_materiais = self.df_input_itens.loc[self.df_input_itens['EC'] == row['EC'], ['Código', 'Quantidade']]
            self.df_materiais.to_csv(self.path_ec_script + "\\" + self.itens_ec, sep='\t', index=False, header=False)
            
        def zp058():
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nzp058"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subSC_CONJ1:ZPPR045:0101/ctxtP_WERKS1").text = "5001"
            self.session.findById("wnd[0]/usr/subSC_OPTIONS:ZPPR045:0103/ctxtPC_VARI").text = "/ROBÔ_EC"
            self.session.findById("wnd[0]/usr/btn%#AUTOTEXT005").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.path_ec_script
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = self.itens_ec
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(1)
            try:
                self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
            except:
                time.sleep(1)
                self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").pressToolbarContextButton("&MB_EXPORT")
            self.session.findById("wnd[0]/usr/shell/shellcont[0]/shell").selectContextMenuItem("&XXL")
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.pasta_arquivos_zp058
            nome_arquivo = row['EC'].replace('/', '_')
            self.lista_excel_aberto.append(nome_arquivo)
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo + '.xlsx'
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

            time.sleep(3)
            try:
                workbook = xw.Book(nome_arquivo + '.xlsx')
                workbook.close()
            except:
                pass

        #Limpar a pasta antes de extrair os arquivos
        for extracoes in self.arquivos_zp058:
            try:
                os.remove(extracoes)
            except:
                continue
        
        self.lista_excel_aberto = []

        for index, row in self.df_input_ec_email.iterrows():
            print('Extraindo:', row['EC'])
            try:
                criar_txt_materiais()
                zp058()
            except:
                
                print('Erro ao extrair:', row['EC'], 'Indo para a próxima EC.')
                self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'ENVIADO E-MAIL'] = 'ERRO ZP058'
                self.df_etapas_script.loc[self.df_etapas_script[['EC']] == row['EC'], 'ENVIADO EMAIL - ITENS DEPM'] = 'ERRO ZP058' 
                with pd.ExcelWriter(self.path_ec_script + "\\" + self.file_macro, mode='a', if_sheet_exists='replace') as writer:
                    self.df_input_ec.to_excel(writer, index=False)

                with pd.ExcelWriter(self.path_ec + "\\" + self.file_ec_script, mode='a', if_sheet_exists='replace') as writer:
                    self.df_etapas_script.to_excel(writer, index=False)
                
                continue
        
        try:
            #Segunda tentativa em fechar os arquivos
            for arquivo in self.lista_excel_aberto:
                try:
                    workbook = xw.Book(arquivo + '.xlsx')
                    workbook.close()
                except:
                    continue
        except:
            pass
        
    def enviar_email(self):
        def veririca_ecs_materiais_unitarios():
            self.unitario = False
            nome_arquivo = row['EC'].replace('/', '_')
            self.df_tabela_email = pd.read_excel(self.pasta_arquivos_zp058 + "\\" + nome_arquivo + ".xlsx")

            if self.df_tabela_email['Componente'].count() == 1:
                self.unitario = True
            
        def montar_tabela_unitarios():
            self.df_input_itens_ec = self.df_input_itens[self.df_input_itens['EC'] == row['EC']]
            self.df_tabela_email['Componente'] = self.df_tabela_email['Componente'].astype(str)
            self.df_input_itens_ec['Código'] = self.df_input_itens_ec['Código'].astype(str)
            self.df_tabela_email = self.df_tabela_email.rename(columns={'Planejador MRP': 'Planejador'})
            self.df_tabela_email = self.df_tabela_email.merge(self.df_input_itens_ec[['Código','Lance', 'Medida']], how='left', left_on='Componente', right_on='Código')
            self.df_tabela_email = self.df_tabela_email.drop(columns='Código')
            self.df_tabela_email = self.df_tabela_email.merge(self.df_planejador_mrp[['Planejador', 'Programador']], how='left', on='Planejador')
            self.df_tabela_email = self.df_tabela_email.fillna(0)

        def montar_tabela_nao_unitarios():
            self.df_tabela_email = self.df_tabela_email.rename(columns={'Texto breve objeto': 'Texto breve', 'Qtd. 1': 'Quantidade', 'UM 1': 'Unidade',
                                                            'Planejador MRP': 'Planejador'})
            self.df_tabela_email = self.df_tabela_email.merge(self.df_planejador_mrp[['Planejador', 'Programador']], how='left', on='Planejador')
            self.df_tabela_email = self.df_tabela_email.sort_values(by='Estoque')     
        
        def definir_destinatarios():
            self.lista_destinatarios = self.df_tabela_email.merge(self.df_email_programadores, how='left', on='Programador')
            self.lista_destinatarios = self.lista_destinatarios['Email']
            self.lista_destinatarios = self.lista_destinatarios.drop_duplicates()
            self.lista_destinatarios = self.lista_destinatarios.dropna()
            self.lista_destinatarios = self.lista_destinatarios.tolist()

        def enviar_email():
            email = self.outlook.CreateItem(0)
            email.Subject = "EC " + row['EC'] + " Obra " + str(self.df_input_ec_email[(self.df_input_ec_email['EC'] == row['EC'])]['Obra'].values[0])

            self.df_tabela_email = self.df_tabela_email.to_html(index=False)
            
            #Enviar o e-mail para o analista responsável
            email.To = ';'.join(self.lista_destinatarios)
            email.CC = self.email_cc #Cópia padrão: para Camila Bitelo Moura

            #Cria o corpo do email
            lista_email = ['---Mensagem Automática - Departamento de Planejamento e Controle de Produção (DPCP)--']
            lista_email.append('<p>Prezados(as),</p>')
            lista_email.append(f'<p>A EC em assunto foi programada para o dia {row['Data Planejada'].strftime('%d/%m/%Y')} </p>')
            lista_email.append(self.df_tabela_email)
            lista_email.append('<br>')
            lista_email.append('<br>')
            lista_string = '\n'.join(lista_email)

            email.HTMLBody = lista_string

            #Envia o e-mail
            email.Send()

        for index, row in self.df_input_ec_email.iterrows():
            print("\nIniciando o envio de email da EC: ", row['EC'])
            veririca_ecs_materiais_unitarios()

            if self.unitario == True:
                montar_tabela_unitarios()
            else:
                montar_tabela_nao_unitarios()
            
            definir_destinatarios()
            print(row['EC'], " -> Destinatários definidos")
            enviar_email()
            print(row['EC'], " -> Email enviado")
            self.df_input_ec.loc[self.df_input_ec['EC'] == row['EC'], 'ENVIADO E-MAIL'] = 'SIM'
            self.df_etapas_script.loc[self.df_etapas_script['EC'] == row['EC'], 'ENVIADO EMAIL - ITENS DEPM'] = 'ENVIADO E-MAIL DEPM'
        
    def gravar_relatorio(self):
        with pd.ExcelWriter(self.path_ec_script + "\\" + self.file_macro, mode='a', if_sheet_exists='replace') as writer:
            self.df_input_ec.to_excel(writer, index=False)

        with pd.ExcelWriter(self.path_ec + "\\" + self.file_ec_script, mode='a', if_sheet_exists='replace') as writer:
            self.df_etapas_script.to_excel(writer, index=False)
            
def main():
    print("Iniciando o script...")
    programar_ec = ProgramarSAP()
    programar_ec.verificando_planilhas_abertas()
    print("Conectando com SAP...")
    programar_ec.conectar_sap()
    print("Lendo input...")
    programar_ec.ler_input()
    print("Verificando incoerência entre planilhas...")
    programar_ec.verificar_incoerencias()
    print("Tratando dados...")
    programar_ec.tratar_dados()
    print("Montando a tabela de report ECs realizadas pelo script...")
    programar_ec.mesclar_tabelas()
    print("Iniciando a programação SAP...")
    programar_ec.va02_cn33_cj20n_md51()
    print("Convertendo ordens de produção na COHV...")
    programar_ec.converter_ordens()
    print("Liberando ordens de produção na COHV...")
    programar_ec.liberar_ordens()
    print("Atualizando os relatórios...")
    programar_ec.gravar_relatorio()

    print("\nIniciando o processo de envio de e-mail...")
    enviar_email = EnviarEmail(programar_ec.path_ec, programar_ec.path_ec_script, programar_ec.session, programar_ec.df_input_itens, programar_ec.df_etapas_script, programar_ec.file_macro, programar_ec.file_ec_script)
    print("Lendo o input...")
    enviar_email.ler_input()
    print("Filtrando ECs que serão enviados os e-mails...")
    enviar_email.filtra_df()
    print("Lendo os dados dos programadores...")
    enviar_email.le_dados_programadores()
    print("Compulando os dados dos programadores...")
    enviar_email.compila_dados_programadores()
    print("Iniciando a extração de dados na ZP058")
    enviar_email.extrair_zp058()
    print("Enviando o e-mail...")
    enviar_email.enviar_email()
    print("\nGravando os relatórios finais...")
    enviar_email.gravar_relatorio()
    print("Script finalizado com sucesso. Verifique se todos os e-mails foram enviados corretamentes ao DEPM")
if __name__ == '__main__':
    main()
