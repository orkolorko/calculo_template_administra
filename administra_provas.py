#!/usr/bin/env python
# -*- coding: utf-8 -*-
u"programa para ler turmas de alunos, gerar cadernos para cabeçalho de provas \
e gerar tabelas para lançamento de notas"
import sys, os, os.path, time
import traceback ,re
import codecs, platform
import datetime
from modulo_alunos_cabecalho import GrupoAlunos
from modulo_alunos_cabecalho import CabecalhoProva
from modulo_alunos_cabecalho import ParaMaiuscula
from modulo_manuseia_odf import ManuseiaArquivoOdf 
from modulo_manuseia_xml import ManuseiaXml 
import wx
import wx.lib
import wx.lib.scrolledpanel
import wx.grid
import wx.py

#import modulo_alunos_cabecalho
import wx.lib.sized_controls
import wx.lib.filebrowsebutton
import time, thread
import wx.lib.delayedresult 
#---------------------------------------------------------------------------
# Create and set a help provider.  Normally you would do this in
# the app's OnInit as it must be done before any SetHelpText calls.
PROVIDER = wx.SimpleHelpProvider()
wx.HelpProvider.Set(PROVIDER)


class MyFrame(wx.Frame):
    "classe para gerenciar as janelas do programa"
    ID_ABRE_PDF = 100
    ID_EXIT  = 101
    ID_SALVA = 102
    ID_SALVA_COMO = 103
    ID_ABRE_ALUNOS = 104
    ID_GERA_CAD = 105
    ID_COMPILA_TEX = 106
    ID_VER_PDF = 107
    ID_NOVO_ALUNOS = 109
    ID_AT_DE_PLAN = 110
    ID_AT_DE_PDF = 111
    ID_GERA_PLANILHA = 112
    ID_ABOUT = 113
    ID_HELP = 114
    def __init__(self, parent, ident, title):
        self.title = title
            
        wx.Frame.__init__(self, parent, ident, self.title, wx.DefaultPosition,
                          wx.Size(1000, 600)) 
#       para testar usando o pycrust
#       shell = wx.py.crust.CrustFrame( self.frame, 
#               locals=dict(wx=wx, frame=self.frame, app=self))
#       shell.Show()
#
        today = datetime.datetime.today()
        today = str(today.day)+ "_" + str(today.month)+ \
            "_" + str(today.year) + "__" + str(today.hour) +\
            "h" + str(today.minute) + "m" + str(today.second) + "s"
        self.CreateStatusBar()
        self.SetStatusText(u"Espaço de avisos")

        menubar = wx.MenuBar()
 
        menu = wx.Menu()
        self.arquivo_csv = False
        menu.Append(MyFrame.ID_NOVO_ALUNOS, u"&Novo...",
                    u"Abre malha vazia para editar alunos")
        menu.Append(MyFrame.ID_ABRE_ALUNOS, u"&Abre...",
                    u"Abre a lista com os alunos")
        menu.Append(MyFrame.ID_ABRE_PDF, u"&Importar alunos de pdf...",
                     u"Adiciona alunos das turmas a partir da pauta do SIGA")
        menu.AppendSeparator()
        self.mn_it_salva = menu.Append(MyFrame.ID_SALVA, u"&Salva", 
          u"Salva a lista no formato csv com os alunos no arquivo de origem")
        self.mn_it_salva_como = menu.Append(MyFrame.ID_SALVA_COMO,
                                            u"Salva &como...",
          u"Salva a lista no formato csv com os alunos em um local a escolher")
        menu.AppendSeparator()
        menu.Append(MyFrame.ID_EXIT, u"&Fim", u"Finaliza o programa")
        menubar.Append(menu, u"&Arquivo")
        menu1 = wx.Menu()
        menu1.AppendSeparator()
        self.mn_it_gr_tx = menu1.Append(MyFrame.ID_GERA_CAD,
               u"&Gera cadernos prova...",
               u"Gera arquivos latex com os cadernos de prova,"+
               u" lista de presença, distribuição dos alunos por sala etc")
        self.mn_it_gr_pl = menu1.Append(MyFrame.ID_GERA_PLANILHA,
             u"&Gera planilha inicial...",
             u"Gera planilha no formato do OpenOffice com os "+
             u"alunos inscritos e as turmas a partir de uma planilha padrão")
        menu1.AppendSeparator()
        self.mn_at_de_pdf = menu1.Append(MyFrame.ID_AT_DE_PDF,
              u"&Atualiza inscrições a partir de pdf e da planilha...", 
              u"Importa arquivos .pdf com as pautas das turmas para "
              u"atualizar a situação da inscrição dos alunos neste  "
              u"programa e, também, na planilha de notas" )
        self.mn_at_de_plan = menu1.Append(MyFrame.ID_AT_DE_PLAN,
              u"&Atualiza aprovados/reprovados de planilha...",
              u"Importa arquivo .ods do OpenOffice onde serão procurados "+
              u"os alunos já aprovados ou reprovados")
#       self.mn_it_cp_tx = menu1.Append(MyFrame.ID_COMPILA_TEX,
#              u"&Compila tex...", u"Compila o arquivo latex")
#       self.mn_it_vs_tx = menu1.Append(MyFrame.ID_VER_PDF, 
#                  u"&Visualiza pdf...", u"Visualiza o arquivo pdf gerado")
        menubar.Append(menu1, u"&Ferramentas")

        menu2 = wx.Menu()
        self.mn_help = menu2.Append(MyFrame.ID_HELP, u"&Ajuda", 
                                    u"Descrição do programa, guia de uso etc")
        self.mn_about = menu2.Append(MyFrame.ID_ABOUT, 
              u"&Sobre", u"Indica dados do programa, autor(es), licença etc")
        menubar.Append(menu2, u"&Ajuda")
 
        self.SetMenuBar(menubar)

#       callbacks para os itens do menu
        wx.EVT_MENU(self, MyFrame.ID_NOVO_ALUNOS, self.novos_alunos)
        wx.EVT_MENU(self, MyFrame.ID_ABRE_ALUNOS, self.abre_alunos)
        wx.EVT_MENU(self, MyFrame.ID_SALVA, self.salva_alunos)
        wx.EVT_MENU(self, MyFrame.ID_SALVA_COMO, self.salva_alunos_como)
        wx.EVT_MENU(self, MyFrame.ID_ABRE_PDF, self.abre_pdf)
        wx.EVT_MENU(self, MyFrame.ID_EXIT,  self.timetoquit)

        wx.EVT_MENU(self, MyFrame.ID_GERA_CAD, self.gera_tex)
        wx.EVT_MENU(self, MyFrame.ID_GERA_PLANILHA, self.gera_planilha)
        wx.EVT_MENU(self, MyFrame.ID_AT_DE_PLAN, self.atualiza_aprovacao_de_planilha)
        wx.EVT_MENU(self, MyFrame.ID_AT_DE_PDF, self.atualiza_de_pdf)
        wx.EVT_MENU(self, MyFrame.ID_HELP, self.ajuda)
        wx.EVT_MENU(self, MyFrame.ID_ABOUT, self.about)

        self.grid = TableGrid(self)
        self.trava_itens_menu_arquivo()
        self.trava_itens_menu_ferramentas()
        diretorio_do_python, dummy = os.path.split(os.path.abspath(sys.argv[0]))
        nome_arquivo_config = os.path.join(diretorio_do_python, \
                   ".administra_provas.xml" )
        try:
            print nome_arquivo_config
            self.file_arquivo_recentes = open ( nome_arquivo_config, "w+" )
        except:
            print "Problemas com a criação do arquivo de uso recente"
            print nome_arquivo_config
            print "Não será usado arquivo de uso recente"
            self.SetStatusText(u"Arquivo com os usos recentes não pôde ser criado/alterado")
            self.file_arquivo_recentes = False
        if self.file_arquivo_recentes:
            self.arquivos_recentes = ManuseiaXml ( self.file_arquivo_recentes,
                                                  "administra_provas" )
        else:
            self.arquivos_recentes = False

    def novos_alunos (self, event):
        u"metodo para inicializar a malha, liberando para edição"
        self.SetStatusText(u"Inicializando um novo conjunto de alunos")
        grupo_alunos_aux = GrupoAlunos()
        self.grid.carrega_grupo_de_alunos ( grupo_alunos_aux )
        self.libera_itens_menu_arquivo()
        self.trava_itens_menu_ferramentas()
        self.SetStatusText(u"Liberada a malha para edição")

    def abre_alunos (self, event):
        "metodo para abrir arquivo csv com os alunos das turmas"
        self.SetStatusText(u"Abrindo arquivo com os dados dos alunos")
        dlg = wx.FileDialog( self,
            message=u"Selecione arquivo com a lista de alunos",
            defaultDir=os.getcwd(), defaultFile=u"", wildcard=u"*.csv",
            style=wx.OPEN | wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            self.SetStatusText(u"Aberto arquivo com a lista dos alunos")
            grupo_alunos_aux = GrupoAlunos()
            self.arquivo_csv = os.path.abspath(dlg.GetPath())
            grupo_alunos_aux.le_alunos_arquivo_txt( self.arquivo_csv )
            self.grid.carrega_grupo_de_alunos ( grupo_alunos_aux )
            self.libera_itens_menu_arquivo()
            self.libera_itens_menu_ferramentas()
        else:
            self.SetStatusText(u"Espaço de avisos")
        dlg.Destroy()

    def abre_pdf (self, event):
        u"call back para abrir um ou mais arquivos pdf" 
        if self.arquivo_csv:
             dl1 = wx.MessageBox(u'Se você carregar os alunos da '+
                 u'pauta do SIGA perderá a numeração original dos alunos, '+
                 u'o que é muito sério se já tiver lançado as notas de '+
                 u'alguma prova.  A troca do número de algum aluno ao longo'+
                 u' do período dificultará a digitação das notas pois estas'+
                 u' deverão passar a ser controladas pelo nome e não pelo' +
                 u' número.', 
                 u'Advertência sobre a numeração dos alunos') 
#                wx.OK | wx.ICON_INFORMATION) 
                 #wx.YES_NO|wx.NO_DEFAULT|wx.CANCEL|wx.ICON_INFORMATION)
#            dl1.ShowModal()
#            dl1.Destroy()
        dlg = wx.FileDialog( self, message=u"Selecione arquivo(s) pdf",
             defaultDir=os.getcwd(), defaultFile=u"", wildcard=u"*.pdf",
             style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            grupo_alunos_aux = self.grid.retorna_grupo_de_alunos ( )
#           alunos_antes = grupo_alunos_aux.get_numero_de_alunos()
            for path in paths:

                dlg_aux = wx.TextEntryDialog( self, 
                    u'De, com poucos caracteres, o nome da turma '+
                    u'correspondente ao arquivo \n '+path,
                    u'Nome turma', u'Python')

                dlg_aux.SetValue(u"")

                if dlg_aux.ShowModal() == wx.ID_OK:
                    turma_atual = dlg_aux.GetValue()
                    turma_atual = turma_atual.replace("_", "/")
                    grupo_alunos_aux.le_alunos_arquivo_pdf(path, 
                                                           turma_atual, True) 
                    dlg_aux.Destroy()
            self.grid.carrega_grupo_de_alunos ( grupo_alunos_aux )
            self.libera_itens_menu_arquivo()
            self.trava_itens_menu_ferramentas()
            self.SetStatusText(
               u"Carregada uma turma a partir de arquivo(s) .pdf")

        dlg.Destroy()

    def salva_alunos (self, event):
        u"""callback para salvar a relação de alunos no arquivo
        que foi carregado"""
        if self.arquivo_csv:
            self.grava_arquivo_csv ( self.arquivo_csv )
            self.libera_itens_menu_ferramentas()
            file_path, nome_arq = os.path.split(self.arquivo_csv)
            self.SetStatusText(u"Arquivo " + nome_arq + " com os alunos foi salvo")
        else:
            self.salva_alunos_como ( event )

    def salva_alunos_como (self, event):
        u"""callback para salvar a relação de alunos em arquivo 
            a ser escolhido pelo usuario"""
        dlg = wx.FileDialog( self, 
                  message=u"Selecione arquivo para salvar a lista de alunos",
                  defaultDir=os.getcwd(), defaultFile=u"", wildcard=u"*.csv",
                  style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT | wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            dlg.GetPath()
            dlg.GetPath()
            dlg.GetPath()
            dlg.GetPath()
            dlg.GetPath()
            nome_arq = os.path.abspath(dlg.GetPath())
            nome_arq = os.path.abspath(dlg.GetPath())
            nome_arq = os.path.abspath(dlg.GetPath())
            nome_arq = os.path.abspath(dlg.GetPath())
            nome_arq = os.path.abspath(dlg.GetPath())
            nome_arq = os.path.abspath(dlg.GetPath())
            nome_arq = os.path.abspath(dlg.GetPath())
            nome_arq = os.path.abspath(dlg.GetPath())
            nome_arq = os.path.abspath(dlg.GetPath())
            if nome_arq[-4:] != ".csv":
                nome_arq = nome_arq + ".csv"
            self.grava_arquivo_csv ( nome_arq )
            self.libera_itens_menu_ferramentas()
        dlg.Destroy()
        self.SetStatusText(
            u"Arquivo .csv com os alunos foi salvo com o nome escolhido")


    def grava_arquivo_csv ( self, arq ):    
        u"metodo auxiliar para mandar gravar o arquivo csv" 
        grupo_aux = self.grid.retorna_grupo_de_alunos ( )
        grupo_aux.grava_alunos_arquivo_txt(arq)
        if self.arquivos_recentes:
            self.arquivos_recentes.set_atributo ( arq, "dia_gravacao", time.strftime("%d-%m-%Y %Hh%Mm%Ss" ) )
            self.arquivos_recentes.grava_arvore ( )
        self.arquivo_csv = arq


    def gera_tex (self, event):
        u"""callback para chamar as janelas que permitirão gerar arquivos
        latex para impressão dos cabeçalhos, folhas de presença etc"""
        alunos_efetivos = self.grid.retorna_num_alunos_para_prova ( )
        dlg = DialogoProva ( alunos_efetivos )
#       dlg.CenterOnScreen()

        # this does not return until the dialog is closed.
        val = dlg.ShowModal()
    
        if val == wx.ID_OK:
            self.SetStatusText(
              u"Preparando para gerar arquivos latex para os cadernos de prova")
            disciplina, idia, mes, iano, prova, nome_arq_modelo_prova, \
            nome_arq_caderno_prova, nome_arq_lista_alunos, nome_arq_irregular, \
            nome_arq_presenca, nome_arq_salas_usadas, \
            num_questoes,  dic_salas = dlg.retorna_dados()
            dia = u"%2.2d" % (idia)
#           mes = u"%2.2d" % (imes)
            ano = u"%4.4d" % (iano)
            grupo_alunos_aux = self.grid.retorna_grupo_de_alunos ( )
            CabecalhoProva.gera_cadernos_prova ( nome_arq_modelo_prova,
                   nome_arq_caderno_prova, grupo_alunos_aux, dia, mes, ano,
                   disciplina, prova, num_questoes, dic_salas )
            CabecalhoProva.gera_listas_alunos ( nome_arq_lista_alunos,
                   nome_arq_irregular, nome_arq_presenca, nome_arq_salas_usadas,
                   grupo_alunos_aux, dia, mes, ano, disciplina, prova, dic_salas )
            self.SetStatusText(
                    u"Gerados arquivos Latex com cadernos de prova "+
                    u"e listas de alunos")
#       else:
#           self.SetStatusText(u"Espaço de avisos")
        dlg.Destroy()
        return

    def atualiza_de_pdf (self, event):
        u'''callback para atualizar o status dos alunos através de pauta
        da turma e acrescentar alunos, se necessario.  Alem disso, atualiza
        a planilha e sincroniza as informacoes da planilha e das pautas'''

#       selecionando e processando arquivos com as pautas atualizadas
        dlg = wx.FileDialog( self,
             message=u"Selecione arquivo(s) pdf com as pautas",
             defaultDir=os.getcwd(), defaultFile=u"", wildcard=u"*.pdf",
             style=wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            grupo_alunos_old = self.grid.retorna_grupo_de_alunos ( )
            grupo_alunos_aux = GrupoAlunos()
            conjunto_de_turmas = set()
#           lista_turmas = grupo_alunos_aux.ret_lista_com_turmas( )
#           alunos_antes = grupo_alunos_aux.get_numero_de_alunos()
            for path in paths:
                turma_lida = grupo_alunos_aux.le_atualiza_alunos_arquivo_pdf (
                                path, grupo_alunos_old, False) 
                conjunto_de_turmas.add(turma_lida)
            self.SetStatusText(
               u"Atualizada(s) turma(s) a partir de arquivo(s) .pdf")
        dlg.Destroy()

#       selecionando e processando a planilha com as notas dos alunos
        self.SetStatusText(
            u"Abrindo planilha do OpenOffice com as notas dos alunos")
        dlg = wx.FileDialog( self,
            message = u"Selecione arquivo com as notas dos " \
                    u"alunos para ser atualizado",
                    defaultDir=os.getcwd(), defaultFile=u"", wildcard=u"*.ods",
                    style=wx.OPEN | wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            arquivo_ods = os.path.abspath(dlg.GetPath())
            dlg.Destroy()
            self.progress = ProgressDialog(self,
                        'Processando planilha com notas',
                        "Aguarde: Processando planilha com notas.")
            self.progress.pulse ("Abrindo arquivo com as notas")
            self.SetStatusText(u"Abrindo arquivo com as notas")
            self.timer = wx.Timer(self)
            self.Bind(wx.EVT_TIMER, self.atualiza_progress_dialog, self.timer)
            self.timer.Start(100) 
            self.recebe_mensagem_thread ("Abrindo a planilha a ser atualizada")
            thread.start_new_thread ( self.atualiza_de_pdf_thread, (arquivo_ods,
                                      grupo_alunos_aux, conjunto_de_turmas) )  
            print "chamei a thread"

        else:
            self.SetStatusText(u"Planilha com as notas não foi selecionada")
            dlg2 = wx.MessageDialog (self,
                 u"Como nao foi selecionada a planilha com as notas " \
                 u"o conjunto com os alunos nao será alterado",  
                 wx.OK | wx.ICON_EXCLAMATION) 
            dlg2.ShowModal()
            dlg2.destroy()
            dlg.Destroy()





    def atualiza_de_pdf_thread (self, arquivo_ods, grupo_alunos_aux,
                              conjunto_de_turmas ):

        u'''metodo para realizar a atualização da planilha callback para atualizar o status dos alunos através de pauta
        da turma e acrescentar alunos, se necessario.  Alem disso, atualiza
        a planilha e sincroniza as informacoes da planilha e das pautas'''
        try:
            print "entrei na thread"

            msg = u"Abrindo arquivo com as notas"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            arq_odf = ManuseiaArquivoOdf ( arquivo_ods, sys.stdout )

            msg = u"Iniciando a leitura da planilha"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
#           procurando a coluna do status            
            nome_prim_tabela = arq_odf.retorna_nome_tabela (  u"content.xml", 1)
            print "nome_tabela:", nome_prim_tabela
            nome_arquivo = u"content.xml"
            ind_tabela = 1
            linha_inic = 1
            max_colunas = 3
            linha_lida = arq_odf.retorna_linha_tabela ( nome_arquivo,
                                      ind_tabela, linha_inic, max_colunas )
            msg = u"Descobrindo as colunas a recuperar"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            situacao = linha_lida[1]
            try:
                num_alunos_planilha_orig = int(linha_lida[2])
            except:
                msg_err = u"A célula A2 da planilha com as notas\n" +\
                          u"deve conter o número total de alunos\n" +\
                          u"Conserte o arquivo e torne a executar.\n" +\
                          u"O conteúdo da linha encontrado foi\n"  +\
                          unicode( str(linha_lida), "utf-8" )
                print msg_err
                wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                return
            linha_labels = 2
            max_colunas = 20
            linha_lida = arq_odf.retorna_linha_tabela ( nome_arquivo,
                                      ind_tabela, linha_labels, max_colunas )
            print "linha lida", linha_lida
            col_num = 0
            col_nome = 0
            col_dre = 0
            col_turma = 0
            col_insc = 0
            for col, valor in enumerate(linha_lida):
                if ( valor == "Num." ):
                    col_num = col + 1
                elif ( valor == "Nome" ):
                    col_nome = col + 1
                elif ( valor == "DRE" ):
                    col_dre = col + 1
                elif ( valor == "Turma" ):
                    col_turma = col + 1
                elif ( valor == "Insc."):
                    col_insc = col + 1
                if ( col_num * col_nome * col_dre * col_turma * col_insc > 0 ):
                    break
            if ( col_num == 0 ):
                msg_err = \
                     "Não foi encontrada a coluna com o número dos alunos.\n" +\
                     "Ela tem que ficar na segunda linha e ter o label 'Num.',\n" +\
                     "(sem os plics).  Conserte o arquivo e torne a executar.\n" +\
                     "O conteúdo da linha encontrado foi\n" +\
                          unicode( str(linha_lida), "utf-8" )
                wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                return
            if ( col_nome == 0 ):
                msg_err = \
                     "Não foi encontrada a coluna com o nome dos alunos.\n" +\
                     "Ela tem que ficar na segunda linha e ter o label 'Nome',\n" +\
                     "(sem os plics).  Conserte o arquivo e torne a executar.\n" +\
                     "O conteúdo da linha encontrado foi\n" +\
                          unicode( str(linha_lida), "utf-8" )
                print msg_err
                wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                return
            if ( col_dre == 0 ):
                msg_err = \
                     "Não foi encontrada a coluna com o DRE dos alunos.\n" +\
                     "Ela tem que ficar na segunda linha e ter o label 'DRE',\n" +\
                     "(sem os plics).  Conserte o arquivo e torne a executar.\n" +\
                     "O conteúdo da linha encontrado foi\n" +\
                          unicode( str(linha_lida), "utf-8" )
                print msg_err
                wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                return
            if ( col_turma == 0 ):
                msg_err = \
                   "Não foi encontrada a coluna com a turma dos alunos.\n" +\
                   "Ela tem que ficar na segunda linha e ter o label 'Turma',\n" +\
                   "(sem os plics).  Conserte o arquivo e torne a executar.\n" +\
                   "O conteúdo da linha encontrado foi\n" +\
                          unicode( str(linha_lida), "utf-8" )
                print msg_err
                wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                return
            if ( col_insc == 0 ):
                msg_err = \
                   "Não foi encontrada a coluna com a situacao da inscrição " +\
                   "dos alunos.\n" +\
                   "Ela tem que ficar na segunda linha e ter o label 'Insc.',\n" +\
                   "(sem os plics).  Conserte o arquivo e torne a executar.\n" +\
                   "O conteúdo da linha encontrado foi\n" +\
                          unicode( str(linha_lida), "utf-8" )
                print msg_err
                wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                return
            max_col = max(col_num, col_nome, col_dre, col_turma, col_insc) 
            dict_alunos_planilha = dict()
            lin_inic = 3
            print "Carregando os alunos da planilha"
            msg = u"Carregando os alunos da planilha"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            linhas_lidas = arq_odf.retorna_lista_linhas_tabela ( 
                           nome_arquivo, ind_tabela, lin_inic,
                           lin_inic + num_alunos_planilha_orig - 1, max_col )
            ilin = lin_inic
            for linha_lida in linhas_lidas:
                num_aux = linha_lida[col_num-1]
                try:
                    inum_aux = int(num_aux)
                except:
                    inum_aux = 0
                nome_aux = linha_lida[col_nome-1]
                dre_aux = linha_lida[col_dre-1]
                turma_aux = linha_lida[col_turma-1]
                insc_aux = linha_lida[col_insc-1]
                if nome_aux != "Extra":
                    if not dre_aux or dre_aux.strip() == "":
                        msg_err = u"O aluno de nome '" + nome_aux + \
                        u"' na linha " + str(ilin) + \
                        u" nao teve o seu DRE digitado." +\
                        u"Inclua DRE do aluno na planilha.\n"
                        print msg_err
                        wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                        return
                    elif not re.match('[0-9]{9}', dre_aux.strip()):
                        msg_err = u"O aluno de nome '" + nome_aux + \
                        u"' na linha " + str(ilin) + \
                        u" tem o DRE inválido dado por '" + dre_aux +\
                        u"'.\nInclua o DRE correto do aluno na planilha.\n"
                        print msg_err
                        wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                        return
                    elif dict_alunos_planilha.has_key(dre_aux):
                        msg_err = u"O aluno com DRE " + dre_aux + \
                        u" apareceu mais de uma vez na planilha." +\
                        u"Altere a planilha.\n"
                        print msg_err
                        wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                        return
                    elif inum_aux != (ilin - lin_inic + 1): 
                        print dre_aux
                        print inum_aux
                        msg_err = u"O numero do aluno que está na linha " +\
                        str(ilin) + u" da planilha não corresponde ao número "+\
                        u"da linha mais dois, como era esperado.\n" +\
                        u"Este número pode não ter sido digitado ou a planilha"+\
                        u" pode ter sido reordenada.\n  Altere a planilha.\n"  
                        print msg_err
                        wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                        return
                    else:
                        dict_alunos_planilha[dre_aux] = [num_aux, nome_aux,
                                                     turma_aux, insc_aux, False]
                    ilin += 1
                else:
                    break
                ultima_linha = ilin - 1

#           atualizando situacao alunos que constam da pauta atual. na planilha
            dict_alunos_atualiza =  grupo_alunos_aux.get_dict_alunos()
            msg = u"Atualizando os alunos da planilha"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
#           ordenando alunos por nome
            lista_aux = dict_alunos_atualiza.items()
            lista_aux.sort((lambda x, y:cmp(x[1][0], y[1][0])), None)
            for (dre, lista_dados) in lista_aux:
                if dict_alunos_planilha.has_key(dre):
                    dict_alunos_planilha[dre][4] = True
                    if ( lista_dados[0] != dict_alunos_planilha[dre][1] ):
                        print "Alterando nome do aluno na planilha. DRE:",dre,\
                            "  Nome alterado:",  lista_dados[0],\
                            "  Nome anterior:",  dict_alunos_planilha[dre][1]
                        arq_odf.altera_celula_por_posicao ( u"content.xml",
                            nome_prim_tabela, 
                            int(dict_alunos_planilha[dre][0])+2,
                            col_nome, dict_alunos_planilha[dre][1],
                            lista_dados[0] )
                        dict_alunos_planilha[dre][1] = lista_dados[0]
                    if ( lista_dados[1] != dict_alunos_planilha[dre][2] ):
                        print "Alterando turma aluno na planilha. DRE:",dre,\
                            "  Nome:",  lista_dados[0],\
                            "  Turma alterada:", lista_dados[1],\
                            "  Turma anterior:", dict_alunos_planilha[dre][2]
                        arq_odf.altera_celula_por_posicao ( u"content.xml",
                            nome_prim_tabela, 
                            int(dict_alunos_planilha[dre][0])+2,
                            col_turma, dict_alunos_planilha[dre][2],
                            lista_dados[1] )
                        dict_alunos_planilha[dre][2] = lista_dados[1]
                    if ( lista_dados[2] != dict_alunos_planilha[dre][3] ):
                        print "Alterando insc. do aluno na planilha. DRE:",dre,\
                            "  Nome:",  lista_dados[0],\
                            "  Insc. alterada:", lista_dados[2],\
                            "  Insc. anterior:", dict_alunos_planilha[dre][3]
                        arq_odf.altera_celula_por_posicao ( u"content.xml",
                            nome_prim_tabela, 
                            int(dict_alunos_planilha[dre][0])+2,
                            col_insc, dict_alunos_planilha[dre][3],
                            lista_dados[2] )
                        dict_alunos_planilha[dre][3] = lista_dados[2]
                else:
                    ultima_linha += 1
                    dict_alunos_planilha[dre] = [ultima_linha - 2,
                         lista_dados[0],
                         lista_dados[1],
                         lista_dados[2], True ]
                    print "Acrescentando aluno na planilha. DRE:",dre,\
                        "  Nome:",  lista_dados[0],\
                        "  Turma:", lista_dados[1],\
                        "  Situacao da inscricao:", lista_dados[2]
                    linha_extra = arq_odf.retorna_linha_tabela ( nome_arquivo,
                           ind_tabela, ultima_linha, max_col )
                    arq_odf.altera_celula_por_posicao ( u"content.xml",
                            nome_prim_tabela, ultima_linha,  col_nome,
                            linha_extra[col_nome-1],
                            lista_dados[0] )
                    arq_odf.altera_celula_por_posicao ( u"content.xml",
                            nome_prim_tabela, ultima_linha,  col_dre,
                            linha_extra[col_dre-1],
                            dre )
                    arq_odf.altera_celula_por_posicao ( u"content.xml",
                            nome_prim_tabela, ultima_linha,  col_turma,
                            linha_extra[col_turma-1],
                            lista_dados[1] )
                    arq_odf.altera_celula_por_posicao ( u"content.xml",
                            nome_prim_tabela, ultima_linha,  col_insc,
                            linha_extra[col_insc-1],
                            lista_dados[2] )

#           atualizando a situação dos alunos que sumiram da pauta
            lista_aux = dict_alunos_planilha.items()
            lista_aux.sort((lambda x, y:cmp(x[1][1], y[1][1])), None)
            for (dre, lista_dados) in lista_aux:
                if ( not lista_dados[4] and 
                    lista_dados[2] in conjunto_de_turmas and
                    dict_alunos_planilha[dre][3] != u"NÃO INSCRITO" ):
                    print \
                        "Alterando insc. de aluno que nao consta das ", \
                        "novas pautas. DRE:",dre,\
                        "  Nome:",  lista_dados[1],\
                        "  Insc. alterada:", u"NÃO INSCRITO",\
                        "  Insc. anterior:", dict_alunos_planilha[dre][3]
                    arq_odf.altera_celula_por_posicao ( u"content.xml",
                        nome_prim_tabela, 
                        int(dict_alunos_planilha[dre][0])+2,
                        col_insc, dict_alunos_planilha[dre][3],
                        u"NÃO INSCRITO" )
                    dict_alunos_planilha[dre][3] = u"NÃO INSCRITO"

            print u"gravando a planilha gerada"
            msg = u"Gravando a planilha gerada"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            planilha_saida = arquivo_ods + ".tmp"
            arq_odf.grava_arvore ( planilha_saida )
            print u"liberando a memória"
            arq_odf.libera_memoria ( )
            arq_odf.encerra_objeto ( )
            today = datetime.datetime.today()
            horario = str(today.day)+ "_" + str(today.month)+ \
            "_" + str(today.year) + "__" + str(today.hour) +\
            "h" + str(today.minute) + "m" + str(today.second) + "s"
            nome_alterado = arquivo_ods + "_" + horario + ".bkp"
            os.rename(arquivo_ods, nome_alterado)
            os.rename(planilha_saida, arquivo_ods)
#           copia as colunas da planilha atualizada para a malha deste programa
            wx.CallAfter (self.carrega_alunos_no_grid, dict_alunos_planilha) 
            msg = u"PLANILHA " + os.path.basename(arquivo_ods) +\
                  u" PARA LANÇAMENTO DAS " + u"NOTAS FOI ATUALIZADA.  " +\
                  u"O arquivo antigo foi salvo com sufixo bkp"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            wx.CallAfter(self.abre_dialog_thread, 
             u"Relação de alunos atualizada com sucesso na planilha " +
              arquivo_ods + u"\n A planilha original foi salva com o nome " + 
             nome_alterado + "\n.  Verifique a planilha criada e, caso " +
             u"tenha ocorrido algum problema, renomeie a planilha original",
             u"Atualização de alunos realizada com sucesso")
        except Exception, err:
            print 'print_exc():'
            traceback.print_exc(file=sys.stdout)
            print
            print 'print_exc(1):'
            traceback.print_exc(limit=1, file=sys.stdout)
            msg_err =  u"Erro de execução, ver mensagem no terminal" 
            wx.CallAfter(self.abre_dialog_thread, 
                          msg_err,u"Mensagem de ERRO")
        return 


    def atualiza_aprovacao_de_planilha (self, event):
        u"""callback para atualizar o status dos alunos ( aprovados, 
        reprovados ) através de arquivo de planilha"""
        self.SetStatusText(
            u"Abrindo planilha do OpenOffice com as notas dos alunos")
        dlg = wx.FileDialog( self,
            message=u"Selecione arquivo com as notas dos alunos",
            defaultDir=os.getcwd(), defaultFile=u"", wildcard=u"*.ods",
            style=wx.OPEN | wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            arquivo_ods = os.path.abspath(dlg.GetPath())
            dlg.Destroy()
            self.progress = ProgressDialog(self, 'Processando planilha com notas',
                        "Aguarde: Processando planilha com notas.\n"\
                        "Este procedimento pode demorar um pouco...") 
            self.timer = wx.Timer(self)
            self.Bind(wx.EVT_TIMER, self.atualiza_progress_dialog, self.timer)
            self.timer.Start(100) 
            self.recebe_mensagem_thread ("Abrindo arquivo com as notas")
            thread.start_new_thread ( self.atualiza_aprovacao_de_planilha_thread, 
                             (arquivo_ods,) )  
#           chama por thread o metodo atualiza_aprovacao_de_planilha_thread, com
#           os parametros a seguir.  No final da execução encerrado o ProgressDialog
#           atraves  do metodo abre_dialog_thread chamado ao final do programa em thread.
#           Muitas vezes  a execução se encerrou com core dump e parece ter a ver
#           com encerrar  o progress dialog sem ter processado todos os eventos
#           chamados pelo timer (e, talvez, pelo thread).  
            print "chamei a thread"

        else:
            self.SetStatusText(u"Planilha com as notas não foi selecionada")
            dlg2 = wx.MessageDialog (self,
                 u"Como nao foi selecionada a planilha com as notas " \
                 u"o conjunto com os alunos nao será alterado",  
                 wx.OK | wx.ICON_EXCLAMATION) 
            dlg2.ShowModal()
            dlg2.destroy()
            dlg.Destroy()


    def atualiza_aprovacao_de_planilha_thread (self, arquivo_ods):
        u"""atualizar o status dos alunos ( aprovados, 
        reprovados ) através de arquivo de planilha"""
        try:
            print u"Abrindo o arquivo com as notas e descomprimindo"
            arq_odf = ManuseiaArquivoOdf ( arquivo_ods, sys.stdout )
#           procurando a coluna do status            
            nome_prim_tabela = arq_odf.retorna_nome_tabela (  u"content.xml", 1)
            print "nome_tabela:", nome_prim_tabela
            nome_arquivo = u"content.xml"
            linha_labels = 2
            ind_tabela = 1
            max_colunas = 20
            msg =  u"Carregando as linhas da planilha"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            linha_lida = arq_odf.retorna_linha_tabela ( nome_arquivo,
                                      ind_tabela, linha_labels, max_colunas )
            print "linha lida", linha_lida
            col_sit = 0
            col_dre = 0
            for col, valor in enumerate(linha_lida):
                if ( valor == "Sit." ):
                    col_sit = col + 1
                elif ( valor == "DRE" ):
                    col_dre = col + 1
                if ( col_sit * col_dre > 0 ):
                    break
            if ( col_sit == 0 ):
                msg_err = \
                   "Não foi encontrada a coluna com a situação dos alunos.\n" +\
                   "Ela tem que ficar na segunda linha e ter o label 'Sit.',\n" +\
                   "(sem os plics).  Conserte o arquivo e torne a executar.\n" +\
                   "O conteúdo da linha encontrado foi\n" +\
                          unicode( str(linha_lida), "utf-8" )
                print msg_err
                wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                return
            if ( col_dre == 0 ):
                msg_err = \
                   "Não foi encontrada a coluna com o DRE dos alunos.\n" +\
                   "Ela tem que ficar na segunda linha e ter o label 'DRE',\n" +\
                   "(sem os plics).  Conserte o arquivo e torne a executar.\n" +\
                   "O conteúdo da linha encontrado foi\n" +\
                          unicode( str(linha_lida), "utf-8" )
                print msg_err
                wx.CallAfter(self.abre_dialog_thread, 
                             msg_err,u"Mensagem de ERRO")
                return
            msg = u"Procurando alunos aprovados, reprovados ou não inscritos"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            print u"Procurando alunos aprovados, reprovados ou não inscritos"
            lista_dre_ap = arq_odf.ret_col_de_valores_com_teste ( nome_arquivo,
                       ind_tabela, linha_labels+1, col_sit, col_dre, "AP" )
            lista_dre_rep = arq_odf.ret_col_de_valores_com_teste ( nome_arquivo,
                       ind_tabela, linha_labels+1, col_sit, col_dre, "RM" )
            lista_dre_rmf = arq_odf.ret_col_de_valores_com_teste ( nome_arquivo,
                       ind_tabela, linha_labels+1, col_sit, col_dre, "RMF" )
            lista_dre_ni = arq_odf.ret_col_de_valores_com_teste ( nome_arquivo,
                       ind_tabela, linha_labels+1, col_sit, col_dre, "NI" )
            lista_dre_faz_prova = arq_odf.ret_col_de_valores_com_teste (
                       nome_arquivo, ind_tabela, linha_labels+1, col_sit,
                       col_dre, "--" )
            print "Aprovados:"
            print lista_dre_ap
            print "Reprovados:"
            print lista_dre_rep
            print "Nao inscritos:"
            print lista_dre_ni
            wx.CallAfter ( self.grid.altera_obs_e_status_por_dre, 
                           lista_dre_ap, "APROVADO", 'NAO' )
            wx.CallAfter ( self.grid.altera_obs_e_status_por_dre, 
                           lista_dre_rep, "REPROVADO", 'NAO' )
            wx.CallAfter ( self.grid.altera_obs_e_status_por_dre,
                           lista_dre_rmf, "REPROVADO POR FALTA E MÉDIA", 'NAO' )
            wx.CallAfter ( self.grid.altera_obs_e_status_por_dre, 
                           lista_dre_ni, "NÃO INSCRITO", 'NAO' )
            wx.CallAfter ( self.grid.altera_obs_e_status_por_dre, 
                           lista_dre_faz_prova, "--", 'S' )
            wx.CallAfter(self.abre_dialog_thread, 
                u"Relação dos alunos aprovados e reprovados foi atualizada\n"+
                u"com sucesso nesta aplicação.  Verifique os resultados e\n" +
                u"depois salve o arquivo e gere os cadernos de provas",
                u"Atualização de alunos realizada com sucesso")
        except Exception, err:
            print 'print_exc():'
            traceback.print_exc(file=sys.stdout)
            print
            print 'print_exc(1):'
            traceback.print_exc(limit=1, file=sys.stdout)
            msg_err =  u"Erro de execução, ver mensagem no terminal" 
            wx.CallAfter(self.abre_dialog_thread, 
                          msg_err,u"Mensagem de ERRO")
        return 
        

    def gera_planilha (self, event):
        "callback para chamar o diálogo que irá permitir gerar uma planilha"
#       alunos_efetivos = self.grid.retorna_num_alunos_para_prova ( )
        try:
            dlg = DialogoPlanilha ( )

            # this does not return until the dialog is closed.
            val = dlg.ShowModal()
        
            if val == wx.ID_OK:
                nome_turma, planilha_template, planilha_saida = dlg.retorna_dados()
                dlg.Destroy()
                grupo_alunos_aux = self.grid.retorna_grupo_de_alunos ( )
                self.progress = ProgressDialog(self,
                            'Processando planilha com notas',
                            "Aguarde: Processando planilha com notas.")
                self.timer = wx.Timer(self)
                self.Bind(wx.EVT_TIMER, self.atualiza_progress_dialog, self.timer)
                self.timer.Start(100) 
                self.recebe_mensagem_thread (
                                   "Abrindo o template da planilha com as notas" )
                thread.start_new_thread ( self.gera_planilha_thread, 
                    (planilha_template, planilha_saida, grupo_alunos_aux) )  
                print "voltei"
            else:
                dlg.Destroy()
        except:
            print "Erro"


    def gera_planilha_thread (self, planilha_template, planilha_saida, 
                            grupo_alunos_aux ):
        "thread para gerar uma planilha a partir do template"
        try:
            print u"Iniciando a geração da planilha"
            lista_alunos_aux = grupo_alunos_aux.get_lista_de_alunos()
            
            file_sai = codecs.open(u"aaa", 'w', 'utf-8')
            arq_odf = ManuseiaArquivoOdf ( planilha_template, sys.stdout )
#       arq_odf = ManuseiaArquivoOdf ( planilha_template, file_sai )
#       arq_odf.imp_nomes_dos_arquivos_no_odf()
#       arq_odf.imprime_arquivo ( u'mimetype' )
#       arq_odf.imprime_arquivo ( u'META-INF/manifest.xml' )
#       arq_odf.imprime_arquivo ( u'styles.xml' )
#       arq_odf.imprime_arquivo ( u'meta.xml' )
#       arq_odf.imprime_arquivo ( u'content.xml' )
#       arq_odf.imprime_nos_e_elementos(u"content.xml",
#            set([u"table:table-row", u"table:table-cell", u"text:P"]))  
#       arq_odf.imprime_nos_e_elementos(u"content.xml", set([]))  


#       o que se espera encontrar no template de planilha             
            msg = u"Apagando os alunos do template que estão em excesso"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            dict_template = {"numero_alunos_template":1500, 
                             "numero_alunos_turma_template":150, 
                             "numero_turmas_template":20,
                             "formato_numero":u"%d", 
                             "formato_aluno":u"Aluno %4.4d",
                             "formato_dre":u"DRE%4.4d",
                             "formato_turma":u"Turma%4.4d",
                             "formato_obs":u"OK",
                             "formato_folha_turma":u"Turma_%d"}
            primeiro_excluido = len(lista_alunos_aux)+1
            lista_colunas = [ 1, 2, 3, 4, 5 ]
            lista_linhas_a_excluir = [ 
                [ dict_template["formato_numero"] % iaux,
                  dict_template["formato_aluno"] % iaux,
                  dict_template["formato_dre"] % iaux,
                  dict_template["formato_turma"] %
                    ((iaux - 1) % dict_template["numero_turmas_template"] + 1),
                  dict_template["formato_obs"] ] 
                  for iaux in range(primeiro_excluido,
                                  dict_template["numero_alunos_template"] + 1)]
            arq_odf.apaga_linhas_na_arvore ( u"content.xml",
                           lista_linhas_a_excluir, lista_colunas )


#       altera as linhas na forma vista anteriormente para os dados
#       corretos dos alunos 
            msg = u"Acertando os dados dos alunos"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            lista_valores_a_substituir = [ 
                [ dict_template["formato_numero"] % iaux,
                  dict_template["formato_aluno"] % iaux,
                  dict_template["formato_dre"] % iaux,
                  dict_template["formato_turma"] % 
                     ((iaux - 1) % dict_template["numero_turmas_template"] + 1),
                  dict_template["formato_obs"] ] 
                  for iaux in range(1, primeiro_excluido)]
            lista_valores_novos = [ 
                [ u"%d" % ( iaux + 1 ), nome, dre, turma, obs ]
                for iaux, (nome, dre, turma, obs, status) in
                enumerate(lista_alunos_aux) ]
            arq_odf.substitui_valores_na_arvore ( u"content.xml",
               lista_valores_a_substituir, lista_valores_novos, lista_colunas )

            
            msg = u"Montando as turmas"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
#       acerta numero de alunos por turma
            dic_alunos_por_turma = \
                    grupo_alunos_aux.ret_dic_com_alunos_por_turma()
            lista_turmas = dic_alunos_por_turma.keys()
            lista_turmas.sort()
            for ind, turma in enumerate(lista_turmas):
                iaux = ind + 1
                arq_odf.altera_celula_por_posicao ( u"content.xml",
                            dict_template["formato_folha_turma"] %iaux, 1, 1,
                            dict_template["formato_turma"] %iaux, turma )
                primeiro_a_apagar = dic_alunos_por_turma[turma] + 12
                alunos_a_apagar = \
                        dict_template["numero_alunos_turma_template"] -\
                        primeiro_a_apagar + 1
                if alunos_a_apagar > 0:
                    arq_odf.apaga_linhas_na_tabela ( u"content.xml",
                       dict_template["formato_folha_turma"] %iaux,
                       primeiro_a_apagar, alunos_a_apagar )
                  

#       apaga tabelas relativas a turmas excedentes
            msg = u"Apaga turmas em excesso"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            lista_tabelas_a_excluir = [ 
                dict_template["formato_folha_turma"] % iaux
                for iaux in range(len(lista_turmas) + 1,
                                dict_template["numero_turmas_template"] + 1)]
            arq_odf.apaga_tabelas ( u"content.xml", lista_tabelas_a_excluir )

#       gera a planilha            
            msg = u"Salva a nova planilha"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            arq_odf.grava_arvore ( planilha_saida )
            print u"liberando a memória"
            arq_odf.libera_memoria ( )
            arq_odf.encerra_objeto ( )
            msg =  u"PLANILHA " + planilha_saida + u" PARA LANÇAMENTO DAS " + u"NOTAS FOI GERADA"
            wx.CallAfter(self.recebe_mensagem_thread, msg)
            wx.CallAfter(self.abre_dialog_thread, 
            u"Planilha com a lista de alunos para lançamento das notas criada com sucesso.\n"+\
             u"O nome do arquivo é " + planilha_saida, 
             u"\nGeração de planilha inicial foi bem sucedida" )
            print "saida"
            return
        except Exception, err:
            print 'print_exc():'
            traceback.print_exc(file=sys.stdout)
            print
            print 'print_exc(1):'
            traceback.print_exc(limit=1, file=sys.stdout)
            msg_err =  u"Erro de execução, ver mensagem no terminal" 
            wx.CallAfter(self.abre_dialog_thread, 
                          msg_err,u"Mensagem de ERRO")
            return

        return
        
    def libera_itens_menu_arquivo ( self ):
        "método para liberar os itens do menu depois de criar alunos"
        self.mn_it_salva.Enable(True)
        self.mn_it_salva_como.Enable(True)
        self.grid.Enable(True)

    def libera_itens_menu_ferramentas ( self ):
        "método para liberar os itens do menu ferramentas depois de salvar o arquivo"
        self.mn_it_gr_tx.Enable(True)
        self.mn_it_gr_pl.Enable(True)
        self.mn_at_de_plan.Enable(True)
        self.mn_at_de_pdf.Enable(True)
        file_path, nome_arq = os.path.split(self.arquivo_csv)
        self.SetTitle ( self.title + "( arquivo: " + nome_arq + " )" )

    def trava_itens_menu_arquivo ( self ):    
        "método para travar itens do menu quando não tem alunos"
        self.mn_it_salva.Enable(False)
        self.mn_it_salva_como.Enable(False)
        self.grid.Enable(False)
       
    def trava_itens_menu_ferramentas ( self ):    
        "método para travar itens do menu ferramentas enquanto não salvar o arquivo"
        self.mn_it_gr_tx.Enable(False)
        self.mn_it_gr_pl.Enable(False)
        self.mn_at_de_plan.Enable(False)
        self.mn_at_de_pdf.Enable(False)
        self.SetTitle ( self.title )
       
    def atualiza_progress_dialog(self, e):
        u"rotina para mandar um pulso para dialogo de progresso"
        self.progress.pulse(self.msg)
       
    def recebe_mensagem_thread(self, msg):
       u"rotina para receber mensagem de processo em outra thread e entrar em contato com o dialogo de progresso"
       self.msg = msg
       wx.SafeYield(None, True)
       self.progress.pulse ( self.msg )
       self.SetStatusText( self.msg )

    def abre_dialog_thread(self, msg_err, msg_tit):
        u"metodo chamado por uma thread para abrir um dialogo na tela e encerrar o ProgressDialog"
        try:
            print "abre_dialog_thread"
            print msg_err
#           tentando assegurar que nao serao gerados novos eventos e que os ja foram gerados serão processados 
#           antes de  encerrar o ProgressDialog, de toda a forma, tem parametro no ProgressDialog para evitar
#           este problema
            self.timer.Stop()
#           wx.SafeYield(None, True)
            self.progress.destroy()
            dlg2 = wx.MessageDialog(self, msg_err, msg_tit, wx.OK | wx.ICON_EXCLAMATION) 
            dlg2.ShowModal()
            dlg2.Destroy()
        except:
            print "Erro 1"
        return

    def carrega_alunos_no_grid(self, dict_alunos_planilha): 
        u"método para carregar a lista dos alunos no grid visível.  Este método é chamado a partir de uma thread"
        self.grid.carrega_dict_de_alunos ( dict_alunos_planilha )

    def timetoquit(self, event):
        u"encerra a execução da aplicação gráfica"
        if self.file_arquivo_recentes:
            self.file_arquivo_recentes.close()
        self.Close(True)


    def ajuda(self, event ):
        u"callback para chamar o método de ajuda"
        self.about(event)

    def about(self, event):
        u"call back para mostrar informações sobre o programa"
        description = \
 u"""SAPU, Sistema para Auxílio ao gerenciamento de Provas Unificadas da UFRJ.
 Este sistema carrega os dados dos alunos inscritos a partir das pautas das
 turmas disponíveis no SIGA.  Com os dados carregados, pode realizar diversas
 funções,como, por exemplo,\ndistribuir os alunos por sala, gerar cadernos
 de prova, gerar planilha para lançamento das notas, ler as planilhas para
 identificar alunos já aprovados, etc """ 
        licence = \
 u"""Este programa é um software livre; você pode redistribuí-lo e/ou
 modificá-lo de acordo com os termos da Licença Pública Geral do GNU,
 tal como publicada pela Free Software Foundation; versão 3 da licença,
 ou ainda (à sua escolha) qualquer versão posterior.
 Este programa é distribuído na esperança de que ele seja útil,mas
 SEM NENHUMA GARANTIA; sem mesmo a garantia implicita de COMERCIALIZAÇÃO
 ou de ADAPTAÇÃO A UM OBJETO PARTICULAR.  Para mais detalhes, consulte a
 Licença Pública Geral do GNU.( http://www.gnu.org/licenses/licenses.html )"""


        info = wx.AboutDialogInfo()
        info.SetName(u'SAPU')
        info.SetVersion(u'1.0')
        info.SetDescription(description)
        info.SetCopyright(u'(C) 2008 Ivo Lopez')
#       info.SetWebSite(u'http://www.zetcode.com')
        info.SetLicence(licence)
#       info.AddDeveloper(u'jan bodnar')
#       info.AddDocWriter(u'jan bodnar')
#       info.AddArtist(u'The Tango crew')
#       info.AddTranslator(u'jan bodnar')
        wx.AboutBox(info)


class ProgressDialog:
    def __init__(self, gui, title, mensagem):    # gui is a wxPanel instance
        self.gui = gui
        self.mensagem = mensagem
        self.progress = wx.ProgressDialog(title, mensagem,
            maximum = 5, parent=gui, 
            style =  wx.PD_APP_MODAL | wx.PD_ELAPSED_TIME)
 
    def pulse(self, msg):
        if self.progress:
            self.progress.Pulse(self.mensagem + "\n" + msg )
            self.progress.Refresh()
        else:
            print "pulse chamado depois de estar fechado"
 
    def destroy(self):
        self.progress.Destroy()






class DialogoProva ( wx.Dialog ):
    "classe para obter os dados necessarios para gerar os cabeçalhos de prova"
#   dir_corrente = os.path.abspath(os.path.dirname(sys.argv[0]))
    dir_corrente = '.'
    mes_por_extenso = [
        u"janeiro", u"fevereiro", u"março", u"abril", u"maio", u"junho",
        u"julho", u"agosto", u"setembro", u"outubro", u"novembro", u"dezembro"]
    dic_disciplinas_usuais = {
        u"Calculo I":"C\\'alculo I", u"Calculo II":"C\\'alculo II", 
        u"Calculo III":"C\\'alculo III", u"Calculo IV":"C\\'alculo IV",
        u"EDP I":"EDP I", u"EDP II":"EDP II",
        u"Algebra Linear":"\\'Algebra Linear",
        u"Analise Real":"An\\'alise Real",
        u"Analise Complexa":"An\\'alise Complexa",
        u"Calculo Diferencial e Integral I":
            "C\\'alculo Diferencial e Integral I",
        u"Calculo Diferencial e Integral II":
            "C\\'alculo Diferencial e Integral II",
       u"Calculo Diferencial e Integral III":
            "C\\'alculo Diferencial e Integral III",
       u"Calculo Diferencial e Integral IV":
            "C\\'alculo Diferencial e Integral IV",
       u"Equacoes Diferenciais Parciais I":
            "Equa\\c{c}\\~oes Diferenciais Parciais I", 
       u"Equacoes Diferenciais Parciais II":
            "Equa\\c{c}\\~oes Diferenciais Parciais II",
       u"Calculo de uma Variavel I":"C\\'alculo de uma Vari\\'avel I",
       u"Calculo de uma Variavel II":"C\\'alculo de uma Vari\\'avel II",
       u"Calculo de varias Variaveis I":
            "C\\'alculo de v\\'arias Vari\\'aveis I",
       u"Calculo de varias Variaveis II":
            "C\\'alculo de v\\'arias Vari\\'aveis II",
       u"Probabilidade e Estatística":
            "Probablidade e Estat\\'{\\i}stica"}
    nomes_disciplinas_usuais = dic_disciplinas_usuais.keys()
    nomes_disciplinas_usuais.sort()
    numero_de_questoes = ["1", u"2", u"3", u"4", u"5", u"6", u"7"]
    provas_usuais = ["P1", u"P2", u"P3", u"PF", u"2CH", u"Primeira Prova",
                     u"Segunda Prova", u"Terceira Prova", u"Prova Final",
                     u"Prova de Segunda Chamada"]
    nome_arq_modelo_prova = "modelo_caderno_prova.tex"
    nome_arq_caderno_prova = [ u"lista_cadernos_prova_p1.tex",
          u"lista_cadernos_prova_p2.tex", u"lista_cadernos_prova_p3.tex",
          u"lista_cadernos_prova_pf.tex", u"lista_cadernos_prova_p2ch.tex",
          u"lista_cadernos_prova_p1.tex", u"lista_cadernos_prova_p2.tex",
          u"lista_cadernos_prova_p3.tex", u"lista_cadernos_prova_pf.tex",
          u"lista_cadernos_prova_p2ch.tex", u"lista_cadernos_prova.tex" ]
    nome_arq_salas_usadas = [ u"lista_salas_usadas_p1.tex",
          u"lista_salas_usadas_p2.tex", u"lista_salas_usadas_p3.tex",
          u"lista_salas_usadas_pf.tex", u"lista_salas_usadas_p2ch.tex",
          u"lista_salas_usadas_p1.tex", u"lista_salas_usadas_p2.tex",
          u"lista_salas_usadas_p3.tex", u"lista_salas_usadas_pf.tex",
          u"lista_salas_usadas_p2ch.tex", u"lista_salas_usadas.tex" ]
    nome_arq_lista_alunos = [ u"lista_alunos_p1.tex", u"lista_alunos_p2.tex",
         u"lista_alunos_p3.tex", u"lista_alunos_pf.tex", 
         u"lista_alunos_p2ch.tex", u"lista_alunos_p1.tex",
         u"lista_alunos_p2.tex", u"lista_alunos_p3.tex", 
         u"lista_alunos_pf.tex", u"lista_alunos_p2ch.tex", u"lista_alunos.tex" ]
    nome_arq_irregular = [ u"lista_alunos_irreg_p1.tex",
         u"lista_alunos_irreg_p2.tex", u"lista_alunos_irreg_p3.tex",
         u"lista_alunos_irreg_pf.tex", u"lista_alunos_irreg_p2ch.tex",
         u"lista_alunos_irreg_p1.tex", u"lista_alunos_irreg_p2.tex",
         u"lista_alunos_irreg_p3.tex", u"lista_alunos_irreg_pf.tex",
         u"lista_alunos_irreg_p2ch.tex", u"lista_alunos_irreg.tex" ]
    nome_arq_presenca = [ u"lista_presenca_p1.tex", u"lista_presenca_p2.tex",
       u"lista_presenca_p3.tex", u"lista_presenca_pf.tex",
       u"lista_presenca_p2ch.tex", u"lista_presenca_p1.tex",
       u"lista_presenca_p2.tex", u"lista_presenca_p3.tex", 
       u"lista_presenca_pf.tex", u"lista_presenca_p2ch.tex", 
       u"lista_presenca.tex" ]
    nome_arq_saida_lista = u"lista_alunos_p1.tex"
#   salas disponiveis, lotacao e prioridade de cada sala
#   Prioridades: 2 (ótima), 3 (salas da EQ primeiro andar), 4(quente), 
#                5(sem uso), 6(anfiteatro), 7(muito baixa lotação),
#                8(curso noturno), 9(curso noturno anfiteatro) 
    salas_usuais = {
        u"A201":[60, 2],
        u"A202":[60, 2],
        u"A203":[60, 2],
        u"A205":[60, 2], 
        u"A206":[60, 2],
        u"B110":[50, 8], 
        u"C205b":[65, 2],
        u"C207a":[65, 2],
        u"D105":[60, 2],
        u"D107":[35, 7],
        u"D109":[35, 7],
        u"D111":[35, 7],
        u"D112":[55, 6],
        u"D114":[45, 6],
        u"D116":[60, 6],
        u"D120":[60, 2],
        u"D213":[50, 8], 
        u"D214":[80, 9], 
        u"D215":[50, 8], 
        u"D217":[80, 8], 
        u"D218":[80, 9], 
        u"D219":[70, 8], 
        u"E118":[55, 3],
        u"E120":[65, 3],
        u"E124":[60, 3],
        u"E126":[60, 3],
        u"E129":[60, 3],
        u"E131":[60, 3],
        u"E216":[90, 2],
        u"E217":[60, 8],
        u"E220":[90, 2],
        u"E221":[60, 8],
        u"F112":[60, 6],
        u"F116":[45, 7],
        u"F119":[50, 7],
        u"F222":[60, 6],
        u"F223":[60, 2],
        u"F227":[65, 2], 
        u"G212":[64, 6],
        u"G213":[55, 4],
        u"G215":[55, 4],
        u"G217":[55, 4],
        u"G220":[64, 6],
        u"H204b":[95, 6],
        u"H208":[55, 6],
        u"H209":[60, 6],
        u"H211":[50, 7],
        u"H213":[60, 6],
        u"H214":[70, 6],
        u"H222":[55, 2],
        u"H224":[60, 2],
        u"H226":[60, 2],
        u"H228":[60, 2],
        u"H230b":[50, 7], 
        u"F3-01(CCMN)":[125, 3], 
        u"F3-02(CCMN)":[90, 3], 
        u"F3-03(CCMN)":[70, 3], 
        u"F3-04(CCMN)":[60, 3], 
        u"F3-05(CCMN)":[60, 3], 
        u"F3-06(CCMN)":[60, 3], 
        u"F3-07(CCMN)":[140, 3], 
        u"F3-08(CCMN)":[50, 3], 
        u"F3-09(CCMN)":[80, 3], 
        u"F3-10(CCMN)":[60, 3], 
        u"F3-11(CCMN)":[70, 3], 
        u"F3-12(CCMN)":[90, 3], 
        u"F3-13(CCMN)":[60, 3], 
        u"F3-14(CCMN)":[90, 3], 
        u"F3-15(CCMN)":[100, 3], 
        u"F3-16(CCMN)":[60, 3], 
		} 
    text_lotacao = [ ]


    def __init__(self, alunos_efetivos):

        self.alunos_efetivos = alunos_efetivos
        wx.Dialog.__init__(self, None, -1, u"Geração dos cadernos de prova",
                           style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
#     parent = wx.Panel(self, -1, style = wx.TAB_TRAVERSAL |
#                       wx.CLIP_CHILDREN | wx.FULL_REPAINT_ON_RESIZE)
        parent = self
          

        self.flag_disciplina = False
        self.flag_dir_modelo = False

        # row 1
        fgs = wx.FlexGridSizer ( 6, 2, 12, 12 )

        fgs.Add(wx.StaticText(parent, -1, u"Disciplina"), flag=wx.EXPAND )
        # This combobox is created with a preset list of values.
        self.nome_disc = wx.ComboBox ( parent, 500, u"Nome da disciplina",
           (-1, -1), (-1, -1), self.nomes_disciplinas_usuais, wx.CB_DROPDOWN ) 
        fgs.Add(self.nome_disc, flag=wx.EXPAND )
#     self.Bind(wx.EVT_COMBOBOX, self.EvtComboBox, self.nome_disc)
        self.Bind(wx.EVT_TEXT, self.evttext, self.nome_disc)
#     self.Bind(wx.EVT_TEXT_ENTER, self.evttextenter, self.nome_disc)

#      # row 2
        fgs.Add(wx.StaticText(parent, -1, u"Data da prova"), flag=wx.EXPAND )
        self.data_prova = wx.DatePickerCtrl(parent, size=(-1, -1),
                                            style=wx.DP_DEFAULT)
        fgs.Add(self.data_prova, flag=wx.EXPAND)
#
#     # row 3
        fgs.Add ( wx.StaticText(parent, -1, u"Prova"), flag=wx.EXPAND )
        self.prova = wx.ComboBox(parent, 500, u"P1", (-1, -1), (-1, -1),
            DialogoProva.provas_usuais, wx.CB_DROPDOWN 
            ) #| wx.TE_PROCESS_ENTER #| wx.CB_SORT)
        self.prova.SetHelpText(
            u"Selecione a prova que irá realizar.  Se for diferente das "+
            u"opções disponíveis, digite o texto que irá constar na prova ")
        fgs.Add ( self.prova, flag=wx.EXPAND ) 
#      
#     # row 4
        fgs.Add( wx.StaticText(parent, -1, u"Número de Questões"),
                flag=wx.EXPAND )
        self.num_quest_prova = wx.ComboBox(parent, 500, u"4", (-1, -1),
               (-1, -1), DialogoProva.numero_de_questoes, wx.CB_DROPDOWN
                                 ) #| wx.TE_PROCESS_ENTER #| wx.CB_SORT)
        self.num_quest_prova.SetEditable(False)
        self.num_quest_prova.SetHelpText(
            u"Selecione o número de questões da prova, para que o programa " +
            u" possa gerar cadernos com o número correto de páginas")
        fgs.Add ( self.num_quest_prova, flag=wx.EXPAND )
#      
#
#     # row 5
        fgs.Add ( wx.StaticText(parent, -1, u"Dir. modelos de cabeçalho "),
                 flag=wx.EXPAND )
        self.dir_modelo = wx.lib.filebrowsebutton.DirBrowseButton(parent,
            pos=(-1, -1),size=(-1, -1),  labelText="",
            toolTip=
             u"Escolha diretório que contém modelos de cabeçalho das provas",
             changeCallback=self.evt_dir_modelo,
             style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST
                      ) #| wx.DD_CHANGE_DIR
        self.dir_modelo.SetHelpText(
            u"Selecione diretório que contém modelo de caderno de provas \n"+
            u"(arquivo\n" + self.nome_arq_modelo_prova + 
            u" )\n"+ 
            u"Neste diretório serão gerados os arquivos latex de saída, \n"+
            u"por ex: "+ self.nome_arq_caderno_prova[0] + u", " +
            self.nome_arq_lista_alunos[0] + u", "+ self.nome_arq_irregular[0] +
            u", "+ self.nome_arq_presenca[0] + 
            u", "+ self.nome_arq_salas_usadas[0])
        fgs.Add ( self.dir_modelo, flag=wx.EXPAND )
        fgs.AddGrowableCol(1)
        fgs.SetFlexibleDirection(wx.BOTH)
#
#     # row 6
        fgs.Add ( wx.StaticText(parent, -1, 
             u"Prefixo dos arquivos lista (OPCIONAL)"), flag=wx.EXPAND )
        self.prefixo_lista = wx.TextCtrl(parent, -1, "", size=(-1, -1))
        self.prefixo_lista.SetHelpText(
          u"Selecione o prefixo que deverá ser usado para os nomes dos\n"+
          u"arquivos com os cadernos de prova, lista de presenca etc.\n"+
          u"Ao nome padrão será acrescentado o prefixo.  Esta opção é\n"+
          u"útil para incluir uma turma separada sem alterar os arquivos\n"+
          u"das turmas já geradas")
        self.prefixo_lista.Enable(True)
        fgs.Add ( self.prefixo_lista, flag=wx.EXPAND )
        fgs.AddGrowableCol(1)
        fgs.SetFlexibleDirection(wx.BOTH)
#

#       ordenando, dando prioridade ao campo prioridade, depois à capacidade
#       da sala, depois à posição da sala
        lista_salas = self.retorna_salas_ordenadas()
        alunos_na_sala = len(DialogoProva.salas_usuais)*[0]
        alunos_que_restam = alunos_efetivos
        for (iaux, sala) in enumerate(lista_salas):
            alunos_na_sala[iaux] = min ( 
                DialogoProva.salas_usuais[sala][0], alunos_que_restam )
            alunos_que_restam -= alunos_na_sala[iaux]
            if alunos_que_restam <= 0:
                break
                   
         
        panel1 = wx.lib.scrolledpanel.ScrolledPanel(parent, -1,
                 style = wx.TAB_TRAVERSAL|wx.SUNKEN_BORDER, name="panel1" )
        fgs1 = wx.FlexGridSizer(cols=6, vgap=12, hgap=12)

        
        fgs1.Add ( (10, 10) )
        label = wx.StaticText(panel1, -1, u"SALA", size=(-1, -1))
        fgs1.Add(label,
             flag=wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL |
                  wx.RIGHT, border=10)
        label = wx.StaticText(panel1, -1, u"Lotação", size=(-1, -1))
        fgs1.Add(label, flag=wx.ALIGN_RIGHT |
                 wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
        label = wx.StaticText(panel1, -1, u"Prioridade", size=(-1, -1))
        fgs1.Add(label, flag=wx.ALIGN_RIGHT |
                 wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
        label = wx.StaticText(panel1, -1, u"Alunos", size=(-1, -1))
        fgs1.Add(label, flag=wx.ALIGN_RIGHT |
                 wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
        fgs1.Add ( (10, 10) )
        salas_usadas = 0
        self.txtc = []
        for (ipos, sala) in enumerate(lista_salas):
            fgs1.Add ( (10, 10) )
            label = wx.StaticText(panel1, -1, sala, size=(-1, -1))
            fgs1.Add(label, flag=wx.ALIGN_RIGHT |
                     wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
            label = wx.StaticText(panel1, -1,
                     unicode(str(DialogoProva.salas_usuais[sala][0]),
                     "utf-8"), size=(-1, -1))
            fgs1.Add(label, flag=wx.ALIGN_RIGHT |
                     wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
            label = wx.StaticText(panel1, -1,
                   unicode(str(DialogoProva.salas_usuais[sala][1]),
                   "utf-8"), size=(-1, -1))
            fgs1.Add(label, flag=wx.ALIGN_RIGHT |
                     wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
            alunos_nesta_sala = alunos_na_sala[ipos]
            self.txtc.append(wx.TextCtrl(panel1, -1,
                     unicode(str(alunos_nesta_sala), "utf-8"), size=(-1, -1)))
            if ( alunos_nesta_sala > 0 ):
                salas_usadas += 1
            fgs1.Add(self.txtc[-1], flag=wx.RIGHT, border=10)
            fgs1.Add ( (10, 10) )
#     sala extra         
        fgs1.Add ( (10, 10) )
        self.sala_extra = wx.TextCtrl(panel1, -1, u" ", size=(-1, -1))
        fgs1.Add(self.sala_extra, flag=wx.ALIGN_RIGHT |
                 wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
        label = wx.StaticText(panel1, -1, u"", size=(-1, -1))
        fgs1.Add(label, flag=wx.ALIGN_RIGHT |
                 wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
        label = wx.StaticText(panel1, -1, u" ", size=(-1, -1))
        fgs1.Add(label, flag=wx.ALIGN_RIGHT |
                 wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
        self.alunos_sala_extra = wx.TextCtrl(panel1, -1, u"0", size=(-1, -1))
        fgs1.Add(self.alunos_sala_extra, flag=wx.RIGHT, border=10)
        fgs1.Add ( (10, 10) )

#       label = wx.StaticText(panel1, -1, u"Total:", size=(-1, -1))
#       fgs1.Add ( (10, 10) )
#       fgs1.Add(label, flag=wx.ALIGN_RIGHT |
#                wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
#       label = wx.StaticText(panel1, -1, u"", size=(-1, -1))
#       fgs1.Add(label, flag=wx.ALIGN_RIGHT |
#                wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
#       label = wx.StaticText(panel1, -1, u"", size=(-1, -1))
#       fgs1.Add(label, flag=wx.ALIGN_RIGHT |
#                wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
#       label = wx.StaticText(panel1, -1,
#            unicode(str(salas_usadas), "utf-8")+u" salas", size=(-1, -1))
#       fgs1.Add(label, flag=wx.ALIGN_RIGHT |
#                wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, border=10)
#       fgs1.Add ( (10, 10) )
        fgs1.AddGrowableCol(3)
        fgs1.AddGrowableCol(0)
        fgs1.Layout()
#     panel1.SetMinSize((-1, 400)
        panel1.SetMaxSize((-1, 400))
        panel1.SetSizerAndFit( fgs1 )
        panel1.SetAutoLayout(1)
        panel1.SetupScrolling()

        btnsizer = wx.StdDialogButtonSizer()
        if wx.Platform != "__WXMSW__":
            btn = wx.ContextHelpButton(parent)
            btnsizer.AddButton(btn)

        self.botao_ok = wx.Button(parent, wx.ID_OK)
        self.botao_ok.SetHelpText(
            u"Este botão inicia geração dos cadernos de prova "+
            u"e tabela com salas de prova")
        self.botao_ok.SetDefault()
        self.botao_ok.Enable(False)
        self.Bind(wx.EVT_BUTTON, self.evt_botao_ok, self.botao_ok)
        btnsizer.AddButton(self.botao_ok)

        botao_cancel = wx.Button(parent, wx.ID_CANCEL)
        botao_cancel.SetHelpText(
            "Este botão cancela a geração dos cadernos de prova")
        btnsizer.AddButton(botao_cancel)
        btnsizer.Realize()


        box_hor1 = wx.BoxSizer(wx.HORIZONTAL)
        box_hor1.Add ( (10, 10) )
        box_hor1.Add ( fgs, flag=wx.GROW|wx.ALL )
        box_hor1.Add ( (10, 10) )

        box_hor5 = wx.BoxSizer(wx.HORIZONTAL)
        box_hor5.Add ( (10, 10) )
        box_hor5.Add( wx.StaticText(parent, -1,
          u"Relação das salas a serem usadas pelos " +
          unicode(str(alunos_efetivos),
          "utf-8") + " alunos", (-1, -1)) )
        box_hor5.Add ( (10, 10) )

        box_hor6 = wx.BoxSizer(wx.HORIZONTAL)
        box_hor6.Add ( (10, 10) )
        box_hor6.Add(panel1, 0, wx.GROW|wx.ALL)
        box_hor6.Add ( (10, 10) )

        box_hor7 = wx.BoxSizer(wx.HORIZONTAL)
        box_hor7.Add ( (10, 10) )
        box_hor7.Add(btnsizer, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL)
        box_hor7.Add ( (10, 10) )

        box_vert = wx.BoxSizer(wx.VERTICAL)
        box_vert.Add ( (10, 10) )
        box_vert.Add(box_hor1, flag=wx.GROW|wx.ALL )
        box_vert.Add ( (10, 10) )
        box_vert.Add(box_hor5, flag=wx.GROW|wx.ALL )
        box_vert.Add ( (10, 10) )
        box_vert.Add(box_hor6, flag=wx.GROW|wx.ALL )
        box_vert.Add ( (10, 10) )
        box_vert.Add(box_hor7, flag=wx.GROW|wx.ALL )
        box_vert.Add ( (10, 10) )

        parent.SetSizerAndFit(box_vert)
        parent.SetAutoLayout(1)
        self.SetClientSize(parent.GetSize())



    # When the user selects something, we go here.
#  def EvtComboBox(self, evt):
#      cb = evt.GetEventObject()
#      data = cb.GetClientData(evt.GetSelection())
#      print 'EvtComboBox: %s\nClientData: %s\n' % (evt.GetString(), data)

#      if evt.GetString() == 'one':
#          print "You follow directions well!\n\n"

    # Capture events every time a user hits a key in the text entry field.
    def evttext(self, evt):
        u"captura eventos de edição de texto"
#       print 'evttext: %s\n' % evt.GetString()
        self.flag_disciplina = True
        if self.flag_dir_modelo:
            self.botao_ok.Enable(True)
        evt.Skip()

    # Capture events every time a user selects a directory
    def evt_dir_modelo(self, evt):
        u"""callback ativada quando seleciona um diretório""" 
        print 'evt_dir_modelo: %s\n' % evt.GetString()
        self.flag_dir_modelo = True
        if self.flag_disciplina:
            self.botao_ok.Enable(True)

    # Capture events when the user types something into the control then
    # hits ENTER.
#   def evttextenter(self, evt):
#       u"trata os eventos de entrada de texto"
#       print  'evttextenter: %s' % evt.GetString() 
#       evt.Skip()

    def evt_botao_ok(self, evt):
        """verifica se o numero de alunos nas salas coincide com o numero de
        alunos para fazer prova"""
        dic_salas = self.retorna_alunos_por_sala()
        alunos_por_sala =  dic_salas.values()
        alunos_com_sala = sum ( alunos_por_sala )
        salas_usadas = sum( map ( lambda x:cmp(x,0), alunos_por_sala ) )
        if ( alunos_com_sala > self.alunos_efetivos ): 
            dlg = wx.MessageDialog(self, u"'O numero de vagas nas salas ( "+
               unicode(str(alunos_com_sala), "utf-8") +
               u" ) está maior que o número de alunos efetivos ( " +
               unicode(str(self.alunos_efetivos), "utf-8") +
               u" ).\nA diferença é de " +
               unicode(str(alunos_com_sala-self.alunos_efetivos), "utf-8") +
               u" alunos.\nO número de salas que está solicitando é de " +
               unicode(str(salas_usadas), "utf-8") +
               u" salas.\nPor favor, ajuste as lotações das turmas", 
               u'Erro no total de vagas nas turmas',
               wx.OK | wx.ICON_INFORMATION)
            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()
        elif ( alunos_com_sala < self.alunos_efetivos ): 
            dlg = wx.MessageDialog(self, u"'O numero de vagas nas salas ( "+
               unicode(str(alunos_com_sala), "utf-8") +
               u" ) está menor que o número de alunos efetivos ( " +
               unicode(str(self.alunos_efetivos),
               "utf-8") + u" ).\nA diferença é de " +
               unicode(str(self.alunos_efetivos-alunos_com_sala), "utf-8") +
               u" alunos.\nO número de salas que está solicitando é de " +
               unicode(str(salas_usadas), "utf-8") +
               u" salas.\nPor favor, ajuste as lotações das turmas", 
               u'Erro no total de vagas nas turmas',
               wx.OK | wx.ICON_INFORMATION )
            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()
        else:
#         identificando o arquivo do modelo da prova
            dir_modelo = self.dir_modelo.GetValue()
            if ( not os.path.isfile(dir_modelo+"/"+self.nome_arq_modelo_prova) ):
                dlg = wx.MessageDialog(self, 
                 u'Erro: O arquivo com o modelo do caderno de provas não foi'+
                 u'encontrado  O nome procurado foi ' +
                 self.nome_arq_modelo_prova + 
                 u'.\n Por favor, coloque o arquivo no diretório selecionado.',
                 u'Erro no nome do diretório que contém modelos de '+
                 u'caderno de provas', wx.OK | wx.ICON_INFORMATION
                 ) #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION)
                dlg.ShowModal()
                dlg.Destroy()
#          tudo está OK                
            else:
                evt.Skip()



    def retorna_dados ( self ):
        u"""método que retorna as informações coletadas no diálogo aberto
        para gerar o cabeçalho de prova"""
        disciplina =  DialogoProva.dic_disciplinas_usuais[
                                              self.nome_disc.GetValue()]
        data = self.data_prova.GetValue()
        dia = data.GetDay()
        mes = DialogoProva.mes_por_extenso[data.GetMonth()]
        ano = data.GetYear()
        prova = self.prova.GetValue()
        try:
            iaux = self.provas_usuais.index(prova)
            nome_arq_caderno_prova = self.nome_arq_caderno_prova[iaux]
            nome_arq_lista_alunos = self.nome_arq_lista_alunos[iaux]
            nome_arq_irregular = self.nome_arq_irregular[iaux]
            nome_arq_presenca = self.nome_arq_presenca[iaux]
            nome_arq_salas_usadas = self.nome_arq_salas_usadas[iaux]
        except ValueError:
            nome_arq_caderno_prova = self.nome_arq_caderno_prova[-1]
            nome_arq_lista_alunos = self.nome_arq_lista_alunos[-1]
            nome_arq_irregular = self.nome_arq_irregular[-1]
            nome_arq_presenca = self.nome_arq_presenca[-1]
            nome_arq_salas_usadas = self.nome_arq_salas_usadas[-1]
        prefixo_arquivos_lista = self.prefixo_lista.GetValue()
        if prefixo_arquivos_lista != "":
            nome_arq_caderno_prova = prefixo_arquivos_lista + u"_" +\
                                     nome_arq_caderno_prova 
            nome_arq_lista_alunos = prefixo_arquivos_lista + u"_" +\
                                     nome_arq_lista_alunos 
            nome_arq_irregular = prefixo_arquivos_lista + u"_" +\
                                     nome_arq_irregular 
            nome_arq_presenca = prefixo_arquivos_lista + u"_" +\
                                     nome_arq_presenca 
            nome_arq_salas_usadas = prefixo_arquivos_lista + u"_" +\
                                     nome_arq_salas_usadas 
        num_questoes = self.num_quest_prova.GetValue()
        dir_modelo = self.dir_modelo.GetValue()
        dic_salas = self.retorna_alunos_por_sala()
        print dic_salas 
        return [ disciplina, dia, mes, ano, prova,
                dir_modelo+"/"+self.nome_arq_modelo_prova,
                dir_modelo+"/"+nome_arq_caderno_prova,
                dir_modelo+"/"+nome_arq_lista_alunos,
                dir_modelo+"/"+nome_arq_irregular,
                dir_modelo+"/"+nome_arq_presenca, 
                dir_modelo+"/"+nome_arq_salas_usadas, 
                int(num_questoes), dic_salas ] 
#---------------------------------------------------------------------------

    def retorna_alunos_por_sala (self ):
        u"""rotina auxiliar para retornar um dicionario com as salas e o 
        número de alunos alocadas"""
        lista_salas = self.retorna_salas_ordenadas()
        dic_salas = {}
        for (icont, sala) in enumerate(lista_salas):
            try:
                alunos_sala = int(self.txtc[icont].GetValue())
            except ValueError:
                alunos_sala = 0
            if alunos_sala > 0:
                dic_salas[sala] =  alunos_sala
        alunos_sala_extra = int(self.alunos_sala_extra.GetValue())
        if ( alunos_sala_extra > 0 ):
            sala_extra = self.sala_extra.GetValue()
            dic_salas[sala_extra] = alunos_sala_extra
     
        return dic_salas 

    def retorna_salas_ordenadas ( self ):
        """Método para retornar uma lista com as salas do CT ordenadar de forma
           que as salas com mais prioridade fiquem à frente, depois as com maior
           lotação e, finalmente, em ordem alfabética"""
        lista_salas = sorted ( DialogoProva.salas_usuais.keys(), 
                key = lambda sala:
#               ( DialogoProva.salas_usuais[sala][1] * 400 -
#                 DialogoProva.salas_usuais[sala][0] ) * 100000 +
                 ( ord(sala[0]) - ord('A') ) * 1000 +
                 ( ord(sala[1]) - ord('0') ) * 100 +
                 ( ord(sala[2]) - ord('0') ) * 10 +
                   ord(sala[3]) - ord('0') ) 
        return lista_salas

class DialogoPlanilha ( wx.Dialog ):
    """classe para obter o diretorio com o template para a planilha e para
       escolher o arquivo para salvar a planilha"""
    import datetime
#  dir_corrente = os.path.abspath(os.path.dirname(sys.argv[0]))
    dir_corrente = '.'
    dic_disciplinas_usuais = {"Calculo I":"calc1", u"Calculo II":"calc2",
        u"Calculo III":"calc3", u"Calculo IV":"calc4", u"EDP I":"edp1",
        u"EDP II":"edp2", u"Algebra Linear":"alin", 
        u"Analise Complexa":"AComplexa", u"Analise Real":"AReal", 
        u"Calculo Diferencial e Integral I":"calc1", 
        u"Calculo Diferencial e Integral II":"calc2", 
        u"Calculo Diferencial e Integral III":"calc3", 
        u"Calculo Diferencial e Integral IV":"calc4", 
        u"Equacoes Diferenciais Parciais I":"edp1",
        u"Equacoes Diferenciais Parciais II":"edp2", 
        u"Calculo de uma Variavel I":"c1v1",
        u"Calculo de uma Variavel II":"c1v2",
        u"Calculo de varias Variaveis I":"cvv1",
        u"Calculo de varias Variaveis II":"cvv2",
        u"Probabilidade e Estatística":"probest"}
    nomes_disciplinas_usuais = dic_disciplinas_usuais.keys()
    nomes_disciplinas_usuais.sort()
    dic_nomes_periodos = {u"Primeiro periodo":"1",
                          u"Segundo periodo":"2", u"Terceiro periodo":"3"}
    nomes_periodos = dic_nomes_periodos.keys()
    nome_arq_template_notas = "template_notas_alunos.odt"


    def __init__(self):
        wx.Dialog.__init__(self, None, -1, u"Geração da planilha de notas",
                   style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
#     parent = wx.Panel(self, -1, style = wx.TAB_TRAVERSAL | 
#              wx.CLIP_CHILDREN | wx.FULL_REPAINT_ON_RESIZE)
        parent = self
          

        self.flag_disciplina = False
        self.flag_planilha_template = False
        self.flag_planilha_saida = False

        fgs = wx.FlexGridSizer ( 4, 2, 12, 12 )

        # row 1
        fgs.Add(wx.StaticText(parent, -1, u"Disciplina"), flag=wx.EXPAND )
        # This combobox is created with a preset list of values.
        self.nome_disc = wx.ComboBox ( parent, 500, u"Nome da disciplina",
          (-1, -1), (-1, -1), self.nomes_disciplinas_usuais, wx.CB_DROPDOWN ) 
        fgs.Add(self.nome_disc, flag=wx.EXPAND )
        self.Bind(wx.EVT_TEXT, self.evttext, self.nome_disc)
#
        # row 2
        fgs.Add(wx.StaticText(parent, -1, u"Período"), flag=wx.EXPAND )
        # This combobox is created with a preset list of values.
        self.periodo = wx.ComboBox ( parent, 500, self.nomes_periodos[0],
              (-1, -1), (-1, -1), self.nomes_periodos, wx.CB_DROPDOWN ) 
        fgs.Add(self.periodo, flag=wx.EXPAND )
        self.Bind(wx.EVT_TEXT, self.evttext1, self.periodo)

#     # row 3
        fgs.Add ( wx.StaticText(parent, -1,
           u"Template para geração da planilha"), flag=wx.EXPAND )
        self.planilha_template = wx.lib.filebrowsebutton.FileBrowseButton(
            parent, pos=(-1, -1), size=(-1, -1),  labelText="", 
            toolTip= \
                 "Escolha o arquivo no formato de template do OpenOffice Calc"+
                 u" ( ots ) que será usado como base da planilha de notas",
            changeCallback=self.evt_planilha_template, fileMode=wx.OPEN,
            fileMask="*.ots", style=wx.FD_DEFAULT_STYLE|wx.CHANGE_DIR ) 
        self.planilha_template.SetHelpText(
           u"Selecione arquivo que será o template para geração da planilha"+
           u"de notas.  Deverá conter 1500 linhas de alunos, 150 linhas de"+
           u"alunos por turma e outras regras, para que seja um modelo válido")
        fgs.Add ( self.planilha_template, flag=wx.EXPAND )
        fgs.AddGrowableCol(1)
        fgs.SetFlexibleDirection(wx.BOTH)
#
#     # row 4
        fgs.Add ( wx.StaticText(parent, -1, u"Planilha a ser gerada"),
                 flag=wx.EXPAND )
        self.planilha_saida = wx.lib.filebrowsebutton.FileBrowseButton(parent, 
              pos=(-1, -1), size=(-1, -1), startDirectory='.',  labelText="", 
              toolTip="Escolha o nome do arquivo com a planilha a ser gerada",
              changeCallback=self.evt_planilha_saida,  fileMode=wx.SAVE,
              fileMask="*.ods", style=wx.FD_DEFAULT_STYLE|wx.CHANGE_DIR)
        self.planilha_saida.SetHelpText(
          u"Selecione arquivo que será a planilha para lançamento das notas."+
          u" A planilha está no formato OpenOffice e deverá conter todas as"+
          u" regras para cálculo das médias, "+
          u"dependendo do modelo em que se baseou")
        self.planilha_saida.Enable(False)
        fgs.Add ( self.planilha_saida, flag=wx.EXPAND )
        fgs.AddGrowableCol(1)
        fgs.SetFlexibleDirection(wx.BOTH)
#

        btnsizer = wx.StdDialogButtonSizer()
        if wx.Platform != "__WXMSW__":
            btn = wx.ContextHelpButton(parent)
            btnsizer.AddButton(btn)

        self.botao_ok = wx.Button(parent, wx.ID_OK)
        self.botao_ok.SetHelpText(
         u"Este botão inicia geração da planilha para lançamento das notas")
        self.botao_ok.SetDefault()
        self.botao_ok.Enable(False)
#     self.Bind(wx.EVT_BUTTON, self.evt_botao_ok, self.botao_ok)
        btnsizer.AddButton(self.botao_ok)

        botao_cancel = wx.Button(parent, wx.ID_CANCEL)
        botao_cancel.SetHelpText(
        u"Este botão cancela geração da planilha para lançamento das notas")
        btnsizer.AddButton(botao_cancel)
        btnsizer.Realize()


        box_hor1 = wx.BoxSizer(wx.HORIZONTAL)
        box_hor1.Add ( (10, 10) )
        box_hor1.Add ( fgs, flag=wx.GROW|wx.ALL )
        box_hor1.Add ( (10, 10) )

        box_hor2 = wx.BoxSizer(wx.HORIZONTAL)
        box_hor2.Add ( (10, 10) )
        box_hor2.Add(btnsizer, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL)
        box_hor2.Add ( (10, 10) )

        box_vert = wx.BoxSizer(wx.VERTICAL)
        box_vert.Add ( (10, 10) )
        box_vert.Add(box_hor1, flag=wx.GROW|wx.ALL )
        box_vert.Add ( (10, 10) )
        box_vert.Add(box_hor2, flag=wx.GROW|wx.ALL )

        parent.SetSizerAndFit(box_vert)
        parent.SetAutoLayout(1)
        self.SetClientSize(parent.GetSize())
        self.nome_default = False

     # When the user selects something, we go here.
#  def EvtComboBox(self, evt):
#      cb = evt.GetEventObject()
#      data = cb.GetClientData(evt.GetSelection())
#      print 'EvtComboBox: %s\nClientData: %s\n' % (evt.GetString(), data)
#      if evt.GetString() == 'one':
#          print "You follow directions well!\n\n"


    def nome_unico ( self, arq_orig, sufixo ):
        u"""Método para alterar o nome de um arquivo caso já exista,
        retornado o nome criado"""
        print os.path.isfile(arq_orig)
        if ( os.path.isfile(arq_orig) ):
#         arquivo existe
            iaux = 1
            comp = len(sufixo)
            while os.path.isfile(arq_orig[:-comp]+"("+str(iaux)+")"+sufixo):
                iaux += 1 
            print "novo nome: "+ arq_orig[:-comp]+"("+str(iaux)+")"+sufixo
            return arq_orig[:-comp]+"("+str(iaux)+")"+sufixo
        else:
            return arq_orig

    def nome_planilha ( self ):
        u"""Método para criar o nome padrão da planilha, armazenando-o em
        self.nome_default e atualizando a tela"""
        import datetime
        disciplina = self.nome_disc.GetValue()
        if disciplina in DialogoPlanilha.dic_disciplinas_usuais:
            abrev_disc =  DialogoPlanilha.dic_disciplinas_usuais[disciplina]
        else:    
            abrev_disc =  "disc"
        periodo = self.periodo.GetValue()
        abrev_periodo = DialogoPlanilha.dic_nomes_periodos[periodo]
        sufixo = ".ods"
        arq_orig = "notas_" + abrev_disc +  "_" + "%4.4d" \
           % ( (datetime.date.today() ).year) + "_" + abrev_periodo + sufixo
        self.nome_default = self.nome_unico ( arq_orig, sufixo )
        self.planilha_saida.SetValue(self.nome_default)
        return self.nome_default


    def evttext(self, evt):
        u"callback para tratar uma edição de texto"
#       print 'evttext: %s\n' % evt.GetString()
        self.flag_disciplina = True
        self.planilha_saida.Enable(True)
        if ( self.flag_disciplina and
            self.flag_planilha_template and self.flag_planilha_saida ):
            self.botao_ok.Enable(True)
        evt.Skip()

    def evttext1(self, evt):
        u"""callback para capturar o nome da planilha"""
#       print 'evttext: %s\n' % evt.GetString()
        self.nome_planilha()
        evt.Skip()

    def evt_planilha_template(self, evt):
        u"""callback para ativar o botão de OK caso tenha sido selecionada o 
        modelo para a palnilha"""
#       print 'evt_planilha_template: %s\n' % evt.GetString()
        self.flag_planilha_template = True
        if ( self.flag_disciplina and
            self.flag_planilha_template and self.flag_planilha_saida ):
            self.botao_ok.Enable(True)
        evt.Skip()

    def evt_planilha_saida(self, evt):
        u"""callback para tratar o evento de edição do nome da planilha que
        será gravada"""
#       print 'evt_planilha_saida: %s\n' % evt.GetString()
        self.flag_planilha_saida = True
        if ( self.flag_disciplina and
            self.flag_planilha_template and self.flag_planilha_saida ):
            self.botao_ok.Enable(True)
        nome_arq = self.planilha_saida.GetValue()
        self.nome_default = self.nome_unico ( nome_arq, ".ods" )
        print "nome_arq: " +  nome_arq + " nome_default: " + self.nome_default
        if ( nome_arq !=  self.nome_default ):
            dlg = wx.MessageDialog(self, 
             u"'Para evitar gravações indevidas, a planilha gerada nunca é"+
             u" gravada sobre planilha pré-existente. No entanto o arquivo "+
             nome_arq+u" já existe.  Por isto o nome foi alterado para "+
             self.nome_default, u'Alteração no nome da planilha', wx.OK |
             wx.ICON_INFORMATION 
             ) #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()
            self.planilha_saida.SetValue(self.nome_default)
        evt.Skip()

    def retorna_dados ( self ):
        u"""Método para retornar os dados selecionados no diálogo de 
        geração da planilha"""
        return [self.nome_default, self.planilha_template.GetValue(),
                self.planilha_saida.GetValue()]
     
#---------------------------------------------------------------------------


class TableGrid(wx.grid.Grid): 
    u"""Classe para armazenar e processar os dados da tabela visível"""
    labels = [u'Numero', u'Nome',  u'DRE', u'Turma', u'Observacao', u'Status' ]
    larg_col = [ 100, 400, 100, 100, 200, 100 ]  
    cor_default = wx.BLACK
    cor_editado = wx.RED
    def __init__(self, parent):
        wx.grid.Grid.__init__(self, parent, -1)
        self.parent = parent
        self.datatypes = [wx.grid.GRID_VALUE_NUMBER, wx.grid.GRID_VALUE_STRING,
           wx.grid.GRID_VALUE_STRING, wx.grid.GRID_VALUE_STRING,
           wx.grid.GRID_VALUE_CHOICE + 
           u':OK,MAIS DE 32 CREDITOS,MENOS DE 6 CREDITOS,MATRICULA TRANCADA,'+
           u'INSCRICAO TRANCADA,TRANSFERIDO DE TURMA,APROVADO,REPROVADO,'+
           u'NÃO INSCRITO,REPROVADO POR MÉDIA,REPROVADO POR FALTA,'+
           u'REPROVADO POR FALTA E MÉDIA,--,OUTROS',
          wx.grid.GRID_VALUE_CHOICE + ':S,' +'NAO' ]
#          u'REPROVADO POR FALTA E MÉDIA, OUTROS', wx.grid.GRID_VALUE_BOOL ]
        self.data = []
        for iaux1 in range(0, 100):
            self.data.append( [ iaux1 + 1, '', '', '', '', 'NAO'] )
        self.tabela = DadosTabela ( parent, TableGrid.labels, 
                                    self.datatypes, self.data )
        self.SetTable(self.tabela, True) 
        self.ordem_coluna = []
        self.linhas_alteradas = []
        for iaux1 in range(0, len(TableGrid.labels)):
            self.SetColSize(iaux1, TableGrid.larg_col[iaux1])
            self.ordem_coluna.append(False)
        self.SetRowLabelSize(0)
        self.SetMargins(0, 0)
#       self.AutoSizeColumns(False)
        self.Bind(wx.grid.EVT_GRID_LABEL_LEFT_CLICK, self.OnLabelLeftClick)
        self.Bind(wx.grid.EVT_GRID_CELL_CHANGE, self.OnCellChange)

    def OnLabelLeftClick(self, evt):
        u"""callback quando seleciona o label com o botão esquerdo do mouse"""
#       print "OnLabelLeftClick: (%d,%d) %s\n" %(evt.GetRow(),\
#             evt.GetCol(), evt.GetPosition())
        col = evt.GetCol()
        self.ordena_grid(col)

    def ordena_grid_por_numero ( self ):
        u"""rotina para ordenar os alunos por número"""
        self.ordem_coluna[0] = False
        return self.ordena_grid( 0 )

    def ordena_grid ( self, col ):
        u"""reordena o grid com a lista de alunos ordenada pela coluna col e
        retorna esta lista"""
        lista_de_alunos = self.retorna_lista_de_alunos ()
        if col >= 0:
            lista_de_alunos.sort((lambda x, y:cmp(x[col], y[col])), None,
                           self.ordem_coluna[col] )
            self.ordem_coluna[col] = not self.ordem_coluna[col]
        iaux1 = 0
        dict_aux = dict()
        for numero, nome, dre, turma, obs, status in lista_de_alunos:
            self.grava_linha ( iaux1, numero, nome, dre, turma, obs, status )
            dict_aux[numero]=iaux1
            iaux1 += 1
        for iaux2 in range(iaux1, self.tabela.GetNumberRows()):
            self.grava_linha ( iaux2, iaux2+1, '', '', '', '', 'NAO' )
        for numero, coluna in self.linhas_alteradas:
            if dict_aux.has_key(numero):
               self.SetCellTextColour(dict_aux[numero], coluna, self.cor_editado)
        wx.SafeYield(None, True)
        self.parent.Refresh()
        return lista_de_alunos

    def OnCellChange(self, evt):
        u"""método para alterar o foreground das células editadas"""
        print "editada a celula(", evt.GetRow()," , ", evt.GetCol(), ")"
        self.le_celula ( evt.GetRow(), 0 )
        self.linhas_alteradas.append([ self.le_celula ( evt.GetRow(), 0 ), evt.GetCol()])
        self.SetCellTextColour(evt.GetRow(), evt.GetCol(), self.cor_editado)
        self.SetCellBackgroundColour(evt.GetRow(),
                                     evt.GetCol(), (240, 240, 240))

    def grava_linha ( self, ipos, numero, nome, dre, turma, obs, status ):
        u"""Método para gravar uma linha na tabela visível"""
        self.grava_celula ( ipos, 0, numero )
        self.grava_celula ( ipos, 1, nome )
        self.grava_celula ( ipos, 2, dre )
        self.grava_celula ( ipos, 3, turma )
        self.grava_celula ( ipos, 4, obs )
        self.grava_celula ( ipos, 5, status )

    def grava_celula ( self, ipos, jpos, valor ):
        u"""Método para gravar um valor em uma célula da tabela visível"""
        self.tabela.SetValue ( ipos, jpos, valor )
        self.SetCellTextColour(ipos, jpos, self.cor_default)

    def le_celula ( self, ipos, jpos ):
        u"""Método para ler o valor de uma célula da tabela visível"""
        texto = self.tabela.GetValue ( ipos, jpos )
        return ParaMaiuscula.converte(texto)

    def le_linha ( self, ipos ):
        u"""Método para ler uma linha da tabela visível"""
        return [ self.le_celula ( ipos, 0 ), self.le_celula ( ipos, 1 ),
                 self.le_celula ( ipos, 2 ), self.le_celula ( ipos, 3 ), 
                 self.le_celula ( ipos, 4, ), self.le_celula ( ipos, 5 ) ]


    def altera_obs_e_status_por_dre ( self, lista_dre, obs, status ):
        u"""Método para alterar a observação e o status de uma lista de 
        alunos identificados pelo DRE"""
        set_dre = set(lista_dre)
        num_alunos = self.retorna_num_total_de_alunos()
        for iaux1 in  range(0, num_alunos ):
            linha_lida = self.le_linha ( iaux1 )
            dre_lido = unicode(str(linha_lida[2]))
            if ( dre_lido in set_dre ):
                self.grava_celula ( iaux1, 4, obs )  
                self.grava_celula ( iaux1, 5, status )  

    def retorna_lista_de_alunos ( self ):
        "retorna uma lista das linhas com a ordem  atual"
        lista_alunos = []
        for iaux1 in range(0, self.tabela.GetNumberRows()):
            linha = self.le_linha ( iaux1 )
            if linha[1] != "":
                lista_alunos.append(linha)
        return lista_alunos

    def retorna_grupo_de_alunos ( self ):
        "retorna um grupo de alunos ordenado por numero"
        lista_alunos = self.ordena_grid_por_numero ( )
        grupo_alunos_aux = GrupoAlunos()
        for (numero, nome, dre, turma, obs, status ) in lista_alunos:
            grupo_alunos_aux.adiciona_altera_aluno ( nome, dre, turma, obs,
                                                     status )
        return grupo_alunos_aux

    def retorna_num_total_de_alunos ( self ):
        u"retorna o número de alunos com dre"
        lista_alunos = self.ordena_grid_por_numero ( )
        alunos = 0
        for lista_aux in lista_alunos:
            if  lista_aux[2] != "":
                alunos += 1
        return alunos

    def retorna_num_alunos_para_prova ( self ):
        u"retorna o número de alunos que poderá fazer a prova"
        lista_alunos = self.ordena_grid_por_numero ( )
        alunos = 0
        for lista_aux in lista_alunos:
            if  lista_aux[5] != 'NAO':
                alunos += 1
        return alunos

    def carrega_grupo_de_alunos ( self, grupo_alunos ):
        u"""carrega na tela um objeto da classe GrupoAlunos correspondente
        aos alunos na tabela visível"""
        self.limpa_celulas() 
        ipos = 0
        for  ( nome, dre, turma, obs, status ) in (
                                      grupo_alunos.get_lista_de_alunos()):
            self.grava_linha ( ipos, ipos+1, nome, dre, turma, obs, status )
            ipos += 1
        for iaux1 in range(ipos, ipos + 100):
            self.grava_linha ( iaux1, iaux1+1, u"", u"", u"", u"", 'NAO' )
        wx.SafeYield(None, True)

    def carrega_dict_de_alunos ( self, dict_alunos ):
        u"""carrega na tela um dicionario indexado por DRE e que associa a 
        cada DRE uma lista com "numero do aluno", "nome", "turma", "sit.insc".
        O status para a prova é assumido ser fazer prova, a menos que o aluno
        tenha se transferido de turma ou trancado a inscricao."""
        self.limpa_celulas() 
        ipos = 0
        max_lin = 0
        for dre in dict_alunos:
            (cnum, nome, turma, obs, dummy ) = dict_alunos[dre]
            num = int(cnum)
            if obs == "INSCRICAO TRANCADA" or obs == "TRANSFERIDO DE TURMA" \
               or obs == "MATRICULA TRANCADA" or obs == "MATRICULA CANCELADA" \
               or obs == "APROVADO" or obs =="REPROVADO" \
               or obs ==  u"NÃO INSCRITO" or obs == u"REPROVADO POR MÉDIA" \
               or obs == u"REPROVADO POR FALTA" or obs == u"REPROVADO POR FALTA E MÉDIA":
               status = 'NAO'
            else:
                status = 'S'
            self.grava_linha ( num-1, num, nome, dre, turma, obs, status )
            max_lin = max ( num, max_lin )
        for iaux1 in range(max_lin+1, max_lin + 101):
            self.grava_linha ( iaux1, iaux1+1, u"", u"", u"", u"", 'NAO' )
        wx.SafeYield(None, True)

#  
    def limpa_celulas ( self ):
        u"""Método para apagar todas as infirmações da tabrla visível"""
        cor_def = self.GetDefaultCellBackgroundColour()
        cor_txt_def = self.GetDefaultCellTextColour()
        for ipos in range(0, self.tabela.GetNumberRows()):
            self.grava_linha ( ipos, ipos+1, u"", u"", u"", u"", 'NAO' )
            for jpos in range(0, self.tabela.GetNumberCols()):
                if self.GetCellBackgroundColour(ipos, jpos) != cor_def: 
                    self.SetCellBackgroundColour(ipos, jpos, cor_def) 
                    self.SetCellTextColour(ipos, jpos, cor_txt_def) 
        wx.SafeYield(None, True)


#---------------------------------------------------------------------------




class DadosTabela(wx.grid.PyGridTableBase):
    u"""Classe para armazenar e processar os dados da tabela visível"""
    def __init__(self, parent, collabels, datatypes, data ):
        wx.grid.PyGridTableBase.__init__(self)
        self.labels = collabels
        self.datatypes = datatypes
        self.data = data

    #--------------------------------------------------
    # required methods for the wxPyGridTableBase interface
    def GetNumberRows(self):
        u"""Retorna o número de linhas visíveis"""
        return len(self.data)
    def GetNumberCols(self):
        u"""Retorna o número de colunas visíveis"""
        return len(self.data[0])
    def IsEmptyCell(self, row, col):
        u"""Retorna falso se a célcula não está vazia"""
        try:
            return not self.data[row][col]
        except IndexError:
            return True
    # Get/Set values in the table.  The Python version of these
    # methods can handle any data-type, (as long as the Editor and
    # Renderer understands the type too,) not just strings as in the
    # C++ version.
    def GetValue(self, row, col):
        u"""retorna o valor da célula"""
        try:
            return self.data[row][col]
        except IndexError:
            return ''
    def SetValue(self, row, col, value):
        u"""Grava o valor da célula"""
        try:
            self.data[row][col] = value
        except IndexError:
            # add a new row
            self.data.append([''] * self.GetNumberCols())
            self.SetValue(row, col, value)
            # tell the grid we've added a row
            msg = wx.grid.GridTableMessage(self,            # The table
                    wx.grid.GRIDTABLE_NOTIFY_ROWS_APPENDED, # what we did to it
                    1                                       # how many
                    )
            self.GetView().ProcessTableMessage(msg)
    #--------------------------------------------------
    # Some optional methods
    # Called when the grid needs to display labels
    def GetColLabelValue(self, col):
        u"""Não sei ao certo"""
        return self.labels[col]
    # Called to determine the kind of editor/renderer to use by
    # default, doesn't necessarily have to be the same type used
    # natively by the editor/renderer if they know how to convert.
    def GetTypeName(self, row, col):
        u"""Não sei ao certo"""
        return self.datatypes[col]
    # Called to determine how the data can be fetched and stored by the
    # editor and renderer.  This allows you to enforce some type-safety
    # in the grid.
    def CanGetValueAs(self, row, col, typename):
        u"""Não sei ao certo"""
        coltype = self.datatypes[col].split(':')[0]
        if typename == coltype:
            return True
        else:
            return False
    def CanSetValueAs(self, row, col, typename):
        u"""Não sei ao certo"""
        return self.CanGetValueAs(row, col, typename)


 
class MyApp(wx.App):
    u"""Classe para chamar a classe que faz todo o processamento"""
    def OnInit(self):
        u"""Método que o wxPython usa para inicializar uma apliação"""

#       self.locale = wx.Locale(wxLANGUAGE_PORTUGUESE_BRAZILIAN)
#       locale.setlocale(locale.LC_ALL, 'PT_BR')
 
        frame = MyFrame(None, -1,
            u"Aplicacao para gerar lista de alunos a partir de arquivos pdf")
        self.SetTopWindow(frame)
        frame.Show(True)
        return True




APP1 = MyApp(0)
APP1.MainLoop()
