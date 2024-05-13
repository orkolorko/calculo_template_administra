#!/usr/bin/env python
# -*- coding: utf-8 -*-
u"""Módulo para manusear grupos de alunos"""
import sys, subprocess, re, codecs, unicodedata, os, os.path

class ParaMaiuscula:
    u"""classe para converter textos, retirando acentuacao e trocando todas as
    letras para letras maiusculas"""
    def converte ( texto ):
        u"""Método para passar um texto para maiúsculas e retirar toda a 
        acentuação"""
        if isinstance(texto, basestring):
            if isinstance(texto, unicode):
                texto_aux = unicodedata.normalize("NFKD",
                            texto ).encode("ASCII", 'ignore')
            else:
                texto_aux = unicodedata.normalize("NFKD",
                            unicode(texto, 'utf8')).encode("ASCII", 'ignore')
            return texto_aux.upper()
        return texto
    converte = staticmethod ( converte )



class GrupoAlunos:
    "Classe que contem a relacao de alunos e processa-os"
#  algumas variaveis da classe para ativar os padroes de busca
    numeros_aux = 'zero|um|dois|tres|quatro|cinco|seis|sete|oito|nove|dez'
    lista_num_aux = numeros_aux.split('|')
    aux_nome_dre = re.compile(
    r"(?P<ordem>[1-9][0-9]{0,2}) (?P<nome>.*) (?P<dre>[0-9]{9,9})(?P<resto>.*)")
    aux_aluno_aprovado = re.compile(
                     r"\"(?P<nome>[A-Z][A-Z. ]{0,100})\".*(?P<aprovado>\,AP\,)")
    aux_aluno = re.compile(r"\"(?P<nome>[A-Z][A-Z. ]{0,100})\".*")
    aux_motivo = re.compile(r"(?P<motivo> [A-Z*]+.*)" )
    lista_dos_motivos_validos = ['INSCRICAO TRANCADA', 'TRANSFERIDO DE TURMA',
         'DISCIPLINA JA CURSADA',  'MATRICULA TRANCADA',
         'MATRICULA CANCELADA', 'LOTACAO ESGOTADA',  'FALTA REQUISITO', 
         'HORARIO SOBREPOSTO', 'ABANDONO JUSTIFICADO', 'MENOS DE 6 CREDITOS', 
         'MAIS DE 32 CREDITOS', 'ULTRAPASSA 1/3 FORA DO CURSO', 'INCOMPLETO',
         '***************', 'PREVISAO PENDENTE' ]
    set_dos_motivos_validos = set ( lista_dos_motivos_validos )
    set_motivos_nao_faz_prova = set (['INSCRICAO TRANCADA',
                             'TRANSFERIDO DE TURMA','DISCIPLINA JA CURSADA'])
    dic_campos = dict()
    dic_campos["nome"] = 0
    dic_campos["dre"] = 1
    dic_campos["turma"] = 2
    dic_campos["observacao"] = 3
    dic_campos["status"] = 4


    def __init__(self):
        self.lista_de_alunos = []
        self.dict_dre = dict()

    def adiciona_altera_aluno ( self, nome, dre, turma, obs, status='S' ):
        u"""Método para adicionar um aluno ao grupo de alunos.  Adiciona
            aluno ao final da lista.  Se ja existir altera os dados para
            os informados a menos que seja aluno transferido de turma"""
        if dre == "":
            print u'''Erro!! O programa nao permite adicionar aluno sem
                      o numero do DRE'''
            sys.exit(15)
        if self.dict_dre.has_key(dre):
            ipos = self.dict_dre[dre]
            if ( obs.find("TRANSFERIDO DE TURMA") == -1 ):
                self.lista_de_alunos[ipos][0] = nome
                self.lista_de_alunos[ipos][4] = status
                self.lista_de_alunos[ipos][2] = turma
                if ( obs == "" ):
                    self.lista_de_alunos[ipos][3] = u'OK'
                else:
                    self.lista_de_alunos[ipos][3] = obs
            return
        self.lista_de_alunos.append([nome, dre, turma, obs, status])
        self.dict_dre[dre] = len(self.lista_de_alunos)-1

    def get_lista_de_alunos ( self ):
        u"""Método para retornar a lista com os alunos"""
        return self.lista_de_alunos

    def get_dict_alunos ( self ):
        u"""Método para retornar um dicionário cujo índice é o DRE e que a 
            este associa uma lista contendo: (0) nome; (1) turma;
            (2) situação da inscrição; (3) indica se faz prova"""
        dict_alunos = dict()
        for dre in self.dict_dre:
            lista_aux = self.lista_de_alunos[self.dict_dre[dre]]
            dict_alunos[dre] = [ lista_aux[0], lista_aux[2], lista_aux[3],
                                 lista_aux[4] ]
        return dict_alunos


    def get_numero_de_alunos ( self ):
        u"Método para retornar o número de alunos no grupo""" 
        return len(self.lista_de_alunos)

    def ret_dic_com_alunos_por_turma ( self ):
        u"""retorna um dicionario em que os indices são os nomes da turmas e
        os valores são o número de alunos da turma"""
        dic_aux = dict()
        for lista_aux in self.lista_de_alunos:
            turma = lista_aux[2]
            if turma in dic_aux:
                dic_aux[turma] += 1
            else:
                dic_aux[turma] = 0
        return dic_aux
           
    def ret_lista_com_turmas ( self ):
        u"""retorna uma lista com todas as turmas a que pertencem este grupo de alunos"""
        conj_turmas_aux = set()
        for lista_aux in self.lista_de_alunos:
            conj_turmas_aux.add(lista_aux[2])
        lista_turmas_aux = list(conj_turmas_aux)
        lista_turmas_aux.sort()
        return lista_turmas_aux

    def limpa_lista_de_alunos ( self ):
        u"""Método para retirar todos os alunos da lista"""
        self.lista_de_alunos = []
        self.dict_dre = dict()

    def reordena_lista ( self, campo='nome', reverse=False ):
        u'''rotina para reordenar os alunos.  O campo pode ser "nome", "dre",
        "turma", "obs", "status". Altera a lista de alunos original'''
        if ( not GrupoAlunos.dic_campos.has_key(campo) ):
            print "O campo informado para ordenacao nao eh um campo valido."+\
            u"  O campo informado foi " + campo + " e os campos validos sao: "
            print GrupoAlunos.dic_campos.keys()
            print "A execucao sera encerrada."
            sys.exit(112)
        iaux = GrupoAlunos.dic_campos[campo]
        lista_aux = self.lista_de_alunos[:]
        lista_aux.sort((lambda x, y:cmp(x[iaux], y[iaux])), None, reverse)
        self.lista_de_alunos = lista_aux
#       reindexa o dicionario com o dre
        for ipos, (nome, dre, turma, obs, status) in enumerate(self.lista_de_alunos):
            self.dict_dre[dre] = ipos
        return 

    def le_alunos_arquivo_pdf ( self, arq_pdf, turma, reordena=True ):
        u'''método para ler os alunos de um arquivo pdf e inserir em um grupo'''
        linhas_alunos = self.__le_arquivo_pauta (arq_pdf)
        self.__processa_alunos_pdf ( linhas_alunos, turma )
        if reordena:
            self.reordena_lista ( "nome", False )
        return turma

    def le_atualiza_alunos_arquivo_pdf ( self, arq_pdf, grupo_alunos_old, reordena=True ):
        u'''método para ler os alunos de um arquivo pdf, identificar a turma baseado
        em um arquivo com as pautas antigas e acrescentar os alunos'''
        linhas_alunos = self.__le_arquivo_pauta (arq_pdf)
        turma = self.__identifica_turma_alunos_pdf ( linhas_alunos, grupo_alunos_old )
        self.__processa_alunos_pdf ( linhas_alunos, turma )
        if reordena:
            self.reordena_lista ( "nome", False )
        return turma

    def __le_arquivo_pauta ( self, arq_pdf ):
        u'''método para ler os alunos de um arquivo pdf, usando o programa pdftotex
        para fazer a conversão para texto.  Retorna o conteudo pre-processado'''
        nome_pdftotext_local_linux = "/usr/bin/pdftotext"
        diretorio_do_python, dummy = os.path.split(os.path.abspath(sys.argv[0]))
        nome_pdftotext_local_windows = os.path.join(diretorio_do_python, \
                   "pdftotext.exe" )
        if os.path.exists(nome_pdftotext_local_linux):
             programa_pdftotext = nome_pdftotext_local_linux
        elif os.path.exists(nome_pdftotext_local_windows):
             programa_pdftotext = nome_pdftotext_local_windows
        elif os.path.exists("C:\Arquivos de programas\Xpdf\pdftotext.exe"):
             programa_pdftotext = "C:\Arquivos de programas\Xpdf\pdftotext.exe"
        elif os.path.exists("/usr/bin/pdftotext"):
             programa_pdftotext = "/usr/bin/pdftotext"
        else:
            print "Nao foi encontrado o programa pdftotext que faz"
            print "a leitura de arquivos no formato pdf.  A execucao"
            print "sera encerrada"
#           sys.exit(120)
        print programa_pdftotext
        print arq_pdf
        linhas_alunos_aux = subprocess.Popen([programa_pdftotext,
            "-enc", "UTF-8", "-nopgbrk", "-layout", arq_pdf, "-"],
            stdout=subprocess.PIPE)
        linhas_alunos = self.__preprocessa_arquivo ( linhas_alunos_aux.stdout )
        print "Arquivo pdf sendo processado: ", arq_pdf
        return linhas_alunos


    def le_alunos_arquivo_txt ( self, arq_txt ):
        "método para ler os alunos de um arquivo txt"
        arq_alunos_txt = open ( arq_txt, "r" )   
        linhas_alunos_aux = arq_alunos_txt.readlines()
        linhas_alunos = self.__preprocessa_arquivo ( linhas_alunos_aux )
        self.__processa_alunos_txt ( linhas_alunos )
        arq_alunos_txt.close()

    def grava_alunos_arquivo_txt ( self, arq_txt ):
        u"""método para gravar os alunos em um arquivo txt.  Grava nome,
        dre, turma, obs, status"""
        arq_alunos_txt = open ( arq_txt, "w" )   
        for (nome, dre, turma, obs, status ) in self.get_lista_de_alunos():
            if ( obs == "" ):
                obs = u'OK'
            arq_alunos_txt.write("%s, %s, %s, %s, %s\n"  %
                                 (nome, dre, turma, obs, status))
        arq_alunos_txt.close()

    def __identifica_turma_alunos_pdf ( self, linhas_alunos, grupo_alunos_old ):
        u"""método especifico para encontrar o nome da turma de um arquivo com pauta  
        no formato pdf"""
        lista_turmas = grupo_alunos_old.ret_lista_com_turmas()
        dic_alunos_na_turma = dict()
        for turma in lista_turmas:
            dic_alunos_na_turma[turma] = 0
        dic_alunos = grupo_alunos_old.get_dict_alunos( )
        total_alunos = 0
        for linha in linhas_alunos:
            linha_aux = ParaMaiuscula.converte(linha)
            aux1 = GrupoAlunos.aux_nome_dre.match(linha_aux)
            if aux1:
                dre = aux1.group('dre')
                if dre in dic_alunos:
                    if dic_alunos[dre][2] != "TRANSFERIDO DE TURMA":
                        turma_aluno = dic_alunos[dre][1]
                        dic_alunos_na_turma[turma_aluno] += 1
                        total_alunos += 1
        turma_escolhida = ""
        max_alunos = 0
        for turma in dic_alunos_na_turma:
            if dic_alunos_na_turma[turma] > max_alunos:
                turma_escolhida = turma
                max_alunos = dic_alunos_na_turma[turma]     
        if 100.00 * max_alunos / max(total_alunos,1) < 80.0:
            print "Turma selecionada tem ", max_alunos,\
                  " alunos, sendo que na pauta temos ", total_alunos 
        print "Turma escolhida, alunos na pauta, alunos na turma",\
              turma_escolhida, max_alunos, total_alunos, "\n"
        return turma_escolhida




    def __processa_alunos_pdf ( self, linhas_alunos, turma="" ):
        u"""método especifico para gerar os dados dos alunos correspondentes 
        a uma pauta no formato pdf"""
        for linha in linhas_alunos:
            linha_aux = ParaMaiuscula.converte(linha)
            aux1 = GrupoAlunos.aux_nome_dre.match(linha_aux)
            if aux1:
                nome = ParaMaiuscula.converte(aux1.group('nome'))
                dre = aux1.group('dre')
#               print "resto: ###"+ aux1.group('resto') + "###"
                aux2 = GrupoAlunos.aux_motivo.match(aux1.group('resto'))
                motivo = u'OK'
                if aux2:
#                   print aux2
                    motivo = (aux2.group('motivo')).strip()
#                   print "motivo" + motivo
#                   procura o motivo para nao constar na pauta
                    if motivo:
                        if motivo not in GrupoAlunos.set_dos_motivos_validos:
                            achei = 0
#                           pode ter mais de um motivo ou motivo com
#                           outros argumentos ( por exemplo,
#                           abandono justificado )
                            for texto in GrupoAlunos.lista_dos_motivos_validos:
                                if motivo.find(texto) > -1:
                                    achei = 1
                            if not achei:
                                print \
                                  u"O motivo do abandono do aluno nao esta "+\
                                  u"entre os conhecidos. A linha contem:"
                                print linha_aux
                                print "O motivo interpretado eh"
                                print "####" + motivo + "####"
                                sys.exit(1)
                        if ( motivo == "PREVISAO PENDENTE" ):
                            self.adiciona_altera_aluno ( nome, dre, turma,
                                                        u'OK', 'S' )
                        elif motivo in self.set_motivos_nao_faz_prova:
                            self.adiciona_altera_aluno ( nome, dre, turma,
                                                        motivo, 'NAO')
                        else:
                            self.adiciona_altera_aluno ( nome, dre, turma,
                                                        motivo, 'S')
                    else:
                        print \
                          u"O motivo do abandono do aluno nao esta "+\
                          u"entre os conhecidos e nem tem forma "+\
                          u"valida. A linha contem:"
                        print linha_aux
                        sys.exit(112)
                else:
                    self.adiciona_altera_aluno ( nome, dre, turma, 
                                                 motivo, 'S' )

    def __processa_alunos_txt ( self, linhas_alunos ):
        u"""método especifico para gerar dados dos alunos correspondentes a um
        arquivo no formato csv. 
        Espera que venha nome, dre, turma, obs, status"""
#     print "entrei na processa alunos txt"
        if ( len(linhas_alunos) > 0 ):
            for linha in linhas_alunos:
                lista_aux = linha.split(',')
                if ( len(lista_aux) > 0 ):
                    for iaux in range(0, len(lista_aux)):
                        lista_aux[iaux] = ParaMaiuscula.converte(
                                                       lista_aux[iaux].strip())
                    for iaux in range(len(lista_aux), 5):
                        lista_aux.append("")
                    if lista_aux[4] == "":
                        self.adiciona_altera_aluno ( lista_aux[0],
                           lista_aux[1], lista_aux[2], lista_aux[3], 'S' )
                    else:
                        self.adiciona_altera_aluno ( lista_aux[0],
                           lista_aux[1], lista_aux[2], lista_aux[3], lista_aux[4] )



    def __preprocessa_arquivo ( self, linhas_convertidas ):
        "le o arquivo, elimina brancos superfluos e linhas vazias"
        linhas_uteis = []
        elimina_brancos = re.compile("[ ]+")
        for linhas in linhas_convertidas:
#        retirando linhas vazias, brancos multiplos, brancos no inicio, 
#        brancos no fim e o newline que o readlines coloca.  Apaga linhas
#       so com virgulas.
            if len ( linhas ) > 1:
                linhas_uteis.append(elimina_brancos.sub(' ', linhas.strip()))
        return linhas_uteis





class CabecalhoProva:
    u"""Classe que gera o arquivo latex com os nomes dos alunos escolhidos. 
    Eh basicamente uma classe estatica"""
    preambulo_latex = """\documentclass[10pt,brazil,a4paper]{article} \n
\usepackage[T1]{fontenc}\n\usepackage[utf8]{inputenc}\n
\usepackage{ae,aecompl}\n\usepackage{geometry}\n
\usepackage{lscape,longtable}\n\geometry{a4paper,tmargin=15mm,
bmargin=15mm,lmargin=10mm,rmargin=10mm}\n\\usepackage{fancyhdr}\n
\pagestyle{fancy}\n\usepackage[brazil]{babel}\n \n \n \\begin{document}\n \n """
    preambulo_landscape_latex = """\documentclass[10pt,brazil,a4paper]{article}
\n\usepackage[T1]{fontenc}\n\usepackage[utf8]{inputenc}\n
\usepackage{ae,aecompl}\n\usepackage[landscape]{geometry}\n
\usepackage{longtable}\n\geometry{a4paper,tmargin=25mm,bmargin=15mm,
lmargin=10mm,rmargin=10mm}\n\usepackage{fancyhdr}\n\pagestyle{fancy}\n
\usepackage[brazil]{babel}\n\\renewcommand{\\arraystretch}{1.8} \n \n 
\\begin{document}\n \n """

    def gera_cadernos_prova ( nome_arq_modelo_prova, nome_arq_caderno_prova,
                              grupo_de_alunos, dia, mes, ano, disc, prova,
                              num_questoes, dic_salas ):
        u"""Método para gerar arquivos Latex para o caderno de provas,
        lista de presença, lista com a sala de cada aluno e lista com
        alunos irregularmente inscritos"""
        arq_ent = codecs.open(nome_arq_modelo_prova, "r", "utf-8")   
        arq_sai = codecs.open(nome_arq_caderno_prova, "w", "utf-8" )   
        linhas_aux = arq_ent.read()
        arq_ent.close()
        linhas = linhas_aux.splitlines()
        #print linhas
        linhas_corpo = []
        estagio = "preambulo"
        for linha in linhas:
            linha.strip("\r\n")
            if (estagio == "preambulo" ):
                if ( linha.find("begin{document}") != -1 ):
                    estagio = "corpo"
                    arq_sai.write(
                     "\setcounter{numquestoes}{"+"%d" % num_questoes + "}" + "\n")
                    arq_sai.write(
                     "\setcounter{numfolhas}{"+"%d" % num_questoes + "}" + "\n")
                arq_sai.write(linha+"\n")
            else:
                linhas_corpo.append(linha)
        lista_salas = dic_salas.keys()
        lista_salas.sort()
        sala_aluno = []
        for sala in lista_salas:
            for iaux in range(0, dic_salas[sala]):
                sala_aluno.append(sala)
        lista_aux = CabecalhoProva.retorna_lista_ordenada ( grupo_de_alunos )
        ipos = 0
        for ( nome_aluno, num_aluno, dre_aluno, turma, motivo, status ) in lista_aux:
            if status == 'S':
                numal = "%.3d" % num_aluno
                sala = sala_aluno[ipos]
                ipos += 1
                for linha in linhas_corpo:
                    if len(nome_aluno) < 32:
                        nome_e_fonte = "\large{"+nome_aluno+"}"
                    elif len(nome_aluno) < 40:
                        nome_e_fonte = "\\normalsize{"+nome_aluno+"}"
                    else:
                        nome_e_fonte = "\small{"+nome_aluno+"}"
                    linha = linha.replace("NOMEDOALUNO", nome_e_fonte )
                    linha = linha.replace("NUMAL", numal )
                    linha = linha.replace("DREAL", dre_aluno )
                    linha = linha.replace("DIAM", dia )
                    linha = linha.replace("MESAN", mes )
                    linha = linha.replace("NUMAN", ano )
                    linha = linha.replace("NOMEDISC", disc )
                    linha = linha.replace("NOMETURM", turma )
                    linha = linha.replace("NOMESALA", sala )
                    linha = linha.replace("NUMPROVA", prova )
                    linha = linha.replace("end{document}", 'newpage' )
                    arq_sai.write(linha+"\n")
#                   print linha
        arq_sai.write("\\end{document}\n")  
        arq_sai.close()
    gera_cadernos_prova = staticmethod ( gera_cadernos_prova )

    def gera_listas_alunos ( nome_arq_lista_alunos, nome_arq_irregular,
                             nome_arq_presenca, nome_arq_salas_usadas, grupo_de_alunos,
                             dia, mes, ano, disc, prova, dic_salas ):
        u"""Método para gerar arquivos latex com as listas dos alunos e das
        salas em que farão prova"""
        arq_sai_lista_alunos = codecs.open(nome_arq_lista_alunos,
                                           "w", "utf-8" )   
        arq_sai_irregular = codecs.open(nome_arq_irregular, "w", "utf-8" )   
        arq_sai_presenca = codecs.open(nome_arq_presenca, "w", "utf-8" )   
        arq_sai_salas_usadas = codecs.open(nome_arq_salas_usadas, "w", "utf-8" )   

        arq_sai_lista_alunos.write(CabecalhoProva.preambulo_latex)
        arq_sai_irregular.write(CabecalhoProva.preambulo_latex)
        arq_sai_presenca.write(CabecalhoProva.preambulo_landscape_latex)
        arq_sai_salas_usadas.write(CabecalhoProva.preambulo_landscape_latex)
        arq_sai_lista_alunos.write(u" {\chead{\\textbf{Lista com os alunos de "+
                       disc + u" que farão a  " + prova + u" no dia " + dia +
                       u"/" + mes + "/" + ano + u"}}}\n \n" )
        arq_sai_irregular.write(
           u"{\chead{\\textbf{Lista com os alunos de " + disc + 
           u" cujo nome consta na pauta mas cuja inscrição " +
           u"não está regular}}}\n\n")
        arq_sai_lista_alunos.write(
          u"\\begin{longtable}{rp{9cm}llp{5cm}} Número & Nome "+
              u"& DRE & Sala &"+ u" Observação \\\\ \n \endhead \n") 
        arq_sai_irregular.write(u"\\begin{longtable}{rp{8cm}lp{6cm}} \n "+
                  u"Número & Nome & DRE & Observação \\\\ \n \endhead \n") 
        arq_sai_salas_usadas.write(
                        u" \n \\newpage \chead{\large{\\textbf{Lista de "+
                        u" salas para a " +
                        prova + u" no dia " + dia + u"/" + mes + u"/" +
                        ano + " }}}\n \n" )
        arq_sai_salas_usadas.write(u"\n \\begin{center} \n"+
                       u" \\begin{longtable}{|p{1cm}|p{2cm}|p{2cm}|p{5cm}|"+
                       u"p{5cm}|p{8cm}|} \n \hline \n  & Sala & "+
                       u"Alunos & Professor & Monitor & Observações \\\\ \hline \hline"+
                       u" \n \endhead \n")
#       arq_sai_irregular.write( 
#            u"\\begin{longtable}{rllll}\n\caption*{Alunos de " + disc +
#            u" cuja inscrição não está regular para a " +
#            prova + 
#            u" }\\\\ \n Número & Nome & DRE & "+
#            u"Sala & Observação \\\\ \n \endhead \n") 
        lista_salas = dic_salas.keys()
        lista_salas.sort()
        sala_aluno = []
        for sala in lista_salas:
            for iaux in range(0, dic_salas[sala]):
                sala_aluno.append(sala)
#       for lista in (grupo_de_alunos.get_lista_de_alunos()):
#           print lista
        lista_aux = CabecalhoProva.retorna_lista_ordenada ( grupo_de_alunos )
        alunos_na_prova = len(lista_aux)
        ipos = 0
        cont_salas = 0
        sala_old = "inicio"
        for ( nome_aluno, num_aluno, dre_aluno, turma, motivo_aluno, status
                                                                 ) in lista_aux:
            if status == 'S':
                numal = "%.3d" % num_aluno
                sala = sala_aluno[ipos]
                ipos += 1
                arq_sai_lista_alunos.write(numal + " & " + nome_aluno +
                      " & " + dre_aluno + " & " + sala + " & " + motivo_aluno +
                      " \\\\ \n" )
                if ( motivo_aluno != "" and motivo_aluno != u'OK' ):
                    arq_sai_irregular.write(numal + " & " + nome_aluno +
                         " & " + dre_aluno + " & " + motivo_aluno + " \\\\ \n" )
                if ( sala != sala_old ):
                    if sala_old != "inicio":
                        arq_sai_presenca.write(
                                  u"\end{longtable} \n \end{center} \n " )
                        arq_sai_salas_usadas.write( str(cont_salas) + " & " + sala_old +
                             " & " + str(icont) + " &  &  & \\\\ \hline \n" )
                    icont = 0
                    cont_salas += 1
                    arq_sai_presenca.write(
                        u" \n \\newpage \chead{\large{\\textbf{Lista de "+
                        u" presença dos alunos de " + disc + u" que fizeram "+
                        u"a " + prova + u" no dia " + dia + u"/" + mes + "/" +
                        ano + u" na sala " + sala + " }}}\n \n" )
                    arq_sai_presenca.write(u"\n \\begin{center} \n"+
                       u" \\begin{longtable}{|p{0.8cm}|p{7.2cm}||p{1.7cm}|"+
                       u"p{7cm}|p{8cm}|} \n \hline \n  & Nome & DRE & "+
                       u"Correio Eletrônico & Assinatura \\\\ \hline \hline"+
                       u" \n \endhead \n")
                    sala_old = sala
                icont += 1   
                arq_sai_presenca.write( str(icont) + " & " + nome_aluno +
                             " & " + dre_aluno + " &  & \\\\ \hline \n" )
            else:
                alunos_na_prova -= 1
        arq_sai_lista_alunos.write(u"\end{longtable} \n\end{document} \n\n")
        arq_sai_lista_alunos.close()
        arq_sai_irregular.write(u"\end{longtable} \n\end{document} \n\n")
        arq_sai_irregular.close()
        arq_sai_presenca.write(
             u"\end{longtable} \n \end{center} \n \end{document} \n \n" )
        arq_sai_presenca.close()
        arq_sai_salas_usadas.write( str(cont_salas) + " & " + sala +
                             " & " + str(icont) + " &  &  & \\\\ \hline \n" )
        arq_sai_salas_usadas.write(
             u"\end{longtable} \n \end{center} \n \n" + 
             u" Total: " + str(cont_salas) + 
             u" salas e " + str(alunos_na_prova) + u" alunos" +
             u"\n\end{document}")
        arq_sai_salas_usadas.close()
    gera_listas_alunos = staticmethod ( gera_listas_alunos )

    def retorna_lista_ordenada ( grupo_de_alunos ):
        u"""Ordena o grupo de alunos por ordem alfabetica e retorna lista 
        ordenada contendo os dados dos alunos, incluindo o numero"""
        num_aluno = 0
        lista_aux = []
        for ( nome_aluno, dre_aluno, turma, motivo_aluno,
             status ) in (grupo_de_alunos.get_lista_de_alunos()):
            num_aluno += 1
            lista_aux.append([nome_aluno, num_aluno, dre_aluno, turma,
                              motivo_aluno, status])
        lista_aux.sort((lambda x, y:cmp(x[0], y[0])), None)
        return lista_aux
    retorna_lista_ordenada = staticmethod ( retorna_lista_ordenada )


