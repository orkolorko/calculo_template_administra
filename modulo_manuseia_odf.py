#!/usr/bin/env python
# -*- coding: utf-8 -*-
u"""Módulo para manusear arquivos do OPenOffice Calc, formato .ods"""
#import os, string, codecs
#import xml.dom
#import xml.dom.ext
#import xml.parsers.expat
#import StringIO
#from xml.dom import minidom, Node

import sys, codecs
import zipfile
import xml.dom.minidom
#import tempfile

class ManuseiaArquivoOdf ( ):
    'classe para manusear arquivo odf, incluindo editar e salvar'

    def __init__ ( self, odf_file, outfile=sys.stdout ):
        "recebe o nome do arquivo no formato odf do OPenOffice e abre o zipfile"
        self.odf_zipfile = zipfile.ZipFile(odf_file, 'r')
        self.outfile = outfile
        self.dict_dom_arq_editados = {}
        for info in self.odf_zipfile.infolist() :
            self.dict_dom_arq_editados[info.filename] = False 
   
#============================================================================
#============================================================================
# rotinas auxiliares para manuseio dos arquivos no formato odf 

    def retorna_arquivo_do_zip (self, nome_arq) :
        u"""Retorna o conteudo de um arquivo dentro do arquivo odf ( que é um
        arquivo comprimido constituido de diversos outros arquivos )"""
        component = self.odf_zipfile.read(nome_arq)
        return component

    def retorna_dom ( self, nome_arq ):
        u'''retorna arvore carregada correspondente ao arquivo nome_arq. 
        A arvore fica aberta'''
        if ( not self.dict_dom_arq_editados.has_key(nome_arq) ):
            print >> self.outfile, u"Arquivo ", nome_arq, \
            u" não está entre os arquivos disponíveis no arquivo zipado."+\
            u"  A execução será encerrada"
            sys.exit(12)
        if self.dict_dom_arq_editados[nome_arq]:
            dom =  self.dict_dom_arq_editados[nome_arq]
        else:
            print >> self.outfile, u"Lendo o arquivo zipado"
            codigo_xml = self.retorna_arquivo_do_zip(nome_arq)
            print >> self.outfile, \
                 u"Iniciando a carga da arvore relativa ao arquivo ", nome_arq
            dom = xml.dom.minidom.parseString(codigo_xml)
            self.dict_dom_arq_editados[nome_arq] = dom 
            print >> self.outfile, u"Árvore relativa ao arquivo ", \
              nome_arq, u" está carregada"
        return dom   

    def libera_memoria ( self ):
        u"""libera a memoria usada nas arvores dom.  Note que a árvore pode
        ficar desatualizada, pois futuros acessos não contarão com a árvore
        alterada anteriomente disponível na memória""" 
        for info in self.odf_zipfile.infolist():
            dom = self.dict_dom_arq_editados[info.filename]
            if dom:
                dom.unlink()
                self.dict_dom_arq_editados[info.filename] = False


    def encerra_objeto (self):
        u"""Metodo para fechar o objeto do tipo ZipFile.  Se nao quiser salvar,
        chamar antes o metodo libera_memoria"""  
        self.odf_zipfile.close()
             

    def retorna_lista_de_linhas( self, node_parent ):
        "retorna lista de linhas filhas do elemento atual"
        return node_parent.getElementsByTagName("table:table-row")

    def retorna_iesimo_filho ( self, node_elemento, nome_filho,
                               nome_attr_filho, contador ):
        u"""metodo auxiliar que procura entre os filhos de node_elemento,
        os que tem nome dado por nome_filho.  Retorna o filho na posição
        contador.  Para levar em conta facilidades do openoffice para grupar
        linhas e colunas, pode contar elementos como se fossem varios, usando
        o nome_attr_filho"""
        lista_elementos_validos = node_elemento.getElementsByTagName(nome_filho)
        conta_elementos_validos = 0
        for elemento_valido in lista_elementos_validos:
            attrs = elemento_valido.attributes         
            if ( nome_attr_filho in set(attrs.keys()) ):
                attrnode = attrs.get(nome_attr_filho)
                attrvalue = attrnode.nodeValue
                conta_elementos_validos += int(attrvalue)
            else:
                conta_elementos_validos += 1
            if ( conta_elementos_validos >= contador ):
                return elemento_valido
        return False

    def retorna_lista_filhos ( self, node_elemento, nome_filho,
                               nome_attr_filho, cont_inic, cont_fim ):
        u"""metodo auxiliar que procura entre os filhos de node_elemento,
        os que tem nome dado por nome_filho.  Retorna uma lista de filhos
        da posicao cont_ini ate a posicao cont_fim.
        Para levar em conta facilidades do openoffice para grupar
        linhas e colunas, pode contar elementos como se fossem varios, usando
        o nome_attr_filho"""
#       print "Entrei retorna_lista_filhos"
        lista_elementos_validos = node_elemento.getElementsByTagName(nome_filho)
        conta_elementos_validos = 0
        lista_elementos_saida = []
        for elemento_valido in lista_elementos_validos:
            attrs = elemento_valido.attributes         
            if ( nome_attr_filho in set(attrs.keys()) ):
                attrnode = attrs.get(nome_attr_filho)
                attrvalue = attrnode.nodeValue
                conta_elementos_validos += int(attrvalue)
            else:
                conta_elementos_validos += 1
#           print "aaa",  conta_elementos_validos, cont_inic, cont_fim, len(lista_elementos_saida)
            if conta_elementos_validos >= cont_inic:
                for ind in range ( cont_inic + len(lista_elementos_saida), 
                    min(cont_fim, conta_elementos_validos) + 1): 
                    lista_elementos_saida.append(elemento_valido)
            if conta_elementos_validos >= cont_fim:
                return lista_elementos_saida
        return False

    def procura_filho_pelo_nome (self, node_elemento, nome_filho,
                                 nome_attr_filho, valor):
        u"""método auxiliar que procura entre os filhos de node_elemento, os
        que tem nome do elemento dado por nome_filho.  Retorna o filho cujo 
        atributo nome_attr_filho for igual ao valor"""
        lista_elementos_validos = node_elemento.getElementsByTagName(nome_filho)
        for elemento_valido in lista_elementos_validos:
            attrs = elemento_valido.attributes         
            if ( nome_attr_filho in set(attrs.keys()) ):
                attrnode = attrs.get(nome_attr_filho)
                attrvalue = attrnode.nodeValue
                if ( attrvalue == valor ):
                    return elemento_valido
        return False

    def procura_filho_pela_ordem (self, node_elemento, nome_filho, 
                                  nome_attr_filho, ordem_filho):
        u"""metodo auliliar que procura entre os filhos de node_elemento, o
        de ordem ordem_filho dentre os que tem atributo  nome_attr_filho. 
        Retorna o valor ( nome ) e o ponteiro ( node )"""
        lista_elementos_validos = node_elemento.getElementsByTagName(nome_filho)
        iaux1 = 0
        for elemento_valido in lista_elementos_validos:
            attrs = elemento_valido.attributes         
            if ( nome_attr_filho in set(attrs.keys()) ):
                iaux1 += 1
                if ( iaux1 == ordem_filho ):     
                    attrnode = attrs.get(nome_attr_filho)
                    attrvalue = attrnode.nodeValue
                    return attrvalue, elemento_valido
        return False

    def retorna_texto_da_celula ( self, node_cell ):
        "metodo auxiliar para retornar o texto em uma celula"
        if ( node_cell.hasChildNodes() ):
            texto = node_cell.childNodes[0]
            if ( texto.hasChildNodes() ):
                texto = texto.childNodes[0]
                if texto.nodeType == xml.dom.Node.TEXT_NODE:
                    return texto.nodeValue
        return False

    def altera_texto_da_celula ( self, node_cell, novo_texto ):
        "metodo auxiliar para alterar o texto em uma celula"
#       print novo_texto
        if ( node_cell.hasChildNodes() ):
#           print "tem filhos"
            texto = node_cell.childNodes[0]
            texto = texto.childNodes[0]
            if texto.nodeType == xml.dom.Node.TEXT_NODE:
#               print "eh texto"
                texto.nodeValue = novo_texto
                return texto
        return False

    def retorna_no_iesima_linha ( self, node_tabela, num_linha ):
        u"""metodo que retorna o no correspondente à linha da tabela
        contida na tabela node_tabela"""
        return self.retorna_iesimo_filho ( node_tabela, u"table:table-row",
                                    u"table:number-rows-repeated", num_linha )

    def retorna_lista_nos_linhas (self, node_tabela, num_lin_inic, num_lin_fim):
        u"""metodo que retorna o no correspondente à linha da tabela
        contida na tabela node_tabela"""
#       print "Entrei retorna_lista_nos_linhas"
        return self.retorna_lista_filhos ( node_tabela, u"table:table-row",
                    u"table:number-rows-repeated", num_lin_inic, num_lin_fim )

    def retorna_iesima_celula_na_linha ( self, node_linha, num_coluna ):
        u"""metodo que retorna o no correspondente à coluna da linha contida
        na linha node_linha"""
        return self.retorna_iesimo_filho ( node_linha, u"table:table-cell",
                                u"table:number-columns-repeated", num_coluna )

    def procura_tabela_por_ordem ( self, dom, ordem_tabela ):
        u"""metodo que retorna o no correspondente à tabela com a ordem pedida 
        na arvore dom"""
        parent = dom.documentElement
        return self.procura_filho_pela_ordem ( parent, u"table:table",
                                               u"table:name", ordem_tabela )

    def procura_tabela_por_nome ( self, dom, nome_tabela ):
        "metodo que retorna o no correspondente à tabela pedida na arvore dom"
        parent = dom.documentElement
        return self.procura_filho_pelo_nome ( parent, u"table:table",
                                             u"table:name", nome_tabela )

    def printlevel( self, level):
        u"""Método para imprimir espaços em branco, para fazer identação"""
        self.outfile.write(level*u'    ')

    def walk ( self, parent, level, elementos_aceitos):  
        u"""Método para percorrer recursivamente a árvore e imprimir os 
        principais dados sobre ela"""
        for node in parent.childNodes:
            print >> self.outfile, u"classe do node: ", node.nodeName, \
              node.__class__
            if ( node.nodeType == xml.dom.Node.ELEMENT_NODE ):
                if ( ( len(elementos_aceitos) == 0) or
                     ( node.nodeName in elementos_aceitos ) ):
                    # Write out the element name.
                    self.printlevel( level)
                    self.outfile.write(u'Element: %s\n' % node.nodeName)
                    # Write out the attributes.
                    attrs = node.attributes         
                    for attrname in attrs.keys():
                        attrnode = attrs.get(attrname)
                        attrvalue = attrnode.nodeValue
                        self.printlevel( level + 2)
                        self.outfile.write(
                             u'Attribute -- Name: %s  Value: %s\n' % \
                            (attrname, attrvalue))
                    # Walk over any text nodes in the current node.
                    content = []                   
                    iaux = 0
                    for child in node.childNodes:
                        if child.nodeType == xml.dom.Node.TEXT_NODE:
                            content.append(child.nodeValue)
                            iaux += 1
                    if content:
                        strcontent = u''.join(content)
                        self.printlevel( level)
                        self.outfile.write(u'Content: "')
                        self.outfile.write(strcontent)
                        self.outfile.write(u'"\n')
                # Walk the child nodes.
                self.walk(node, level+1, elementos_aceitos)
    
# fim das rotinas auxiliares para manuseio dos arquivos no formato odf 
#============================================================================
#============================================================================




#============================================================================
#============================================================================
# rotinas para impressao das informações dos arquivos

    def imp_nomes_dos_arquivos_no_odf(self):
        u"""Imprime nomes dos arquivos que constituem o arquivo odf, bem como
        tamanho, tamanho comprimido."""
        print >> self.outfile, u'\nOs arquivos no ODF são os seguintes:'
        total_compressed_size = 0
        total_uncompressed_size = 0
        print >> self.outfile, u"  %-45s %12s %12s %12s\n" % (u'Name',
                            u'Compressed', u'Uncompressed', u'Date')
        print >> self.outfile, u"  %-45s %12s %12s %12s\n\n" % (u' ',
                            u'Size', u'Size', u'        ' )
        for info in self.odf_zipfile.infolist() :
            total_compressed_size += info.compress_size
            total_uncompressed_size += info.file_size
            print >> self.outfile,  u"  %-45s %12d %12d %12s\n" % (
                 info.filename, info.compress_size, info.file_size,
                 info.date_time )
        print >> self.outfile, u"  %-45s %12s %12s\n" % (
                                               u' ', u'--------', u'--------') 
        print >> self.outfile, u"  %-45s %12d %12d\n" % (
                     u'Total', total_compressed_size, total_uncompressed_size)
        return total_compressed_size


    def imprime_arquivo ( self, nome_arq ):
        """Imprime o conteudo do arquivo nome_arq"""
        print >> self.outfile, 80*u"="
        print >> self.outfile, u'Conteudo do arquivo ' + nome_arq + u':\n' 
        if ( nome_arq[-4:] == u'.xml' ):
            dom = self.retorna_dom(nome_arq)
            print >> self.outfile, dom.toprettyxml() 
        else:
            codigo_xml = self.retorna_arquivo_do_zip(nome_arq)
            print >> self.outfile, codigo_xml
        print >> self.outfile, 80*u"="
        print >> self.outfile, u"\n" 
        return 

    def imprime_nos_e_elementos(self, nome_arq, elementos_aceitos=set()):  
        u"""metodo que percorre a arvore de um arquivo xml imprimindo o nome 
        dos nos, seus elementos e seus valores.  Chama o metodo walk 
        recursivamente e este faz o trabalho todo"""
        dom = self.retorna_dom(nome_arq)
        rootnode = dom.documentElement
        level = 0
        self.walk ( rootnode, level, elementos_aceitos ) 
#     dom.unlink()


# fim das rotinas para impressao das informações dos arquivos
#============================================================================
#============================================================================



#============================================================================
#============================================================================
# rotinas para manusear as tabelas, alterando, apagando ou capturando 
# informação

    def apaga_linhas_na_tabela ( self, nome_arq, nome_tabela,
                                indice_primeira_linha, numero_de_linhas ):
        u'''edita a arvore de xml,  apaga as linhas da tabela com nome
        nome_tabela, a partir da linha indice_primeira_linha ate a linha 
        indice_primeira_linha+numero_de_linhas-1'''  
        dom = self.retorna_dom(nome_arq)
        node_tabela = self.procura_tabela_por_nome ( dom, nome_tabela )
        node_primeira_linha = self.retorna_no_iesima_linha ( node_tabela,
                                                      indice_primeira_linha )
        node_ultima_linha = self.retorna_no_iesima_linha ( node_tabela,
                                   indice_primeira_linha+numero_de_linhas-1 )
        if ( not node_primeira_linha ) or ( not node_ultima_linha ):
            print >> self.outfile,\
                    u"Não foram encontradas todas as linhas a apagar."+\
                    u" A execucao será encerrada"
            sys.exit(21)
        lista_linhas = self.retorna_lista_de_linhas ( node_tabela )
        apagando = False
        for linha in lista_linhas:
            if ( linha == node_primeira_linha ):
                apagando = True
            elif ( linha == node_ultima_linha ):
                no_pai = linha.parentNode
                no_pai.removeChild(linha)
                break
            if apagando:
                no_pai = linha.parentNode
                no_pai.removeChild(linha)

    def apaga_linhas_na_arvore ( self, nome_arq, lista_linhas_a_excluir,
                                lista_colunas ):
        u"""edita a arvore de xml,  apaga as linhas identificadas pela lista
        lista_linhas_a_excluir.  Cada elemento da lista é uma lista com os
        valores para identificar linhas a serem excluidas"""  
        dom = self.retorna_dom(nome_arq)
        parent = dom.documentElement
        num_colunas = len(lista_colunas)
        ilin = 0
        total_de_linhas = len( lista_linhas_a_excluir )
        lista_linhas = self.retorna_lista_de_linhas ( parent )
        print >> self.outfile, u"numero de linhas ", len(lista_linhas)
#       print u"lista_linhas_a_excluir", lista_linhas_a_excluir
        for linha in lista_linhas:
            ipos = 0
            achou = False
#           print u"entrei na table-row", ipos, lista_colunas[ipos] 
            for ipos, coluna in enumerate(lista_colunas):
                node_cell = self.retorna_iesima_celula_na_linha ( 
                                                          linha, coluna )
                if ( not node_cell ):
                    achou = False
                    break
                else:
                    valor_antigo = self.retorna_texto_da_celula ( node_cell )
                    if ( valor_antigo == lista_linhas_a_excluir[ilin][ipos] ):
                        ipos += 1
                        achou = True
                        if ipos >= num_colunas:
                            break
                    else:
                        achou = False
                        break
            if achou:
                no_pai = linha.parentNode
                no_pai.removeChild(linha)
                ilin += 1
                if ilin >= total_de_linhas:
                    break
        if ilin < total_de_linhas:
            print >>self.outfile,\
              u"Não foram encontradas todas as linhas.  Foram encontradas,"+\
              u" ao todo, ", ilin, u" linhas.  A execucao será encerrada"
            sys.exit(10)

    def substitui_valores_na_arvore ( self, nome_arq, 
          lista_valores_a_substituir, lista_valores_novos, lista_colunas ):
        u'''edita a arvore de xml,  substituindo os valores nas celulas dados
        por lista_valores_a_substituir pela lista_valores_novos.  Cada elemento
        das listas corresponde a uma linha da planilha. Por isso, cada elemento
        das listas é também uma lista com os valores a serem alterados
        ( inseridos )'''  
        dom = self.retorna_dom(nome_arq)
        parent = dom.documentElement
        num_colunas = len(lista_colunas)
        ilin = 0
        total_de_linhas = len( lista_valores_novos )
        lista_linhas = self.retorna_lista_de_linhas ( parent )
        for linha in lista_linhas:
            ipos = 0
            achou = True
            for ipos, coluna in enumerate(lista_colunas):
                node_cell = self.retorna_iesima_celula_na_linha (
                            linha, coluna )
                if ( not node_cell ):
                    achou = False
                    break
                else:
                    valor_antigo = self.retorna_texto_da_celula ( node_cell )
                    if valor_antigo == lista_valores_a_substituir[ilin][ipos]:
#                       print valor_antigo,lista_valores_a_substituir[ilin][ipos], lista_valores_novos[ilin][ipos]
                        self.altera_texto_da_celula ( node_cell,
                                      lista_valores_novos[ilin][ipos] )
                    else:
                        achou = False
                        break
            if achou:
                ilin += 1
                if ilin >= total_de_linhas:
                    break
        if ilin < total_de_linhas:
            print >> self.outfile,\
            u"Não foram encontradas todas as linhas."+ \
            u" Foram encontradas, ao todo, ", \
            ilin, u" linhas.  A execucao será encerrada"
            sys.exit(10)

    def altera_celula_por_posicao ( self, nome_arq,  nome_tabela, linha, coluna,
                                    valor_antigo, valor_novo ):
        "metodo para alterar uma celula a partir da sua posicao"
#       print "altera_celula_por_posicao", linha, coluna, valor_antigo, valor_novo
        dom = self.retorna_dom(nome_arq)
        node_tabela = self.procura_tabela_por_nome ( dom, nome_tabela )
        if ( not node_tabela ):
            print >> self.outfile,  u"Não foi encontrada a tabela ", \
                         nome_tabela, u".  A execucao será encerrada"
            sys.exit(14)
        node_linha = self.retorna_no_iesima_linha ( node_tabela, linha )
        if ( not node_linha ):
            print >> self.outfile,  u"Não foi encontrada a linha ", linha, \
                                   u".  A execucao será encerrada"
            sys.exit(15)
        node_celula = self.retorna_iesima_celula_na_linha ( node_linha, coluna )
        if ( not node_celula ):
            print >> self.outfile, u"Não foi encontrada a celula na coluna ", \
                       coluna,  u".  A execucao será encerrada"
            sys.exit(16)
        valor_antigo_lido = self.retorna_texto_da_celula ( node_celula )
        if ( valor_antigo_lido == valor_antigo ):
            self.altera_texto_da_celula ( node_celula, valor_novo )
        else:
            print >> self.outfile,  u"O valor antigo na celula: ", \
                 valor_antigo_lido, u" é diferente do valor esperado: ",\
                 valor_antigo,  u".  A execucao será encerrada"
            sys.exit(16)

    def retorna_nome_tabela ( self, nome_arq, ordem_tabela ):
        u"""Método para retornar o nome da tabela de ordem ordem_tabela 
        ( começa por 1 ), que pode ter sido alterado pelo usuário"""
        dom = self.retorna_dom(nome_arq)
        nome_tabela, node_tabela = self.procura_tabela_por_ordem ( 
                                                        dom, ordem_tabela )
        return nome_tabela

    def retorna_linha_tabela ( self, nome_arq, ordem_tabela, ordem_linha,
                              max_colunas ):
        u"""Método para retornar a linha de ordem ordem_linha da tabela de
        ordem ordem_tabela ate um máximo de max_colunas"""
        dom = self.retorna_dom(nome_arq)
        nome_tabela, node_tabela = self.procura_tabela_por_ordem (
                                                        dom, ordem_tabela )
        node_linha = self.retorna_no_iesima_linha ( node_tabela, ordem_linha )
        linha_lida = []
        for coluna in range(1, max_colunas+1):
            node_celula = self.retorna_iesima_celula_na_linha ( node_linha,
                                                                coluna )
            if ( not node_celula ):
                return linha_lida
            else:
                valor_lido = self.retorna_texto_da_celula ( node_celula )
                linha_lida.append(valor_lido)
        return linha_lida

    def retorna_lista_linhas_tabela ( self, nome_arq, ordem_tabela,
                                      ordem_lin_inic, ordem_lin_fim,
                                      max_colunas ):
        u"""Método para retornar uma lista com as linhas de ordem 
        ordem_lin_inic até ordem_lin_fim da tabela de
        ordem ordem_tabela ate um máximo de max_colunas"""
#       print "Entrei retorna_lista_linhas_tabela"
        dom = self.retorna_dom(nome_arq)
        nome_tabela, node_tabela = self.procura_tabela_por_ordem (
                                                        dom, ordem_tabela )
        lista_nodes_linha = self.retorna_lista_nos_linhas ( node_tabela, 
                                            ordem_lin_inic, ordem_lin_fim )
        linhas_lidas = []
        for no in lista_nodes_linha:
            linha_aux = []
            for coluna in range(1, max_colunas+1):
                node_celula = self.retorna_iesima_celula_na_linha ( no, coluna )
                if ( not node_celula ):
                    return False
                else:
                    valor_lido = self.retorna_texto_da_celula ( node_celula )
                    linha_aux.append(valor_lido)
            linhas_lidas.append(linha_aux)
        return linhas_lidas

    def ret_col_de_valores_com_teste ( self, nome_arq, ordem_tabela,
           indice_primeira_linha, col_busca, col_retorno, valor_a_encontrar ):
        u"""retorna os valores que estão na coluna col_retorno da tabela de
        ordem ordem_tabela.  Retorna somente os valores das linhas em que o
        valor da coluna col_busca coincide com valor_a_encontrar.  A busca
        começa na linha indice primeira_linha e vai até o fim da tabela"""
        dom = self.retorna_dom(nome_arq)
        nome_tabela, node_tabela = self.procura_tabela_por_ordem (
                                                       dom, ordem_tabela )
        node_primeira_linha = self.retorna_no_iesima_linha ( node_tabela,
                                                      indice_primeira_linha )
        if ( not node_primeira_linha ):
            print >> self.outfile,\
                  u"Não foi encontrada a primeira linha a apagar."+\
                  u" A execucao será encerrada"
            sys.exit(22)
        lista_linhas = self.retorna_lista_de_linhas ( node_tabela )
        inicio = False
        lista_retorno = []
        for node_linha in lista_linhas:
            if ( node_linha == node_primeira_linha ):
                inicio = True
            if ( inicio ):
                node_celula = self.retorna_iesima_celula_na_linha ( node_linha,
                                                                  col_busca )
                if ( node_celula ):
                    valor_lido = self.retorna_texto_da_celula ( node_celula )
                    if ( valor_lido == valor_a_encontrar ):
                        node_celula = self.retorna_iesima_celula_na_linha (
                                                    node_linha, col_retorno )
                        if ( node_celula ):
                            valor_lido = self.retorna_texto_da_celula ( 
                                                                node_celula )
                            lista_retorno.append(valor_lido)
        return lista_retorno


    def le_celula_por_posicao ( self, nome_arq,  nome_tabela, linha, coluna ):
        "metodo para ler o valor de uma celula a partir da sua posicao"
        dom = self.retorna_dom(nome_arq)
        node_tabela = self.procura_tabela_por_nome ( dom, nome_tabela )
        if ( not node_tabela ):
            print >> self.outfile,  u"Não foi encontrada a tabela ", \
                         nome_tabela, u".  A execucao será encerrada"
            sys.exit(14)
        node_linha = self.retorna_no_iesima_linha ( node_tabela, linha )
        if ( not node_linha ):
            print >> self.outfile,  u"Não foi encontrada a linha ", linha, \
                                   u".  A execucao será encerrada"
            sys.exit(15)
        node_celula = self.retorna_iesima_celula_na_linha ( node_linha, coluna )
        if ( not node_celula ):
            print >> self.outfile, u"Não foi encontrada a celula na coluna ", \
                       coluna,  u".  A execucao será encerrada"
            sys.exit(16)
        return self.retorna_texto_da_celula ( node_celula )

    def apaga_tabelas ( self, nome_arq, lista_tabelas_a_excluir ):
        u"""edita a arvore de xml,  apaga as tabelas cujos nomes 
        estejam na lista_tabelas_a_excluir """  
        dom = self.retorna_dom(nome_arq)
        for nome_tabela in lista_tabelas_a_excluir:
            tabela = self.procura_tabela_por_nome ( dom, nome_tabela )
            if ( not  tabela ):
                print >> self.outfile,  u"Não foi encontrada a tabela ", \
                  nome_tabela, u".  A execucao será encerrada"
                sys.exit(13)
            else:
                no_pai = tabela.parentNode
                no_pai.removeChild(tabela)
                print >> self.outfile, u"removeu tabela ", nome_tabela


    def grava_arvore ( self, nome_arq_sai ):
        'metodo para guardar a arvore atual em um arquivo e fechar a arvore'
        zipfile_arq_sai = zipfile.ZipFile (nome_arq_sai, 'w' )
        for info in self.odf_zipfile.infolist():
            if self.dict_dom_arq_editados[info.filename]:
                texto =  self.dict_dom_arq_editados[info.filename].toxml()
                zipfile_arq_sai.writestr(info.filename, texto.encode('utf-8'))
            else:
                zipfile_arq_sai.writestr(info.filename,
                       self.retorna_arquivo_do_zip(info.filename))
        zipfile_arq_sai.close()

# fim das rotinas para manusear as tabelas, alterando, apagando ou capturando 
# informação
#============================================================================
#============================================================================
  

def exemplo_manuseio_odf():
    u"""Rotina para testar a classe manuseia_arquivo_odf, com alguns exemplos de
    uso a partir de um template"""
    if len(sys.argv) > 1 :
        odf_file = sys.argv[1]
        file_sai = codecs.open(u"aaa", 'w', 'utf-8')
        arq_odf = ManuseiaArquivoOdf ( odf_file, file_sai )
#       arq_odf = ManuseiaArquivoOdf ( odf_file, sys.stdout )
        arq_odf.imp_nomes_dos_arquivos_no_odf()
#       arq_odf.imprime_arquivo ( 'mimetype' )
#       arq_odf.imprime_arquivo ( 'META-INF/manifest.xml' )
#       arq_odf.imprime_arquivo ( 'styles.xml' )
#       arq_odf.imprime_arquivo ( 'meta.xml' )
        arq_odf.imprime_arquivo ( 'content.xml' )

#       arq_odf.imprime_nos_e_elementos("content.xml",
#               set(["table:table-row","table:table-cell","text:P"]))  
#       arq_odf.imprime_nos_e_elementos( "content.xml", set([]))  
        return

        lista_tabelas_a_excluir = [ u"Turma_%d" % iaux for iaux in range(5, 16)]
        arq_odf.apaga_tabelas ( u"content.xml", lista_tabelas_a_excluir )

        lista_colunas = [ 1, 2, 3, 4, 5 ]
        lista_linhas_a_excluir =  [ [ u"%d" % iaux, u"Aluno %4.4d" % iaux,
               u"108%6.6d" % iaux, u"Turma%4.4d" % ( (iaux - 1) % 15 + 1 ),
               u"OK" ] for iaux in range(501,1501)]
        arq_odf.apaga_linhas_na_arvore ( u"content.xml", lista_linhas_a_excluir,
                                        lista_colunas )

        lista_valores_a_substituir =  [ [ u"%d" % iaux, u"Aluno %4.4d" % iaux,
                u"108%6.6d" % iaux, u"Turma%4.4d" % ( ( iaux - 1 ) % 15 + 1 ),
                u"OK" ] for iaux in range(1, 501)]
        lista_valores_novos =  [ [ u"%6.6d" % iaux, u"Aluno %4d" % iaux,
              u"108%4d" % iaux, u"Turma%2.2d" % (iaux % 4+1), u"OK%4.4d" % iaux
                                            ] for iaux in range(1,501)]
        arq_odf.substitui_valores_na_arvore ( u"content.xml",
             lista_valores_a_substituir, lista_valores_novos, lista_colunas )

        for iaux in range(1, 5):
            arq_odf.altera_celula_por_posicao ( u"content.xml", 
               u"Turma_%d"%iaux, 1, 1, u"Turma%4.4d"%iaux, u"Turma%2.2d"%iaux )
            arq_odf.apaga_linhas_na_tabela ( u"content.xml",
                                             u"Turma_%d" % iaux, 138, 15 )

        arq_odf.grava_arvore ( u'saida.ods' )

        arq_odf.libera_memoria ( )

    else:
        print u"""É necessário o nome da arquivo com o template.  """+\
              u"""A execução será encerrada"""

#exemplo_manuseio_odf()  

