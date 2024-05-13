#!/usr/bin/env python
# -*- coding: utf-8 -*-
u"""Módulo para manusear arquivos xml, em especial, arquivos auxiliares de um programa"""

import sys, codecs
import zipfile
#import xml.dom.minidom
import xml.etree
import xml.etree.ElementTree
#import tempfile

class ManuseiaXml ( ):
    'classe para manusear arquivo xml, incluindo criar, editar e salvar'

    def __init__ ( self, xml_file, nome_raiz, outfile=sys.stdout ):
        "recebe o objeto xml_file existente ou a ser criado"
        self.xml_file = xml_file
        self.outfile = outfile
        xml_file.seek(0,2)
        if xml_file.tell() > 0:
            xml_file.seek(0,0)
            print >> self.outfile, \
                u"Iniciando a carga da arvore relativa ao arquivo ",xml_file.name
            try:
                arvore_objeto = xml.etree.ElementTree.parse(xml_file)
                self.arvore = arvore_objeto.getroot()
#               xml.etree.ElementTree.dump(self.arvore)
                print >> self.outfile, u"Árvore relativa ao arquivo ", \
                    xml_file.name, u" está carregada"
                return
            except:
                print >> self.outfile,u"Arquivo XML tem erro" + \
                    u" de sintaxe ou não pode ser aberto"
        print >> self.outfile, \
              u"Criando uma arvore relativa ao arquivo ", xml_file.name
        self.arvore = xml.etree.ElementTree.Element(nome_raiz)
        print >> self.outfile, u"Árvore  está criada"
   
    def set_atributo ( self, nome_arquivo, nome_atributo, valor ):
        u'''atualiza ou cria, se não existir, o atributo nome_atributo, cujo valor é valor no elemento correspondente ao arquivo nome_arquivo'''
        achou = False
        for element in self.arvore:
            if element.text ==  nome_arquivo:
                achou = True
                break
        if not achou:
            element = xml.etree.ElementTree.SubElement (self.arvore, "arquivo")
            element.text = nome_arquivo
        element.set ( nome_atributo, valor )

    def set_atributo_subelemento ( self, nome_arquivo, nome_subelemento, nome_atributo, valor ):
        u'''atualiza ou cria, se não existir, o atributo nome_atributo, cujo valor é valor no subelemento de nome nome_subelemento correspondente ao arquivo nome_arquivo'''
        achou = False
        for element in self.arvore:
            if element.text ==  nome_arquivo:
                achou = True
                break
        if not achou:
            element = xml.etree.ElementTree.SubElement (self.arvore, "arquivo")
            element.text = nome_arquivo
        achou = False
        for subelement in element:
            if subelement.tag ==  nome_subelemento:
                achou = True
                break
        if not achou:
            subelement = xml.etree.ElementTree.SubElement (element, nome_subelemento )
        subelement.set ( nome_atributo, valor )


    def get_atributo ( self, nome_arquivo, nome_atributo ):
        u'''retorna o atributo nome_atributo do elemento correspondente ao arquivo nome_arquivo'''
        for element in self.arvore:
            if element.text ==  nome_arquivo:
                return element.get ( nome_atributo, False )
        return False

    def get_atributo_subelemento ( self, nome_arquivo, nome_subelemento, nome_atributo ):
        u'''retorna o valor do atributo nome_atributo, do subelemento de nome nome_subelemento correspondente ao arquivo nome_arquivo'''
        achou = False
        for element in self.arvore:
            if element.text ==  nome_arquivo:
                achou = True
                break
        if not achou:
            return False
        for subelement in element:
            if subelement.tag ==  nome_subelemento:
                return subelement.get ( nome_atributo, False )
        return False

    def grava_arvore ( self ):
        u'''atualiza o arquivo com a arvore já com as modificações feitas'''
#       print xml.etree.ElementTree.tostring ( self.arvore,"utf-8" )
        self.xml_file.seek(0,0)
        try:
           print >> self.xml_file, xml.etree.ElementTree.tostring ( self.arvore,"utf-8" )
        except:
            print >> self.outfile, u"Problemas com a gravação do arquivo", self.xml_file.name



def exemplo_manuseio_xml(argv):
    u"""Rotina para testar a classe ManuseiaXml"""
    if len(sys.argv) > 1 :
        arquivo_xml = argv[1]
#       file_sai = codecs.open(u"aaa", 'w', 'utf-8')
        file_xml = open( arquivo_xml, "r+" )
#       arq_xml = ManuseiaXml ( file_xml, "config", file_sai )
        arq_xml = ManuseiaXml ( file_xml, "config" )
        arq_xml.set_atributo ( "/home/ivolopez/aaa.csv", "prova", "PF" )
        arq_xml.set_atributo_subelemento ( "/home/ivolopez/aaa.csv", "salas", "D-226", "70" )
        arq_xml.set_atributo_subelemento ( "/home/ivolopez/aaa.csv", "salas", "D-220", "50" )
        arq_xml.set_atributo_subelemento ( "/home/ivolopez/aaa.csv", "salas", "D-246", "40" )
        print arq_xml.get_atributo ( "/home/ivolopez/aaa.csv", "prova" )
        arq_xml.grava_arvore ( )


    else:
        print u"""É necessário o nome do arquivo de configuração.  A execução será encerrada"""

#exemplo_manuseio_xml( sys.argv )

