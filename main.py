# -*- coding: utf-8 -*-
"""
Ponto de entrada principal da aplicação Leitor de XML NF-e Pro.
Executa a interface gráfica moderna para leitura de notas fiscais e exportação para Excel.
"""

import sys
import os

# Adiciona o diretório atual ao sys.path para garantir importações relativas limpas
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app_gui import iniciar_aplicacao

if __name__ == "__main__":
    iniciar_aplicacao()
