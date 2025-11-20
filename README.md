# Exportador Automático: Google Sheets para Excel (XLSX)

Este projeto contém um script para Google Apps Script que exporta uma aba específica de uma planilha Google Sheets para um arquivo Excel (.xlsx) no Google Drive. O script foi desenhado para ser modular e de fácil adaptação ("Plug & Play").

## 🚀 Funcionalidades

* **Conversão Limpa:** Copia apenas valores e formatação (remove fórmulas para evitar erros no Excel).
* **Organização Inteligente:** Salva o arquivo no Google Drive criando automaticamente uma estrutura de pastas: `Pasta Raiz > Ano > Mês`.
* **Nomenclatura Dinâmica:** Nomeia o arquivo com prefixo + data + hora (ex: `Export-2023-11-19_14h30.xlsx`).
* **Log de Auditoria:** Cria e preenche automaticamente uma aba de "Log" na planilha com Data, Nome do Arquivo, Usuário responsável e Link direto.
* **Interface:** Adiciona um menu personalizado "⚡ Automação" na barra superior da planilha.

## ⚙️ Instalação

1.  Abra sua planilha Google e vá em **Extensões > Apps Script**.
2.  Crie dois arquivos de script: `Config.gs` e `Code.gs`.
3.  Copie os códigos deste repositório para os respectivos arquivos.
4.  Edite o arquivo `Config.gs` com suas definições (ID da Pasta do Drive, Nome da Aba, etc).
5.  Salve e atualize a planilha.

## 📋 Pré-requisitos

* Uma conta Google com acesso ao Google Drive.
* Permissão para executar scripts (na primeira execução, será necessário autorizar o acesso ao Drive e Planilhas).

---
*Desenvolvido para integração de dados e rotinas de importação de ERP.*
