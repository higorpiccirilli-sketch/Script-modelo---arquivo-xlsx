// ============================================================================
// ⚙️ ARQUIVO DE CONFIGURAÇÃO (Config.gs)
// ============================================================================
// Edite APENAS o que está dentro das aspas " " ou após o sinal de =
// ============================================================================

var CONFIG = {
  
  // 1. NOME DA ABA DE ORIGEM
  // Digite exatamente o nome da aba que contém os dados que você quer exportar.
  SOURCE_SHEET_NAME: "Dados", 

  // 2. ID DA PASTA MÃE NO GOOGLE DRIVE
  // Abra a pasta no Drive. O ID é a sequência de letras/números no final da URL.
  // Exemplo: drive.google.com/drive/folders/1aBcD_xYz... -> O ID é "1aBcD_xYz..."
  ROOT_FOLDER_ID: "COLE_O_ID_DA_SUA_PASTA_AQUI",

  // 3. PREFIXO DO NOME DO ARQUIVO
  // O nome final será: PREFIXO + DATA + HORA.xlsx
  // Exemplo: "Relatorio-" vira "Relatorio-2025-11-19_14h30.xlsx"
  FILE_PREFIX: "Exportacao-",

  // 4. CONFIGURAÇÃO DO LOG
  // Nome da aba onde o histórico será salvo. O script cria se não existir.
  LOG_SHEET_NAME: "Log_Geracao",
  
  // Linha onde o cabeçalho do Log deve ficar (conforme seu pedido: linha 3)
  LOG_HEADER_ROW: 3,

  // 5. FUSO HORÁRIO
  // Use "GMT-3" para horário de Brasília.
  TIMEZONE: "GMT-3"
};
