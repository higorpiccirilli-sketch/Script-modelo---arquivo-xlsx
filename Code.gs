/**
 * Adiciona o menu personalizado ao abrir a planilha.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Automação')
    .addItem('📤 Exportar para Excel', 'mainExportProcess')
    .addToUi();
}

/**
 * Função Principal: Orquestra todo o processo.
 */
function mainExportProcess() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Validar se a aba de origem existe
    var sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      ui.alert('ERRO: A aba "' + CONFIG.SOURCE_SHEET_NAME + '" não foi encontrada. Verifique o nome no arquivo Config.gs');
      return;
    }

    // Notificar inicio (opcional, bom para arquivos grandes)
    ss.toast('Iniciando exportação...', 'Aguarde', 5);

    // 2. Gerar Nomes e Datas
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd_HH'h'mm");
    var fileName = CONFIG.FILE_PREFIX + formattedDate + ".xlsx";
    
    var yearStr = Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy");
    var monthStr = Utilities.formatDate(now, CONFIG.TIMEZONE, "MM-MMM"); // Ex: 11-Nov

    // 3. Gerenciar Pastas (Ano > Mês)
    var targetFolder = getTargetFolder(CONFIG.ROOT_FOLDER_ID, yearStr, monthStr);

    // 4. Criar Planilha Temporária (Apenas Valores)
    // Isso é necessário para remover as fórmulas antes de converter para Excel
    var tempSpreadsheet = createTempSpreadsheet(sourceSheet, fileName);
    
    // 5. Converter para Blob (XLSX)
    var xlsxBlob = getExcelBlob(tempSpreadsheet);
    
    // 6. Salvar Arquivo na Pasta Destino
    var file = targetFolder.createFile(xlsxBlob);
    
    // 7. Excluir Planilha Temporária
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);

    // 8. Gravar Log
    logExecution(ss, file.getName(), file.getUrl(), now);

    // 9. Sucesso
    ss.toast('Arquivo salvo na pasta: ' + yearStr + '/' + monthStr, 'Sucesso!');
    
  } catch (e) {
    console.error(e);
    ui.alert('Ocorreu um erro: ' + e.toString());
  }
}

/**
 * Navega ou cria a estrutura de pastas Ano > Mês
 */
function getTargetFolder(rootId, year, month) {
  var rootFolder = DriveApp.getFolderById(rootId);
  var yearFolder, monthFolder;

  // Verifica/Cria pasta do Ano
  var yearFolders = rootFolder.getFoldersByName(year);
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = rootFolder.createFolder(year);
  }

  // Verifica/Cria pasta do Mês
  var monthFolders = yearFolder.getFoldersByName(month);
  if (monthFolders.hasNext()) {
    monthFolder = monthFolders.next();
  } else {
    monthFolder = yearFolder.createFolder(month);
  }

  return monthFolder;
}

/**
 * Cria uma planilha temporária copiando apenas VALORES e FORMATOS
 */
function createTempSpreadsheet(sourceSheet, name) {
  // Cria planilha temporária
  var tempSS = SpreadsheetApp.create(name);
  var tempSheet = tempSS.getSheets()[0];
  
  // Pega dados e formatação da origem
  var range = sourceSheet.getDataRange();
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  var fontColors = range.getFontColors();
  var fontWeights = range.getFontWeights();
  
  // Cola na temporária
  var targetRange = tempSheet.getRange(1, 1, values.length, values[0].length);
  targetRange.setValues(values); // COLA APENAS VALORES (Remove Fórmulas)
  targetRange.setBackgrounds(backgrounds);
  targetRange.setFontColors(fontColors);
  targetRange.setFontWeights(fontWeights);
  
  // Ajustar largura das colunas (opcional, visual)
  for (var i = 1; i <= values[0].length; i++) {
    tempSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
  }

  SpreadsheetApp.flush(); // Força a atualização
  return tempSS;
}

/**
 * Converte a planilha temporária para XLSX usando a API do Drive
 */
function getExcelBlob(spreadsheet) {
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=xlsx";
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  return response.getBlob().setName(spreadsheet.getName());
}

/**
 * Gerencia a aba de Log
 */
function logExecution(ss, fileName, fileUrl, dateObj) {
  var logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  
  // Se não existe, cria
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    // Se criou agora, precisamos garantir que o cabeçalho vá na linha configurada
  }

  // Verifica se precisa montar cabeçalho (se a linha configurada está vazia)
  var headerRange = logSheet.getRange(CONFIG.LOG_HEADER_ROW, 1, 1, 4);
  if (headerRange.isBlank()) {
    var headers = [["Data e Hora", "Nome do Arquivo", "Usuário", "Link"]];
    headerRange.setValues(headers);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#d9d9d9"); // Cinza leve
  }

  // Determina a próxima linha vazia. 
  // Se a última linha com dados for menor que a linha do cabeçalho, a próxima é a (Header + 1)
  var lastRow = logSheet.getLastRow();
  var nextRow = lastRow < CONFIG.LOG_HEADER_ROW ? CONFIG.LOG_HEADER_ROW + 1 : lastRow + 1;

  var formattedDate = Utilities.formatDate(dateObj, CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm:ss");
  var userEmail = Session.getActiveUser().getEmail();
  var hyperlinkFormula = '=HYPERLINK("' + fileUrl + '"; "Abrir Arquivo")';

  logSheet.getRange(nextRow, 1, 1, 4).setValues([[
    formattedDate,
    fileName,
    userEmail,
    hyperlinkFormula
  ]]);
}
