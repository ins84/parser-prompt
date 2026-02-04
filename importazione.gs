function parseMarkdownFromLocal() {
  var html = HtmlService.createHtmlOutputFromFile('FilePicker')
    .setWidth(500)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, '🎬 Sora Parser - Carica File dal PC');
}

function processMarkdownContent(markdownText, fileName) {
  var ui = SpreadsheetApp.getUi();
  
  markdownText = markdownText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  
  var data = [['Prompt', 'NomeFile']];
  var sceneCount = 0;
  
  var lines = markdownText.split('\n');
  var currentPrompt = '';
  var currentScene = '';
  var isCapturing = false;
  var skipNextLines = 0;
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    
    if (line.indexOf('SORA 2 VIDEO PROMPT - SCENE') > -1) {
      var match = line.match(/SCENE (\d+_\d+)/);
      if (match) {
        isCapturing = true;
        currentScene = match[1];
        skipNextLines = 1;
        continue;
      }
    }
    
    if (skipNextLines > 0) {
      skipNextLines--;
      continue;
    }
    
    if (isCapturing && line.indexOf('===') > -1) {
      if (i + 1 < lines.length && lines[i+1].indexOf('END OF SCENE') > -1) {
        data.push([currentPrompt.trim(), 'SCENE ' + currentScene]);
        sceneCount++;
        
        isCapturing = false;
        currentPrompt = '';
        currentScene = '';
        continue;
      }
    }
    
    if (isCapturing) {
      currentPrompt += line + '\n';
    }
  }
  
  if (data.length <= 1) {
    return {
      success: false,
      message: '⚠️ Nessuna scena trovata!\n\nFile: ' + fileName + '\nDimensione: ' + markdownText.length + ' caratteri'
    };
  }
  
  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Prompts') || 
                    SpreadsheetApp.getActiveSpreadsheet().insertSheet('Prompts');
  
  outputSheet.clear();
  outputSheet.getRange(1, 1, data.length, 2).setValues(data);
  
  var headerRange = outputSheet.getRange(1, 1, 1, 2);
  headerRange.setFontWeight('bold')
             .setBackground('#4285F4')
             .setFontColor('white')
             .setFontSize(12);
  
  outputSheet.setColumnWidth(1, 800);
  outputSheet.setColumnWidth(2, 150);
  outputSheet.setRowHeight(1, 30);
  
  if (data.length > 1) {
    var dataRange = outputSheet.getRange(2, 1, data.length - 1, 2);
    dataRange.setWrap(true)
             .setVerticalAlignment('top')
             .setFontSize(10);
    
    for (var i = 2; i <= data.length; i++) {
      if (i % 2 == 0) {
        outputSheet.getRange(i, 1, 1, 2).setBackground('#f3f3f3');
      }
    }
  }
  
  return {
    success: true,
    message: '✅ SUCCESSO!\n\n' + 
             'File: ' + fileName + '\n' +
             'Scene estratte: ' + sceneCount + '\n\n' +
             'Controlla il foglio "Prompts"'
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎬 Sora Parser')
    .addItem('📥 Importa dal PC', 'parseMarkdownFromLocal')
    .addToUi();
}
