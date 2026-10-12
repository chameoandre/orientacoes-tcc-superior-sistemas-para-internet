/**
 * ============================================================================
 * GOOGLE APPS SCRIPT: BACKEND DE INTEGRAÇÃO DO DASHBOARD DE TCC (IFSC GAROPABA)
 * ============================================================================
 * 
 * Este script deve ser colado em:
 * Extensões > Apps Script da planilha do Google Sheets:
 * https://docs.google.com/spreadsheets/d/193Uemkt_-eAkzlQxUuAznmwLxzzO4FW6mjw1e_DURhA/edit
 * 
 * Após colar:
 * 1. Clique em "Salvar" (ícone de disquete).
 * 2. Clique em "Implantar" > "Gerenciar implantações".
 * 3. Clique no ícone de lápis (Editar) na implantação existente.
 * 4. Em "Versão", selecione "Nova versão".
 * 5. Garanta que "Quem pode acessar" esteja como "Qualquer pessoa" (Anyone).
 * 6. Clique em "Implantar".
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  // Aguarda até 10 segundos para evitar concorrência
  lock.tryLock(10000);

  try {
    var rawData = e.postData ? e.postData.contents : null;
    if (!rawData) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: 'Nenhum dado recebido no corpo da requisição.' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var data = JSON.parse(rawData);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ========================================================================
    // CASO 1: ATUALIZAR DADOS CADASTRAIS DO ORIENTANDO (Aba PRÉ TCC / Gid 0)
    // ========================================================================
    if (data.acao === 'atualizar_aluno') {
      var sheetAlunos = ss.getSheets()[0]; // Primeira aba: PRÉ TCC
      var values = sheetAlunos.getDataRange().getValues();
      var targetRow = -1;

      var searchName = (data.nome || '').toString().trim().toLowerCase();
      var searchFirstWord = searchName.split(' ')[0];

      for (var i = 0; i < values.length; i++) {
        var rowName = (values[i][0] || '').toString().trim().toLowerCase();
        if (rowName && (rowName.indexOf(searchFirstWord) !== -1 || searchName.indexOf(rowName.split(' ')[0]) !== -1)) {
          targetRow = i + 1; // Linha 1-indexada no Sheets
          break;
        }
      }

      if (targetRow !== -1) {
        // Coluna 4 (D): Dashboard de Avanços
        if (data.linkDashboard !== undefined) {
          sheetAlunos.getRange(targetRow, 4).setValue(data.linkDashboard);
        }
        // Coluna 7 (G): Prazo de Entrega
        if (data.prazoEntrega !== undefined) {
          sheetAlunos.getRange(targetRow, 7).setValue(data.prazoEntrega);
        }
        // Coluna 8 (H): Status TCC
        if (data.status !== undefined) {
          sheetAlunos.getRange(targetRow, 8).setValue(data.status);
        }
        // Coluna 9 (I): Status Detalhe / Próximos Passos
        if (data.statusDetalhe !== undefined) {
          sheetAlunos.getRange(targetRow, 9).setValue(data.statusDetalhe);
        }
        // Coluna 18 (R): Link do Paper Overleaf
        if (data.linkOverleaf !== undefined) {
          sheetAlunos.getRange(targetRow, 18).setValue(data.linkOverleaf);
        }
        // Coluna 19 (S): GitHub Repository
        if (data.linkGithub !== undefined) {
          sheetAlunos.getRange(targetRow, 19).setValue(data.linkGithub);
        }

        return ContentService.createTextOutput(JSON.stringify({ 
          status: 'success', 
          message: 'Orientando ' + data.nome + ' atualizado com sucesso na linha ' + targetRow 
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ 
          status: 'error', 
          message: 'Orientando ' + data.nome + ' não foi localizado na planilha principal.' 
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ========================================================================
    // CASO 2: DISPARO DE E-MAIL ISOLADO
    // ========================================================================
    if (data.acao === 'enviar_email') {
      enviarAtaPorEmail(data);
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'E-mail de ata enviado com sucesso.' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ========================================================================
    // CASO 3: REGISTRO DE NOVA REUNIÃO DE ORIENTAÇÃO (Aba reunioes-registro)
    // ========================================================================
    var sheetReunioes = ss.getSheetByName('reunioes-registro');
    if (!sheetReunioes) {
      sheetReunioes = ss.insertSheet('reunioes-registro');
      sheetReunioes.appendRow([
        'Data', 'Hora', 'Aluno', 'Tipo', 'Pauta', 'Deliberações', 'Próximos Passos', 'Data Registro'
      ]);
    }

    sheetReunioes.appendRow([
      data.data || '',
      data.hora || '',
      data.alunoNome || '',
      data.tipo || '',
      data.pauta || '',
      data.deliberacoes || '',
      data.proximosPassos || '',
      new Date()
    ]);

    if (data.enviarEmail && data.alunoEmail) {
      enviarAtaPorEmail(data);
    }

    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'success', 
      message: 'Reunião registrada com sucesso na aba reunioes-registro.' 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Função utilitária para envio da ata de reunião ao orientando
 */
function enviarAtaPorEmail(data) {
  var assunto = 'Registro de Orientação • IFSC Garopaba — ' + (data.alunoNome || 'Orientando');
  var corpo = 'Olá,\n\n' +
    'Segue o registro da reunião de orientação realizada em ' + (data.data || '') + ' às ' + (data.hora || '') + ':\n\n' +
    '• Orientando(a): ' + (data.alunoNome || '') + '\n' +
    '• Tipo: ' + (data.tipo || 'Alinhamento') + '\n' +
    '• Pauta: ' + (data.pauta || 'Não informada') + '\n\n' +
    'Deliberações e Acordos Firmados:\n' +
    (data.deliberacoes || 'Conforme alinhado em reunião.') + '\n\n' +
    'Próximos Passos:\n' +
    (data.proximosPassos || 'Dar continuidade às entregas previstas.') + '\n\n' +
    '---\n' +
    'Atenciosamente,\n' +
    'Prof. Dr. André Moraes\n' +
    'Instituto Federal de Santa Catarina (IFSC) — Câmpus Garopaba';

  MailApp.sendEmail(data.alunoEmail, assunto, corpo);
}
