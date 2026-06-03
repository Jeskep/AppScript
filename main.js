function enviarEmailAtrasos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const colunaEmail = 1;   // Coluna A - E-mail do avaliador
  const colunaEquipamento = 2;    // Coluna B - Produto ou descrição do item
  const colunaNumeroReclamacao = 3;  // Coluna C - Número do item
  const colunaDefeitoAlegado = 4; // Coluna D - Número de série
  const colunaDiasAguardando = 20; // Tempo de análise
  const colunaStatusAnalise = 21; // Coluna do status de atraso ("Atrasado" ou "No prazo")
  const colunaStatus = 9; // Nova coluna onde estão "Concluído", "Em Execução" e "Bloqueado"
  const ultimaLinha = sheet.getLastRow();
  let atrasosPorAvaliador = {};  // Objeto para armazenar atrasos agrupados por e-mail

  for (let i = 2; i <= ultimaLinha; i++) {
    const emailAvaliador = sheet.getRange(i, colunaEmail).getValue();
    const item = sheet.getRange(i, colunaEquipamento).getValue();
    const numeroItem = sheet.getRange(i, colunaNumeroReclamacao).getValue();
    const numeroSerie = sheet.getRange(i, colunaDefeitoAlegado
).getValue();
    const status = sheet.getRange(i, colunaStatus).getValue(); 
    const diasAguardando = sheet.getRange(i, colunaDiasAguardando).getValue(); 
    const statusAnalise = sheet.getRange(i, colunaStatusAnalise).getValue();

  }

  // **Envia e-mails somente para itens atrasados e não concluídos**
  for (let email in atrasosPorAvaliador) 
  {
      if (!email) continue;

      let corpoMensagem = '<p><strong>Os seguintes itens estão atrasados e ainda não foram concluídos:</strong></p>';
      corpoMensagem += '<table border="1" cellpadding="5" cellspacing="1">';
      corpoMensagem += '<tr><th>Equipamento</th><th>N° da Reclamação</th><th>Defeito Alegado</th><th>Dias Aguardando</th><th>Status</th></tr>';

      atrasosPorAvaliador[email].forEach(atraso => {   
      corpoMensagem += `<tr>
                          <td>${atraso.item}</td>
                          <td>${atraso.numeroItem}</td>
                          <td>${atraso.numeroSerie}</td>
                          <td style="color: red;">${atraso.diasAguardando}</td>
                          <td>${atraso.status}</td>
                          </tr>`;
      });

      corpoMensagem += '</table>';

      try 
      {
        MailApp.sendEmail
          ({
          to: "matheus.jeske@lifemed.com.br",
          subject: "Vai trabalhar gordinho",
          htmlBody: corpoMensagem
          });
        Logger.log(`E-mail enviado para ${email}`);
      
      } catch (error) 
        { 
          Logger.log(`Falha ao enviar e-mail para ${email}: ${error.message}`);
        }

  } 

  let _corpoMensagem = ""

  for (var _email in atrasosPorAvaliador) 
    {
      if (!_email) continue;

      _corpoMensagem += `<p><strong>Os seguintes itens estão atrasados e ainda não foram concluídos para o colaborador ${_email}:</strong></p>`;
      _corpoMensagem += '<table border="1" cellpadding="5" cellspacing="1">';
      /*
      _corpoMensagem += ` <tr>
                            <th style="text-align: left;">E-mail do Colaborador</th>
                            <td colspan="4" style="text-align: left;">${_email}</td>
                          </tr>`
      */
      _corpoMensagem += '<tr><th>Equipamento</th><th>N° da Reclamação</th><th>Defeito Alegado</th><th>Dias Aguardando</th><th>Status</th></tr>';

      atrasosPorAvaliador[_email].forEach(atraso => 
      {   
      _corpoMensagem += `<tr>
                          <td>${atraso.item}</td>
                          <td>${atraso.numeroItem}</td>
                          <td>${atraso.numeroSerie}</td>
                          <td style="color: red;">${atraso.diasAguardando}</td>
                          <td>${atraso.status}</td>
                          </tr>`;
      });
      _corpoMensagem += '</table>';
    }

  _corpoMensagem += '</table>';

  try 
  {
      MailApp.sendEmail({
      to: "email do colaborador",
      subject: "Resumo de Atividades",
      htmlBody: _corpoMensagem
      });

      Logger.log(`E-mail enviado para ${_email}`);
      
  } catch (error) 
    {
          Logger.log(`Falha ao enviar e-mail para ${_email}: ${error.message}`);
    }


  Logger.log("Execução finalizada.");
}
