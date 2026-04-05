// Configuração do CORS (Vital para os formulários conseguirem comunicar com o Apps Script)
function doOptions(e) {
  // O Apps Script não precisa de headers manuais para OPTIONS
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    // Fazer parse dos dados enviados pelo formulário web
    const payload = JSON.parse(e.postData.contents);
    
    // Pasta onde vamos salvar os anexos
    const folderName = "EmanaForms_Uploads";
    let folders = DriveApp.getFoldersByName(folderName);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    let linksFicheiros = [];

    // 1. Lidar com os Anexos (Converter Base64 para Arquivo no Drive)
    if (payload.attachments && payload.attachments.length > 0) {
      payload.attachments.forEach(function(att) {
        // Criar ficheiro binário
        let blob = Utilities.newBlob(Utilities.base64Decode(att.data), att.mimeType, att.fileName);
        let file = folder.createFile(blob);
        
        // Mudar a permissão para qualquer pessoa com o link poder ver
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // Guardar o nome e URL gerado
        linksFicheiros.push({
          name: att.fileName,
          url: file.getUrl()
        });
      });
    }

    // 2. Construir o Design do E-mail em HTML
    let emailHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #2563eb; color: white; padding: 20px; text-align: center;">
          <h2 style="margin: 0; font-size: 20px;">Nova Submissão Recebida</h2>
          <p style="margin: 5px 0 0 0; opacity: 0.9;">Tabela: <b>${payload.tableName}</b></p>
        </div>
        <div style="padding: 24px; background-color: #f8fafc;">
          <table style="width: 100%; border-collapse: collapse;">
    `;

    // Inserir os dados de texto
    for (let key in payload.data) {
      emailHtml += `
        <tr>
          <td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #475569; width: 40%;">${key}:</td>
          <td style="padding: 10px; border-bottom: 1px solid #e2e8f0; color: #0f172a;">${payload.data[key]}</td>
        </tr>
      `;
    }
    emailHtml += `</table></div>`;

    // Inserir os Links dos Anexos no E-mail (se existirem)
    if (linksFicheiros.length > 0) {
      emailHtml += `
        <div style="padding: 24px; background-color: #ffffff; border-top: 2px dashed #e2e8f0;">
          <h3 style="margin-top: 0; color: #2563eb;">📎 Ficheiros Anexados (Drive)</h3>
          <ul style="list-style-type: none; padding: 0; margin: 0;">
      `;
      
      linksFicheiros.forEach(function(link) {
        emailHtml += `
          <li style="margin-bottom: 10px;">
            <a href="${link.url}" target="_blank" style="display: inline-block; padding: 10px 15px; background-color: #eff6ff; color: #1d4ed8; text-decoration: none; border-radius: 6px; font-weight: 500; border: 1px solid #bfdbfe;">
              Ver ${link.name}
            </a>
          </li>
        `;
      });
      emailHtml += `</ul></div>`;
    }

    emailHtml += `
        <div style="padding: 15px; text-align: center; font-size: 12px; color: #94a3b8; background-color: #f1f5f9;">
          Este e-mail foi gerado automaticamente pelo motor V-mailJS / EmanaForms Pro.
        </div>
      </div>
    `;

    // 3. Enviar o E-mail
    let destEmails = payload.emails.join(",");
    if (destEmails) {
      MailApp.sendEmail({
        to: destEmails,
        subject: "Nova Submissão: " + payload.tableName,
        htmlBody: emailHtml
      });
    }

    // 4. Retornar resposta de Sucesso para o Formulário web (com CORS habilitado)
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // Em caso de erro
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
