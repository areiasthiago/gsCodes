//Crie uma planilha no Google Planilhas e nomeie a "Página 1" como "Sitemap"
function sitemapChecker() {
  //EDITAR
  var sitemap_url = "https://www.oiplace.com.br/sitemap_index.xml"; // Insira a URL do arquivo XML que você deseja importar
  var nomePlanilha = "Sitemap"; // Insira o nome da planilha (aba) que você deseja escrever os dados do XML
  //EDITAR
  var response = UrlFetchApp.fetch(sitemap_url);
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomePlanilha); // Obtém a planilha pelo nome
  //Checa se a planilha existe
  if (!planilha) {
    // Verifica se a planilha foi encontrada
    // Logger.log("A planilha com o nome '" + nomePlanilha + "' não foi encontrada.");
    return;
  }
  //Checa se a planilha já foi completamente preenchida hoje
  var ultimaLinha = planilha.getLastRow();
  var dataAtual = new Date(); // Obtém a data atual
  if (ultimaLinha !== 0) {
    var valorUltimaCelulaA = planilha.getRange(ultimaLinha, 1).getValue();
    var valorUltimaCelulaB = planilha.getRange(ultimaLinha, 2).getValue();
    if (valorUltimaCelulaA instanceof Date && valorUltimaCelulaA.toDateString() === dataAtual.toDateString() && valorUltimaCelulaB instanceof Date && valorUltimaCelulaB.toDateString() === dataAtual.toDateString()) {
      // Logger.log("A planilha com o nome '" + nomePlanilha + "' já foi preenchida hoje.");
      var cellB1 = planilha.getRange(1,2).getValue();
      if(cellB1 !== "Códigos"){
        planilha.getRange(1,2).setValue("Códigos")
        sumarioSitemap(planilha);
      }
      return;
    }
  }
  //Checa se o sitemap existe e retorna 200
  if (response.getResponseCode() == 200) {
    //Importa as urls do sitemap para a planilha 
    var xml = UrlFetchApp.fetch(sitemap_url).getContentText();
    var document = XmlService.parse(xml);
    var root = document.getRootElement();
    //Checa se tem a data atual está preenchida na última célula da Coluna A, se não, importa os dados
    if (!(valorUltimaCelulaA instanceof Date && valorUltimaCelulaA.toDateString() === dataAtual.toDateString())) {
      // Limpa os dados existentes na planilha
      planilha.clearContents();
      // Logger.log("Planilha foi completamente limpa. Iniciando os trabalhos.");
      //Insere os dados na planilha
      importarXML(sitemap_url,planilha,root,xml);
    }
    //Escreve os status das urls na coluna B
    escreverStatusURLs(planilha);
  }else{
    // Logger.log("Falha ao obter o sitemap. Status code: " + response.getResponseCode());
  }
}

function importarXML(sitemap_url,planilha,root,xml) {
  var ns = XmlService.getNamespace("http://www.sitemaps.org/schemas/sitemap/0.9"); // Define o namespace do XML
  var locs = []; // Array para armazenar os valores de loc
  if (xml.indexOf("<sitemap>") !== -1) { 
    var sitemaps = root.getChildren("sitemap", ns); // Obtém todos os elementos filho "url" do elemento raiz, usando o namespace
    // Logger.log("Iniciando sitemap índice - "+sitemap_url);
    // Loop para percorrer os elementos url e obter os valores de loc
    for (var j = 0; j < sitemaps.length; j++) {
      var urlElementSitemap = sitemaps[j]; // Obtém o elemento url atual
      var locElementSitemap = urlElementSitemap.getChild("loc", ns); // Obtém o elemento loc dentro do elemento url, usando o namespace
      var locValueSitemap = locElementSitemap.getValue(); // Obtém o valor de texto do elemento loc
      // Logger.log("Iniciando o sitemap único - "+locValueSitemap);
      var xmlFilho = UrlFetchApp.fetch(locValueSitemap).getContentText(); // Faz a solicitação do URL e obtém o conteúdo em formato de texto
      var documentFilho = XmlService.parse(xmlFilho); // Analisa o XML obtido em um objeto Document
      var rootFilho = documentFilho.getRootElement(); // Obtém o elemento raiz do XML
      var urls = rootFilho.getChildren("url", ns); // Obtém todos os elementos filho "url" do elemento raiz, usando o namespace
      // Logger.log(urls.length);
      // Loop para percorrer os elementos url e obter os valores de loc
      for (var i = 0; i < urls.length; i++) {
        var urlElement = urls[i]; // Obtém o elemento url atual
        var locElement = urlElement.getChild("loc", ns); // Obtém o elemento loc dentro do elemento url, usando o namespace
        var locValue = locElement.getValue(); // Obtém o valor de texto do elemento loc
        if (locValue.toString().indexOf("smplaces.") != -1) {
          locValue = locValue.toString().replace("smplaces.", "www.");
        }
        // Logger.log(locValue);
        locs.push([locValue]); // Adiciona o valor de loc ao array locs
      }
      // Logger.log("Finalizando o sitemap único - "+locValueSitemap);
    }
    // Logger.log("Finalizando sitemap índice - "+sitemap_url);
  }else{
    // Logger.log("Iniciando sitemap único - "+sitemap_url);
    var urls = root.getChildren("url", ns); // Obtém todos os elementos filho "url" do elemento raiz, usando o namespace
    // Loop para percorrer os elementos url e obter os valores de loc
    for (var i = 0; i < urls.length; i++) {
      var urlElement = urls[i]; // Obtém o elemento url atual
      var locElement = urlElement.getChild("loc", ns); // Obtém o elemento loc dentro do elemento url, usando o namespace
      var locValue = locElement.getValue(); // Obtém o valor de texto do elemento loc
      if (locValue.toString().indexOf("smplaces.") != -1) {
        locValue = locValue.toString().replace("smplaces.", "www.");
      }
      // Logger.log(locValue);
      locs.push([locValue]); // Adiciona o valor de loc ao array locs
    }
    // Logger.log("Finalizando sitemap único - "+sitemap_url);
  }
  planilha.getRange(1,1).setValue("URLs");
  planilha.getRange(2, 1, locs.length, locs[0].length).setValues(locs);
  var dataAtual = new Date();
  planilha.getRange(locs.length+2,1).setValue(dataAtual); //escreve a data na primeira linha depois do último status
}

function escreverStatusURLs(planilha) {
  var colunaA = planilha.getRange("A:A").getValues();
  var ultimaLinhaA = colunaA.filter(String).length-1;
  var colunaB = planilha.getRange("B:B").getValues();
  var ultimaLinhaB = colunaB.filter(String).length;
  if(ultimaLinhaB===0){
    planilha.getRange(1,2).setValue("Status Code")
    var linhaInicial = 2;
  }else{
    var linhaInicial = ultimaLinhaB+1;
  }
  var urls = planilha.getRange("A"+linhaInicial+":A" + ultimaLinhaA).getValues();
  // Logger.log("Iniciando captura de status code das URLs");
  // Logger.log("Começando na linha: "+linhaInicial);
  // Logger.log("Total de urls: "+urls.length);
  for (var i = 0; i < urls.length; i++) {
    var url = urls[i][0];
    var options = {
      'muteHttpExceptions': true,
      'followRedirects': false
    };
    var url_trimmed = url.trim();
    var response = UrlFetchApp.fetch(url_trimmed, options);
    var headerStatus = response.getResponseCode();
    planilha.getRange(linhaInicial+i, 2).setNumberFormat('@')
    planilha.getRange(linhaInicial+i, 2).setValue(headerStatus);
    planilha.getRange(linhaInicial+i, 2).setNumberFormat('0')
    // Logger.log(url+" - "+headerStatus);
  }
  var dataAtual = new Date();
  planilha.getRange(ultimaLinhaA+1,2).setValue(dataAtual); //escreve a data na primeira linha depois do último status
  //Faz o backup das informação eu uma outra planilha
  registroDeSitemap();
  // Logger.log("Todos os status code das URLs foram escritos na planilha.");
}

function sumarioSitemap(planilha) {
  // Definir a planilha e a coluna a serem verificadas
  var colunaB = planilha.getRange("B:B").getValues();
  var ultimaLinha = colunaB.filter(String).length;
  var penultimaLinha = colunaB.filter(String).length-1;
  var dados = planilha.getRange("B2:B"+penultimaLinha).getValues();
  // Encontrar o valor na última linha
  var valorUltimaLinha = planilha.getRange("B"+ultimaLinha+":B"+ultimaLinha).getValues();
  var fusoHorario = Session.getScriptTimeZone(); // Obtém o fuso horário do script
  var dataCell = new Date(valorUltimaLinha); // Obtém a data atual
  var dataExtraction = Utilities.formatDate(dataCell, fusoHorario, 'dd/MM/yyyy');
  // Obter o total de linhas preenchidas
  var totalLinhasPreenchidas = dados.filter(function (valor) {
    return valor[0] !== '';
  }).length;
  var totalFormatado = totalLinhasPreenchidas.toLocaleString('pt-BR');
  // Obter os valores únicos encontrados
  var valoresUnicos = [...new Set(dados.flat())];
  // Calcular a porcentagem de cada valor único referente ao total de linhas preenchidas
  var resultados = valoresUnicos.map(function (valorUnico) {
    var totalOcorrencias = dados.filter(function (valor) {
      return valor[0] === valorUnico;
    }).length;
    var porcentagem = (totalOcorrencias / totalLinhasPreenchidas) * 100;
    return { valor: valorUnico, total: totalOcorrencias, porcentagem: porcentagem };
  });;
  // Obter a data e hora atual
  var dataHoraAtual = new Date();
  var dataChecker = Utilities.formatDate(dataHoraAtual, fusoHorario, 'dd/MM/yyyy');
  var horaChecker = Utilities.formatDate(dataHoraAtual, fusoHorario, 'HH:mm');
  // Construir a mensagem com os resultados
  var mensagem = 'Data da extração do sitemap: ' + dataExtraction + '\n' +
    'URLs encontradas: ' + totalFormatado + '\n' +
    '\n<b>Status codes encontrados</b>\n';
  resultados.forEach(function (resultado) {
    mensagem += resultado.valor + ': ' + resultado.porcentagem.toFixed(2).replace(/\./g, ',') + '% (' + resultado.total.toLocaleString('pt-BR') + ')\n';
  });
  mensagem += '\n<b>Checagem realizada em ' + dataChecker+' às '+horaChecker+'</b>';
  // Enviar mensagem pelo Telegram
  alertaSitemapTelegram(mensagem);
}

function alertaSitemapTelegram(mensagem) {
  //EDITAR
  var botToken = "5878777413:AAHTOabO4RlSJimIbp0tdFhOWaKIaCCUnfM"; // Insira o token do seu bot do Telegram
  var chatId = "-932152817"; // Insira o ID do chat do Telegram para onde a mensagem será enviada
  //EDITAR
  var url = "https://api.telegram.org/bot" + botToken + "/sendMessage";
  var payload = {
    "chat_id": chatId,
    "text": mensagem,
    "parse_mode": "HTML"
  };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  try {
    UrlFetchApp.fetch(url, options);
  } catch (error) {
    Logger.log("Erro ao enviar mensagem para o Telegram: " + error.toString());
  }
}

function registroDeSitemap() {
  //EDITAR
  var planilhaOrigem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sitemap");
  var planilhaDestino = SpreadsheetApp.openById('10wAX0xkGS_IU8AXpNSIbTRMLUqq25zLSNFAFNjOs2Vg');
  //EDITAR
  var hoje = new Date();
  var nomeAba = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd-MM-yyyy");
  var aba = planilhaDestino.getSheetByName(nomeAba);
  if (!aba) {
    aba = planilhaDestino.insertSheet(nomeAba);  
    if(planilhaOrigem){
      var valores = planilhaOrigem.getDataRange().getValues();
      aba.getRange(1, 1, valores.length, valores[0].length).setValues(valores);
    }
  }
}
