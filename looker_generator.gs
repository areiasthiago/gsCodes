var arquivo = SpreadsheetApp.getActiveSpreadsheet();
var planilhas = arquivo.getSheets();
var sumario = arquivo.getSheetByName("Sumário");

function criarSumario() {
  var nomesPlanilhas = [];
  for (var i = 0; i < planilhas.length; i++) {
    var planilha = planilhas[i];
    var padraoData = /^\d{2}-\d{2}-\d{4}$/;
    if (padraoData.test(planilha.getName())) {
      nomesPlanilhas.push(planilha.getName());
    }
  }

  var colunaA = sumario.getRange("A:A").getValues();
  var ultimaLinha = colunaA.filter(String).length + 1;

  if(ultimaLinha === 1){
    sumario.getRange("A1").setValue("Data");
    var ultimaLinha = 2;
  }

  nomesPlanilhas.sort(function(a, b) {
    var dataA = getDataDaPlanilha(a);
    var dataB = getDataDaPlanilha(b);
    if (dataA < dataB) {
      return -1;
    }
    if (dataA > dataB) {
      return 1;
    }
    return 0;
  });
  
  var colunaAValores = colunaA.flat();
  var nomesPlanilhas = converterDatas(nomesPlanilhas);
  var novosNomes = nomesPlanilhas.filter(function(nome) {
    var nomeString = nome.toString();
    return colunaAValores.findIndex(function(data) {
      return data.toString() === nomeString;
    }) === -1;
  });
  

  if (novosNomes.length > 0) {
    var range = sumario.getRange(ultimaLinha, 1, novosNomes.length, 1);
    range.setValues(novosNomes.map(function(nome) { return [nome]; }));
  }

  preencherPlanilha("B");
  preencherCelulasVazias();
  apagarAbas();
}

function getDataDaPlanilha(nome) {
  var partes = nome.split("-");
  var dia = partes[0];
  var mes = partes[1];
  var ano = partes[2];
  return new Date(ano, mes - 1, dia);
}

function converterDatas(array) {
  var datasConvertidas = [];
  
  for (var i = 0; i < array.length; i++) {
    var dataString = array[i];
    var partesData = dataString.split('-');
    
    // O mês em JavaScript começa em 0, então subtraímos 1 do valor do mês
    var data = new Date(partesData[2], partesData[1] - 1, partesData[0]);
    
    datasConvertidas.push(data);
  }
  
  return datasConvertidas;
}

function preencherPlanilha(colunaAlvo) {
  var lastRowA = sumario.getRange("A:A").getValues().filter(String).length;
  var lastRowB = sumario.getRange("B:B").getValues().filter(String).length+1;
  if(lastRowB === 1){
    var lastRowB = 2;
  }else if(lastRowB > lastRowA){
    var lastRowB = lastRowB-1;
  }
  var dataRange = sumario.getRange("A"+lastRowB+":A"+lastRowA);
  var datas = dataRange.getValues();


  for (var i = 0; i < datas.length; i++) {
    var data = Utilities.formatDate(datas[i][0], Session.getScriptTimeZone(), "dd-MM-yyyy");
    var planilha = arquivo.getSheetByName(data);
    var linhaEscrita = i+lastRowB;

    if (planilha) {

      var valoresColuna = planilha.getRange(colunaAlvo + "2:" + colunaAlvo).getValues();

      var valoresUnicos = [];

      for (var j = 0; j < valoresColuna.length; j++) {
        var valor = valoresColuna[j][0];
        var statusType = valor.toString()[0] + "xx";
        if (valor !== "" && valoresUnicos.indexOf(statusType) === -1 && !(valor instanceof Date)) {
          valoresUnicos.push(statusType);
        }
      }

      sumario.getRange(1, 2).setValue("Total");

      for (var k = 0; k < valoresUnicos.length; k++) {
        var valorUnico = valoresUnicos[k];
        // var colunaEscrita = k + 2;        

        var total = -1;
        var contador = [];
        contador[k] = 0;

        var statusTypes = [];

        for (var j = 0; j < valoresColuna.length; j++) {
          var valor = valoresColuna[j][0];
          var statusType = valor.toString()[0] + "xx";
          
          if (statusType === valorUnico) {
            contador[k]++;
            if(statusTypes.indexOf(statusType) === -1){
              statusTypes.push(statusType);
            } 
          }

          total++;

        }

        
        var novaColunaEscrita = checkHeaders(statusTypes[0]);

        if(statusTypes.length > 0 && novaColunaEscrita){
          
          sumario.getRange(linhaEscrita, novaColunaEscrita).setValue(contador[k]);
        }else{
          var colunaEscrita = sumario.getLastColumn();
          sumario.getRange(1, colunaEscrita + 1).setValue(valorUnico);
          sumario.getRange(linhaEscrita, colunaEscrita + 1).setValue(contador[k]);
        }
        sumario.getRange(linhaEscrita, 2).setValue(total);
      }
    }
  }
}

function checkHeaders(statusType){
  var headers = sumario.getRange(1, 1, 1, sumario.getLastColumn()).getValues()[0];
  for (var coluna = 0; coluna < headers.length; coluna++) {
    var valorCelula = headers[coluna];

    if (valorCelula === statusType) {
      var colunaEncontrada = coluna + 1;
      return colunaEncontrada;
    }
  }
}

function apagarAbas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var planilhas = ss.getSheets();
  var datas = [];
  var padraoData = /^\d{2}-\d{2}-\d{4}$/;

  // Filtrar as abas que são nomeadas como datas
  for (var i = 0; i < planilhas.length; i++) {
    var aba = planilhas[i];
    if (padraoData.test(aba.getName())) {
      datas.push(new Date(aba.getName().replace(/(\d{2})-(\d{2})-(\d{4})/, "$2/$1/$3")));
    }
  }

  // Encontrar a data mais recente
  var dataMaisRecente = new Date(Math.max.apply(null, datas));
  var dataMaisRecenteFormatada = Utilities.formatDate(dataMaisRecente, ss.getSpreadsheetTimeZone(), "dd-MM-yyyy");

  // Excluir as abas que são nomeadas como datas, exceto a data mais recente
  for (var j = 0; j < planilhas.length; j++) {
    var abaAtual = planilhas[j];
    var nomeAbaAtual = abaAtual.getName();

    if (padraoData.test(nomeAbaAtual)) {
      var dataAbaAtual = new Date(nomeAbaAtual.replace(/(\d{2})-(\d{2})-(\d{4})/, "$2/$1/$3"));
      var dataAbaAtualFormatada = Utilities.formatDate(dataAbaAtual, ss.getSpreadsheetTimeZone(), "dd-MM-yyyy");

      if (dataAbaAtualFormatada !== dataMaisRecenteFormatada) {
        ss.setActiveSheet(abaAtual);
        ss.deleteActiveSheet();
        Logger.log(dataAbaAtualFormatada+" foi deletada!");
      }
    }
  }
}

function preencherCelulasVazias() {
  var sumario = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ultimaColuna = sumario.getLastColumn();
  var ultimaLinha = sumario.getLastRow();
  var valores = sumario.getDataRange().getValues();
  
  for (var i = 0; i < ultimaLinha; i++) {
    for (var j = 0; j < ultimaColuna; j++) {
      if (valores[i][j] === "") {
        valores[i][j] = 0;
      }
    }
  }
  sumario.getDataRange().setValues(valores);
}
