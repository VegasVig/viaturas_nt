/*******************************************************************
 *  VEGAS FROTA — BACKEND (Google Apps Script)
 *  Banco de dados central em Google Sheets + Fotos no Google Drive
 *
 *  COMO INSTALAR (passo a passo no guia PDF/TXT que acompanha):
 *   1. Crie uma planilha nova no Google Sheets.
 *   2. Menu  Extensões > Apps Script.
 *   3. Apague o conteúdo e cole TODO este arquivo.
 *   4. Salve. Rode a função "primeiraInstalacao" uma vez (autorize).
 *   5. Implantar > Nova implantação > tipo "App da Web".
 *        - Executar como: Eu mesmo
 *        - Quem pode acessar: Qualquer pessoa
 *   6. Copie a URL do app da Web (termina em /exec) e cole no app HTML.
 *******************************************************************/

// ====== CONFIG ======
var SENHA_ADMIN = "Vegas4747@";       // senha do painel (igual à do app)
var PASTA_FOTOS = "VEGAS_FROTA_FOTOS"; // pasta criada no seu Drive p/ as fotos

// Abas (tabelas) do banco
var ABAS = ["usuarios","motoristas","veiculos","destinos","retiradas",
            "devolucoes","abastecimentos","manutencoes","ocorrencias","auditoria","meta"];

// ====== INSTALAÇÃO (rodar 1x) ======
function primeiraInstalacao(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ABAS.forEach(function(nome){
    var sh = ss.getSheetByName(nome);
    if(!sh){ sh = ss.insertSheet(nome); }
    // cada aba guarda: coluna A = id, coluna B = JSON do registro completo
    if(sh.getLastRow()===0){ sh.appendRow(["id","json"]); }
  });
  // remove a aba padrão "Página1"/"Sheet1" se existir e estiver vazia
  ["Página1","Sheet1","Planilha1"].forEach(function(n){
    var s=ss.getSheetByName(n);
    if(s && s.getLastRow()<=1 && ABAS.indexOf(n)<0){ try{ss.deleteSheet(s);}catch(e){} }
  });
  // pasta de fotos
  pastaFotos_();
  // meta com _seq
  var meta = ss.getSheetByName("meta");
  if(meta.getLastRow()<=1){ meta.appendRow(["_seq", JSON.stringify({v:100})]); }
  return "Instalação concluída. Agora publique como App da Web.";
}

function pastaFotos_(){
  var it = DriveApp.getFoldersByName(PASTA_FOTOS);
  var pasta = it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_FOTOS);
  // garante que a pasta é pública para leitura (fotos abrem em qualquer aparelho)
  try{ pasta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
  return pasta;
}

/* ===== TESTE RÁPIDO (rode no editor do Apps Script) =====
   Cria uma foto de teste e devolve o link. Depois de rodar:
   1) veja no "Registro de execução" o link gerado;
   2) cole o link no navegador — deve abrir a imagem (um quadrado colorido).
   Se abrir, o backend está correto. */
function testarFoto(){
  // 1x1 pixel PNG vermelho em base64
  var dataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
  var url = salvarFoto_("TESTE_"+new Date().getTime(), dataUrl);
  Logger.log("Link da foto de teste: " + url);
  return url;
}

// ====== ROTEADOR HTTP ======
function doGet(e){  return handle_(e); }
function doPost(e){ return handle_(e); }

function handle_(e){
  var out = {ok:false};
  try{
    var p = (e && e.parameter) ? e.parameter : {};
    var body = {};
    if(e && e.postData && e.postData.contents){
      try{ body = JSON.parse(e.postData.contents); }catch(err){}
    }
    var acao = body.acao || p.acao || "ping";

    if(acao==="ping"){ out = {ok:true, msg:"VEGAS FROTA backend online"}; }

    else if(acao==="pull"){                 // baixa o banco inteiro
      out = {ok:true, db: pullDB_()};
    }
    else if(acao==="push"){                 // grava/atualiza um registro
      // body: {aba, registro:{...}}
      upsert_(body.aba, body.registro);
      out = {ok:true};
    }
    else if(acao==="pushMuitos"){           // grava vários de uma vez
      (body.itens||[]).forEach(function(it){ upsert_(it.aba, it.registro); });
      out = {ok:true};
    }
    else if(acao==="apagar"){               // apaga um registro por id
      // body: {aba, id}
      apagar_(body.aba, body.id);
      out = {ok:true};
    }
    else if(acao==="substituirTabela"){     // troca a tabela inteira de uma vez
      // body: {aba, registros:[...]}  -> apaga tudo e regrava só o que veio
      substituirTabela_(body.aba, body.registros||[]);
      out = {ok:true};
    }
    else if(acao==="setSeq"){
      setMeta_("_seq", {v: body.v});
      out = {ok:true};
    }
    else if(acao==="foto"){                 // salva foto no Drive
      // body: {nome, dataUrl}
      var url = salvarFoto_(body.nome, body.dataUrl);
      out = {ok:true, url:url};
    }
    else if(acao==="login"){                // valida senha admin
      out = {ok: (body.senha===SENHA_ADMIN)};
    }
    else { out = {ok:false, erro:"ação desconhecida: "+acao}; }

  }catch(err){
    out = {ok:false, erro:String(err)};
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== BANCO (planilha) ======
function pullDB_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var db = {};
  ABAS.forEach(function(nome){
    if(nome==="meta") return;
    var sh = ss.getSheetByName(nome);
    var arr = [];
    if(sh && sh.getLastRow()>1){
      var vals = sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
      vals.forEach(function(r){
        if(r[1]){ try{ arr.push(JSON.parse(r[1])); }catch(e){} }
      });
    }
    db[nome] = arr;
  });
  db._seq = (getMeta_("_seq")||{v:100}).v;
  return db;
}

function upsert_(aba, reg){
  if(ABAS.indexOf(aba)<0) throw "aba inválida: "+aba;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aba);
  var id = reg.id;
  var last = sh.getLastRow();
  if(last>1){
    var ids = sh.getRange(2,1,last-1,1).getValues();
    for(var i=0;i<ids.length;i++){
      if(String(ids[i][0])===String(id)){
        sh.getRange(i+2,2).setValue(JSON.stringify(reg));
        return;
      }
    }
  }
  sh.appendRow([id, JSON.stringify(reg)]);
}

/* Apaga um registro pelo id (remove a linha da planilha). */
function apagar_(aba, id){
  if(ABAS.indexOf(aba)<0) throw "aba inválida: "+aba;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aba);
  var last = sh.getLastRow();
  if(last<=1) return;
  var ids = sh.getRange(2,1,last-1,1).getValues();
  for(var i=ids.length-1;i>=0;i--){          // de baixo p/ cima ao apagar
    if(String(ids[i][0])===String(id)){ sh.deleteRow(i+2); }
  }
}

/* Substitui a tabela inteira: apaga todas as linhas de dados e regrava
   apenas os registros enviados. Ideal para corrigir a lista de motoristas
   de uma vez, para todos os aparelhos (o banco central fica correto). */
function substituirTabela_(aba, registros){
  if(ABAS.indexOf(aba)<0) throw "aba inválida: "+aba;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(aba);
  var last = sh.getLastRow();
  if(last>1){ sh.deleteRows(2, last-1); }     // apaga tudo menos o cabeçalho
  (registros||[]).forEach(function(reg){
    if(reg && reg.id!=null){ sh.appendRow([reg.id, JSON.stringify(reg)]); }
  });
}

function getMeta_(chave){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("meta");
  if(!sh || sh.getLastRow()<2) return null;
  var vals = sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
  for(var i=0;i<vals.length;i++){
    if(vals[i][0]===chave){ try{return JSON.parse(vals[i][1]);}catch(e){return null;} }
  }
  return null;
}
function setMeta_(chave,obj){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("meta");
  var vals = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1,2).getValues() : [];
  for(var i=0;i<vals.length;i++){
    if(vals[i][0]===chave){ sh.getRange(i+2,2).setValue(JSON.stringify(obj)); return; }
  }
  sh.appendRow([chave, JSON.stringify(obj)]);
}

// ====== FOTOS (Drive) ======
function salvarFoto_(nome, dataUrl){
  var pasta = pastaFotos_();
  var partes = dataUrl.split(",");
  var meta = partes[0];             // data:image/jpeg;base64
  var b64  = partes[1];
  var tipo = (meta.match(/data:(.*?);/)||[])[1] || "image/jpeg";
  var bytes = Utilities.base64Decode(b64);
  var blob = Utilities.newBlob(bytes, tipo, nome+".jpg");
  var arq = pasta.createFile(blob);
  arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Link que funciona diretamente em <img> (o formato antigo uc?export=view
  // foi descontinuado pelo Google e não carrega mais em tags de imagem).
  return "https://lh3.googleusercontent.com/d/" + arq.getId();
}
