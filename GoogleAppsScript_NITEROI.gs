/*******************************************************************
 *  VEGAS FROTA — BACKEND (Google Apps Script)  ·  versão 6
 *  Banco de dados central em Google Sheets + Fotos no Google Drive
 *
 *  COMO INSTALAR (primeira vez):
 *   1. Crie uma planilha nova no Google Sheets.
 *   2. Menu  Extensões > Apps Script.
 *   3. Apague o conteúdo e cole TODO este arquivo.
 *   4. Salve. Rode a função "primeiraInstalacao" uma vez (autorize).
 *   5. Implantar > Nova implantação > tipo "App da Web".
 *        - Executar como: Eu mesmo
 *        - Quem pode acessar: Qualquer pessoa
 *   6. Copie a URL do app da Web (termina em /exec) e cole no app HTML.
 *
 *  COMO ATUALIZAR DA VERSÃO 5 PARA A 6 (planilha que já está em uso):
 *   1. Cole este arquivo no lugar do antigo e salve.
 *   2. Rode a função "atualizarParaV6" uma vez (cria a aba "baixas" e os
 *      índices que deixam a gravação rápida). NENHUM dado é apagado.
 *   3. Implantar > Gerenciar implantações > lápis (editar) >
 *      Versão: "Nova versão" > Implantar.  (a URL /exec continua a mesma)
 *
 *  O QUE MUDOU NA VERSÃO 6
 *   - Gravação muito mais rápida: o servidor não lê mais a planilha
 *     inteira a cada registro (usa colunas de índice).
 *   - Sincronização incremental: os aparelhos baixam só o que mudou.
 *   - Login validado no servidor + permissões no servidor (motorista não
 *     consegue usar funções administrativas, mesmo alterando o app).
 *   - Baixa administrativa de veículo não entregue (aba "baixas"),
 *     com histórico que NÃO pode ser alterado nem apagado.
 *******************************************************************/

// ====== CONFIG ======
var VERSAO_BACKEND = 6;
var SENHA_ADMIN = "Vegas4747@";       // senha do painel (usada só na 1ª instalação, antes de existir a aba "usuarios")
var PASTA_FOTOS = "VEGAS_FROTA_FOTOS"; // pasta criada no seu Drive p/ as fotos
var TOKEN_DIAS  = 7;                   // validade da sessão (dias)
// Níveis do painel que podem dar baixa administrativa em veículo não entregue.
// Para deixar SÓ o Administrador, troque por: ["Administrador"]
var NIVEIS_BAIXA = ["Administrador","Gestor"];

// Abas (tabelas) do banco
var ABAS = ["usuarios","motoristas","veiculos","destinos","retiradas",
            "devolucoes","abastecimentos","manutencoes","ocorrencias","calibragens",
            "baixas","auditoria","meta"];

/* Colunas de cada aba (exceto "meta"):
   A id | B json (registro completo) | C atualizadoEm (ms) |
   D veiculoId | E km | F data (ms) | G aberta (retirada sem entrega = 1) | H motoristaId
   As colunas C–H são índices: permitem achar retiradas abertas e calcular
   a quilometragem sem ler (e decodificar) a planilha inteira. */
var CABECALHO = ["id","json","atualizadoEm","veiculoId","km","data","aberta","motoristaId"];
var NCOL = CABECALHO.length;
var IMUTAVEIS = ["baixas","auditoria"];          // histórico: nunca é alterado nem apagado
var SEM_LISTA_IDS = ["auditoria","baixas"];      // tabelas que só crescem (sincronização incremental)
var CAMPO_KM = {retiradas:"kmInicial", devolucoes:"kmFinal", abastecimentos:"km", manutencoes:"km", calibragens:"km"};

// ====== CACHE POR REQUISIÇÃO (cada chamada ao app da Web começa vazia) ======
var _SH = {}, _IDX = {}, _TAB = {}, _IND = {};
function limparCache_(aba){ delete _IDX[aba]; delete _TAB[aba]; delete _IND[aba]; }

/* Pega a aba; se não existir (ex.: "baixas", criada nesta versão),
   cria na hora com o cabeçalho. Antes, uma aba faltando fazia a gravação
   falhar e o registro se perdia. */
function aba_(nome){
  if(_SH[nome]) return _SH[nome];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nome);
  if(!sh){
    sh = ss.insertSheet(nome);
    if(nome==="meta") sh.appendRow(["id","json"]);
    else sh.getRange(1,1,1,NCOL).setValues([CABECALHO]);
  } else if(nome!=="meta" && (sh.getLastRow()===0 || sh.getLastColumn()<NCOL)){
    sh.getRange(1,1,1,NCOL).setValues([CABECALHO]);
  }
  _SH[nome] = sh;
  return sh;
}

// ====== INSTALAÇÃO (rodar 1x) ======
function primeiraInstalacao(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ABAS.forEach(function(nome){ aba_(nome); });
  // remove a aba padrão "Página1"/"Sheet1" se existir e estiver vazia
  ["Página1","Sheet1","Planilha1"].forEach(function(n){
    var s=ss.getSheetByName(n);
    if(s && s.getLastRow()<=1 && ABAS.indexOf(n)<0){ try{ss.deleteSheet(s);}catch(e){} }
  });
  // pasta de fotos
  var pasta = pastaFotos_();
  try{ pasta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
  // meta com _seq
  var meta = ss.getSheetByName("meta");
  if(meta.getLastRow()<=1){ meta.appendRow(["_seq", JSON.stringify({v:100})]); }
  atualizarIndices();
  segredo_();
  return "Instalação concluída. Agora publique como App da Web.";
}

/* Atualização da versão 5 para a 6 (rodar 1x no editor). Não apaga nada:
   cria a aba "baixas", completa o cabeçalho e preenche os índices. */
function atualizarParaV6(){
  ABAS.forEach(function(nome){ aba_(nome); });
  var r = atualizarIndices();
  segredo_();
  return "Atualizado para a versão 6. " + r + " Agora reimplante (Gerenciar implantações > Nova versão).";
}

/* Preenche as colunas de índice das linhas antigas. Pode rodar quantas vezes quiser. */
function atualizarIndices(){
  var total = 0;
  ABAS.forEach(function(nome){
    if(nome==="meta") return;
    total += lerIndices_(nome).length;
  });
  return total + " registro(s) indexado(s).";
}

function pastaFotos_(){
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("PASTA_FOTOS_ID");
  if(id){ try{ return DriveApp.getFolderById(id); }catch(e){} }
  var it = DriveApp.getFoldersByName(PASTA_FOTOS);
  var pasta = it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_FOTOS);
  // garante que a pasta é pública para leitura (fotos abrem em qualquer aparelho)
  try{ pasta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
  try{ props.setProperty("PASTA_FOTOS_ID", pasta.getId()); }catch(e){}
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

// ====== SESSÃO (token assinado — não precisa guardar nada no servidor) ======
function segredo_(){
  var p = PropertiesService.getScriptProperties();
  var s = p.getProperty("TOKEN_SECRET");
  if(!s){ s = Utilities.getUuid()+Utilities.getUuid(); p.setProperty("TOKEN_SECRET", s); }
  return s;
}
function assinar_(txt){
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(txt, segredo_(), Utilities.Charset.UTF_8));
}
function criarToken_(d){
  d.exp = Date.now() + TOKEN_DIAS*864e5;
  var corpo = Utilities.base64EncodeWebSafe(JSON.stringify(d), Utilities.Charset.UTF_8);
  return corpo + "." + assinar_(corpo);
}
function lerToken_(tok){
  if(!tok || typeof tok!=="string") return null;
  var p = tok.split(".");
  if(p.length!==2 || assinar_(p[0])!==p[1]) return null;
  var d = null;
  try{ d = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString("UTF-8")); }catch(e){ return null; }
  if(!d || !d.exp || d.exp < Date.now()) return null;
  return d;   // {t:"admin"|"motorista", id, n:nome, nv:nível, exp}
}
function exigirSessao_(sess){ if(!sess) throw "SESSAO_EXPIRADA"; }
function exigirAdmin_(sess, motivo){
  exigirSessao_(sess);
  if(sess.t!=="admin") throw "SEM_PERMISSAO: "+(motivo||"função exclusiva do administrador");
  if(sess.nv==="Consulta") throw "SEM_PERMISSAO: perfil somente consulta";
}

/* Bloqueio simples contra tentativa de adivinhar senha: 8 erros em 10 min. */
function tentativasLogin_(chave, somar){
  var c = CacheService.getScriptCache(), k = "lg_"+String(chave||"").toLowerCase();
  var n = Number(c.get(k))||0;
  if(somar){ n++; c.put(k, String(n), 600); }
  return n;
}

function login_(body){
  if(body.tipo==="motorista"){
    var m = acharPorId_("motoristas", body.motoristaId);
    if(!m) return {ok:false, erro:"Motorista não encontrado. Atualize a lista e tente de novo."};
    if(m.status && m.status!=="Ativo") return {ok:false, erro:"Motorista inativo — procure o gestor da frota."};
    var sm = {t:"motorista", id:m.id, n:m.nome, nv:"Motorista"};
    return {ok:true, versao:VERSAO_BACKEND, token:criarToken_(sm),
      sessao:{id:m.id, nome:m.nome, nivel:"Motorista", tipo:"motorista"}};
  }
  if(body.login!==undefined){
    if(tentativasLogin_(body.login,false)>=8) return {ok:false, erro:"Muitas tentativas. Aguarde 10 minutos e tente de novo."};
    var us = lerTabela_("usuarios"), u = null;
    us.forEach(function(x){ if(String(x.login)===String(body.login) && String(x.senha)===String(body.senha)) u = x; });
    // planilha ainda sem usuários (1ª instalação): aceita o admin padrão
    if(!u && !us.length && String(body.login)==="admin" && body.senha===SENHA_ADMIN) u = {id:"u1", nome:"Administrador", nivel:"Administrador"};
    if(!u){ tentativasLogin_(body.login,true); return {ok:false, erro:"Login ou senha inválidos."}; }
    var sa = {t:"admin", id:u.id, n:u.nome, nv:u.nivel||"Gestor"};
    return {ok:true, versao:VERSAO_BACKEND, token:criarToken_(sa),
      sessao:{id:u.id, nome:u.nome, nivel:sa.nv, tipo:"admin"}};
  }
  // formato antigo (só a senha)
  return {ok:(body.senha===SENHA_ADMIN)};
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
    var sess = lerToken_(body.token || p.token);

    // Trava de escrita: impede que dois aparelhos gravem ao mesmo tempo
    // (sem isso, gravações simultâneas podiam se perder ou duplicar).
    var ESCRITA = ["push","pushMuitos","apagar","substituirTabela","setSeq","retirar","devolver","baixaAdministrativa"];
    var lock = null;
    if(ESCRITA.indexOf(acao)>=0){
      lock = LockService.getScriptLock();
      if(!lock.tryLock(25000)) throw "servidor ocupado, tente novamente";
    }
    try{

    if(acao==="ping"){ out = {ok:true, msg:"VEGAS FROTA backend online", versao:VERSAO_BACKEND, agora:Date.now()}; }

    else if(acao==="login"){                // valida usuário/senha (admin) ou identifica o motorista
      out = login_(body);
    }
    else if(acao==="listaMotoristas"){      // lista pública p/ a tela de login (só nome e id)
      out = {ok:true, versao:VERSAO_BACKEND, motoristas: lerTabela_("motoristas").map(function(m){
        return {id:m.id, nome:m.nome, status:m.status||"Ativo"}; })};
    }
    else if(acao==="pull"){                 // baixa o banco (inteiro ou só o que mudou)
      exigirSessao_(sess);
      out = pullDB_(Number(body.desde)||0, sess);
      out.ok = true; out.versao = VERSAO_BACKEND;
    }
    else if(acao==="push"){                 // grava/atualiza um registro
      // body: {aba, registro:{...}}
      exigirSessao_(sess);
      upsert_(body.aba, body.registro, {sess:sess});
      out = {ok:true};
    }
    else if(acao==="pushMuitos"){           // grava vários de uma vez
      // grava todos os que puder; se algum falhar, avisa (o app reenvia)
      exigirSessao_(sess);
      var falhas = [];
      (body.itens||[]).forEach(function(it){
        try{ upsert_(it.aba, it.registro, {sess:sess}); }
        catch(e){ falhas.push((it.registro&&it.registro.id)+": "+e); }
      });
      out = falhas.length ? {ok:false, erro:"falhou: "+falhas.join(" | ")} : {ok:true};
    }
    else if(acao==="apagar"){               // apaga um registro por id
      // body: {aba, id}
      exigirAdmin_(sess);
      apagar_(body.aba, body.id);
      out = {ok:true};
    }
    else if(acao==="substituirTabela"){     // troca a tabela inteira de uma vez
      // body: {aba, registros:[...]}  -> apaga tudo e regrava só o que veio
      exigirAdmin_(sess);
      substituirTabela_(body.aba, body.registros||[]);
      out = {ok:true};
    }
    else if(acao==="setSeq"){
      exigirSessao_(sess);
      var atual = (getMeta_("_seq")||{v:100}).v;
      setMeta_("_seq", {v: Math.max(Number(atual)||100, Number(body.v)||100)});
      out = {ok:true};
    }
    else if(acao==="foto"){                 // salva foto no Drive
      // body: {nome, dataUrl}
      exigirSessao_(sess);
      var url = salvarFoto_(body.nome, body.dataUrl);
      out = {ok:true, url:url};
    }
    else if(acao==="retirar"){              // retirada validada no servidor
      out = retirar_(body, sess);
    }
    else if(acao==="devolver"){             // devolução validada no servidor
      out = devolver_(body, sess);
    }
    else if(acao==="baixaAdministrativa"){  // baixa de veículo não entregue (só administrador)
      out = baixaAdministrativa_(body, sess);
    }
    else { out = {ok:false, erro:"ação desconhecida: "+acao}; }
    } finally { if(lock){ SpreadsheetApp.flush(); lock.releaseLock(); } }

  }catch(err){
    out = {ok:false, erro:String(err)};
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== BANCO (planilha) ======
/* desde = 0  -> banco inteiro (primeira vez no aparelho)
   desde > 0  -> só os registros alterados depois desse horário + a lista de
                 ids existentes (para o aparelho saber o que foi removido). */
function pullDB_(desde, sess){
  var agora = Date.now();
  var db = {}, ids = {};
  ABAS.forEach(function(nome){
    if(nome==="meta") return;
    if(nome==="usuarios" && sess.t!=="admin") return;   // motorista não recebe a lista de usuários
    var sh = aba_(nome), last = sh.getLastRow(), arr = [];
    if(last>1){
      var n = last-1;
      if(!desde){
        sh.getRange(2,1,n,2).getValues().forEach(function(r){
          if(r[1]){ try{ arr.push(JSON.parse(r[1])); }catch(e){} }
        });
      } else {
        var col = sh.getRange(2,1,n,1).getValues(), ts = sh.getRange(2,3,n,1).getValues();
        var mudou = [], lista = [];
        for(var i=0;i<n;i++){
          if(col[i][0]==="") continue;
          lista.push(String(col[i][0]));
          if((Number(ts[i][0])||0) >= desde) mudou.push(i);
        }
        if(SEM_LISTA_IDS.indexOf(nome)<0) ids[nome] = lista;
        if(mudou.length){
          var ini = mudou[0], fim = mudou[mudou.length-1];
          var js = sh.getRange(2+ini,2,fim-ini+1,1).getValues();
          mudou.forEach(function(i){ var t = js[i-ini][0]; if(t){ try{ arr.push(JSON.parse(t)); }catch(e){} } });
        }
      }
    } else if(desde && SEM_LISTA_IDS.indexOf(nome)<0){ ids[nome] = []; }
    if(nome==="usuarios") arr = arr.map(function(u){ var c = JSON.parse(JSON.stringify(u)); delete c.senha; return c; });
    db[nome] = arr;
  });
  db._seq = (getMeta_("_seq")||{v:100}).v;
  var out = {db:db, agora:agora, delta:!!desde};
  if(desde) out.ids = ids;
  return out;
}

/* Valores das colunas de índice (D–H) de um registro. */
function indicesDe_(aba, r){
  var campo = CAMPO_KM[aba];
  var data = r && r.data ? (new Date(r.data).getTime()||"") : "";
  return [
    r && r.veiculoId ? String(r.veiculoId) : "",
    campo ? (Number(r[campo])||0) : "",
    data,
    aba==="retiradas" ? (r.devolvida ? 0 : 1) : "",
    r ? String(r.motoristaId || r.responsavelId || "") : ""
  ];
}

/* Mapa id -> número da linha (lê só a coluna A). */
function idIndex_(aba){
  if(_IDX[aba]) return _IDX[aba];
  var sh = aba_(aba), last = sh.getLastRow(), m = {};
  if(last>1){
    var ids = sh.getRange(2,1,last-1,1).getValues();
    for(var i=0;i<ids.length;i++){ var k = String(ids[i][0]); if(k!=="" && !(k in m)) m[k] = i+2; }
  }
  _IDX[aba] = m;
  return m;
}
function lerLinha_(aba, linha){
  try{ return JSON.parse(aba_(aba).getRange(linha,2).getValue()); }catch(e){ return null; }
}

/* Índices de todas as linhas: [{id,row,ts,vid,km,data,aberta,mid}].
   Se encontrar linhas antigas sem índice, preenche (uma vez só). */
function lerIndices_(aba){
  if(_IND[aba]) return _IND[aba];
  var sh = aba_(aba), last = sh.getLastRow(), out = [];
  if(last<=1){ _IND[aba] = out; return out; }
  var n = last-1;
  var ids = sh.getRange(2,1,n,1).getValues();
  var ind = sh.getRange(2,3,n,NCOL-2).getValues();
  var falta = false;
  for(var i=0;i<n;i++){ if(ids[i][0]!=="" && (ind[i][0]==="" || ind[i][0]===null)){ falta = true; break; } }
  if(falta){
    var js = sh.getRange(2,2,n,1).getValues();
    for(var j=0;j<n;j++){
      if(ind[j][0]!=="" && ind[j][0]!==null) continue;
      var r = null; try{ r = JSON.parse(js[j][0]); }catch(e){}
      ind[j] = [1].concat(r ? indicesDe_(aba, r) : ["","","","",""]);   // 1 = registro antigo
    }
    sh.getRange(2,3,n,NCOL-2).setValues(ind);
  }
  for(var k=0;k<n;k++){
    if(ids[k][0]==="") continue;
    out.push({id:String(ids[k][0]), row:k+2, ts:Number(ind[k][0])||0, vid:String(ind[k][1]||""),
      km:Number(ind[k][2])||0, data:Number(ind[k][3])||0, aberta:String(ind[k][4])==="1", mid:String(ind[k][5]||"")});
  }
  _IND[aba] = out;
  return out;
}

/* Grava/atualiza um registro.
   opts.sess presente = gravação vinda do aplicativo (passa pelas permissões). */
function upsert_(aba, reg, opts){
  if(ABAS.indexOf(aba)<0 || aba==="meta") throw "aba inválida: "+aba;
  if(!reg || reg.id==null || reg.id==="") throw "registro sem id";
  var sh = aba_(aba);
  var id = String(reg.id);
  var idx = idIndex_(aba);
  var linha = idx[id] || -1, antigo = null;
  if(linha>0) antigo = lerLinha_(aba, linha);
  if(opts && opts.hasOwnProperty("sess")){
    reg = autorizar_(opts.sess, aba, reg, antigo);
    if(reg===null) return antigo;              // nada a alterar (ex.: histórico já gravado)
  }
  if(antigo){
    if(IMUTAVEIS.indexOf(aba)>=0) return antigo;   // histórico nunca é sobrescrito
    reg = protegerRegistro_(aba, antigo, reg);
  }
  var limpo = limpaBase64_(reg);
  var json = JSON.stringify(limpo);
  // limite do Google Sheets: 50.000 caracteres por célula
  if(json.length > 49000) throw "registro grande demais ("+json.length+" caracteres)";
  var valores = [json, Date.now()].concat(indicesDe_(aba, limpo));
  if(linha>0) sh.getRange(linha,2,1,valores.length).setValues([valores]);
  else { sh.appendRow([id].concat(valores)); idx[id] = sh.getLastRow(); }
  delete _IND[aba];
  if(_TAB[aba]){
    var t = _TAB[aba], achou = false;
    for(var i=0;i<t.length;i++){ if(String(t[i].id)===id){ t[i] = limpo; achou = true; break; } }
    if(!achou) t.push(limpo);
  }
  return limpo;
}

/* PERMISSÕES NO SERVIDOR — valem mesmo que alguém altere o aplicativo.
   Retorna o registro a gravar, null (ignorar) ou lança erro. */
function autorizar_(sess, aba, reg, antigo){
  exigirSessao_(sess);
  if(aba==="baixas") throw "SEM_PERMISSAO: a baixa administrativa só pode ser registrada pela função própria";
  if(sess.t==="admin"){
    if(sess.nv==="Consulta") throw "SEM_PERMISSAO: perfil somente consulta";
    if(aba==="usuarios" && sess.nv!=="Administrador") throw "SEM_PERMISSAO: somente o Administrador gerencia usuários";
    return reg;
  }
  if(sess.t!=="motorista") throw "SEM_PERMISSAO";
  var meu = function(o){
    return !!o && (String(o.motoristaId||"")===String(sess.id) || String(o.registradoPorId||"")===String(sess.id) ||
                   String(o.responsavelId||"")===String(sess.id));
  };
  switch(aba){
    case "auditoria":
      if(antigo) return null;
      if(reg.usuario!==sess.n){ reg.enviadoPor = sess.n; }
      return reg;
    case "veiculos": {
      // motorista só atualiza KM e o vínculo de uso; nunca o cadastro
      if(!antigo) throw "SEM_PERMISSAO: motorista não cadastra veículos";
      var v = JSON.parse(JSON.stringify(antigo));
      ["km","kmEm","motoristaAtualId","retiradaAtualId","ultimoMotoristaId"].forEach(function(k){ if(reg[k]!==undefined) v[k] = reg[k]; });
      var livre = function(s){ return !s || s==="Em uso" || s==="Disponível"; };
      if(livre(antigo.status) && livre(reg.status) && reg.status) v.status = reg.status;
      return v;
    }
    case "retiradas": case "devolucoes": case "abastecimentos": case "calibragens": case "ocorrencias":
      if(antigo){
        if(!meu(antigo)) return null;
        if(aba==="ocorrencias") reg.status = antigo.status;   // status da ocorrência é do gestor
        return reg;
      }
      if(!meu(reg)) throw "SEM_PERMISSAO: motorista só registra em seu próprio nome";
      return reg;
    default:
      throw "SEM_PERMISSAO: função exclusiva do administrador";
  }
}

/* Proteções contra aparelhos com dados velhos regravando por cima:
   - retirada já DEVOLVIDA (ou com baixa administrativa) nunca volta a ficar aberta;
   - a quilometragem do veículo nunca diminui (exceto correção explícita
     feita pelo gestor, marcada com kmCorrigidoEm);
   - a senha do usuário não se perde quando o aparelho não a conhece. */
function protegerRegistro_(aba, antigo, novo){
  var r = JSON.parse(JSON.stringify(novo));
  if(aba==="retiradas" && antigo.devolvida && !r.devolvida){
    r.devolvida = true;
    r.devolucaoId = r.devolucaoId || antigo.devolucaoId;
    r.devolvidaEm = r.devolvidaEm || antigo.devolvidaEm;
  }
  if(aba==="retiradas" && antigo.baixaId){
    r.devolvida = true;
    r.baixaId = antigo.baixaId;
    r.baixaAdministrativa = antigo.baixaAdministrativa;
    r.entregaNaoRegistrada = true;
    r.devolvidaEm = r.devolvidaEm || antigo.devolvidaEm;
  }
  if(aba==="veiculos"){
    var kmA = Number(antigo.km)||0, kmN = Number(r.km)||0;
    var correcaoNova = r.kmCorrigidoEm && r.kmCorrigidoEm !== antigo.kmCorrigidoEm;
    if(kmN < kmA && !correcaoNova){ r.km = kmA; r.kmEm = antigo.kmEm || r.kmEm; }
  }
  if(aba==="usuarios" && !r.senha && antigo.senha){ r.senha = antigo.senha; }
  return r;
}

/* Tira fotos em base64 (data:...) que tenham vindo por engano dentro do
   registro — elas estouravam o limite da célula e a gravação falhava. */
function limpaBase64_(o){
  if(typeof o==="string") return (o.indexOf("data:")===0 && o.length>2000) ? null : o;
  if(Array.isArray(o)) return o.map(limpaBase64_).filter(function(x){ return x!==null; });
  if(o && typeof o==="object"){
    var r={}; for(var k in o){ r[k]=limpaBase64_(o[k]); } return r;
  }
  return o;
}

/* Apaga um registro pelo id (remove a linha da planilha).
   Histórico (baixas e auditoria) é protegido: não pode ser apagado. */
function apagar_(aba, id){
  if(ABAS.indexOf(aba)<0 || aba==="meta") throw "aba inválida: "+aba;
  if(IMUTAVEIS.indexOf(aba)>=0) throw "HISTORICO_PROTEGIDO: registros de "+aba+" não podem ser apagados";
  var sh = aba_(aba);
  var last = sh.getLastRow();
  if(last<=1) return;
  var ids = sh.getRange(2,1,last-1,1).getValues();
  for(var i=ids.length-1;i>=0;i--){          // de baixo p/ cima ao apagar
    if(String(ids[i][0])===String(id)){ sh.deleteRow(i+2); }
  }
  limparCache_(aba);
}

/* Substitui a tabela inteira: apaga todas as linhas de dados e regrava
   apenas os registros enviados. Ideal para corrigir a lista de motoristas
   de uma vez, para todos os aparelhos (o banco central fica correto).
   Só para cadastros — nunca para movimentos ou histórico. */
function substituirTabela_(aba, registros){
  if(["motoristas","destinos","veiculos"].indexOf(aba)<0) throw "HISTORICO_PROTEGIDO: a tabela "+aba+" não pode ser substituída";
  var sh = aba_(aba);
  var last = sh.getLastRow();
  if(last>1){ sh.deleteRows(2, last-1); }     // apaga tudo menos o cabeçalho
  var agora = Date.now();
  var linhas = (registros||[]).filter(function(reg){ return reg && reg.id!=null; }).map(function(reg){
    var limpo = limpaBase64_(reg);
    return [reg.id, JSON.stringify(limpo), agora].concat(indicesDe_(aba, limpo));
  });
  if(linhas.length) sh.getRange(2,1,linhas.length,NCOL).setValues(linhas);
  limparCache_(aba);
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
  var sh = aba_("meta");
  var vals = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1,2).getValues() : [];
  for(var i=0;i<vals.length;i++){
    if(vals[i][0]===chave){ sh.getRange(i+2,2).setValue(JSON.stringify(obj)); return; }
  }
  sh.appendRow([chave, JSON.stringify(obj)]);
}

// ====== RETIRADA / DEVOLUÇÃO VALIDADAS NO SERVIDOR ======
/* Estas ações rodam dentro da trava (LockService). Assim, mesmo com
   vários celulares ao mesmo tempo, é IMPOSSÍVEL existir duas retiradas
   abertas para o mesmo veículo, ou devolver um veículo que não está retirado. */

/* Quilometragem atual: maior leitura registrada (veículo, retiradas,
   devoluções, abastecimentos, manutenções, calibragens). Se o gestor fez uma
   correção manual (kmCorrigidoEm), só contam as leituras posteriores a ela.
   Versão 6: usa as colunas de índice (não decodifica a planilha inteira). */
function kmAtualServidor_(vid, veiculo){
  var corte = veiculo && veiculo.kmCorrigidoEm ? new Date(veiculo.kmCorrigidoEm).getTime() : 0;
  var km = veiculo ? (Number(veiculo.km)||0) : 0;
  vid = String(vid);
  Object.keys(CAMPO_KM).forEach(function(aba){
    lerIndices_(aba).forEach(function(x){
      if(x.vid!==vid || !x.km) return;
      if(corte && (x.data||0) <= corte) return;
      if(x.km>km) km = x.km;
    });
  });
  return km;
}
function acharPorId_(aba, id){
  if(id==null || id==="") return null;
  var linha = idIndex_(aba)[String(id)];
  return linha ? lerLinha_(aba, linha) : null;
}
function nomeMotorista_(id){ var m = acharPorId_("motoristas", id); return m ? m.nome : ""; }
function retiradasAbertasIdx_(vid){
  vid = String(vid);
  return lerIndices_("retiradas").filter(function(x){ return x.vid===vid && x.aberta; })
    .sort(function(a,b){ return (b.data||0)-(a.data||0); });
}

function retirar_(body, sess){
  exigirSessao_(sess);
  var reg = body.retirada;
  if(!reg || !reg.id || !reg.veiculoId) return {ok:false, erro:"Retirada inválida"};
  if(sess.t==="admin" && sess.nv==="Consulta") return {ok:false, erro:"SEM_PERMISSAO: perfil somente consulta"};
  if(sess.t==="motorista" && String(reg.motoristaId)!==String(sess.id))
    return {ok:false, erro:"SEM_PERMISSAO: o motorista só pode retirar viatura em seu próprio nome"};
  var veiculo = acharPorId_("veiculos", reg.veiculoId);
  if(!veiculo) return {ok:false, erro:"VEICULO_NAO_ENCONTRADO"};
  if(idIndex_("retiradas")[String(reg.id)]) return {ok:true, jaRegistrada:true, veiculo:veiculo};  // reenvio
  var abertas = retiradasAbertasIdx_(reg.veiculoId);
  if(abertas.length){
    var aberta = lerLinha_("retiradas", abertas[0].row) || {};
    return {ok:false, erro:"VEICULO_EM_USO", conflito:{retiradaId:aberta.id, motoristaId:aberta.motoristaId,
      motoristaNome: aberta.motoristaNome || nomeMotorista_(aberta.motoristaId), data:aberta.data, kmInicial:aberta.kmInicial}};
  }
  var kmAt = kmAtualServidor_(reg.veiculoId, veiculo);
  if((Number(reg.kmInicial)||0) < kmAt) return {ok:false, erro:"KM_INVALIDO", kmAtual:kmAt};

  reg.devolvida = false;
  upsert_("retiradas", reg);
  veiculo.status = "Em uso";
  veiculo.km = Number(reg.kmInicial)||kmAt; veiculo.kmEm = reg.data;
  veiculo.motoristaAtualId = reg.motoristaId;
  veiculo.retiradaAtualId = reg.id;
  upsert_("veiculos", veiculo);
  gravarExtras_(body.extras, sess);
  return {ok:true, veiculo:acharPorId_("veiculos", reg.veiculoId)};
}

function devolver_(body, sess){
  exigirSessao_(sess);
  var dev = body.devolucao;
  if(!dev || !dev.id || !dev.retiradaId) return {ok:false, erro:"Devolução inválida"};
  if(sess.t==="admin" && sess.nv==="Consulta") return {ok:false, erro:"SEM_PERMISSAO: perfil somente consulta"};
  var r = acharPorId_("retiradas", dev.retiradaId);
  if(!r && body.retirada && body.retirada.id===dev.retiradaId){
    // retirada feita sem internet que ainda não tinha chegado à planilha
    r = body.retirada; r.devolvida = false;
    if(sess.t==="motorista" && String(r.motoristaId)!==String(sess.id)) return {ok:false, erro:"SEM_PERMISSAO"};
    upsert_("retiradas", r);
  }
  if(!r) return {ok:false, erro:"RETIRADA_NAO_ENCONTRADA"};
  if(sess.t==="motorista" && String(r.motoristaId)!==String(sess.id) && String(r.registradoPorId||"")!==String(sess.id))
    return {ok:false, erro:"SEM_PERMISSAO: esta viatura está no nome de outro motorista"};
  if(r.devolvida){
    if(r.devolucaoId===dev.id) return {ok:true, jaRegistrada:true, veiculo:acharPorId_("veiculos", r.veiculoId)};
    if(r.baixaId) return {ok:false, erro:"BAIXA_ADMINISTRATIVA", baixaId:r.baixaId};
    return {ok:false, erro:"JA_DEVOLVIDA", devolucaoId:r.devolucaoId};
  }
  if((Number(dev.kmFinal)||0) < (Number(r.kmInicial)||0)) return {ok:false, erro:"KM_INVALIDO", kmAtual:Number(r.kmInicial)||0};

  upsert_("devolucoes", dev);
  r.devolvida = true; r.devolucaoId = dev.id; r.devolvidaEm = dev.data;
  upsert_("retiradas", r);

  var veiculo = acharPorId_("veiculos", r.veiculoId);
  if(veiculo){
    var aindaAberta = retiradasAbertasIdx_(r.veiculoId).length>0;
    if(!aindaAberta){
      if(veiculo.status==="Em uso" || !veiculo.status) veiculo.status = "Disponível";
      veiculo.motoristaAtualId = ""; veiculo.retiradaAtualId = "";
    }
    if((Number(dev.kmFinal)||0) > (Number(veiculo.km)||0)){ veiculo.km = Number(dev.kmFinal); veiculo.kmEm = dev.data; }
    veiculo.ultimoMotoristaId = r.motoristaId;
    upsert_("veiculos", veiculo);
  }
  gravarExtras_(body.extras, sess);
  return {ok:true, veiculo: veiculo ? acharPorId_("veiculos", veiculo.id) : null};
}
function gravarExtras_(extras, sess){
  (extras||[]).forEach(function(it){
    if(it && it.aba && it.registro && it.aba!=="retiradas" && it.aba!=="devolucoes") upsert_(it.aba, it.registro, {sess:sess});
  });
}

// ====== BAIXA ADMINISTRATIVA — VEÍCULO NÃO ENTREGUE ======
/* Registra que a entrega deveria ter acontecido e não foi registrada.
   - Quem fez, data e hora vêm do SERVIDOR (não podem ser forjados pelo app).
   - Fecha a retirada aberta (ela continua no histórico, marcada).
   - O registro da baixa é imutável: não pode ser alterado nem apagado. */
function baixaAdministrativa_(body, sess){
  exigirAdmin_(sess, "somente o administrador pode dar baixa em veículo não entregue");
  if(NIVEIS_BAIXA.indexOf(sess.nv)<0) return {ok:false, erro:"SEM_PERMISSAO: seu perfil ("+sess.nv+") não pode dar baixa administrativa"};
  var b = body.baixa || {};
  if(!b.id || !b.veiculoId) return {ok:false, erro:"Informe o veículo."};
  if(!String(b.motivo||"").trim()) return {ok:false, erro:"Informe o motivo da não entrega."};

  var ja = acharPorId_("baixas", b.id);
  if(ja) return {ok:true, jaRegistrada:true, baixa:ja,
    retirada: ja.retiradaId ? acharPorId_("retiradas", ja.retiradaId) : null, veiculo: acharPorId_("veiculos", ja.veiculoId)};

  var v = acharPorId_("veiculos", b.veiculoId);
  if(!v) return {ok:false, erro:"VEICULO_NAO_ENCONTRADO"};
  var r = null;
  if(b.retiradaId){ r = acharPorId_("retiradas", b.retiradaId); if(r && String(r.veiculoId)!==String(b.veiculoId)) r = null; }
  if(!r){
    var abertas = retiradasAbertasIdx_(b.veiculoId);
    if(abertas.length) r = lerLinha_("retiradas", abertas[0].row);
  }
  if(r && r.devolvida){
    if(r.baixaId) return {ok:false, erro:"JA_BAIXADA", baixaId:r.baixaId};
    return {ok:false, erro:"JA_DEVOLVIDA: a entrega desta retirada já foi registrada pelo motorista."};
  }

  var agora = new Date().toISOString();
  var motId = String(b.motoristaId || (r && r.motoristaId) || "");
  var statusApos = ["Disponível","Manutenção","Indisponível"].indexOf(b.statusVeiculoApos)>=0 ? b.statusVeiculoApos : "Disponível";
  var reg = {
    id: String(b.id),
    tipo: "Baixa administrativa",
    situacao: "Não entregue",
    status: "Baixa realizada",
    veiculoId: v.id, placa: v.placa||"", veiculoInterno: v.interno||"",
    motoristaId: motId,
    motoristaNome: nomeMotorista_(motId) || b.motoristaNome || (r && r.motoristaNome) || "",
    dataReferencia: b.dataReferencia || (r && r.data) || agora,
    motivo: String(b.motivo).trim(),
    observacao: String(b.observacao||"").trim(),
    statusVeiculoApos: statusApos,
    retiradaId: r ? r.id : "",
    retiradaData: r ? (r.data||"") : "",
    kmInicial: r ? (Number(r.kmInicial)||0) : "",
    destinoNome: r ? (r.destinoNome||"") : "",
    motoristaNaoRegistrouEntrega: true,
    responsavelBaixaId: sess.id,
    responsavelBaixaNome: sess.n,
    responsavelBaixaNivel: sess.nv,
    registradoEm: agora,
    registradoEmCliente: b.registradoEmCliente || ""
  };
  upsert_("baixas", reg);

  if(r){
    r.devolvida = true;
    r.devolvidaEm = agora;
    r.baixaId = reg.id;
    r.entregaNaoRegistrada = true;
    r.baixaAdministrativa = {id:reg.id, por:sess.n, porId:sess.id, em:agora, motivo:reg.motivo};
    upsert_("retiradas", r);
  }
  var aindaAberta = retiradasAbertasIdx_(v.id).length>0;
  if(!aindaAberta){ v.status = statusApos; v.motoristaAtualId = ""; v.retiradaAtualId = ""; }
  if(r) v.ultimoMotoristaId = r.motoristaId;
  upsert_("veiculos", v);

  var aud = {id:"aud_"+reg.id, usuario:sess.n, nivel:sess.nv, data:agora,
    acao:"Baixa administrativa — veículo não entregue",
    registro:(v.interno||"")+" · "+(v.placa||"")+" · motorista: "+(reg.motoristaNome||"—")+" · motivo: "+reg.motivo+
      (reg.observacao ? " · obs: "+reg.observacao : "")};
  upsert_("auditoria", aud);
  return {ok:true, baixa:reg, retirada:r, veiculo:acharPorId_("veiculos", v.id), auditoria:aud};
}

// ====== FOTOS (Drive) ======
function salvarFoto_(nome, dataUrl){
  var pasta = pastaFotos_();
  var partes = String(dataUrl||"").split(",");
  var meta = partes[0];             // data:image/jpeg;base64
  var b64  = partes[1];
  if(!b64) throw "foto inválida";
  var tipo = (meta.match(/data:(.*?);/)||[])[1] || "image/jpeg";
  var bytes = Utilities.base64Decode(b64);
  var blob = Utilities.newBlob(bytes, tipo, nome+".jpg");
  var arq = pasta.createFile(blob);
  arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Link que funciona diretamente em <img> (o formato antigo uc?export=view
  // foi descontinuado pelo Google e não carrega mais em tags de imagem).
  return "https://lh3.googleusercontent.com/d/" + arq.getId();
}

/*******************************************************************
 *  CORREÇÃO AUTOMÁTICA DO STATUS DA FROTA (a cada 3 horas)
 *  Faz exatamente o que o botão "Corrigir status da frota" faz no app,
 *  mas sozinho, direto na planilha — mesmo com ninguém usando o app.
 *
 *  ATIVAR: no editor do Apps Script, escolha a função
 *          "instalarCorrecaoAutomatica" e clique em Executar (1 vez só).
 *  DESATIVAR: rode "removerCorrecaoAutomatica".
 *******************************************************************/
function instalarCorrecaoAutomatica(){
  removerCorrecaoAutomatica();                 // evita gatilho duplicado
  ScriptApp.newTrigger("corrigirStatusFrotaAuto").timeBased().everyHours(3).create();
  var r = corrigirStatusFrotaAuto();           // já roda uma vez agora
  return "Correção automática ativada (a cada 3 horas). Primeira execução: "+r;
}
function removerCorrecaoAutomatica(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()==="corrigirStatusFrotaAuto") ScriptApp.deleteTrigger(t);
  });
  return "Correção automática desativada.";
}

function lerTabela_(nome){
  if(_TAB[nome]) return _TAB[nome];
  var sh = aba_(nome);
  var arr = [];
  if(sh && sh.getLastRow()>1){
    sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(function(r){
      if(r[1]){ try{ arr.push(JSON.parse(r[1])); }catch(e){} }
    });
  }
  _TAB[nome] = arr;
  return arr;
}

function corrigirStatusFrotaAuto(){
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(60000)) return "servidor ocupado — tenta de novo no próximo ciclo";
  try{
    var veiculos   = lerTabela_("veiculos");
    var retiradas  = lerTabela_("retiradas");
    var devolucoes = lerTabela_("devolucoes");
    var retAlteradas = {}, veiAlterados = {}, detalhes = [];

    veiculos.forEach(function(v){
      var abertas = retiradas.filter(function(r){ return r.veiculoId===v.id && !r.devolvida; });

      /* Só fecha uma retirada aberta se EXISTIR uma devolução que aponta
         para ela (retiradaId). A regra antiga fechava qualquer retirada mais
         antiga que a última devolução do veículo — com relógios diferentes
         entre celulares, isso apagava o vínculo do motorista que estava
         com o veículo de verdade. */
      abertas.forEach(function(r){
        var dev = devolucoes.filter(function(d){ return d.retiradaId===r.id; })[0];
        if(dev){
          r.devolvida = true;
          r.devolucaoId = r.devolucaoId || dev.id;
          r.devolvidaEm = r.devolvidaEm || dev.data;
          retAlteradas[r.id] = r;
          detalhes.push((v.interno||v.id)+": retirada com devolução registrada foi fechada");
        }
      });

      var aindaAberta = retiradas.some(function(r){ return r.veiculoId===v.id && !r.devolvida; });
      var antes = v.status;
      if(aindaAberta){
        if(v.status!=="Manutenção" && v.status!=="Indisponível") v.status = "Em uso";
      } else if(v.status==="Em uso"){
        v.status = "Disponível";
      }
      if(v.status!==antes){
        veiAlterados[v.id] = v;
        detalhes.push((v.interno||v.id)+": "+antes+" → "+v.status);
      }
    });

    Object.keys(retAlteradas).forEach(function(id){ upsert_("retiradas", retAlteradas[id]); });
    Object.keys(veiAlterados).forEach(function(id){ upsert_("veiculos", veiAlterados[id]); });

    if(detalhes.length){
      upsert_("auditoria", {
        id: "auto" + new Date().getTime(),
        usuario: "Sistema (automático)", nivel: "Sistema",
        data: new Date().toISOString(),
        acao: "Correção automática da frota",
        registro: detalhes.length+" ajuste(s): "+detalhes.join("; ")
      });
    }
    SpreadsheetApp.flush();
    Logger.log(detalhes.length ? detalhes.join("\n") : "Nada a corrigir");
    return detalhes.length + " ajuste(s)";
  } finally {
    lock.releaseLock();
  }
}
