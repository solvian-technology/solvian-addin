/*
 * Botão "Formatar Solvian" do Outlook (comando de função, sem painel).
 * Lê o rascunho aberto, aplica a formatação e a assinatura Solvian e escreve o resultado de volta.
 * Nada é enviado: a pessoa revê e carrega em Enviar.
 */
/* global Office, SolvianCore */

var BASE = new URL('.', window.location.href).href; // pasta onde está alojado o suplemento

// Imagens da assinatura: ficheiro na pasta do suplemento -> nome do anexo inline (cid)
var IMAGENS = {
  FAIXA1: { caminho: 'img/faixa-clientes-1.png', nome: 'solvian-faixa-1.png' },
  FAIXA2: { caminho: 'img/faixa-clientes-2.png', nome: 'solvian-faixa-2.png' },
  LOGO: { caminho: 'img/logo-solvian-group.png', nome: 'solvian-logo.png' }
};

Office.onReady(function () {});

// ---------- Office.js em forma de promessa ----------
function chamar(fn) {
  return new Promise(function (resolve, reject) {
    fn(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
      else reject(new Error((r.error && r.error.message) || 'Erro do Outlook'));
    });
  });
}

function aviso(mensagem, tipo) {
  var item = Office.context.mailbox.item;
  var obj = tipo === 'erro'
    ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: mensagem }
    : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, message: mensagem, icon: 'Icon.16x16', persistent: false };
  return new Promise(function (resolve) {
    try { item.notificationMessages.replaceAsync('solvian', obj, function () { resolve(); }); } catch (e) { resolve(); }
  });
}

// ---------- dados ----------
function lerTexto(caminho) {
  return fetch(new URL(caminho, BASE).href, { cache: 'no-store' }).then(function (r) {
    if (!r.ok) throw new Error('Não foi possível carregar ' + caminho);
    return r.text();
  });
}

function lerBase64(caminho) {
  return fetch(new URL(caminho, BASE).href).then(function (r) {
    if (!r.ok) throw new Error('Não foi possível carregar ' + caminho);
    return r.blob();
  }).then(function (blob) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(String(fr.result).split(',')[1]); };
      fr.onerror = reject;
      fr.readAsDataURL(blob);
    });
  });
}

function dadosDaPessoa(equipa) {
  var perfil = Office.context.mailbox.userProfile;
  var email = String(perfil.emailAddress || '').toLowerCase();
  var p = equipa[email] || {};
  return {
    email: email,
    nome: p.nome || perfil.displayName || email,
    cargo: p.cargo || '',
    telefone: p.telefone || '',
    foto: p.foto || 'fotos/generico.png'
  };
}

// Anexa uma imagem como inline (cid) e devolve o src a usar; se falhar, usa o endereço público.
function imagemInline(caminho, nome, jaExistentes) {
  var url = new URL(caminho, BASE).href;
  if (jaExistentes.indexOf(nome) >= 0) return Promise.resolve('cid:' + nome);
  return lerBase64(caminho).then(function (b64) {
    return chamar(function (cb) { Office.context.mailbox.item.addFileAttachmentFromBase64Async(b64, nome, { isInline: true }, cb); });
  }).then(function () { return 'cid:' + nome; }, function () { return url; });
}

// ---------- comando ----------
function formatarSolvian(event) {
  var item = Office.context.mailbox.item;
  var estado = {};
  chamar(function (cb) { item.body.getTypeAsync(cb); })
    .then(function (tipo) {
      if (tipo !== Office.CoercionType.Html) throw new Error('O email está em texto simples. Em Formatar texto, escolha HTML e volte a carregar no botão.');
      return chamar(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); });
    })
    .then(function (html) {
      estado.html = html;
      estado.analise = SolvianCore.analisar(html, DOMParser);
      if (estado.analise.jaTemAssinatura) return null; // só reformata o texto novo
      return Promise.all([lerTexto('equipa.json'), lerTexto('assinatura.template.html'),
        chamar(function (cb) { item.getAttachmentsAsync(cb); })]);
    })
    .then(function (dados) {
      if (!dados) return null;
      var equipa = JSON.parse(dados[0]), modelo = dados[1];
      var existentes = (dados[2] || []).map(function (a) { return a.name; });
      var d = dadosDaPessoa(equipa);
      var fotoNome = 'solvian-foto.png';
      return Promise.all([
        imagemInline(d.foto, fotoNome, existentes),
        imagemInline(IMAGENS.FAIXA1.caminho, IMAGENS.FAIXA1.nome, existentes),
        imagemInline(IMAGENS.FAIXA2.caminho, IMAGENS.FAIXA2.nome, existentes),
        imagemInline(IMAGENS.LOGO.caminho, IMAGENS.LOGO.nome, existentes)
      ]).then(function (s) {
        return SolvianCore.montarAssinatura(modelo, d, { FOTO: s[0], FAIXA1: s[1], FAIXA2: s[2], LOGO: s[3] });
      });
    })
    .then(function (sigHtml) {
      var r = SolvianCore.transformar(estado.html, sigHtml, DOMParser);
      estado.resultado = r;
      return chamar(function (cb) { item.body.setAsync(r.html, { coercionType: Office.CoercionType.Html }, cb); });
    })
    .then(function () {
      var r = estado.resultado, partes = ['Formatação Solvian aplicada.'];
      if (r.assinaturaInserida) partes.push('Assinatura inserida.');
      else if (estado.analise.jaTemAssinatura) partes.push('A assinatura já lá estava.');
      if (!r.historicoDetetado) partes.push('Não encontrei a mensagem citada: formatei o email todo.');
      return aviso(partes.join(' '), 'info');
    })
    .catch(function (e) { return aviso('Solvian: ' + (e && e.message ? e.message : 'não foi possível formatar.'), 'erro'); })
    .then(function () { event.completed(); });
}

Office.actions.associate('formatarSolvian', formatarSolvian);
