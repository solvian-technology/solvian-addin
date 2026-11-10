/*
 * Solvian - formatação de emails (lógica pura, sem Office.js).
 * Portada da macro FormatarEmailSolvian (v3): aplica letra, cores, títulos e setas
 * ao texto novo, retira a despedida duplicada e insere a assinatura por cima do histórico citado.
 * Trabalha sobre uma string HTML (o corpo do rascunho) e devolve outra string HTML.
 * Corre no browser (DOMParser) e em Node (jsdom), para se poder testar fora do Outlook.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SolvianCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // --- Configuração (equivalente às constantes do topo da macro) --------------------------
  var CFG = {
    fontStack: "'Century Gothic', CenturyGothic, AppleGothic, sans-serif",
    fontSizePt: 10,
    corTexto: '#050F17',    // Solvian deep blue (texto)
    corTitulos: '#062B44',  // Solvian medium blue (parágrafos todos a negrito)
    corSetas: '#68EAEA',    // Solvian aqua (setas ►)
    seta: '►',
    // Texto que só existe na assinatura: evita inseri-la duas vezes. (A macro usava o cargo "PARTNER".)
    sigMarker: '@solviangroup',
    despedidas: [
      'com os melhores cumprimentos', 'melhores cumprimentos', 'cumprimentos',
      'best regards', 'kind regards', 'regards', 'atenciosamente', 'atentamente'
    ]
  };

  var BLOCO_FOLHA = 'p,div,li,h1,h2,h3,h4,h5,h6,pre,td,th';
  var BLOCO_QUALQUER = 'p,div,li,ul,ol,table,tr,td,th,h1,h2,h3,h4,h5,h6,blockquote,pre';
  var SEM_ESTILO = { img: 1, br: 1, hr: 1, style: 1, script: 1, head: 1, meta: 1, title: 1, link: 1 };
  var BLOCOS_MARGEM = { p: 1, div: 1, li: 1, h1: 1, h2: 1, h3: 1, h4: 1, h5: 1, h6: 1, ul: 1, ol: 1, pre: 1 };

  var RE_CABECALHO = /^(de|from|von|da)\s?:/i;
  var RE_CAMPO_SEGUINTE = /(^|\n)\s*(enviad[oa]( em)?|sent|para|to|assunto|subject|date|data|cc)\s?:/i;
  var RE_ORIGINAL = /^-{2,}\s*(original message|mensagem original|mensagem inicial)\s*-{2,}$/i;
  var RE_ESCREVEU = /(wrote|escreveu)\s*:?\s*$/i;

  // --- Utilitários -----------------------------------------------------------------------
  function limpar(s) {
    return String(s || '').replace(/[​‌‍﻿]/g, '').replace(/ /g, ' ').trim();
  }

  // Texto de um elemento, com <br> como quebra de linha.
  function texto(el) {
    var c = el.cloneNode(true);
    var brs = c.querySelectorAll('br');
    for (var i = 0; i < brs.length; i++) brs[i].parentNode.replaceChild(el.ownerDocument.createTextNode('\n'), brs[i]);
    return c.textContent || '';
  }

  function eFolha(el) {
    return el.matches(BLOCO_FOLHA) && !el.querySelector(BLOCO_QUALQUER);
  }

  function antes(a, b) { // a precede b e não o contém
    var m = a.compareDocumentPosition(b);
    return !!(m & 4) && !(m & 16);
  }

  function dentroDeAssinatura(el) {
    return !!el.closest('[data-solvian-sig]');
  }

  // --- Onde começa a mensagem citada ------------------------------------------------------
  function confirmaCabecalho(b) {
    if (RE_CAMPO_SEGUINTE.test(texto(b))) return true;
    var s = b.nextElementSibling, n = 0;
    while (s && n < 2) { if (RE_CAMPO_SEGUINTE.test('\n' + texto(s))) return true; s = s.nextElementSibling; n++; }
    // cabeçalho dentro de um bloco com separador (linha horizontal ou borda superior)
    var prev = b.previousElementSibling;
    if (prev && prev.tagName === 'HR') return true;
    var p = b.parentElement;
    return !!(p && /border-top/i.test(p.getAttribute('style') || '') && p.firstElementChild === b);
  }

  function primeiroPor(candidatos) {
    var melhor = null;
    for (var i = 0; i < candidatos.length; i++) {
      var c = candidatos[i];
      if (!c) continue;
      if (!melhor || antes(c, melhor) || c.contains(melhor)) melhor = c;
    }
    return melhor;
  }

  function semTextoAntes(el) { // nenhum texto visível antes de el dentro do seu pai
    for (var n = el.previousSibling; n; n = n.previousSibling) {
      if (n.nodeType === 3 && limpar(n.nodeValue) !== '') return false;
      if (n.nodeType === 1 && limpar(n.textContent) !== '') return false;
    }
    return true;
  }

  function elevarAoSeparador(el, body) {
    // sobe para o bloco que contém o cabeçalho quando este é o seu primeiro filho e o bloco é o separador
    for (;;) {
      var p = el.parentElement;
      if (!p || p === body) break;
      var st = (p.getAttribute('style') || '');
      if (semTextoAntes(el) && (/border-top/i.test(st) || /rplyfwdmsg/i.test(p.id || ''))) el = p; else break;
    }
    return el;
  }

  function encontrarFronteira(doc) {
    var body = doc.body;
    var cOutlook = body.querySelector('[id$="divRplyFwdMsg"]');   // Outlook na web, Outlook novo
    var cCab = null, cOrig = null, cEscreveu = null;

    var folhas = body.querySelectorAll(BLOCO_FOLHA);
    for (var i = 0; i < folhas.length; i++) {
      var b = folhas[i];
      if (!eFolha(b) || dentroDeAssinatura(b)) continue;
      var t = limpar(texto(b));
      if (!cCab && RE_CABECALHO.test(t) && confirmaCabecalho(b)) cCab = b;
      if (!cOrig && RE_ORIGINAL.test(t)) cOrig = b;
      if (cCab && cOrig) break;
    }
    // "No dia ... escreveu:" / "On ... wrote:" seguido de citação (Gmail, Apple Mail)
    var bq = body.querySelector('blockquote');
    if (bq) {
      var ant = bq.previousElementSibling;
      if (ant && RE_ESCREVEU.test(limpar(texto(ant)))) cEscreveu = ant;
      else if ((bq.getAttribute('type') || '').toLowerCase() === 'cite') cEscreveu = bq;
    }
    var f = primeiroPor([cOutlook, cCab, cOrig, cEscreveu]);
    return f ? elevarAoSeparador(f, body) : null;
  }

  // --- Blocos do texto novo -----------------------------------------------------------------
  function folhasNovas(doc, fronteira, incluirAssinatura) {
    var todas = doc.body.querySelectorAll(BLOCO_FOLHA), out = [];
    for (var i = 0; i < todas.length; i++) {
      var b = todas[i];
      if (!eFolha(b) || (!incluirAssinatura && dentroDeAssinatura(b))) continue;
      if (fronteira && !antes(b, fronteira)) continue;
      out.push(b);
    }
    return out;
  }

  function elementosNovos(doc, fronteira) {
    var todos = doc.body.querySelectorAll('*'), out = [];
    for (var i = 0; i < todos.length; i++) {
      var e = todos[i];
      if (SEM_ESTILO[e.tagName.toLowerCase()] || dentroDeAssinatura(e)) continue;
      if (fronteira && (!antes(e, fronteira) || e.contains(fronteira))) continue;
      out.push(e);
    }
    return out;
  }

  // Texto novo, incluindo uma assinatura já inserida (para a encontrar e não a duplicar)
  function textoNovo(doc, fronteira) {
    return folhasNovas(doc, fronteira, true).map(function (b) { return texto(b); }).join('\n');
  }

  function temTexto(b) { return limpar(texto(b)).length > 0 || !!b.querySelector('img'); }

  function ultimoComTexto(blocos) {
    for (var i = blocos.length - 1; i >= 0; i--) if (limpar(texto(blocos[i])).length > 0) return blocos[i];
    return null;
  }

  // --- Passos 1 a 3 da macro ------------------------------------------------------------------
  function estilo(el, prop, valor) { el.style.setProperty(prop, valor); }

  function aplicarBase(els) {
    els.forEach(function (e) {
      var t = e.tagName.toLowerCase();
      estilo(e, 'font-family', CFG.fontStack);
      estilo(e, 'font-size', CFG.fontSizePt + 'pt');
      estilo(e, 'color', CFG.corTexto);
      if (BLOCOS_MARGEM[t]) { estilo(e, 'margin-top', '0'); estilo(e, 'margin-bottom', '0'); estilo(e, 'line-height', 'normal'); }
    });
  }

  function negritoAtivo(no, limite) {
    for (var e = no.parentElement; e; e = e.parentElement) {
      var t = e.tagName.toLowerCase();
      if (t === 'b' || t === 'strong') return true;
      var fw = (e.style && e.style.fontWeight) || '';
      if (fw === 'bold' || fw === 'bolder' || parseInt(fw, 10) >= 600) return true;
      if (e === limite) break;
    }
    // o bloco pode estar dentro de um <b>
    var up = limite && limite.parentElement ? limite.parentElement.closest('b,strong') : null;
    return !!up;
  }

  function eTitulo(b) { // parágrafo inteiro a negrito
    var w = b.ownerDocument.createTreeWalker(b, 4 /* SHOW_TEXT */), n, visto = false;
    while ((n = w.nextNode())) {
      if (limpar(n.nodeValue) === '') continue;
      visto = true;
      if (!negritoAtivo(n, b)) return false;
    }
    return visto;
  }

  function aplicarTitulos(blocos) {
    blocos.forEach(function (b) {
      if (!eTitulo(b)) return;
      estilo(b, 'color', CFG.corTitulos);
      var d = b.querySelectorAll('*');
      for (var i = 0; i < d.length; i++) if (!SEM_ESTILO[d[i].tagName.toLowerCase()]) estilo(d[i], 'color', CFG.corTitulos);
    });
  }

  function aplicarSetas(doc, els) {
    var vistos = [];
    els.forEach(function (e) {
      for (var n = e.firstChild; n; n = n.nextSibling) if (n.nodeType === 3 && n.nodeValue.indexOf(CFG.seta) >= 0 && vistos.indexOf(n) < 0) vistos.push(n);
    });
    vistos.forEach(function (n) {
      var partes = n.nodeValue.split(CFG.seta), pai = n.parentNode, ref = n.nextSibling;
      pai.removeChild(n);
      var frag = doc.createDocumentFragment();
      partes.forEach(function (p, i) {
        if (p) frag.appendChild(doc.createTextNode(p));
        if (i < partes.length - 1) {
          var s = doc.createElement('span');
          s.style.setProperty('color', CFG.corSetas);
          s.textContent = CFG.seta;
          frag.appendChild(s);
        }
      });
      pai.insertBefore(frag, ref);
    });
  }

  function removerDespedida(blocos) {
    var u = ultimoComTexto(blocos);
    if (!u) return false;
    var t = limpar(texto(u)).toLowerCase();
    for (var i = 0; i < CFG.despedidas.length; i++) {
      if (t.indexOf(CFG.despedidas[i]) === 0) { u.parentNode.removeChild(u); return true; }
    }
    return false;
  }

  // --- API ---------------------------------------------------------------------------------
  function analisar(html, DP) {
    var doc = new DP().parseFromString(html, 'text/html');
    var f = encontrarFronteira(doc);
    return {
      historicoDetetado: !!f,
      jaTemAssinatura: textoNovo(doc, f).indexOf(CFG.sigMarker) >= 0
    };
  }

  /**
   * @param {string} html     corpo completo do rascunho (body.getAsync, HTML)
   * @param {string|null} sigHtml  HTML da assinatura (ou null para só formatar)
   * @param {function} DP     construtor DOMParser
   * @returns {{html:string, historicoDetetado:boolean, assinaturaInserida:boolean, despedidaRemovida:boolean}}
   */
  function transformar(html, sigHtml, DP) {
    var doc = new DP().parseFromString(html, 'text/html');
    var fronteira = encontrarFronteira(doc);
    var jaTem = textoNovo(doc, fronteira).indexOf(CFG.sigMarker) >= 0;

    // 1) letra, cor, espaçamento; 2) títulos; 3) setas
    aplicarBase(elementosNovos(doc, fronteira));
    aplicarTitulos(folhasNovas(doc, fronteira));
    aplicarSetas(doc, elementosNovos(doc, fronteira));

    var res = { historicoDetetado: !!fronteira, assinaturaInserida: false, despedidaRemovida: false };

    // 4) assinatura (só se ainda não estiver no texto novo)
    if (sigHtml && !jaTem) {
      var blocos = folhasNovas(doc, fronteira);
      res.despedidaRemovida = removerDespedida(blocos);
      var ult = ultimoComTexto(folhasNovas(doc, fronteira));
      var ins = '<div><br></div><div data-solvian-sig="1">' + sigHtml + '</div>';
      if (ult) ult.insertAdjacentHTML('afterend', ins);
      else if (fronteira) fronteira.insertAdjacentHTML('beforebegin', ins);
      else doc.body.insertAdjacentHTML('beforeend', ins);
      res.assinaturaInserida = true;
    }
    res.html = '<!DOCTYPE html>' + doc.documentElement.outerHTML;
    return res;
  }

  // --- Assinatura a partir do modelo ----------------------------------------------------------------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }

  /**
   * @param {string} modelo  conteúdo de assinatura.template.html
   * @param {{nome,cargo,email,telefone}} d  dados da pessoa (cargo e telefone podem ser vazios)
   * @param {{FOTO,FAIXA1,FAIXA2,LOGO}} srcs  valores de src das imagens (cid: ou https)
   */
  function montarAssinatura(modelo, d, srcs) {
    var h = modelo;
    h = h.replace(/<!--SE:CARGO-->([\s\S]*?)<!--\/SE-->/, d.cargo ? '$1' : '');
    h = h.replace(/<!--SE:TELEFONE-->([\s\S]*?)<!--\/SE-->/, d.telefone ? '$1' : '');
    var tel = String(d.telefone || '');
    var valores = {
      NOME: esc(d.nome), CARGO: esc(d.cargo), EMAIL: esc(d.email),
      TEL_TEXTO: esc(tel), TEL_LINK: esc(tel.replace(/[^\d+]/g, '')),
      FOTO: srcs.FOTO, FAIXA1: srcs.FAIXA1, FAIXA2: srcs.FAIXA2, LOGO: srcs.LOGO
    };
    return h.replace(/\{\{(\w+)\}\}/g, function (m, k) { return valores[k] != null ? valores[k] : m; });
  }

  return { CFG: CFG, analisar: analisar, transformar: transformar, montarAssinatura: montarAssinatura };
});
