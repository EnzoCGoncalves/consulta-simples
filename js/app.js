/* Interface da consulta em lote. A lógica (CNPJ, APIs, Excel, cache) fica em nucleo.js. */
(function () {
  'use strict';

  const N = window.Nucleo;
  const $ = (id) => document.getElementById(id);
  const el = {
    zona: $('zona'),
    arquivo: $('arquivo'),
    modelo: $('modelo'),
    resumo: $('resumo'),
    erro: $('erro'),
    iniciar: $('iniciar'),
    cancelar: $('cancelar'),
    baixar: $('baixar'),
    ibsCbs: $('ibsCbs'),
    pausa: $('pausa'),
    simultaneas: $('simultaneas'),
    reaproveitar: $('reaproveitar'),
    limparCache: $('limparCache'),
    infoCache: $('infoCache'),
    progresso: $('progresso'),
    barra: $('barra'),
    barraTrilho: $('barraTrilho'),
    textoProgresso: $('textoProgresso'),
    detalhe: $('detalhe'),
    blocoIbs: $('blocoIbs'),
    barraIbs: $('barraIbs'),
    barraIbsTrilho: $('barraIbsTrilho'),
    textoIbs: $('textoIbs'),
    detalheIbs: $('detalheIbs'),
    cartaoTabela: $('cartaoTabela'),
    contagem: $('contagem'),
    corpo: $('corpo'),
  };

  function abrirArmazenamento() {
    try {
      const s = window.localStorage;
      s.setItem('consultaSimples:teste', '1');
      s.removeItem('consultaSimples:teste');
      return s;
    } catch (e) {
      return null; // navegador bloqueou: o site funciona, só não reaproveita resultados
    }
  }

  const estado = {
    nomeArquivo: '',
    nomeAba: '',
    linhas: [], // uma por linha da planilha
    unicos: [], // CNPJs válidos sem repetição, na ordem em que aparecem
    porCnpj: new Map(), // cnpj → linhas da planilha com esse CNPJ
    consultas: new Map(), // cnpj → resposta da consulta principal (Simples/SIMEI)
    ibs: new Map(), // cnpj → situação da busca de IBS/CBS na CNPJá
    rodando: false,
    faseSimples: false, // true enquanto a consulta principal está em andamento
    executou: false,
    cancelado: false,
    controle: null,
    consultor: N.criarConsultor(), // um só por página: guarda o horário da próxima vez da CNPJá
    cache: N.criarCache(abrirArmazenamento()),
  };

  const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const INTERVALO_CNPJA_MS = 12500;

  // ------------------------------------------------------------ utilidades

  function mostrarErro(msg) {
    el.erro.textContent = msg;
    el.erro.hidden = !msg;
  }

  function bibliotecasCarregadas() {
    if (typeof window.XLSX === 'undefined') {
      mostrarErro(
        'Não foi possível carregar a biblioteca de planilhas (SheetJS). Verifique a conexão com a internet e recarregue a página.'
      );
      return false;
    }
    return true;
  }

  function formatarDuracao(ms) {
    const s = Math.max(1, Math.round(ms / 1000));
    if (s < 60) return `${s} s`;
    const min = Math.round(s / 60);
    return min < 60 ? `${min} min` : `${Math.floor(min / 60)} h ${min % 60} min`;
  }

  function salvarArquivo(bytes, nome) {
    const blob = new Blob([bytes], { type: MIME_XLSX });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nome;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  function numero(input, padrao, min, max) {
    const v = parseFloat(String(input.value).replace(',', '.'));
    return Number.isFinite(v) ? Math.min(Math.max(v, min), max) : padrao;
  }
  const pausaMs = () => Math.round(numero(el.pausa, 0.5, 0, 10) * 1000);
  const simultaneas = () => Math.round(numero(el.simultaneas, 3, 1, 5));
  const diasReaproveitar = () => Math.round(numero(el.reaproveitar, 3, 0, 7));

  function erroCancelado() {
    const e = new Error('Consulta cancelada');
    e.name = 'AbortError';
    return e;
  }

  function atualizarInfoCache() {
    const n = estado.cache.contar();
    el.infoCache.textContent = n ? `${n} CNPJ${n === 1 ? '' : 's'} guardado${n === 1 ? '' : 's'} neste navegador.` : 'Nenhum resultado guardado.';
    el.limparCache.disabled = estado.rodando || n === 0;
  }

  // ------------------------------------------------------------ carregar planilha

  async function carregarArquivo(arquivo) {
    if (!arquivo || estado.rodando) return;
    mostrarErro('');
    if (!bibliotecasCarregadas()) return;
    if (!/\.(xlsx|xlsm|xls)$/i.test(arquivo.name)) {
      mostrarErro('Envie um arquivo do Excel (.xlsx).');
      return;
    }
    try {
      const dados = new Uint8Array(await arquivo.arrayBuffer());
      const lido = N.lerPlanilha(window.XLSX, dados);
      if (!lido.linhas.length) throw new Error('A planilha não tem nenhuma linha preenchida abaixo do cabeçalho.');
      prepararLinhas(arquivo.name, lido);
    } catch (e) {
      mostrarErro(e && e.message ? e.message : 'Não consegui ler esta planilha.');
    }
  }

  function prepararLinhas(nomeArquivo, lido) {
    estado.nomeArquivo = nomeArquivo;
    estado.nomeAba = lido.nomeAba;
    estado.linhas = lido.linhas.map((l) => Object.assign({}, l, { tr: null, consultando: false }));
    estado.unicos = [];
    estado.porCnpj = new Map();
    estado.consultas = new Map();
    estado.ibs = new Map();
    estado.executou = false;
    estado.cancelado = false;

    for (const l of estado.linhas) {
      if (!l.valido) continue;
      if (!estado.porCnpj.has(l.cnpj)) {
        estado.porCnpj.set(l.cnpj, []);
        estado.unicos.push(l.cnpj);
      }
      estado.porCnpj.get(l.cnpj).push(l);
    }

    const invalidos = estado.linhas.filter((l) => !l.valido).length;
    const validos = estado.linhas.length - invalidos;
    const repetidos = validos - estado.unicos.length;
    const estimativa = Math.ceil(estado.unicos.length / simultaneas()) * (pausaMs() + 400);

    el.resumo.replaceChildren();
    const titulo = document.createElement('strong');
    titulo.textContent = nomeArquivo;
    el.resumo.append(titulo, ` — aba “${lido.nomeAba}”`);
    const lista = document.createElement('ul');
    const plural = (n, s, p) => (n === 1 ? s : p);
    const itens = [
      `${estado.linhas.length} ${plural(estado.linhas.length, 'linha', 'linhas')} com dados`,
      `${estado.unicos.length} ${plural(estado.unicos.length, 'CNPJ', 'CNPJs')} para consultar` +
        (estado.unicos.length ? ` (Simples e SIMEI em ≈ ${formatarDuracao(estimativa)})` : ''),
    ];
    if (repetidos) itens.push(`${repetidos} ${plural(repetidos, 'repetido (consultado uma vez só)', 'repetidos (consultados uma vez só)')}`);
    if (invalidos) itens.push(`${invalidos} ${plural(invalidos, 'CNPJ inválido (não será consultado)', 'CNPJs inválidos (não serão consultados)')}`);
    if (!lido.temColunaEmpresa) itens.push('Coluna EMPRESA não encontrada: o nome ficará em branco');
    for (const t of itens) {
      const li = document.createElement('li');
      li.textContent = t;
      lista.append(li);
    }
    el.resumo.append(lista);
    el.resumo.hidden = false;

    renderizarTabela();
    el.progresso.hidden = true;
    atualizarBotoes();
  }

  // ------------------------------------------------------------ o que falta consultar

  function precisaIbs(cnpj) {
    const c = estado.consultas.get(cnpj);
    if (!c || c.tipo !== 'ok' || c.fonte === 'CNPJá' || !N.optanteHoje(c)) return false; // CNPJá já trouxe os eventos
    const s = estado.ibs.get(cnpj);
    return !s || s.status === 'falha' || s.status === 'cancelado';
  }

  function pendentesSimples() {
    return estado.unicos.filter((c) => {
      const r = estado.consultas.get(c);
      return !r || r.tipo === 'falha';
    });
  }

  function pendentesIbs() {
    return el.ibsCbs.checked ? estado.unicos.filter(precisaIbs) : [];
  }

  // ------------------------------------------------------------ tabela

  function resultadoDe(cnpj) {
    const c = estado.consultas.get(cnpj);
    if (!c) return null;
    const rr = estado.ibs.get(cnpj);
    return N.montarResultado(rr ? Object.assign({}, c, { regimeRegular: rr }) : c);
  }

  /** As 7 colunas de saída de uma linha. Na tela, mostra também o andamento da busca na CNPJá. */
  function valoresSaida(l, paraTela) {
    const obs = [];
    let cols = ['', '', '', ''];
    if (!l.valido) {
      obs.push(l.erro);
    } else {
      const r = resultadoDe(l.cnpj);
      const iguais = estado.porCnpj.get(l.cnpj) || [];
      if (iguais.length > 1) obs.push(`CNPJ repetido na planilha (linhas ${iguais.map((x) => x.linhaExcel).join(', ')})`);
      if (r) {
        cols = [r.simples, r.simei, r.ibs, r.cbs];
        const s = estado.ibs.get(l.cnpj);
        if (paraTela && estado.rodando && s && (s.status === 'fila' || s.status === 'pendente')) {
          cols[2] = cols[3] = s.status === 'pendente' ? 'Buscando na CNPJá…' : 'Na fila da CNPJá…';
        }
        obs.push(r.observacao);
      } else if (estado.cancelado) {
        obs.push('Não consultado (consulta cancelada)');
      } else {
        obs.push(l.consultando ? 'Consultando…' : 'Aguardando consulta');
      }
    }
    return [l.empresa, l.cnpj || l.cnpjBruto, cols[0], cols[1], cols[2], cols[3], obs.filter(Boolean).join('; ')];
  }

  function classeValor(texto) {
    if (/^(Optante|Enquadrado|Opção)/.test(texto)) return 'sim';
    if (/^(NÃO|Não optante)/.test(texto)) return 'nao';
    if (/^Verificar/.test(texto)) return 'atencao';
    if (/^(Buscando|Na fila)/.test(texto)) return 'aguardando';
    return '';
  }

  function criarLinhaTabela(l) {
    const tr = document.createElement('tr');
    const tdNum = document.createElement('td');
    tdNum.className = 'num';
    tdNum.textContent = l.linhaExcel;
    tr.append(tdNum);
    for (let i = 0; i < 7; i++) tr.append(document.createElement('td'));
    tr.children[2].className = 'cnpj';
    l.tr = tr;
    atualizarLinhaTabela(l);
    return tr;
  }

  function atualizarLinhaTabela(l) {
    const v = valoresSaida(l, true);
    const c = l.valido ? estado.consultas.get(l.cnpj) : null;
    const tds = l.tr.children;
    for (let i = 0; i < 7; i++) {
      const td = tds[i + 1];
      td.textContent = v[i];
      if (i >= 2 && i <= 5) td.className = classeValor(v[i]);
    }
    let situacao = 'pendente';
    if (!l.valido) situacao = 'invalido';
    else if (c) situacao = c.tipo === 'ok' ? 'ok' : 'erro';
    else if (l.consultando) situacao = 'consultando';
    else if (estado.cancelado) situacao = 'erro';
    l.tr.className = situacao;
    tds[7].className = situacao === 'invalido' || situacao === 'erro' ? 'obs erro' : 'obs';
  }

  function atualizarLinhasDoCnpj(cnpj) {
    for (const l of estado.porCnpj.get(cnpj) || []) atualizarLinhaTabela(l);
  }

  function renderizarTabela() {
    const frag = document.createDocumentFragment();
    for (const l of estado.linhas) frag.append(criarLinhaTabela(l));
    el.corpo.replaceChildren(frag);
    el.cartaoTabela.hidden = false;
    atualizarContagem();
  }

  function atualizarContagem() {
    let sim = 0;
    let nao = 0;
    let erros = 0;
    for (const c of estado.consultas.values()) {
      if (c.tipo !== 'ok') erros++;
      else if (N.optanteHoje(c)) sim++;
      else nao++;
    }
    const partes = [];
    if (sim) partes.push(`${sim} optante${sim === 1 ? '' : 's'}`);
    if (nao) partes.push(`${nao} não optante${nao === 1 ? '' : 's'}`);
    if (erros) partes.push(`${erros} com erro`);
    el.contagem.textContent = partes.length ? `${partes.join(' · ')} (por CNPJ)` : '';
  }

  // ------------------------------------------------------------ botões

  function atualizarBotoes() {
    const temArquivo = estado.linhas.length > 0;
    const faltam = pendentesSimples().length + pendentesIbs().length;
    el.iniciar.hidden = estado.rodando;
    el.cancelar.hidden = !estado.rodando;
    el.iniciar.disabled = !temArquivo || faltam === 0;
    if (!temArquivo || !estado.executou) el.iniciar.textContent = 'Iniciar consulta';
    else if (faltam) el.iniciar.textContent = `Consultar pendentes (${faltam})`;
    else el.iniciar.textContent = 'Consulta concluída';
    // Durante a busca de IBS/CBS já dá para baixar (Simples e SIMEI prontos).
    el.baixar.disabled = !temArquivo || (estado.rodando ? estado.faseSimples : !estado.executou && estado.unicos.length > 0);
    for (const c of [el.ibsCbs, el.simultaneas, el.reaproveitar, el.arquivo]) c.disabled = estado.rodando;
    el.zona.classList.toggle('bloqueada', estado.rodando);
    atualizarInfoCache();
  }

  // ------------------------------------------------------------ progresso

  function mostrarBarra(barra, trilho, feitos, total) {
    const pct = total ? Math.round((feitos / total) * 100) : 100;
    barra.style.width = `${pct}%`;
    trilho.setAttribute('aria-valuenow', String(pct));
  }

  // ------------------------------------------------------------ consulta

  async function executarConsulta() {
    if (estado.rodando) return;
    const comIbs = el.ibsCbs.checked;
    const filaSimples = pendentesSimples();
    const filaIbs = pendentesIbs(); // optantes já consultadas que ainda não têm o IBS/CBS
    if (!filaSimples.length && !filaIbs.length) return;

    const dias = diasReaproveitar();
    const total = estado.unicos.length;
    const jaFeitos = total - filaSimples.length;
    let feitos = jaFeitos;
    let consultasReais = 0;
    let reaproveitados = 0;
    let ibsFeitos = 0;
    let ibsTotal = filaIbs.length;
    let limite = Math.min(simultaneas(), Math.max(filaSimples.length, 1));
    let proximo = 0;
    const inicio = Date.now();
    let acordar = null;
    const acordarIbs = () => {
      if (acordar) {
        const f = acordar;
        acordar = null;
        f();
      }
    };

    estado.rodando = true;
    estado.faseSimples = filaSimples.length > 0;
    estado.executou = true;
    estado.cancelado = false;
    estado.controle = new AbortController();
    const signal = estado.controle.signal;
    signal.addEventListener('abort', acordarIbs);
    for (const c of filaIbs) estado.ibs.set(c, { status: 'fila' });
    estado.cache.podar(dias);
    for (const l of estado.linhas) if (l.tr) atualizarLinhaTabela(l);
    mostrarErro('');
    el.progresso.hidden = false;
    el.blocoIbs.hidden = !comIbs;
    el.detalhe.textContent = estado.faseSimples ? `Consultando na Minha Receita, ${limite} ao mesmo tempo…` : '';
    el.detalheIbs.textContent = '';
    atualizarBotoes();

    function progressoSimples() {
      mostrarBarra(el.barra, el.barraTrilho, feitos, total);
      if (estado.faseSimples) {
        let t = `Consultando ${Math.min(feitos + 1, total)} de ${total}…`;
        const nestaRodada = feitos - jaFeitos;
        if (nestaRodada > 0 && nestaRodada < filaSimples.length) {
          t += ` (faltam ≈ ${formatarDuracao(((Date.now() - inicio) / nestaRodada) * (filaSimples.length - nestaRodada))})`;
        }
        el.textoProgresso.textContent = t;
      } else {
        const falhas = pendentesSimples().length;
        el.textoProgresso.textContent =
          `Simples Nacional e SIMEI: ${total - falhas} de ${total} CNPJs consultados` +
          (reaproveitados ? ` (${reaproveitados} reaproveitado${reaproveitados === 1 ? '' : 's'})` : '') +
          (falhas ? `; ${falhas} com falha.` : '.');
      }
    }

    function progressoIbs() {
      if (!comIbs) return;
      mostrarBarra(el.barraIbs, el.barraIbsTrilho, ibsFeitos, ibsTotal);
      const restantes = ibsTotal - ibsFeitos;
      let t;
      if (!ibsTotal) t = estado.faseSimples ? 'IBS/CBS na CNPJá: aguardando encontrar optantes do Simples…' : 'IBS/CBS na CNPJá: nenhuma optante para consultar.';
      else {
        t = `IBS/CBS na CNPJá: ${ibsFeitos} de ${ibsTotal} ${ibsTotal === 1 ? 'optante' : 'optantes'}`;
        if (estado.faseSimples) t += ibsTotal === 1 ? ' encontrada até agora' : ' encontradas até agora';
      }
      if (restantes > 0) t += ` (faltam ≈ ${formatarDuracao(restantes * INTERVALO_CNPJA_MS)})`;
      el.textoIbs.textContent = t;
    }

    function avisoSimples(ev) {
      const seg = (ms) => Math.ceil(ms / 1000);
      if (ev.tipo === 'nova_tentativa') {
        let t = `${ev.api}: ${ev.motivo}. Nova tentativa em ${seg(ev.ms)} s (${ev.tentativa} de ${ev.total})…`;
        if (ev.limite && ev.api !== 'CNPJá') {
          // A API pediu calma: passa a consultar uma por vez e aumenta a pausa.
          limite = 1;
          const nova = Math.min(pausaMs() / 1000 + 1, 10);
          el.pausa.value = String(nova);
          t += ` Passando a consultar uma por vez, com pausa de ${nova} s.`;
        }
        el.detalhe.textContent = t;
      } else if (ev.tipo === 'reserva') {
        el.detalhe.textContent = ev.naoEncontrado
          ? `${ev.api} não encontrou um CNPJ; conferindo na ${ev.proxima}…`
          : `${ev.api} não respondeu; tentando a reserva ${ev.proxima}…`;
      } else if (ev.tipo === 'espera') {
        el.detalhe.textContent = `Aguardando ${seg(ev.ms)} s para usar a ${ev.api} (limite de 5 consultas por minuto)…`;
      }
    }

    function avisoIbs(ev) {
      const seg = (ms) => Math.ceil(ms / 1000);
      if (ev.tipo === 'consultando') el.detalheIbs.textContent = 'Buscando na CNPJá…';
      else if (ev.tipo === 'espera') el.detalheIbs.textContent = `Aguardando ${seg(ev.ms)} s (a CNPJá aceita 5 consultas por minuto)…`;
      else if (ev.tipo === 'nova_tentativa') el.detalheIbs.textContent = `CNPJá: ${ev.motivo}. Nova tentativa em ${seg(ev.ms)} s (${ev.tentativa} de ${ev.total})…`;
    }

    function enfileirarIbs(cnpj) {
      if (!comIbs || !precisaIbs(cnpj)) return;
      estado.ibs.set(cnpj, { status: 'fila' });
      filaIbs.push(cnpj);
      ibsTotal++;
      acordarIbs();
    }

    /** Devolve true se consultou de verdade (false se reaproveitou). */
    async function processarSimples(cnpj) {
      const linhas = estado.porCnpj.get(cnpj);
      let consulta = dias > 0 ? estado.cache.lerConsulta(cnpj, dias) : null;
      if (consulta) {
        reaproveitados++;
      } else {
        for (const l of linhas) {
          l.consultando = true;
          atualizarLinhaTabela(l);
        }
        try {
          consulta = await estado.consultor.consultar(cnpj, { signal, avisar: avisoSimples });
        } catch (e) {
          if (e && e.name === 'AbortError') throw e;
          consulta = { tipo: 'falha', falhas: [e && e.message ? e.message : 'erro inesperado'] };
        } finally {
          for (const l of linhas) l.consultando = false;
        }
        consulta.consultadoEm = N.hojeISO();
        estado.cache.gravarConsulta(cnpj, consulta); // só guarda respostas corretas
        consultasReais++;
      }
      estado.consultas.set(cnpj, consulta);
      feitos++;
      enfileirarIbs(cnpj);
      atualizarLinhasDoCnpj(cnpj);
      atualizarContagem();
      progressoSimples();
      progressoIbs();
      return !consulta.doCache;
    }

    async function trabalhadorSimples(id) {
      while (proximo < filaSimples.length) {
        if (signal.aborted) throw erroCancelado();
        if (id >= limite) return; // a API pediu calma: menos consultas ao mesmo tempo
        const cnpj = filaSimples[proximo++];
        const consultou = await processarSimples(cnpj);
        if (consultou && proximo < filaSimples.length) await N.esperar(pausaMs(), signal);
      }
    }

    async function processarIbs(cnpj) {
      let rr = dias > 0 ? estado.cache.lerIbs(cnpj, dias) : null;
      if (!rr) {
        estado.ibs.set(cnpj, { status: 'pendente' });
        atualizarLinhasDoCnpj(cnpj);
        try {
          rr = await estado.consultor.buscarRegimeRegular(cnpj, { signal, avisar: avisoIbs });
        } catch (e) {
          if (e && e.name === 'AbortError') throw e;
          rr = { status: 'falha', motivo: e && e.message ? e.message : 'erro inesperado' };
        }
        rr.consultadoEm = N.hojeISO();
        estado.cache.gravarIbs(cnpj, rr);
      }
      estado.ibs.set(cnpj, rr);
      ibsFeitos++;
      atualizarLinhasDoCnpj(cnpj);
      progressoIbs();
    }

    async function trabalhadorIbs() {
      if (!comIbs) return;
      for (;;) {
        if (signal.aborted) throw erroCancelado();
        if (filaIbs.length) {
          await processarIbs(filaIbs.shift());
          continue;
        }
        if (!estado.faseSimples) return;
        await new Promise((resolve) => {
          acordar = resolve;
        });
      }
    }

    const tarefaIbs = trabalhadorIbs();
    tarefaIbs.catch(() => {}); // o erro é tratado no await abaixo
    progressoSimples();
    progressoIbs();

    try {
      try {
        const trabalhadores = [];
        for (let i = 0; i < limite; i++) trabalhadores.push(trabalhadorSimples(i));
        await Promise.all(trabalhadores);
      } finally {
        estado.faseSimples = false;
        acordarIbs();
      }
      progressoSimples();
      progressoIbs();
      atualizarBotoes();
      el.detalhe.textContent =
        comIbs && ibsFeitos < ibsTotal
          ? 'Simples e SIMEI prontos. Se baixar agora, as linhas com IBS/CBS ainda em consulta saem como “Verificar no portal”; para a planilha completa, espere terminar.'
          : '';

      await tarefaIbs;

      const falhas = pendentesSimples().length + pendentesIbs().length;
      el.textoProgresso.textContent =
        `Consulta concluída: ${total} de ${total} CNPJs` +
        (reaproveitados ? ` (${reaproveitados} reaproveitado${reaproveitados === 1 ? '' : 's'}, ${consultasReais} consultado${consultasReais === 1 ? '' : 's'} agora)` : '') +
        '.' +
        (falhas ? ` ${falhas} com falha — clique em “Consultar pendentes” para tentar de novo.` : '');
      el.detalhe.textContent = 'Clique em “Baixar planilha” para salvar o resultado.';
      el.detalheIbs.textContent = '';
    } catch (e) {
      if (!e || e.name !== 'AbortError') throw e;
      estado.cancelado = true;
      for (const [c, s] of estado.ibs) if (s.status === 'fila' || s.status === 'pendente') estado.ibs.set(c, { status: 'cancelado' });
      for (const l of estado.linhas) if (l.tr) atualizarLinhaTabela(l);
      const ok = total - pendentesSimples().length;
      el.textoProgresso.textContent =
        `Consulta cancelada: ${ok} de ${total} CNPJs consultados` + (comIbs ? `; IBS/CBS: ${ibsFeitos} de ${ibsTotal}.` : '.');
      el.detalhe.textContent = 'Você já pode baixar a planilha com o que foi consultado, ou clicar em “Consultar pendentes” para continuar.';
      el.textoIbs.textContent = `IBS/CBS na CNPJá: ${ibsFeitos} de ${ibsTotal} (cancelado).`;
      el.detalheIbs.textContent = '';
    } finally {
      estado.rodando = false;
      estado.faseSimples = false;
      estado.controle = null;
      for (const l of estado.linhas) if (l.tr) atualizarLinhaTabela(l);
      atualizarBotoes();
    }
  }

  // ------------------------------------------------------------ downloads

  async function baixarResultado() {
    if (!bibliotecasCarregadas() || !estado.linhas.length) return;
    try {
      const linhas = estado.linhas.map((l) => valoresSaida(l, false));
      const bytes = await N.gerarXlsx(window.XLSX, window.JSZip, N.COLUNAS_SAIDA, linhas, {
        nomeAba: estado.nomeAba || 'Consulta Simples',
        colunaTexto: 1,
        larguraMinima: [20, 16],
      });
      salvarArquivo(bytes, N.nomeArquivoResultado());
    } catch (e) {
      mostrarErro(`Não foi possível gerar a planilha: ${e && e.message ? e.message : e}`);
    }
  }

  async function baixarModelo(ev) {
    ev.preventDefault();
    if (!bibliotecasCarregadas()) return;
    const bytes = await N.gerarXlsx(window.XLSX, window.JSZip, ['EMPRESA', 'CNPJ'], [], {
      nomeAba: 'Empresas',
      colunaTexto: 1,
      larguraMinima: [40, 22],
    });
    salvarArquivo(bytes, 'modelo_consulta_simples.xlsx');
  }

  // ------------------------------------------------------------ eventos

  el.arquivo.addEventListener('change', () => {
    const f = el.arquivo.files && el.arquivo.files[0];
    if (f) carregarArquivo(f);
    el.arquivo.value = ''; // permite reenviar o mesmo arquivo
  });

  // Evita que o navegador abra o arquivo se ele for solto fora da área.
  ['dragover', 'drop'].forEach((tipo) => window.addEventListener(tipo, (e) => e.preventDefault()));
  let profundidadeArraste = 0;
  el.zona.addEventListener('dragenter', (e) => {
    e.preventDefault();
    if (estado.rodando) return;
    profundidadeArraste++;
    el.zona.classList.add('arrastando');
  });
  el.zona.addEventListener('dragleave', () => {
    profundidadeArraste = Math.max(0, profundidadeArraste - 1);
    if (!profundidadeArraste) el.zona.classList.remove('arrastando');
  });
  el.zona.addEventListener('drop', (e) => {
    e.preventDefault();
    profundidadeArraste = 0;
    el.zona.classList.remove('arrastando');
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) carregarArquivo(f);
  });

  el.iniciar.addEventListener('click', () => {
    executarConsulta().catch((e) => {
      mostrarErro(`Erro inesperado: ${e && e.message ? e.message : e}`);
    });
  });
  el.cancelar.addEventListener('click', () => {
    if (estado.controle) estado.controle.abort();
  });
  el.baixar.addEventListener('click', baixarResultado);
  el.modelo.addEventListener('click', baixarModelo);
  el.ibsCbs.addEventListener('change', atualizarBotoes);
  el.limparCache.addEventListener('click', () => {
    estado.cache.limpar();
    atualizarInfoCache();
    el.infoCache.textContent = 'Resultados guardados apagados.';
  });
  window.addEventListener('beforeunload', (e) => {
    if (estado.rodando) {
      e.preventDefault();
      e.returnValue = '';
    }
  });

  bibliotecasCarregadas();
  atualizarBotoes();
})();
