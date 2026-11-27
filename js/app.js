/* Interface da consulta em lote. A lógica (CNPJ, APIs, Excel) fica em nucleo.js. */
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
    pausa: $('pausa'),
    ibsCbs: $('ibsCbs'),
    progresso: $('progresso'),
    barra: $('barra'),
    barraTrilho: $('barraTrilho'),
    textoProgresso: $('textoProgresso'),
    detalhe: $('detalhe'),
    cartaoTabela: $('cartaoTabela'),
    contagem: $('contagem'),
    corpo: $('corpo'),
  };

  const estado = {
    nomeArquivo: '',
    nomeAba: '',
    linhas: [], // uma por linha da planilha
    unicos: [], // CNPJs válidos sem repetição, na ordem em que aparecem
    porCnpj: new Map(), // cnpj → linhas da planilha com esse CNPJ
    resultados: new Map(), // cnpj → resultado montado
    rodando: false,
    executou: false,
    cancelado: false,
    controle: null,
    consultor: N.criarConsultor(), // um só por página: lembra o intervalo da CNPJá
  };

  const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  // ------------------------------------------------------------ utilidades de tela

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

  function pausaMs() {
    const v = parseFloat(String(el.pausa.value).replace(',', '.'));
    return Math.round((Number.isFinite(v) ? Math.min(Math.max(v, 0.5), 10) : 1.5) * 1000);
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
    estado.resultados = new Map();
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
    const estimativa = estado.unicos.length * (pausaMs() + 800);

    el.resumo.replaceChildren();
    const titulo = document.createElement('strong');
    titulo.textContent = nomeArquivo;
    el.resumo.append(titulo, ` — aba “${lido.nomeAba}”`);
    const lista = document.createElement('ul');
    const itens = [
      `${estado.linhas.length} ${estado.linhas.length === 1 ? 'linha' : 'linhas'} com dados`,
      `${estado.unicos.length} CNPJ${estado.unicos.length === 1 ? '' : 's'} para consultar` +
        (estado.unicos.length
          ? ` (≈ ${formatarDuracao(estimativa)}, mais até 12,5 s por optante se a busca de IBS/CBS estiver ligada)`
          : ''),
    ];
    if (repetidos) itens.push(`${repetidos} repetido${repetidos === 1 ? '' : 's'} (consultado${repetidos === 1 ? '' : 's'} uma vez só)`);
    if (invalidos) itens.push(`${invalidos} CNPJ${invalidos === 1 ? '' : 's'} inválido${invalidos === 1 ? '' : 's'} (não ser${invalidos === 1 ? 'á' : 'ão'} consultado${invalidos === 1 ? '' : 's'})`);
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

  // ------------------------------------------------------------ tabela

  function pendentes() {
    return estado.unicos.filter((c) => {
      const r = estado.resultados.get(c);
      return !r || r.tipo === 'falha';
    });
  }

  /** As 7 colunas de saída de uma linha (usadas na tela e no Excel). */
  function valoresSaida(l) {
    const obs = [];
    let cols = ['', '', '', ''];
    if (!l.valido) {
      obs.push(l.erro);
    } else {
      const r = estado.resultados.get(l.cnpj);
      const iguais = estado.porCnpj.get(l.cnpj) || [];
      if (iguais.length > 1) obs.push(`CNPJ repetido na planilha (linhas ${iguais.map((x) => x.linhaExcel).join(', ')})`);
      if (r) {
        cols = [r.simples, r.simei, r.ibs, r.cbs];
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
    if (/^(Optante|Enquadrado)/.test(texto)) return 'sim';
    if (/^(NÃO|Não optante)/.test(texto)) return 'nao';
    if (/^Verificar/.test(texto)) return 'atencao';
    if (/^Opção/.test(texto)) return 'sim';
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
    const v = valoresSaida(l);
    const r = l.valido ? estado.resultados.get(l.cnpj) : null;
    const tds = l.tr.children;
    for (let i = 0; i < 7; i++) {
      const td = tds[i + 1];
      td.textContent = v[i];
      if (i >= 2 && i <= 5) td.className = classeValor(v[i]);
    }
    let situacao = 'pendente';
    if (!l.valido) situacao = 'invalido';
    else if (r) situacao = r.tipo === 'ok' ? 'ok' : 'erro';
    else if (l.consultando) situacao = 'consultando';
    else if (estado.cancelado) situacao = 'erro';
    l.tr.className = situacao;
    tds[7].className = situacao === 'invalido' || situacao === 'erro' ? 'obs erro' : 'obs';
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
    for (const r of estado.resultados.values()) {
      if (r.tipo !== 'ok') erros++;
      else if (r.optanteSimples) sim++;
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
    const faltam = pendentes().length;
    el.iniciar.hidden = estado.rodando;
    el.cancelar.hidden = !estado.rodando;
    el.iniciar.disabled = !temArquivo || faltam === 0;
    if (!temArquivo || !estado.executou) el.iniciar.textContent = 'Iniciar consulta';
    else if (faltam) el.iniciar.textContent = `Consultar pendentes (${faltam})`;
    else el.iniciar.textContent = 'Consulta concluída';
    el.baixar.disabled = estado.rodando || !temArquivo || (!estado.executou && estado.unicos.length > 0);
    el.ibsCbs.disabled = estado.rodando;
    el.zona.classList.toggle('bloqueada', estado.rodando);
    el.arquivo.disabled = estado.rodando;
  }

  // ------------------------------------------------------------ progresso

  function mostrarProgresso(feitos, total, texto) {
    const pct = total ? Math.round((feitos / total) * 100) : 100;
    el.barra.style.width = `${pct}%`;
    el.barraTrilho.setAttribute('aria-valuenow', String(pct));
    el.textoProgresso.textContent = texto;
    el.progresso.hidden = false;
  }

  function aviso(ev) {
    const seg = (ms) => Math.ceil(ms / 1000);
    let t = '';
    const ibs = ev.finalidade === 'ibscbs';
    if (ev.tipo === 'consultando') t = ibs ? `Buscando IBS/CBS na ${ev.api}…` : `Consultando na ${ev.api}…`;
    else if (ev.tipo === 'espera') {
      t = `Aguardando ${seg(ev.ms)} s para ${ibs ? 'buscar IBS/CBS na' : 'usar a'} ${ev.api} (limite de 5 consultas por minuto)…`;
    }
    else if (ev.tipo === 'nova_tentativa') {
      t = `${ev.api}: ${ev.motivo}. Nova tentativa em ${seg(ev.ms)} s (${ev.tentativa} de ${ev.total})…`;
      if (ev.limite && ev.api !== 'CNPJá') {
        // A API pediu calma: aumenta a pausa entre consultas para as próximas.
        const nova = Math.min(pausaMs() / 1000 + 1, 10);
        el.pausa.value = String(nova);
        t += ` A pausa entre consultas passou para ${nova} s.`;
      }
    } else if (ev.tipo === 'reserva') {
      t = ev.naoEncontrado
        ? `${ev.api} não encontrou o CNPJ; conferindo na ${ev.proxima}…`
        : `${ev.api} não respondeu; tentando a reserva ${ev.proxima}…`;
    }
    if (t) el.detalhe.textContent = t;
  }

  // ------------------------------------------------------------ consulta

  async function executarConsulta() {
    if (estado.rodando) return;
    const fila = pendentes();
    if (!fila.length) return;

    estado.rodando = true;
    estado.executou = true;
    estado.cancelado = false;
    estado.controle = new AbortController();
    const signal = estado.controle.signal;
    const total = estado.unicos.length;
    let feitos = total - fila.length;
    const inicio = Date.now();
    for (const l of estado.linhas) if (l.tr) atualizarLinhaTabela(l);
    atualizarBotoes();
    mostrarErro('');

    try {
      for (let i = 0; i < fila.length; i++) {
        const cnpj = fila[i];
        const linhas = estado.porCnpj.get(cnpj);
        let texto = `Consultando ${feitos + 1} de ${total}…`;
        if (i > 0) {
          const media = (Date.now() - inicio) / i;
          texto += ` (faltam ≈ ${formatarDuracao(media * (fila.length - i))})`;
        }
        mostrarProgresso(feitos, total, texto);
        el.detalhe.textContent = '';
        for (const l of linhas) {
          l.consultando = true;
          atualizarLinhaTabela(l);
        }
        if (linhas[0].tr && i % 5 === 0) linhas[0].tr.scrollIntoView({ block: 'nearest' });

        let consulta;
        try {
          consulta = await estado.consultor.consultar(cnpj, { signal, avisar: aviso, regimeRegular: el.ibsCbs.checked });
        } catch (e) {
          if (e && e.name === 'AbortError') throw e;
          consulta = { tipo: 'falha', falhas: [e && e.message ? e.message : 'erro inesperado'] };
        } finally {
          for (const l of linhas) l.consultando = false;
        }
        estado.resultados.set(cnpj, N.montarResultado(consulta));
        feitos++;
        for (const l of linhas) atualizarLinhaTabela(l);
        atualizarContagem();
        mostrarProgresso(feitos, total, `Consultando ${Math.min(feitos + 1, total)} de ${total}…`);

        if (i < fila.length - 1) await N.esperar(pausaMs(), signal);
      }
      const falhas = pendentes().length;
      mostrarProgresso(
        total,
        total,
        `Consulta concluída: ${total} de ${total} CNPJs.` +
          (falhas ? ` ${falhas} com falha — clique em “Consultar pendentes” para tentar de novo.` : '')
      );
      el.detalhe.textContent = 'Clique em “Baixar planilha” para salvar o resultado.';
    } catch (e) {
      if (!e || e.name !== 'AbortError') throw e;
      estado.cancelado = true;
      for (const l of estado.linhas) if (l.tr) atualizarLinhaTabela(l);
      const ok = estado.unicos.length - pendentes().length;
      mostrarProgresso(ok, total, `Consulta cancelada: ${ok} de ${total} CNPJs consultados.`);
      el.detalhe.textContent = 'Você já pode baixar a planilha com o que foi consultado, ou continuar depois.';
    } finally {
      estado.rodando = false;
      estado.controle = null;
      atualizarBotoes();
    }
  }

  // ------------------------------------------------------------ downloads

  async function baixarResultado() {
    if (!bibliotecasCarregadas() || !estado.linhas.length) return;
    try {
      const linhas = estado.linhas.map(valoresSaida);
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
  window.addEventListener('beforeunload', (e) => {
    if (estado.rodando) {
      e.preventDefault();
      e.returnValue = '';
    }
  });

  bibliotecasCarregadas();
  atualizarBotoes();
})();
