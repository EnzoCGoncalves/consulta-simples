/*
 * Núcleo da consulta em lote do Simples Nacional.
 * Só funções sem DOM: tratamento de CNPJ, leitura da planilha, consulta às APIs,
 * montagem do resultado e geração do .xlsx. Funciona no navegador (window.Nucleo)
 * e no Node (module.exports), para poder ser testado.
 */
(function (raiz, fabrica) {
  const api = fabrica();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else raiz.Nucleo = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const COLUNAS_SAIDA = [
    'EMPRESA',
    'CNPJ',
    'Situação no Simples Nacional',
    'Situação no SIMEI',
    'Situação no Regime Regular do IBS',
    'Situação no Regime Regular da CBS',
    'Observação',
  ];

  const TEXTO = {
    simplesSim: (data) => (data ? `Optante pelo Simples Nacional desde ${data}` : 'Optante pelo Simples Nacional'),
    simplesNao: 'NÃO optante pelo Simples Nacional',
    meiSim: (data) => (data ? `Enquadrado no SIMEI desde ${data}` : 'Enquadrado no SIMEI'),
    meiNao: 'NÃO enquadrado no SIMEI',
    regimeNaoOptante: 'Não optante',
    regimeVerificar: 'Verificar no portal do Simples Nacional',
    cnpjInvalido: 'CNPJ inválido',
    cnpjInvalidoDV: 'CNPJ inválido (dígito verificador)',
  };

  // ---------------------------------------------------------------- texto e datas

  /** "Cnpj ", "CNPJ:", "empresa" → "CNPJ", "EMPRESA" (sem acento, maiúsculas, sem pontuação). */
  function normalizarCabecalho(valor) {
    return String(valor == null ? '' : valor)
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .toUpperCase()
      .replace(/[^A-Z0-9]+/g, ' ')
      .trim();
  }

  const doisDigitos = (n) => String(n).padStart(2, '0');

  /** Data local de hoje (ou de `d`) em AAAA-MM-DD. */
  function hojeISO(d) {
    const x = d || new Date();
    return `${x.getFullYear()}-${doisDigitos(x.getMonth() + 1)}-${doisDigitos(x.getDate())}`;
  }

  /** Aceita "2020-06-05", "2020-06-05T00:00:00Z", "05/06/2020" ou "20200605". Vazio se não for data. */
  function paraISO(valor) {
    if (valor == null || valor === '') return '';
    const s = String(valor).trim();
    let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
    if (m) return m[1] === '0000' || m[2] === '00' ? '' : `${m[1]}-${m[2]}-${m[3]}`;
    m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    m = /^(\d{4})(\d{2})(\d{2})$/.exec(s);
    if (m && m[1] !== '0000' && m[2] !== '00') return `${m[1]}-${m[2]}-${m[3]}`;
    return '';
  }

  /** "2020-06-05" → "05/06/2020". */
  function isoParaBR(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? `${m[3]}/${m[2]}/${m[1]}` : '';
  }

  // ---------------------------------------------------------------- CNPJ

  const PESOS_DV1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
  const PESOS_DV2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];

  // Valor de cada caractere = código ASCII − 48. Para dígitos dá o próprio número;
  // para letras (CNPJ alfanumérico, IN RFB 2.229/2024) dá A=17, B=18, …
  function digitoVerificador(base, pesos) {
    let soma = 0;
    for (let i = 0; i < pesos.length; i++) soma += (base.charCodeAt(i) - 48) * pesos[i];
    const resto = soma % 11;
    return resto < 2 ? 0 : 11 - resto;
  }

  function dvCnpjValido(cnpj) {
    if (!/^[0-9A-Z]{12}[0-9]{2}$/.test(cnpj)) return false;
    if (/^(.)\1{13}$/.test(cnpj)) return false; // 00000000000000, 11111111111111…
    const base = cnpj.slice(0, 12);
    const dv1 = digitoVerificador(base, PESOS_DV1);
    const dv2 = digitoVerificador(base + dv1, PESOS_DV2);
    return cnpj.slice(12) === `${dv1}${dv2}`;
  }

  /** Converte o valor bruto da célula em texto, sem notação científica para números. */
  function valorCelula(valor) {
    if (valor == null) return '';
    if (typeof valor === 'number') return Number.isFinite(valor) ? BigInt(Math.round(valor)).toString() : '';
    if (valor instanceof Date) return '';
    return String(valor);
  }

  /**
   * Regras: tira tudo que não é número; < 14 dígitos → zeros à esquerda; vazio ou > 14 → inválido;
   * confere os dígitos verificadores. CNPJ alfanumérico (letras nas 12 primeiras posições,
   * emitido desde jul/2026) é mantido com as letras, já que tirá-las geraria outro CNPJ.
   */
  function tratarCnpj(valor) {
    const texto = valorCelula(valor).trim();
    const alfanumerico = texto.toUpperCase().replace(/[^0-9A-Z]/g, '');
    if (/[A-Z]/.test(alfanumerico) && /^[0-9A-Z]{12}[0-9]{2}$/.test(alfanumerico)) {
      return dvCnpjValido(alfanumerico)
        ? { valido: true, cnpj: alfanumerico, alfanumerico: true }
        : { valido: false, cnpj: alfanumerico, erro: TEXTO.cnpjInvalidoDV };
    }
    const digitos = texto.replace(/\D/g, '');
    if (!digitos || digitos.length > 14) return { valido: false, cnpj: digitos, erro: TEXTO.cnpjInvalido };
    const cnpj = digitos.padStart(14, '0');
    if (!dvCnpjValido(cnpj)) return { valido: false, cnpj, erro: TEXTO.cnpjInvalidoDV };
    return { valido: true, cnpj, alfanumerico: false };
  }

  function cpfValido(cpf) {
    if (!/^\d{11}$/.test(cpf) || /^(\d)\1{10}$/.test(cpf)) return false;
    const dv = (n) => {
      let soma = 0;
      for (let i = 0; i < n; i++) soma += Number(cpf[i]) * (n + 1 - i);
      const r = (soma * 10) % 11;
      return r === 10 ? 0 : r;
    };
    return dv(9) === Number(cpf[9]) && dv(10) === Number(cpf[10]);
  }

  /**
   * O valor da coluna CNPJ é, na verdade, um CPF (pessoa física)?
   * 'cpf' = ignorar a linha; 'talvez' = vale como CPF E como CNPJ (consulta como CNPJ);
   * false = não é CPF.
   * Cuidado: CNPJ guardado como número perde os zeros iniciais e pode ficar com 11 dígitos,
   * igual a um CPF. Por isso a decisão usa a máscara e os dígitos verificadores.
   */
  function classificarCpf(valor) {
    const texto = valorCelula(valor).trim();
    // Escrito como CPF: 123.456.789-09 (também dentro de um texto, como "CPF: 123.456.789-09")
    if (/^\d{3}\.?\d{3}\.?\d{3}-\d{2}$/.test(texto) || /(^|\D)\d{3}\.\d{3}\.\d{3}-\d{2}(\D|$)/.test(texto)) return 'cpf';
    if (/[A-Za-z/]/.test(texto)) return false; // letras (CNPJ alfanumérico) ou barra (máscara de CNPJ)
    const d = texto.replace(/\D/g, '');
    if (!d || d.length > 11) return false; // 12 a 14 dígitos só pode ser CNPJ
    if (!cpfValido(d.padStart(11, '0'))) return false;
    return dvCnpjValido(d.padStart(14, '0')) ? 'talvez' : 'cpf';
  }

  /** 36515322000110 → 36.515.322/0001-10 (só para exibir). */
  function formatarCnpj(cnpj) {
    return /^[0-9A-Z]{12}[0-9]{2}$/.test(cnpj || '')
      ? `${cnpj.slice(0, 2)}.${cnpj.slice(2, 5)}.${cnpj.slice(5, 8)}/${cnpj.slice(8, 12)}-${cnpj.slice(12)}`
      : cnpj || '';
  }

  // ---------------------------------------------------------------- leitura da planilha

  // Cabeçalhos aceitos quando não há uma coluna chamada só "CNPJ" (planilhas com CPF e CNPJ juntos).
  const CABECALHOS_MISTOS = ['CPF CNPJ', 'CNPJ CPF', 'CPF OU CNPJ', 'CNPJ OU CPF'];

  /** Procura, nas primeiras linhas, a que tem a coluna CNPJ (e, se houver, EMPRESA). */
  function localizarColunas(matriz, maxLinhas) {
    const limite = Math.min(matriz.length, maxLinhas || 30);
    for (let r = 0; r < limite; r++) {
      const linha = matriz[r] || [];
      let iEmpresa = -1;
      let iCnpj = -1;
      let iMisto = -1;
      for (let c = 0; c < linha.length; c++) {
        const nome = normalizarCabecalho(linha[c]);
        if (nome === 'EMPRESA' && iEmpresa < 0) iEmpresa = c;
        if (nome === 'CNPJ' && iCnpj < 0) iCnpj = c;
        if (CABECALHOS_MISTOS.includes(nome) && iMisto < 0) iMisto = c;
      }
      if (iCnpj < 0) iCnpj = iMisto;
      if (iCnpj >= 0) return { linha: r, iEmpresa, iCnpj };
    }
    return null;
  }

  /** Lê a primeira aba e devolve as linhas com EMPRESA e CNPJ já tratado. */
  function lerPlanilha(XLSX, dados) {
    const wb = XLSX.read(dados, { type: 'array', cellDates: false });
    const nomeAba = wb.SheetNames[0];
    const ws = nomeAba ? wb.Sheets[nomeAba] : null;
    if (!ws || !ws['!ref']) throw new Error('A primeira aba da planilha está vazia.');
    const linhaInicial = XLSX.utils.decode_range(ws['!ref']).s.r;
    const matriz = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });
    const cab = localizarColunas(matriz);
    if (!cab) {
      throw new Error(
        `Não encontrei a coluna "CNPJ" na primeira aba ("${nomeAba}"). ` +
          'Confira se o cabeçalho tem as colunas EMPRESA e CNPJ, ou use a planilha modelo.'
      );
    }
    const linhas = [];
    const cpfsIgnorados = []; // números das linhas com CPF (pessoa física), que ficam fora do resultado
    for (let r = cab.linha + 1; r < matriz.length; r++) {
      const linha = matriz[r] || [];
      const empresa = cab.iEmpresa >= 0 ? valorCelula(linha[cab.iEmpresa]).trim() : '';
      const cnpjBruto = valorCelula(linha[cab.iCnpj]).trim();
      if (!empresa && !cnpjBruto) continue; // linha em branco
      const linhaExcel = linhaInicial + r + 1;
      const cpf = classificarCpf(linha[cab.iCnpj]);
      if (cpf === 'cpf') {
        cpfsIgnorados.push(linhaExcel);
        continue;
      }
      linhas.push(Object.assign({ linhaExcel, empresa, cnpjBruto, talvezCpf: cpf === 'talvez' }, tratarCnpj(linha[cab.iCnpj])));
    }
    return { nomeAba, linhas, cpfsIgnorados, temColunaEmpresa: cab.iEmpresa >= 0 };
  }

  // ---------------------------------------------------------------- interpretação das APIs

  const simNao = (v) => v === true || /^(S|SIM|TRUE)$/i.test(String(v == null ? '' : v).trim());

  /** BrasilAPI e Minha Receita usam o mesmo formato (dados abertos da Receita). */
  function interpretarPadraoReceita(j) {
    if (!j || typeof j !== 'object' || !('opcao_pelo_simples' in j)) return null;
    return {
      razaoSocial: j.razao_social || '',
      situacaoCadastral: j.descricao_situacao_cadastral || '',
      simples: {
        optante: simNao(j.opcao_pelo_simples),
        desde: paraISO(j.data_opcao_pelo_simples),
        exclusao: paraISO(j.data_exclusao_do_simples),
      },
      mei: {
        optante: simNao(j.opcao_pelo_mei),
        desde: paraISO(j.data_opcao_pelo_mei),
        exclusao: paraISO(j.data_exclusao_do_mei),
      },
      eventosRegimeRegular: [],
    };
  }

  /** CNPJá: company.simples / company.simei = { optant, since, history[{from, until, text}], events[] }. */
  function regimeCnpja(r) {
    if (!r || typeof r !== 'object') return { optante: false, desde: '', exclusao: '' };
    const optante = r.optant === true;
    let ultimo = null;
    if (!optante && Array.isArray(r.history)) {
      for (const p of r.history) if (p && (!ultimo || paraISO(p.until) > paraISO(ultimo.until))) ultimo = p;
    }
    return { optante, desde: paraISO(r.since), exclusao: ultimo ? paraISO(ultimo.until) : '' };
  }

  function interpretarCnpja(j) {
    if (!j || typeof j !== 'object' || !j.company || typeof j.company !== 'object' || !('simples' in j.company)) return null;
    const c = j.company;
    const eventos = (c.simples && Array.isArray(c.simples.events) ? c.simples.events : [])
      .filter((e) => e && (e.regime === 'IBS' || e.regime === 'CBS'))
      .map((e) => ({ regime: e.regime, desde: paraISO(e.from), texto: String(e.text || '').trim() }));
    return {
      razaoSocial: c.name || '',
      situacaoCadastral: (j.status && j.status.text) || '',
      atualizadoEm: paraISO(j.updated),
      simples: regimeCnpja(c.simples),
      mei: regimeCnpja(c.simei),
      eventosRegimeRegular: eventos,
    };
  }

  // ---------------------------------------------------------------- resultado (textos da Receita)

  /**
   * Decide se o período está vigente hoje, olhando a data de exclusão:
   * exclusão futura = ainda optante (exclusão prevista); exclusão já ocorrida que encerra
   * o período atual = não optante. Exclusão anterior a uma nova opção = reingressou.
   * Opção e exclusão no mesmo dia (ex.: 01/07/2007 em grandes empresas, início do Simples)
   * é um período de duração zero: a empresa nunca foi optante, então não vale "Excluída".
   */
  function avaliarPeriodo(reg, hoje) {
    let optante = !!reg.optante;
    let nota = null;
    if (reg.exclusao && reg.exclusao === reg.desde) {
      optante = false;
    } else if (reg.exclusao) {
      if (reg.exclusao > hoje) {
        nota = { tipo: optante ? 'prevista' : 'ocorrida', data: reg.exclusao };
      } else if (!optante || !reg.desde || reg.exclusao >= reg.desde) {
        optante = false;
        nota = { tipo: 'ocorrida', data: reg.exclusao };
      }
    }
    return { optante, desde: reg.desde, nota };
  }

  function textoRegimeRegular(eventos, regime) {
    const doRegime = (eventos || []).filter((e) => e.regime === regime);
    if (!doRegime.length) return TEXTO.regimeVerificar;
    const ev = doRegime.sort((a, b) => (a.desde < b.desde ? 1 : -1))[0];
    return `${ev.texto || 'Opção pelo Regime Regular'}${ev.desde ? ` a partir de ${isoParaBR(ev.desde)}` : ''}`;
  }

  /** Transforma a resposta da consulta nas colunas da planilha de saída. */
  function montarResultado(consulta, hoje) {
    const dia = hoje || hojeISO();
    const diaBR = isoParaBR(dia);
    const vazio = { simples: '', simei: '', ibs: '', cbs: '' };

    if (consulta.tipo === 'ok') {
      const d = consulta.dados;
      const sn = avaliarPeriodo(d.simples, dia);
      const mei = avaliarPeriodo(d.mei, dia);
      const obs = [];
      if (sn.nota) {
        obs.push(
          sn.nota.tipo === 'prevista'
            ? `Exclusão do Simples prevista para ${isoParaBR(sn.nota.data)}`
            : `Excluída do Simples em ${isoParaBR(sn.nota.data)}`
        );
      }
      if (mei.nota) {
        obs.push(
          mei.nota.tipo === 'prevista'
            ? `Desenquadramento do SIMEI previsto para ${isoParaBR(mei.nota.data)}`
            : `Desenquadrada do SIMEI em ${isoParaBR(mei.nota.data)}`
        );
      }
      // IBS/CBS: eventos vêm da própria resposta (se a fonte foi a CNPJá) ou da busca extra na CNPJá.
      const rr = consulta.regimeRegular;
      const eventos =
        d.eventosRegimeRegular && d.eventosRegimeRegular.length
          ? d.eventosRegimeRegular
          : rr && rr.status === 'ok'
            ? rr.eventos
            : [];
      if (sn.optante && rr) {
        const quando = rr.doCache && rr.consultadoEm ? ` (consultado em ${isoParaBR(rr.consultadoEm)}, reaproveitado)` : '';
        if (rr.status === 'ok') {
          obs.push(
            eventos.length
              ? `IBS/CBS: fonte CNPJá${quando}`
              : `IBS/CBS: nenhuma opção registrada na CNPJá${quando} – conferir no portal`
          );
        } else if (rr.status === 'fila' || rr.status === 'pendente') {
          obs.push('IBS/CBS: consulta na CNPJá ainda em andamento – conferir no portal');
        } else if (rr.status === 'cancelado') {
          obs.push('IBS/CBS: não consultado na CNPJá (consulta cancelada) – conferir no portal');
        } else {
          obs.push(`IBS/CBS: CNPJá indisponível (${rr.motivo}) – conferir no portal`);
        }
      }
      const situacao = normalizarCabecalho(d.situacaoCadastral);
      if (situacao && situacao !== 'ATIVA') obs.push(`Situação cadastral na Receita: ${String(d.situacaoCadastral).toUpperCase()}`);
      const atualizacao = d.atualizadoEm ? ` (dados de ${isoParaBR(d.atualizadoEm)})` : '';
      const consultadoEm = isoParaBR(consulta.consultadoEm) || diaBR;
      obs.push(
        `Fonte: ${consulta.fonte}${atualizacao} – consultado em ${consultadoEm}` +
          (consulta.doCache ? ' (resultado reaproveitado)' : '')
      );

      return {
        tipo: 'ok',
        optanteSimples: sn.optante,
        simples: sn.optante ? TEXTO.simplesSim(isoParaBR(sn.desde)) : TEXTO.simplesNao,
        simei: mei.optante ? TEXTO.meiSim(isoParaBR(mei.desde)) : TEXTO.meiNao,
        // Só quem é do Simples pode optar pelo Regime Regular do IBS/CBS.
        ibs: sn.optante ? textoRegimeRegular(eventos, 'IBS') : TEXTO.regimeNaoOptante,
        cbs: sn.optante ? textoRegimeRegular(eventos, 'CBS') : TEXTO.regimeNaoOptante,
        observacao: obs.join('; '),
      };
    }

    const falhas = consulta.falhas && consulta.falhas.length ? ` (${consulta.falhas.join('; ')})` : '';
    if (consulta.tipo === 'nao_encontrado') {
      return Object.assign({ tipo: 'nao_encontrado' }, vazio, {
        observacao: `CNPJ não encontrado na base pública de CNPJ da Receita${falhas} – consultado em ${diaBR}`,
      });
    }
    return Object.assign({ tipo: 'falha' }, vazio, {
      observacao: `Falha na consulta${falhas} – tentativa em ${diaBR}`,
    });
  }

  // ---------------------------------------------------------------- consulta às APIs

  // Ordem de tentativa. A BrasilAPI só repassa a consulta para a Minha Receita (ver services/cnpj.js
  // no código dela), então a Minha Receita é chamada direto, sem intermediário; a BrasilAPI fica de
  // reserva caso o acesso direto falhe.
  const APIS = [
    {
      id: 'minhareceita',
      nome: 'Minha Receita',
      url: (cnpj) => `https://minhareceita.org/${cnpj}`,
      interpretar: interpretarPadraoReceita,
      intervaloMinimoMs: 0,
    },
    {
      id: 'brasilapi',
      nome: 'BrasilAPI',
      url: (cnpj) => `https://brasilapi.com.br/api/cnpj/v1/${cnpj}`,
      interpretar: interpretarPadraoReceita,
      intervaloMinimoMs: 0,
    },
    {
      id: 'cnpja',
      nome: 'CNPJá',
      url: (cnpj) => `https://open.cnpja.com/office/${cnpj}`,
      interpretar: interpretarCnpja,
      intervaloMinimoMs: 12500, // API aberta: 5 consultas por minuto por IP
    },
  ];

  const NOVAS_TENTATIVAS = 3;

  function erroCancelado() {
    const e = new Error('Consulta cancelada');
    e.name = 'AbortError';
    return e;
  }

  /** setTimeout que pode ser interrompido pelo botão Cancelar. */
  function esperar(ms, signal) {
    return new Promise((resolve, reject) => {
      if (signal && signal.aborted) return reject(erroCancelado());
      const aoCancelar = () => {
        clearTimeout(timer);
        reject(erroCancelado());
      };
      const timer = setTimeout(() => {
        if (signal) signal.removeEventListener('abort', aoCancelar);
        resolve();
      }, Math.max(0, ms));
      if (signal) signal.addEventListener('abort', aoCancelar, { once: true });
    });
  }

  /** Retry-After em segundos ou data HTTP → milissegundos (máx. 2 min). */
  function retryAfterMs(valor, agora) {
    if (!valor) return 0;
    const seg = Number(valor);
    const ms = Number.isFinite(seg) ? seg * 1000 : Date.parse(valor) - agora;
    return Number.isFinite(ms) && ms > 0 ? Math.min(ms, 120000) : 0;
  }

  /** GET com tempo limite. status 0 = falha de rede ou tempo esgotado. */
  async function requisitar(fetchFn, url, signal, tempoLimiteMs) {
    if (signal && signal.aborted) throw erroCancelado();
    const ctrl = new AbortController();
    let estourou = false;
    const aoCancelar = () => ctrl.abort();
    if (signal) signal.addEventListener('abort', aoCancelar, { once: true });
    const timer = setTimeout(() => {
      estourou = true;
      ctrl.abort();
    }, tempoLimiteMs);
    try {
      const resp = await fetchFn(url, {
        method: 'GET',
        headers: { Accept: 'application/json' },
        credentials: 'omit',
        referrerPolicy: 'no-referrer',
        signal: ctrl.signal,
      });
      let corpo = null;
      if (resp.status === 200) {
        try {
          corpo = JSON.parse(await resp.text());
        } catch (e) {
          corpo = null;
        }
      }
      return { status: resp.status, retryAfter: resp.headers.get('retry-after'), corpo };
    } catch (e) {
      if (signal && signal.aborted) throw erroCancelado();
      return { status: 0, erro: estourou ? 'tempo de resposta esgotado' : 'falha de rede' };
    } finally {
      clearTimeout(timer);
      if (signal) signal.removeEventListener('abort', aoCancelar);
    }
  }

  /**
   * Cria o consultor. Guarda o horário da última chamada de cada API para respeitar
   * o limite da CNPJá entre um CNPJ e outro.
   * opcoes: fetch, esperar, agora, apis, tempoLimiteMs, esperaLimiteMs, esperaRedeMs (para testes).
   */
  function criarConsultor(opcoes) {
    const o = opcoes || {};
    const fetchFn = o.fetch || ((url, init) => fetch(url, init));
    const dormir = o.esperar || esperar;
    const agora = o.agora || (() => Date.now());
    const apis = o.apis || APIS;
    const tempoLimiteMs = o.tempoLimiteMs || 20000;
    const esperaLimiteMs = o.esperaLimiteMs != null ? o.esperaLimiteMs : 5000; // 429: 5 s, 10 s, 20 s
    const esperaRedeMs = o.esperaRedeMs != null ? o.esperaRedeMs : 3000; // rede: 3 s, 6 s, 12 s
    const filaDaApi = new Map(); // api.id → promessa com o horário REAL da última chamada

    // Fila rigorosa: cada chamada espera a anterior realmente sair e só então conta o intervalo.
    // Mesmo com várias consultas em paralelo (e atrasos do relógio do navegador), duas chamadas
    // à CNPJá nunca saem com menos de 12,5 s de diferença.
    async function aguardarVez(api, signal, avisar) {
      if (!api.intervaloMinimoMs) return;
      const anterior = filaDaApi.get(api.id) || Promise.resolve(-Infinity);
      let liberar;
      filaDaApi.set(api.id, new Promise((resolve) => (liberar = resolve)));
      const ultima = await anterior;
      try {
        const falta = ultima + api.intervaloMinimoMs - agora();
        if (falta > 0) {
          avisar({ tipo: 'espera', api: api.nome, ms: falta });
          await dormir(falta, signal);
        }
        liberar(agora()); // esta chamada sai agora
      } catch (e) {
        liberar(ultima); // cancelada: não houve chamada, a fila segue com o horário anterior
        throw e;
      }
    }

    async function consultarNaApi(api, cnpj, signal, avisar) {
      let motivo = 'sem resposta';
      for (let tentativa = 0; tentativa <= NOVAS_TENTATIVAS; tentativa++) {
        await aguardarVez(api, signal, avisar);
        avisar({ tipo: 'consultando', api: api.nome, tentativa });

        const r = await requisitar(fetchFn, api.url(cnpj), signal, tempoLimiteMs);
        let espera;
        if (r.status === 200) {
          const dados = r.corpo ? api.interpretar(r.corpo) : null;
          return dados ? { ok: true, dados } : { ok: false, motivo: 'resposta em formato inesperado' };
        }
        if (r.status === 404) return { ok: false, naoEncontrado: true };
        if (r.status === 429) {
          motivo = 'limite de requisições (erro 429)';
          espera = retryAfterMs(r.retryAfter, agora()) || esperaLimiteMs * 2 ** tentativa;
        } else if (r.status === 0) {
          motivo = r.erro;
          espera = esperaRedeMs * 2 ** tentativa;
        } else if (r.status >= 500) {
          motivo = `erro ${r.status} no servidor`;
          espera = esperaRedeMs * 2 ** tentativa;
        } else {
          return { ok: false, motivo: `erro ${r.status}` }; // 400, 403…: repetir não resolve
        }
        if (tentativa < NOVAS_TENTATIVAS) {
          avisar({
            tipo: 'nova_tentativa',
            api: api.nome,
            motivo,
            ms: espera,
            tentativa: tentativa + 1,
            total: NOVAS_TENTATIVAS,
            limite: r.status === 429,
          });
          await dormir(espera, signal);
        }
      }
      return { ok: false, motivo: `${motivo}, mesmo após ${NOVAS_TENTATIVAS} novas tentativas` };
    }

    /** Consulta extra na CNPJá só para pegar os eventos do Regime Regular do IBS/CBS. */
    async function buscarRegimeRegular(cnpj, signal, avisar) {
      const cnpja = apis.find((a) => a.id === 'cnpja');
      if (!cnpja) return { status: 'falha', motivo: 'CNPJá não configurada' };
      const avisarIbs = avisar || function () {};
      const r = await consultarNaApi(cnpja, cnpj, signal, (ev) => avisarIbs(Object.assign({ finalidade: 'ibscbs' }, ev)));
      if (r.ok) return { status: 'ok', eventos: r.dados.eventosRegimeRegular || [] };
      return { status: 'falha', motivo: r.naoEncontrado ? 'CNPJ não encontrado na CNPJá' : r.motivo };
    }

    /**
     * Tenta a API principal; se falhar (ou não achar o CNPJ), passa para as reservas.
     * Com ctx.regimeRegular = true, faz uma 2ª consulta na CNPJá para as optantes do Simples
     * (só o IBS/CBS; o resto continua vindo da API que respondeu).
     */
    async function consultar(cnpj, ctx) {
      const signal = ctx && ctx.signal;
      const avisar = (ctx && ctx.avisar) || function () {};
      const falhas = [];
      let naoEncontrado = false;
      for (let i = 0; i < apis.length; i++) {
        const api = apis[i];
        const r = await consultarNaApi(api, cnpj, signal, avisar);
        if (r.ok) {
          const resultado = { tipo: 'ok', fonte: api.nome, dados: r.dados, falhas };
          const optante = avaliarPeriodo(r.dados.simples, (ctx && ctx.hoje) || hojeISO()).optante;
          if (ctx && ctx.regimeRegular && api.id !== 'cnpja' && optante) {
            resultado.regimeRegular = await buscarRegimeRegular(cnpj, signal, avisar);
          }
          return resultado;
        }
        if (r.naoEncontrado) naoEncontrado = true;
        else falhas.push(`${api.nome}: ${r.motivo}`);
        if (i < apis.length - 1) {
          avisar({ tipo: 'reserva', api: api.nome, proxima: apis[i + 1].nome, naoEncontrado: !!r.naoEncontrado });
        }
      }
      return naoEncontrado ? { tipo: 'nao_encontrado', falhas } : { tipo: 'falha', falhas };
    }

    return {
      consultar,
      buscarRegimeRegular: (cnpj, ctx) => buscarRegimeRegular(cnpj, ctx && ctx.signal, ctx && ctx.avisar),
    };
  }

  /** A empresa é optante do Simples HOJE (considerando datas de exclusão)? */
  function optanteHoje(consulta, hoje) {
    return !!(consulta && consulta.tipo === 'ok' && avaliarPeriodo(consulta.dados.simples, hoje || hojeISO()).optante);
  }

  // ---------------------------------------------------------------- resultados guardados

  const DIA_MS = 24 * 60 * 60 * 1000;

  /**
   * Guarda no navegador (localStorage) as respostas já consultadas, para não consultar de novo.
   * Só guarda respostas corretas (nunca falhas ou "não encontrado") e só reaproveita
   * as que tiverem até `validadeDias` dias. Tudo protegido por try/catch: se o navegador
   * bloquear o armazenamento, o site só deixa de reaproveitar.
   */
  function criarCache(storage, opcoes) {
    const o = opcoes || {};
    const PREFIXO = 'consultaSimples:v1:';
    const agora = o.agora || (() => Date.now());

    function ler(chave) {
      try {
        const v = storage && storage.getItem(PREFIXO + chave);
        return v ? JSON.parse(v) : null;
      } catch (e) {
        return null;
      }
    }
    function gravar(chave, valor) {
      try {
        if (storage) storage.setItem(PREFIXO + chave, JSON.stringify(valor));
      } catch (e) {
        /* armazenamento cheio ou bloqueado: segue sem guardar */
      }
    }
    function chaves() {
      const lista = [];
      try {
        for (let i = 0; storage && i < storage.length; i++) {
          const k = storage.key(i);
          if (k && k.indexOf(PREFIXO) === 0) lista.push(k);
        }
      } catch (e) {
        /* ignora */
      }
      return lista;
    }
    const valido = (e, dias) => !!(e && typeof e.salvoEm === 'number' && dias > 0 && agora() - e.salvoEm < dias * DIA_MS);

    return {
      lerConsulta(cnpj, dias) {
        const e = ler(`s:${cnpj}`);
        if (!valido(e, dias) || e.tipo !== 'ok' || !e.dados) return null;
        return { tipo: 'ok', fonte: e.fonte, dados: e.dados, falhas: [], consultadoEm: e.consultadoEm, doCache: true };
      },
      gravarConsulta(cnpj, consulta) {
        if (!consulta || consulta.tipo !== 'ok') return;
        gravar(`s:${cnpj}`, {
          tipo: 'ok',
          fonte: consulta.fonte,
          dados: consulta.dados,
          consultadoEm: consulta.consultadoEm,
          salvoEm: agora(),
        });
      },
      lerIbs(cnpj, dias) {
        const e = ler(`i:${cnpj}`);
        if (!valido(e, dias) || e.status !== 'ok') return null;
        return { status: 'ok', eventos: e.eventos || [], consultadoEm: e.consultadoEm, doCache: true };
      },
      gravarIbs(cnpj, rr) {
        if (!rr || rr.status !== 'ok') return;
        gravar(`i:${cnpj}`, { status: 'ok', eventos: rr.eventos || [], consultadoEm: rr.consultadoEm, salvoEm: agora() });
      },
      /** Apaga o que passou do prazo (ou estiver corrompido). */
      podar(dias) {
        for (const k of chaves()) {
          const e = ler(k.slice(PREFIXO.length));
          if (!valido(e, Math.max(dias || 0, 7))) {
            try {
              storage.removeItem(k);
            } catch (err) {
              /* ignora */
            }
          }
        }
      },
      limpar() {
        for (const k of chaves()) {
          try {
            storage.removeItem(k);
          } catch (e) {
            /* ignora */
          }
        }
      },
      contar() {
        return chaves().filter((k) => k.indexOf(`${PREFIXO}s:`) === 0).length;
      },
    };
  }

  // ---------------------------------------------------------------- geração do .xlsx

  function larguraColunas(aoa, minimo, maximo) {
    const larguras = [];
    for (const linha of aoa) {
      linha.forEach((v, c) => {
        const tam = String(v == null ? '' : v).length;
        larguras[c] = Math.max(larguras[c] || 0, tam);
      });
    }
    return larguras.map((l, c) => ({ wch: Math.min(Math.max(l + 2, (minimo && minimo[c]) || 10), maximo || 70) }));
  }

  /**
   * O SheetJS gratuito não grava negrito nem painel congelado. Depois de gerar o arquivo,
   * ajusta o XML: fonte negrita no cabeçalho, 1ª linha congelada e coluna de texto ("@").
   * Se algo der errado, devolve o arquivo original (sem esses detalhes visuais).
   */
  async function formatarXlsx(JSZip, dados, colunaTexto) {
    if (!JSZip) return new Uint8Array(dados);
    try {
      const zip = await JSZip.loadAsync(dados);
      const caminho = Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort()[0];
      let estilos = await zip.file('xl/styles.xml').async('string');
      let planilha = await zip.file(caminho).async('string');

      // 1) fonte negrita = cópia da fonte padrão + <b/>
      const reFontes = /<fonts\b([^>]*)>([\s\S]*?)<\/fonts>/;
      const mF = reFontes.exec(estilos);
      if (!mF) throw new Error('styles.xml sem <fonts>');
      const fontes = mF[2].match(/<font\b[^>]*\/>|<font\b[^>]*>[\s\S]*?<\/font>/g) || [];
      const primeira = fontes[0] && !/\/>$/.test(fontes[0]) ? fontes[0].replace(/^<font\b[^>]*>|<\/font>$/g, '') : '';
      const fonteNegrito = `<font><b/>${primeira.replace(/<b\s*\/>/g, '') || '<sz val="11"/><name val="Calibri"/>'}</font>`;
      const idFonteNegrito = fontes.length;
      estilos = estilos.replace(
        reFontes,
        () => `<fonts${mF[1].replace(/\bcount="\d+"/, `count="${fontes.length + 1}"`)}>${mF[2]}${fonteNegrito}</fonts>`
      );

      // 2) estilos de célula: negrito (cabeçalho) e texto "@" (coluna CNPJ)
      const reXfs = /<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/;
      const mX = reXfs.exec(estilos);
      if (!mX) throw new Error('styles.xml sem <cellXfs>');
      const qtdXfs = (mX[2].match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g) || []).length;
      const idNegrito = qtdXfs;
      const idTexto = qtdXfs + 1;
      const novos =
        `<xf numFmtId="0" fontId="${idFonteNegrito}" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
        '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>';
      estilos = estilos.replace(
        reXfs,
        () => `<cellXfs${mX[1].replace(/\bcount="\d+"/, `count="${qtdXfs + 2}"`)}>${mX[2]}${novos}</cellXfs>`
      );

      // 3) cabeçalho (linha 1) em negrito
      planilha = planilha.replace(
        /<row\b([^>]*\br="1"[^>]*)>([\s\S]*?)<\/row>/,
        (m, attrs, celulas) =>
          `<row${attrs}>${celulas.replace(
            /<c\b([^>]*?)(\/?)>/g,
            (m2, at, fecha) => `<c${at.replace(/\s+s="\d+"/, '')} s="${idNegrito}"${fecha}>`
          )}</row>`
      );

      // 4) congelar a primeira linha
      const painel =
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
      if (/<sheetView\b[^>]*\/>/.test(planilha)) {
        planilha = planilha.replace(/<sheetView\b([^>]*)\/>/, (m, at) => `<sheetView${at}>${painel}</sheetView>`);
      } else if (!/<sheetViews\b/.test(planilha)) {
        planilha = planilha.replace(
          /(<dimension\b[^>]*\/>)/,
          (m) => `${m}<sheetViews><sheetView workbookViewId="0">${painel}</sheetView></sheetViews>`
        );
      }

      // 5) coluna inteira como texto, para quem digitar novos CNPJs não perder o zero inicial
      if (colunaTexto != null) {
        const n = colunaTexto + 1;
        planilha = planilha.replace(/<col\b([^>]*?)\/>/g, (m, at) => {
          const min = Number((/\bmin="(\d+)"/.exec(at) || [])[1]);
          const max = Number((/\bmax="(\d+)"/.exec(at) || [])[1]);
          return min <= n && n <= max ? `<col${at.replace(/\s+style="\d+"/, '')} style="${idTexto}"/>` : m;
        });
      }

      zip.file('xl/styles.xml', estilos);
      zip.file(caminho, planilha);
      return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    } catch (e) {
      if (typeof console !== 'undefined') console.warn('Formatação extra do Excel não aplicada:', e);
      return new Uint8Array(dados);
    }
  }

  /**
   * Gera o .xlsx: cabeçalho + linhas. A coluna `colunaTexto` (CNPJ) é gravada como TEXTO.
   * Devolve Uint8Array.
   */
  async function gerarXlsx(XLSX, JSZip, cabecalho, linhas, opcoes) {
    const o = opcoes || {};
    const aoa = [cabecalho].concat(linhas);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (o.colunaTexto != null) {
      for (let r = 1; r < aoa.length; r++) {
        const ref = XLSX.utils.encode_cell({ r, c: o.colunaTexto });
        const v = aoa[r][o.colunaTexto];
        ws[ref] = { t: 's', v: v == null ? '' : String(v), z: '@' };
      }
    }
    ws['!cols'] = larguraColunas(aoa, o.larguraMinima, 70);
    if (linhas.length) {
      ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: cabecalho.length - 1 } }),
      };
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, String(o.nomeAba || 'Consulta').slice(0, 31));
    const bruto = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true, bookSST: true });
    return formatarXlsx(JSZip, bruto, o.colunaTexto);
  }

  function nomeArquivoResultado(d) {
    return `consulta_simples_${hojeISO(d)}.xlsx`;
  }

  return {
    COLUNAS_SAIDA,
    TEXTO,
    APIS,
    normalizarCabecalho,
    hojeISO,
    paraISO,
    isoParaBR,
    dvCnpjValido,
    cpfValido,
    classificarCpf,
    tratarCnpj,
    formatarCnpj,
    localizarColunas,
    lerPlanilha,
    interpretarPadraoReceita,
    interpretarCnpja,
    montarResultado,
    optanteHoje,
    esperar,
    criarConsultor,
    criarCache,
    gerarXlsx,
    formatarXlsx,
    nomeArquivoResultado,
  };
});
