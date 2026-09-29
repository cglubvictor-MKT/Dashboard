// =====================================================================================
//  COLETOR DO PAINEL DE FATURAMENTO — colar no Console (F12) de uma aba do Fiori my436333
//  Lê o faturamento do SAP a cada 10 min e grava o dados.json no Gist do GitHub.
//  Comandos no Console:  __coletor.status()   __coletor.agora()   __coletor.parar()
//                        __coletor.modelo()   (copia a lista de vendedores para colar no Excel)
//  MODO TESTE: com GIST_ID ainda em COLE_AQUI..., ele só lê o SAP e mostra o resumo por vendedor, sem gravar nada.
// =====================================================================================
(async () => {
  const CFG = {
    GIST_ID: 'COLE_AQUI_O_ID_DO_GIST',
    DIA: '2026-09-29',             // só as notas com data de faturamento neste dia
    INTERVALO_MIN: 10,
    TIPOS_VENDA: ['F2'],           // contam como faturamento
    TIPOS_DEVOLUCAO: ['CBRE'],     // entram negativos (abatem o faturamento)
    TIPOS_ESTORNO: ['S1', 'S1B', 'S2'] // anulam o documento que apontam em CancelledBillingDocument
  };

  if (window.__coletor) window.__coletor.parar();
  const TESTE = !CFG.GIST_ID || CFG.GIST_ID.startsWith('COLE_AQUI');
  const TOKEN = TESTE ? 'teste' : prompt('Token do GitHub (permissão gist). Ele fica só nesta aba e some ao fechar:');
  if (!TOKEN) { console.log('Coletor cancelado: sem token.'); return; }

  const B = '/sap/opu/odata/sap/UI_BILLINGDOCUMENTFS/';
  const P = '/sap/opu/odata/sap/API_PRODUCT_SRV/';
  const GH = { Authorization: 'Bearer ' + TOKEN, Accept: 'application/vnd.github+json' };
  const docs = new Map();        // BillingDocument -> documento resumido
  const litrosUn = new Map();    // "material|unidade" -> litros por unidade
  const nomesCli = {}, nomesVend = {};
  let timer = null, rodando = false, ultimo = null, erros = 0, semSelect = false;

  const log = (...a) => console.log('%c[coletor ' + new Date().toLocaleTimeString('pt-BR') + ']', 'color:#7a9d00;font-weight:bold', ...a);
  const dataSAP = s => new Date(+String(s).match(/-?\d+/)[0]).toISOString().slice(0, 10);
  const dtOData = k => `datetime'${k}T00:00:00'`;

  async function sap(url) {
    const r = await fetch(url + (url.includes('?') ? '&' : '?') + '$format=json', { headers: { Accept: 'application/json' }, credentials: 'include' });
    const tipo = r.headers.get('content-type') || '';
    if (!r.ok || !tipo.includes('json')) throw new Error(`SAP respondeu ${r.status}${tipo.includes('html') ? ' (sessão expirada? faça login de novo nesta aba)' : ''}`);
    return (await r.json()).d;
  }

  // ---------- 1) Faturamentos a partir de uma data, com parceiros e itens ----------
  async function lerFaturamentos(dia) {
    const sel = ['BillingDocument', 'BillingDocumentType', 'BillingDocumentDate', 'SoldToParty', 'CancelledBillingDocument', 'TotalNetAmount',
      'to_Partner/PartnerFunction', 'to_Partner/Personnel', 'to_Partner/FullName',
      'to_Item/Material', 'to_Item/BillingQuantity', 'to_Item/BillingQuantityUnit', 'to_Item/NetAmount'].join(',');
    const lote = 100; let skip = 0, n = 0;
    for (;;) {
      const base = `${B}C_BillingDocumentFs?$filter=BillingDocumentDate eq ${dtOData(dia)}&$orderby=BillingDocument&$top=${lote}&$skip=${skip}&$expand=to_Partner,to_Item`;
      let d;
      try { d = await sap(semSelect ? base : base + '&$select=' + sel); }
      catch (e) { if (semSelect || !/ 400/.test(e.message)) throw e; semSelect = true; log('Serviço não aceita $select, seguindo sem ele'); continue; }
      for (const f of d.results) {
        const parc = f.to_Partner?.results || [];
        const ve = parc.find(p => p.PartnerFunction === 'VE'), ag = parc.find(p => p.PartnerFunction === 'AG');
        if (ve) nomesVend[ve.Personnel] = ve.FullName;
        if (ag) nomesCli[f.SoldToParty] = ag.FullName;
        docs.set(f.BillingDocument, {
          tipo: f.BillingDocumentType, data: dataSAP(f.BillingDocumentDate), cli: f.SoldToParty,
          vend: ve ? ve.Personnel : 'SEM', anula: f.CancelledBillingDocument || '',
          itens: (f.to_Item?.results || []).map(i => ({ m: i.Material, q: +i.BillingQuantity, u: i.BillingQuantityUnit, v: +i.NetAmount }))
        });
        n++;
      }
      if (d.results.length < lote) break;
      skip += lote;
      if (skip % 1000 === 0) log(`... ${skip} documentos lidos`);
    }
    return n;
  }

  // ---------- 2) Litros por unidade, pelo cadastro do produto ----------
  const paraLitros = (vol, un) => un === 'L' ? vol : un === 'ML' ? vol / 1000 : un === 'M3' ? vol * 1000 : vol;
  async function carregarLitros() {
    const faltam = new Set();
    for (const d of docs.values()) for (const i of d.itens) if (!litrosUn.has(i.m + '|' + i.u)) faltam.add(i.m);
    for (const m of faltam) {
      try {
        const u = await sap(`${P}A_Product('${encodeURIComponent(m)}')/to_ProductUnitsOfMeasure`);
        for (const x of u.results) {
          litrosUn.set(m + '|' + x.AlternativeUnit, paraLitros(+x.MaterialVolume || 0, x.VolumeUnit) || 0);
        }
      } catch (e) { log('Sem conversão para litros do material', m, e.message); }
      for (const d of docs.values()) for (const i of d.itens) if (i.m === m && !litrosUn.has(m + '|' + i.u)) litrosUn.set(m + '|' + i.u, 0);
    }
  }

  // ---------- 3) Config (equipes, vendedores, metas) guardada no próprio Gist ----------
  async function lerConfig() {
    const r = await fetch('https://api.github.com/gists/' + CFG.GIST_ID, { headers: GH });
    if (!r.ok) throw new Error('GitHub respondeu ' + r.status + ' ao ler o Gist (ID ou token errado?)');
    const f = (await r.json()).files?.['config.json'];
    if (!f) return { equipes: [], vendedores: [] };
    const txt = f.truncated ? await (await fetch(f.raw_url)).text() : f.content;
    try { return JSON.parse(txt); } catch { log('config.json com erro de formato, ignorado'); return { equipes: [], vendedores: [] }; }
  }

  // ---------- 4) Monta o dados.json ----------
  function montar(cfg) {
    const anulados = new Set();
    for (const d of docs.values()) if (CFG.TIPOS_ESTORNO.includes(d.tipo) && d.anula) anulados.add(d.anula);
    const agg = new Map(); let semLitro = 0;
    for (const [id, d] of docs) {
      const sinal = CFG.TIPOS_VENDA.includes(d.tipo) ? 1 : CFG.TIPOS_DEVOLUCAO.includes(d.tipo) ? -1 : 0;
      if (!sinal || anulados.has(id)) continue;
      const k = d.data + '|' + d.vend + '|' + d.cli;
      const a = agg.get(k) || [d.data, d.vend, d.cli, 0, 0, 0];
      for (const i of d.itens) {
        const l = litrosUn.get(i.m + '|' + i.u) || 0; if (!l) semLitro++;
        a[3] += sinal * i.v; a[4] += sinal * i.q * l;
      }
      if (sinal > 0) a[5] += 1;
      agg.set(k, a);
    }
    const linhas = [...agg.values()].map(a => [a[0], a[1], a[2], Math.round(a[3] * 100) / 100, Math.round(a[4]), a[5]]);

    const cfgV = new Map((cfg.vendedores || []).map(v => [String(v.id), v]));
    const ids = new Set([...cfgV.keys(), ...linhas.map(l => l[1])]);
    const vendedores = [...ids].map(id => {
      const c = cfgV.get(id) || {};
      return { id, nome: c.nome || nomesVend[id] || (id === 'SEM' ? 'Sem vendedor na nota' : 'Vendedor ' + id),
        equipe: c.equipe || 'SEM', canal: c.canal || '', uf: c.uf || '', meta_mes: +c.meta_mes || 0 };
    });
    const equipes = [...(cfg.equipes || [])];
    if (vendedores.some(v => v.equipe === 'SEM') && !equipes.some(e => e.id === 'SEM')) equipes.push({ id: 'SEM', nome: 'Sem equipe' });
    const usados = new Set(linhas.map(l => l[2]));
    const clientes = Object.fromEntries([...usados].map(c => [c, nomesCli[c] || c]));
    return { dados: { gerado_em: new Date().toISOString(), equipes, vendedores, clientes, linhas }, semLitro };
  }

  async function gravar(dados) {
    const r = await fetch('https://api.github.com/gists/' + CFG.GIST_ID, {
      method: 'PATCH', headers: { ...GH, 'Content-Type': 'application/json' },
      body: JSON.stringify({ files: { 'dados.json': { content: JSON.stringify(dados) } } })
    });
    if (!r.ok) throw new Error('GitHub respondeu ' + r.status + ' ao gravar: ' + (await r.text()).slice(0, 200));
  }

  // ---------- Ciclo ----------
  async function ciclo() {
    if (rodando) return; rodando = true;
    const t0 = Date.now();
    try {
      docs.clear();                       // o dia é pequeno: relê tudo a cada ciclo
      const n = await lerFaturamentos(CFG.DIA);
      await carregarLitros();
      const cfg = TESTE ? { equipes: [], vendedores: [] } : await lerConfig();
      const { dados, semLitro } = montar(cfg);
      if (TESTE) {
        const nome = Object.fromEntries(dados.vendedores.map(v => [v.id, v.nome])), t = {};
        for (const [, v, , val, lit, nfs] of dados.linhas) { const x = t[v] ||= { Vendedor: nome[v], 'Valor líquido': 0, Litros: 0, Notas: 0 };
          x['Valor líquido'] += val; x.Litros += lit; x.Notas += nfs; }
        const tab = Object.values(t).sort((a, b) => b['Valor líquido'] - a['Valor líquido']).map(x => ({ ...x, 'Valor líquido': Math.round(x['Valor líquido'] * 100) / 100 }));
        console.table(tab);
        log(`TESTE (nada foi gravado): total R$ ${tab.reduce((s, x) => s + x['Valor líquido'], 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` +
          (semLitro ? ` · ${semLitro} itens sem conversão para litros` : ''));
        rodando = false; return;
      }
      await gravar(dados);
      ultimo = new Date(); erros = 0;
      const kb = Math.round(JSON.stringify(dados).length / 1024);
      log(`OK: ${n} documentos do dia ${CFG.DIA.split('-').reverse().join('/')}, ${dados.linhas.length} linhas, ${kb} KB, ${Math.round((Date.now() - t0) / 1000)} s` +
        (semLitro ? ` · ${semLitro} itens sem conversão para litros` : ''));
    } catch (e) {
      erros++;
      console.error('[coletor] FALHOU (' + erros + 'ª vez seguida):', e.message);
    } finally { rodando = false; }
  }

  window.__coletor = {
    parar() { clearInterval(timer); log('Coletor parado.'); },
    agora: ciclo,
    status() { log(`Último envio: ${ultimo ? ultimo.toLocaleTimeString('pt-BR') : 'nenhum ainda'} · documentos em memória: ${docs.size} · falhas seguidas: ${erros}`); },
    modelo() {
      const linhas = Object.entries(nomesVend).sort((a, b) => a[1].localeCompare(b[1], 'pt-BR'))
        .map(([id, n]) => [id, n, '', '', '', ''].join('\t'));
      const txt = ['ID_SAP\tNome\tEquipe\tCanal\tUF\tMeta_mes'].concat(linhas).join('\n');
      try { copy(txt); log(linhas.length + ' vendedores copiados. Cole no Excel, preencha e mande no chat.'); }
      catch { console.log(txt); log('Copie o texto acima, cole no Excel, preencha e mande no chat.'); }
    }
  };

  log(`Iniciando. Lendo só as notas de ${CFG.DIA.split('-').reverse().join('/')}.`);
  await ciclo();
  if (TESTE) return;
  timer = setInterval(ciclo, CFG.INTERVALO_MIN * 60000);
  log(`Rodando a cada ${CFG.INTERVALO_MIN} min. Deixe esta aba aberta.`);
})();
