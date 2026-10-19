/* ============================================================
   PAINEL DE PRODUÇÃO - REDE CAPS PETROLINA
   Motor (leitura + validação + cálculos) e renderização.
   PASSO 1: motor, Capa, Panorama.
   ============================================================ */
(function () {
  'use strict';

  const MONTHS = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO',
                  'AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
  const MSHORT = {JANEIRO:'Jan',FEVEREIRO:'Fev','MARÇO':'Mar',ABRIL:'Abr',MAIO:'Mai',
                  JUNHO:'Jun',JULHO:'Jul',AGOSTO:'Ago',SETEMBRO:'Set',OUTUBRO:'Out',
                  NOVEMBRO:'Nov',DEZEMBRO:'Dez'};
  const UNIT_NAMES = {
    'CAPSIJ':'CAPS Infanto-Juvenil', 'CAPS II':'CAPS II', 'CAPS AD III':'CAPS AD III',
    'UAA':'Unidade de Acolhimento Adulto', 'UAI':'Unidade de Acolhimento Infantojuvenil'
  };

  // paleta para séries
  const PAL = ['#123a8f','#d23a34','#efb93c','#2f9e5f','#4f73c4','#9b59b6','#e08a1e'];

  // ---------- util ----------
  const $ = (s, r) => (r||document).querySelector(s);
  const el = (t, cls, html) => { const e=document.createElement(t); if(cls)e.className=cls; if(html!=null)e.innerHTML=html; return e; };
  const fmt = n => (n==null?'—':Math.round(n).toLocaleString('pt-BR'));
  const pct = n => (n==null?'—':n.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'%');
  const norm = s => (s==null?'':String(s).trim());
  const isNum = v => { if(v==null||v==='')return false; if(typeof v==='number')return true; return !isNaN(parseFloat(String(v).replace(',','.'))); };
  const toNum = v => isNum(v)?parseFloat(String(v).replace(',','.')):null;
  // hex (#rrggbb) -> rgba(r,g,b,alpha); aceita também já-rgba (retorna como veio)
  function hexA(hex, a) {
    hex = (hex||'').trim();
    const m = hex.match(/^#?([0-9a-f]{6})$/i);
    if (!m) return hex;
    const n = parseInt(m[1],16);
    return `rgba(${(n>>16)&255},${(n>>8)&255},${n&255},${a})`;
  }

  // ========================================================
  //  TEMA (claro / escuro)
  // ========================================================
  const THEME_KEY = 'caps-tema';
  function getTheme() {
    const saved = localStorage.getItem(THEME_KEY);
    if (saved === 'dark' || saved === 'light') return saved;
    return (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
  }
  function applyTheme(t) {
    document.documentElement.setAttribute('data-theme', t);
    const btn = $('#tema-btn');
    if (btn) { btn.innerHTML = t === 'dark' ? ICON.sun : ICON.moon; btn.title = t === 'dark' ? 'Modo claro' : 'Modo escuro'; }
  }
  function toggleTheme() {
    const next = (document.documentElement.getAttribute('data-theme') === 'dark') ? 'light' : 'dark';
    localStorage.setItem(THEME_KEY, next);
    applyTheme(next);
    renderAllCharts(); // redesenha gráficos com as cores do novo tema
  }
  // cores dos gráficos conforme o tema atual (lidas das CSS vars)
  function chartTheme() {
    const cs = getComputedStyle(document.documentElement);
    const v = n => cs.getPropertyValue(n).trim();
    return {
      azul: v('--azul') || '#123a8f',
      amarelo: v('--amarelo') || '#efb93c',
      axis: v('--chart-axis') || '#45586b',
      grid: v('--grid-line') || '#eef2f7',
      line: v('--linha-forte') || '#d0d8e4',
      azul300: v('--azul-300') || '#4f73c4',
      surface: v('--branco') || '#ffffff'
    };
  }
  // aplica tema o quanto antes (evita flash). Chamado no load.
  applyTheme(getTheme());

  // ========================================================
  //  MOTOR: lê a planilha e extrai só as linhas de procedimento
  // ========================================================
  function parseWorkbook(wb) {
    const result = { units: [], warnings: [], errors: [] };

    wb.SheetNames.forEach(sheetName => {
      const ws = wb.Sheets[sheetName];
      const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

      // acha a linha de cabeçalho (começa com PROCEDIMENTO)
      let hr = -1;
      for (let r = 0; r < Math.min(grid.length, 15); r++) {
        const row = grid[r] || [];
        for (let c = 0; c < row.length; c++) {
          if (norm(row[c]).toUpperCase().startsWith('PROCEDIMENTO')) { hr = r; break; }
        }
        if (hr >= 0) break;
      }
      if (hr < 0) {
        result.warnings.push(`Aba "${sheetName}" ignorada: não encontrei a linha de cabeçalho (PROCEDIMENTO).`);
        return;
      }

      // mapeia colunas de mês e TAB pelos NOMES do cabeçalho
      const header = grid[hr] || [];
      const monthCol = {}, tabCol = {};
      header.forEach((h, c) => {
        const H = norm(h).toUpperCase();
        if (MONTHS.includes(H)) monthCol[H] = c;
        else if (H.includes('/TAB')) {
          const abbr = H.split('/')[0];
          const full = MONTHS.find(m => m === abbr || m.startsWith(abbr) || abbr.startsWith(m.slice(0,3)));
          if (full) tabCol[full] = c;
        }
      });

      // lê linhas de procedimento até "Total procedimentos"
      const procedures = [];
      for (let r = hr + 1; r < grid.length; r++) {
        const row = grid[r] || [];
        const label = norm(row[0]);
        if (label === '') continue;
        const up = label.toUpperCase();
        if (up.startsWith('TOTAL PROCEDIMENTOS') || up.startsWith('TOTAL DE PROCEDIMENTOS')) break;
        const m = label.match(/^(\d{2,4})\s*-?\s*(.*)$/);
        if (!m) continue; // não é procedimento (sem código) -> ignora
        const code = m[1], desc = m[2].trim();
        const realizado = {}, tab = {};
        MONTHS.forEach(mo => {
          if (mo in monthCol) {
            const raw = row[monthCol[mo]];
            if (isNum(raw)) realizado[mo] = toNum(raw);
            else if (raw != null && norm(raw) !== '')
              result.warnings.push(`${sheetName} · ${code} · ${MSHORT[mo]}: valor não numérico "${raw}" (ignorado).`);
          }
          if (mo in tabCol) {
            const raw = row[tabCol[mo]];
            if (isNum(raw)) tab[mo] = toNum(raw);
            else if (raw != null && norm(raw) !== '')
              result.warnings.push(`${sheetName} · ${code} · ${MSHORT[mo]}/TAB: valor não numérico "${raw}" (ignorado).`);
          }
        });
        procedures.push({ code, desc, realizado, tab });
      }

      if (procedures.length === 0) {
        result.warnings.push(`Aba "${sheetName}" não tem linhas de procedimento reconhecíveis.`);
        return;
      }

      const hasTab = procedures.some(p => Object.keys(p.tab).length > 0);
      const monthsPresent = MONTHS.filter(m => procedures.some(p => (p.realizado[m]||0) > 0));
      const tabMonths = MONTHS.filter(m => procedures.some(p => (p.tab[m]||0) > 0));

      result.units.push({
        sheet: sheetName,
        name: UNIT_NAMES[sheetName] || sheetName,
        hasTab, monthsPresent, tabMonths, procedures,
        isNew: !(sheetName in UNIT_NAMES)
      });
    });

    if (result.units.length === 0)
      result.errors.push('Nenhuma unidade válida encontrada na planilha.');

    return result;
  }

  // ========================================================
  //  CÁLCULOS derivados (taxas, agregados)
  // ========================================================
  function computeMetrics(data) {
    data.network = { months: [], tabMonths: [], totalReal: 0, grandByUnit: [] };
    // meses presentes na rede
    MONTHS.forEach(m => {
      if (data.units.some(u => u.monthsPresent.includes(m))) data.network.months.push(m);
      if (data.units.some(u => u.tabMonths.includes(m))) data.network.tabMonths.push(m);
    });

    data.units.forEach(u => {
      u.totalReal = u.procedures.reduce((s,p)=> s + Object.values(p.realizado).reduce((a,b)=>a+b,0), 0);
      u.totalTab  = u.procedures.reduce((s,p)=> s + Object.values(p.tab).reduce((a,b)=>a+b,0), 0);

      // taxa calculada SÓ nos meses em que há TAB (comparação justa)
      const tm = u.tabMonths;
      const realInTab = u.procedures.reduce((s,p)=> s + tm.reduce((a,m)=> a+(p.realizado[m]||0),0), 0);
      const tabInTab  = u.procedures.reduce((s,p)=> s + tm.reduce((a,m)=> a+(p.tab[m]||0),0), 0);
      u.taxaGeral = realInTab>0 ? 100*tabInTab/realInTab : null;

      // taxa dos TABULÁVEIS: só procedimentos com algum tab>0 no ano
      const tabuláveis = u.procedures.filter(p => Object.values(p.tab).some(v=>v>0));
      const realTB = tabuláveis.reduce((s,p)=> s + tm.reduce((a,m)=>a+(p.realizado[m]||0),0),0);
      const tabTB  = tabuláveis.reduce((s,p)=> s + tm.reduce((a,m)=>a+(p.tab[m]||0),0),0);
      u.taxaTabulaveis = realTB>0 ? 100*tabTB/realTB : null;
      u.nTabulaveis = tabuláveis.length;
    });

    data.network.totalReal = data.units.reduce((s,u)=> s+u.totalReal, 0);
    data.network.nUnits = data.units.length;

    // ---- mix de procedimentos agregado na rede (por descrição canônica) ----
    const mix = {};
    data.units.forEach(u => u.procedures.forEach(p => {
      const key = p.code + ' - ' + shortDesc(p.desc);
      const r = sumObj(p.realizado);
      if (!mix[key]) mix[key] = { code:p.code, desc:shortDesc(p.desc), real:0 };
      mix[key].real += r;
    }));
    data.network.mix = Object.values(mix).filter(m=>m.real>0).sort((a,b)=>b.real-a.real);

    // ---- por unidade CAPS: perdas de tabulação por procedimento (jan-jul) ----
    data.units.forEach(u => {
      if (!u.hasTab) return;
      const tm = u.tabMonths;
      u.procLosses = u.procedures.map(p => {
        const real = tm.reduce((a,m)=>a+(p.realizado[m]||0),0);
        const tab  = tm.reduce((a,m)=>a+(p.tab[m]||0),0);
        return {
          code:p.code, desc:shortDesc(p.desc), real, tab,
          gap: real - tab,
          taxa: real>0 ? 100*tab/real : null,
          nuncaTab: tab===0 && real>0
        };
      }).filter(x=>x.real>0);
      // classificação: "não tabulado" = nunca teve tab no período
      u.naoTabulados = u.procLosses.filter(x=>x.nuncaTab).sort((a,b)=>b.real-a.real);
      u.maioresPerdas = u.procLosses.filter(x=>!x.nuncaTab && x.gap>0).sort((a,b)=>b.gap-a.gap);
    });

    // ---- alertas automáticos (data-driven) ----
    data.alerts = [];
    data.units.forEach(u => {
      if (u.hasTab && u.taxaGeral!=null && u.taxaGeral < 20)
        data.alerts.push({ nivel:'crit', unidade:u.name,
          texto:`Taxa de tabulação crítica (${pct(u.taxaGeral)}): grande parte da produção não chega ao sistema.` });
      // unidade com volume muito baixo (possível subutilização/subnotificação)
      if (u.totalReal>0 && u.totalReal < 100 && data.network.totalReal>0)
        data.alerts.push({ nivel:'warn', unidade:u.name,
          texto:`Volume de produção muito baixo (${fmt(u.totalReal)} no período). Verificar subutilização ou subnotificação.` });
    });
    // TAB > Realizado: contagem + lista detalhada dos registros
    const tabGtList = [];
    data.units.forEach(u => u.procedures.forEach(p => MONTHS.forEach(m => {
      const r=p.realizado[m], t=p.tab[m];
      if (r!=null && t!=null && t>r) tabGtList.push({ unidade:u.name, code:p.code, desc:shortDesc(p.desc), mes:m, real:r, tab:t });
    })));
    data.network.tabGtCount = tabGtList.length;
    data.network.tabGtList = tabGtList;
  }

  function sumObj(o){ return Object.values(o).reduce((a,b)=>a+b,0); }
  // encurta a descrição do procedimento para rótulos
  function shortDesc(d){
    return d.replace(/\s+EM CENTRO DE ATENÇÃO PSICOSSOCIAL.*$/i,'')
            .replace(/\s+DE PACIENTE.*$/i,'')
            .replace(/\s+PARA PACIENTES.*$/i,'')
            .replace(/\s+POR CENTRO DE ATENÇÃO PSICOSSOCIAL.*$/i,'')
            .trim();
  }

  // ========================================================
  //  RENDER: tela de validação
  // ========================================================
  function renderValidation(data) {
    const box = $('#validacao');
    box.innerHTML = '';
    box.classList.add('show');

    const unitNames = data.units.map(u=>u.name).join(', ');
    const monthsLbl = data.network.months.map(m=>MSHORT[m]).join(', ');
    addVal(box,'ok', `${data.units.length} unidades reconhecidas: ${unitNames}`);
    addVal(box,'ok', `Período com dados: ${monthsLbl}`);
    const capsTab = data.units.filter(u=>u.hasTab).length;
    addVal(box,'ok', `Dados de tabulação (TAB) em ${capsTab} unidade(s), nos meses: ${data.network.tabMonths.map(m=>MSHORT[m]).join(', ')}`);

    data.units.filter(u=>u.isNew).forEach(u =>
      addVal(box,'warn', `Unidade nova detectada: "${u.sheet}" (incluída automaticamente).`));

    data.warnings.forEach(w => addVal(box,'warn', w));
    data.errors.forEach(e => addVal(box,'err', e));

    const btn = el('button','btn-ver','Ver apresentação');
    btn.disabled = data.errors.length > 0;
    btn.onclick = () => showApp(data);
    box.appendChild(btn);
  }
  function addVal(box, kind, text) {
    const icon = kind==='ok'?ICON.check : kind==='warn'?ICON.alert : ICON.x;
    const item = el('div', `val-item val-${kind}`);
    item.innerHTML = `<span class="vi">${icon}</span><span>${text}</span>`;
    box.appendChild(item);
  }

  // ========================================================
  //  RENDER: aplicação (Passo 1 = Capa + Panorama)
  // ========================================================
  function showApp(data) {
    $('#tela-upload').style.display = 'none';
    const app = $('#tela-app');
    app.classList.add('show');

    const periodo = data.network.months.length
      ? `${MSHORT[data.network.months[0]]} a ${MSHORT[data.network.months[data.network.months.length-1]]}/2026` : '2026';

    app.innerHTML = `
      <div class="topbar"><div class="wrap">
        <img src="${LOGO.brasao}" alt="Prefeitura de Petrolina">
        <div class="tb-sep"></div>
        <img src="${LOGO.saude}" alt="Secretaria de Saúde">
        <div class="tb-title">Produção da Rede CAPS</div>
        <div class="tb-period">${periodo}</div>
        <button class="tema-btn" id="tema-btn" type="button" aria-label="Alternar tema"></button>
      </div></div>
      <nav class="nav-mobile" id="nav-mobile"></nav>
      <div class="wrap"><div class="layout">
        <nav class="indice" id="indice"></nav>
        <main id="main"></main>
      </div></div>
      <footer class="rodape">Painel gerado localmente a partir da planilha de produção. Nenhum dado é enviado pela internet.</footer>
    `;

    const main = $('#main', app);
    main.appendChild(secCapa(data, periodo));
    main.appendChild(secPanorama(data));
    main.appendChild(secProducao(data));
    main.appendChild(secTabulacao(data));
    main.appendChild(secQualidade(data));
    main.appendChild(secInsights(data));
    buildIndice([
      ['capa','Capa', ICON.home],
      ['panorama','Panorama da rede', ICON.grid],
      ['producao','Produção', ICON.activity],
      ['tabulacao','Realizado x Tabulado', ICON.percent],
      ['qualidade','Qualidade dos dados', ICON.shield],
      ['insights','Insights', ICON.bulb],
    ]);
    setupScrollSpy();
    applyTheme(getTheme());               // garante ícone correto no botão
    const tbtn = $('#tema-btn');
    if (tbtn) tbtn.onclick = toggleTheme;
  }

  // destaca item do índice conforme a rolagem
  function setupScrollSpy() {
    const links = Array.from(document.querySelectorAll('.indice a'));
    const mlinks = Array.from(document.querySelectorAll('.nav-mobile a'));
    const secs = links.map(a => document.querySelector(a.getAttribute('href')));
    const onScroll = () => {
      let idx = 0;
      secs.forEach((s,i)=>{ if (s && s.getBoundingClientRect().top < 120) idx = i; });
      links.forEach((a,i)=> a.classList.toggle('active', i===idx));
      mlinks.forEach((a,i)=> {
        const on = i===idx; a.classList.toggle('active', on);
        if (on && a.scrollIntoView) a.scrollIntoView({ inline:'center', block:'nearest', behavior:'smooth' });
      });
    };
    window.addEventListener('scroll', onScroll, { passive:true });
    onScroll();
  }

  function buildIndice(items) {
    const nav = $('#indice');
    nav.innerHTML = `<div class="indice-head">Seções</div>`;
    const mob = $('#nav-mobile');
    if (mob) mob.innerHTML = '';
    items.forEach(([id,label,icon]) => {
      const a = el('a'); a.href = '#'+id;
      a.innerHTML = `<span class="ii">${icon}</span><span class="il">${label}</span>`;
      nav.appendChild(a);
      if (mob) {
        const m = el('a'); m.href = '#'+id; m.textContent = label;
        mob.appendChild(m);
      }
    });
  }

  // ---- Seção: Capa ----
  function secCapa(data, periodo) {
    const s = el('section','bloco'); s.id='capa';
    s.innerHTML = `
      <div class="capa">
        <div class="capa-eyebrow">Secretaria de Saúde de Petrolina</div>
        <h2>Produção da Rede de Atenção<br>Psicossocial de Petrolina</h2>
        <p class="capa-sub">Análise da produção de procedimentos das unidades CAPS e de acolhimento, com foco na relação entre o que é realizado e o que é efetivamente tabulado no sistema.</p>
        <div class="capa-meta">
          <div><div class="l">Período</div><div class="v">${periodo}</div></div>
          <div><div class="l">Unidades</div><div class="v num">${data.network.nUnits}</div></div>
          <div><div class="l">Procedimentos realizados</div><div class="v num">${fmt(data.network.totalReal)}</div></div>
        </div>
      </div>`;
    return s;
  }

  // ---- Seção: Panorama ----
  function secPanorama(data) {
    const s = el('section','bloco'); s.id='panorama';
    const totalReal = data.network.totalReal;
    const capsUnits = data.units.filter(u=>u.hasTab);
    const totReal7 = capsUnits.reduce((acc,u)=> acc + u.procedures.reduce((s,p)=> s+u.tabMonths.reduce((a,m)=>a+(p.realizado[m]||0),0),0),0);
    const totTab7  = capsUnits.reduce((acc,u)=> acc + u.procedures.reduce((s,p)=> s+u.tabMonths.reduce((a,m)=>a+(p.tab[m]||0),0),0),0);
    const taxaRede = totReal7>0 ? 100*totTab7/totReal7 : null;
    const topUnit = [...data.units].sort((a,b)=>b.totalReal-a.totalReal)[0];

    s.innerHTML = `
      <div class="bloco-head">
        <div class="bloco-kicker">Visão geral</div>
        <div class="bloco-title">Panorama da rede</div>
        <p class="bloco-lede">A rede realizou <strong>${fmt(totalReal)} procedimentos</strong> no período. A unidade de maior volume é o <strong>${topUnit.name}</strong>. Entre os meses com dado de tabulação, apenas <strong>${pct(taxaRede)}</strong> da produção das unidades CAPS chegou ao sistema.</p>
      </div>
      <div class="kpis">
        ${kpi(ICON.activity, fmt(totalReal), 'Procedimentos realizados', 'Soma de todas as unidades',
          'Total de procedimentos efetivamente realizados pelas equipes no período, somando todas as unidades. Considera apenas o que foi feito (não inclui a tabulação).')}
        ${kpi(ICON.building, data.network.nUnits, 'Unidades na rede', `${capsUnits.length} CAPS com tabulação`,
          'Quantidade de unidades presentes na planilha. Dessas, as unidades CAPS possuem dados de tabulação; as unidades de acolhimento ainda não.')}
        ${kpi(ICON.percent, pct(taxaRede), 'Taxa de tabulação (CAPS)', 'Calculada só em meses com TAB',
          'Dos procedimentos realizados nos meses que têm tabulação, qual percentual chegou ao sistema. Calculada só no período em que Realizado e Tabulado coexistem, para ser justa.')}
        ${kpi(ICON.calendar, data.network.months.length, 'Meses com produção', `${data.network.tabMonths.length} meses com TAB`,
          'Número de meses com produção registrada na planilha. Ao lado, quantos desses meses já possuem dados de tabulação.')}
      </div>
      <div class="grid-2" style="margin-top:16px">
        <div class="card"><div class="card-title">Participação por unidade</div><div class="card-sub">Procedimentos realizados no período (todas as competências)</div><div class="chart" id="ch-participacao"></div></div>
        <div class="card"><div class="card-title">Produção mensal da rede</div><div class="card-sub">Total de procedimentos realizados por mês</div><div class="chart" id="ch-mensal"></div></div>
      </div>`;

    // desenha gráficos depois do DOM existir
    setTimeout(()=>{ chartParticipacao(data); chartMensal(data); }, 0);
    return s;
  }

  // monta um cartão KPI com ícone de informação (tooltip ao passar o mouse / tocar)
  function kpi(icon, val, lab, hint, info) {
    return `<div class="kpi">
      <div class="kpi-top">
        <div class="kpi-ico">${icon}</div>
        <button class="kpi-info" type="button" aria-label="O que é isto?" data-tip="${esc(info)}">${ICON.info}</button>
      </div>
      <div class="kpi-val num">${val}</div>
      <div class="kpi-lab">${lab}</div>
      <div class="kpi-hint">${hint}</div>
    </div>`;
  }
  function esc(s){ return String(s).replace(/"/g,'&quot;'); }

  // ---- gráficos do Panorama ----
  function chartParticipacao(data) {
    const dom = $('#ch-participacao'); if(!dom) return;
    const rows = [...data.units].sort((a,b)=>a.totalReal-b.totalReal);
    regChart(dom, () => { const c = chartTheme(); return {
      grid:{left:8,right:70,top:10,bottom:10,containLabel:true},
      tooltip:{trigger:'axis',axisPointer:{type:'shadow'},valueFormatter:v=>fmt(v)+' proc.'},
      xAxis:{type:'value',show:false,max:v=>v.max*1.12},
      yAxis:{type:'category',data:rows.map(u=>u.name),axisLabel:{fontSize:12,color:c.axis},axisTick:{show:false},axisLine:{lineStyle:{color:c.line}}},
      series:[{type:'bar',data:rows.map(u=>({value:u.totalReal,itemStyle:{color:c.azul,borderRadius:[0,5,5,0]}})),barWidth:'55%',
        label:{show:true,position:'right',formatter:p=>fmt(p.value),color:c.axis,fontSize:11.5}}]
    };});
  }
  function chartMensal(data) {
    const dom = $('#ch-mensal'); if(!dom) return;
    const months = data.network.months;
    const vals = months.map(m => data.units.reduce((s,u)=> s + u.procedures.reduce((a,p)=>a+(p.realizado[m]||0),0),0));
    regChart(dom, () => { const c = chartTheme(); return {
      grid:{left:8,right:16,top:16,bottom:10,containLabel:true},
      tooltip:{trigger:'axis',valueFormatter:v=>fmt(v)+' proc.'},
      xAxis:{type:'category',data:months.map(m=>MSHORT[m]),axisLine:{lineStyle:{color:c.line}},axisLabel:{color:c.axis}},
      yAxis:{type:'value',axisLabel:{formatter:v=>fmt(v),color:c.axis},splitLine:{lineStyle:{color:c.grid}}},
      series:[{type:'line',data:vals,smooth:true,symbol:'circle',symbolSize:7,
        lineStyle:{color:c.azul,width:3},itemStyle:{color:c.azul},
        areaStyle:{color:new echarts.graphic.LinearGradient(0,0,0,1,[{offset:0,color:hexA(c.azul,.18)},{offset:1,color:hexA(c.azul,0)}])}}]
    };});
  }

  // ========================================================
  //  Seção: PRODUÇÃO
  // ========================================================
  function secProducao(data) {
    const s = el('section','bloco'); s.id='producao';
    const mixTop = data.network.mix.slice(0,8);
    const top2 = data.network.mix.slice(0,2);
    const top2share = data.network.totalReal>0 ? 100*(top2.reduce((a,m)=>a+m.real,0))/data.network.totalReal : 0;

    s.innerHTML = `
      <div class="bloco-head">
        <div class="bloco-kicker">O que a rede produz</div>
        <div class="bloco-title">Produção por unidade e por tipo</div>
        <p class="bloco-lede">Dois procedimentos concentram <strong>${pct(top2share)}</strong> de toda a produção da rede: <strong>${top2.map(m=>m.desc.toLowerCase()).join('</strong> e <strong>')}</strong>. Use o seletor para alternar entre produção realizada e tabulada.</p>
        <div class="seg" id="seg-prod">
          <button class="seg-op active" data-modo="realizado" type="button">Realizado</button>
          <button class="seg-op" data-modo="tabulado" type="button">Tabulado</button>
        </div>
      </div>
      <div class="grid-2">
        <div class="card"><div class="card-title">Composição da produção por unidade</div><div class="card-sub">Participação de cada tipo de procedimento dentro da unidade</div><div class="chart" id="ch-composicao"></div></div>
        <div class="card"><div class="card-title">Principais procedimentos da rede</div><div class="card-sub">Top ${mixTop.length} por volume</div><div class="chart" id="ch-mix"></div></div>
      </div>
      <div class="card" style="margin-top:16px">
        <div class="card-title">Evolução mensal por unidade</div>
        <div class="card-sub">Produção por mês. Passe o mouse para ver os valores.</div>
        <div class="chart" id="ch-evolucao" style="height:360px"></div>
      </div>`;

    // estado do toggle (realizado | tabulado) para esta seção
    const estado = { modo: 'realizado' };
    setTimeout(()=>{
      chartComposicao(data, estado); chartMix(data, estado); chartEvolucao(data, estado);
      const seg = $('#seg-prod', s);
      if (seg) seg.querySelectorAll('.seg-op').forEach(b => b.onclick = () => {
        estado.modo = b.dataset.modo;
        seg.querySelectorAll('.seg-op').forEach(x=>x.classList.toggle('active', x===b));
        chartComposicao(data, estado); chartMix(data, estado); chartEvolucao(data, estado);
      });
    },0);
    return s;
  }

  // valor de um procedimento no modo atual (realizado = ano todo; tabulado = meses com TAB)
  function valModo(unit, p, modo) {
    if (modo === 'tabulado') return unit.tabMonths.reduce((a,m)=>a+(p.tab[m]||0),0);
    return Object.values(p.realizado).reduce((a,b)=>a+b,0);
  }

  // instâncias dedicadas desta seção (recriadas ao trocar modo)
  const prodCharts = {};
  function makeProdChart(id, buildOption) {
    const dom = $('#'+id); if(!dom) return;
    if (prodCharts[id]) { prodCharts[id].setOption(buildOption(), true); return; }
    const ch = echarts.init(dom);
    ch.setOption(buildOption());
    prodCharts[id] = ch;
    REG.push({ ch, buildOption });
  }

  // Composição por unidade (barras empilhadas por tipo de procedimento)
  function chartComposicao(data, estado) {
    // top 5 tipos (por realizado na rede) + "Outros"
    const topKeys = data.network.mix.slice(0,5).map(m=>m.code);
    const units = [...data.units].sort((a,b)=>
      b.procedures.reduce((s,p)=>s+valModo(b,p,estado.modo),0) - a.procedures.reduce((s,p)=>s+valModo(a,p,estado.modo),0));
    const tipos = data.network.mix.slice(0,5).map(m=>m.desc);
    makeProdChart('ch-composicao', () => { const c = chartTheme();
      const series = topKeys.map((code,i)=>({
        name: tipos[i], type:'bar', stack:'t', barWidth:'58%',
        itemStyle:{ color: PAL[i % PAL.length] },
        emphasis:{ focus:'series' },
        data: units.map(u => { const p = u.procedures.find(x=>x.code===code); return p?valModo(u,p,estado.modo):0; })
      }));
      // "Outros"
      series.push({ name:'Outros', type:'bar', stack:'t', barWidth:'58%',
        itemStyle:{ color: c.line },
        data: units.map(u => u.procedures.filter(p=>!topKeys.includes(p.code)).reduce((a,p)=>a+valModo(u,p,estado.modo),0)) });
      return {
        color:PAL,
        grid:{left:8,right:12,top:52,bottom:8,containLabel:true},
        tooltip:{trigger:'axis',axisPointer:{type:'shadow'},valueFormatter:v=>fmt(v)+' proc.'},
        legend:{top:0,type:'scroll',textStyle:{color:c.axis,fontSize:10.5},itemWidth:12,itemHeight:8},
        xAxis:{type:'value',show:false},
        yAxis:{type:'category',data:units.map(u=>u.name),axisLabel:{color:c.axis,fontSize:11},axisTick:{show:false},axisLine:{lineStyle:{color:c.line}}},
        series
      };
    });
  }

  function chartMix(data, estado) {
    // recalcula o top por modo
    const agg = {};
    data.units.forEach(u => u.procedures.forEach(p => {
      const key = p.code+' '+shortDesc(p.desc);
      agg[key] = agg[key] || { desc: shortDesc(p.desc), v:0 };
      agg[key].v += valModo(u, p, estado.modo);
    }));
    const rows = Object.values(agg).filter(x=>x.v>0).sort((a,b)=>a.v-b.v).slice(-8);
    makeProdChart('ch-mix', () => { const c = chartTheme(); return {
      grid:{left:8,right:72,top:8,bottom:8,containLabel:true},
      tooltip:{trigger:'axis',axisPointer:{type:'shadow'},valueFormatter:v=>fmt(v)+' proc.'},
      xAxis:{type:'value',show:false,max:v=>v.max*1.12},
      yAxis:{type:'category',data:rows.map(m=>m.desc),axisLabel:{color:c.axis,fontSize:11,width:150,overflow:'truncate'},axisTick:{show:false},axisLine:{lineStyle:{color:c.line}}},
      series:[{type:'bar',barWidth:'60%',data:rows.map((m,i)=>({value:m.v,itemStyle:{color:i===rows.length-1?c.azul:c.azul300,borderRadius:[0,5,5,0]}})),
        label:{show:true,position:'right',formatter:p=>fmt(p.value),color:c.axis,fontSize:11}}]
    };});
  }

  function chartEvolucao(data, estado) {
    const months = estado.modo==='tabulado' ? data.network.tabMonths : data.network.months;
    makeProdChart('ch-evolucao', () => { const c = chartTheme();
      const series=data.units.map((u,i)=>({
        name:u.name, type:'line', smooth:true, symbol:'circle', symbolSize:6,
        lineStyle:{width:2.4}, itemStyle:{color:PAL[i%PAL.length]},
        emphasis:{ focus:'series' },
        data:months.map(m=> u.procedures.reduce((a,p)=>a+(estado.modo==='tabulado'?(p.tab[m]||0):(p.realizado[m]||0)),0))
      }));
      return {
        color:PAL,
        grid:{left:8,right:16,top:40,bottom:8,containLabel:true},
        tooltip:{trigger:'axis',valueFormatter:v=>fmt(v)+' proc.'},
        legend:{top:0,textStyle:{color:c.axis,fontSize:11.5},itemWidth:14,itemHeight:8},
        xAxis:{type:'category',data:months.map(m=>MSHORT[m]),axisLine:{lineStyle:{color:c.line}},axisLabel:{color:c.axis}},
        yAxis:{type:'value',axisLabel:{formatter:v=>fmt(v),color:c.axis},splitLine:{lineStyle:{color:c.grid}}},
        series
      };
    });
  }

  // ========================================================
  //  Seção: REALIZADO x TABULADO (tese central)
  // ========================================================
  function secTabulacao(data) {
    const s = el('section','bloco'); s.id='tabulacao';
    const caps = data.units.filter(u=>u.hasTab);
    const semTab = data.units.filter(u=>!u.hasTab);

    // cartões de semáforo por unidade
    const cards = caps.map(u => {
      const nivel = u.taxaGeral==null?'' : u.taxaGeral<20?'crit' : u.taxaGeral<50?'warn':'ok';
      const cor = nivel==='crit'?'var(--vermelho)':nivel==='warn'?'var(--ambar)':'var(--verde)';
      const badge = nivel==='crit'?'badge-crit':nivel==='warn'?'badge-warn':'badge-ok';
      const rotulo = nivel==='crit'?'Crítico':nivel==='warn'?'Atenção':'Adequado';
      return `<div class="card sem-card">
        <div class="sem-top"><div class="sem-nome">${u.name}</div><span class="badge ${badge}">${rotulo}</span></div>
        <div class="sem-val" style="color:${cor}">${pct(u.taxaGeral)}</div>
        <div class="sem-sub">taxa geral &middot; tabuláveis: <strong>${pct(u.taxaTabulaveis)}</strong></div>
        <div class="sem-bar"><span style="width:${Math.min(100,u.taxaGeral||0)}%;background:${cor}"></span></div>
      </div>`;
    }).join('');

    // unidade de pior taxa, para a frase de leitura
    const pior = [...caps].filter(u=>u.taxaGeral!=null).sort((a,b)=>a.taxaGeral-b.taxaGeral)[0];

    s.innerHTML = `
      <div class="bloco-head">
        <div class="bloco-kicker">A questão central</div>
        <div class="bloco-title">Realizado x Tabulado</div>
        <p class="bloco-lede">Nem tudo o que a equipe produz chega ao sistema. Comparando apenas os meses que têm tabulação, o <strong>${pior?pior.name:'—'}</strong> apresenta a menor taxa (<strong>${pior?pct(pior.taxaGeral):'—'}</strong>). As taxas abaixo são calculadas somente no período em que Realizado e Tabulado coexistem.</p>
      </div>
      <div class="sem-grid">${cards}</div>
      <div class="card" style="margin-top:16px">
        <div class="card-title">Realizado x Tabulado por unidade</div>
        <div class="card-sub">Soma dos meses com tabulação disponível</div>
        <div class="chart" id="ch-realxtab"></div>
      </div>
      <div class="card" style="margin-top:16px">
        <div class="card-title">Onde a produção não chega ao sistema</div>
        <div class="card-sub">Procedimentos com produção registrada mas sem tabulação, ou com maior diferença. Dois grupos: os que podem não ser tabuláveis pela natureza, e os que deveriam ser tabulados.</div>
        <div id="perdas-wrap"></div>
      </div>
    `;
    setTimeout(()=>{ chartRealxTab(data, caps); renderPerdas(data, caps); },0);
    return s;
  }

  function chartRealxTab(data, caps) {
    const dom=$('#ch-realxtab'); if(!dom)return;
    const real = caps.map(u => u.procedures.reduce((a,p)=> a+u.tabMonths.reduce((x,m)=>x+(p.realizado[m]||0),0),0));
    const tab  = caps.map(u => u.procedures.reduce((a,p)=> a+u.tabMonths.reduce((x,m)=>x+(p.tab[m]||0),0),0));
    regChart(dom, () => { const c = chartTheme(); return {
      grid:{left:8,right:16,top:40,bottom:8,containLabel:true},
      tooltip:{trigger:'axis',axisPointer:{type:'shadow'},valueFormatter:v=>fmt(v)+' proc.'},
      legend:{top:0,textStyle:{color:c.axis},itemWidth:14,itemHeight:10},
      xAxis:{type:'category',data:caps.map(u=>u.name),axisLine:{lineStyle:{color:c.line}},axisLabel:{color:c.axis,interval:0}},
      yAxis:{type:'value',axisLabel:{formatter:v=>fmt(v),color:c.axis},splitLine:{lineStyle:{color:c.grid}}},
      series:[
        {name:'Realizado',type:'bar',data:real,barWidth:'32%',itemStyle:{color:c.azul,borderRadius:[4,4,0,0]}},
        {name:'Tabulado', type:'bar',data:tab, barWidth:'32%',itemStyle:{color:c.amarelo,borderRadius:[4,4,0,0]}}
      ]
    };});
  }

  function renderPerdas(data, caps) {
    const wrap=$('#perdas-wrap'); if(!wrap)return;
    caps.forEach(u => {
      const naoTab = u.naoTabulados.slice(0,5);
      const perdas = u.maioresPerdas.slice(0,4);
      if (!naoTab.length && !perdas.length) return;
      const block = el('div','perda-unidade');
      let html = `<div class="perda-titulo">${u.name}</div>`;
      if (naoTab.length) {
        html += `<div class="perda-grupo"><span class="badge badge-crit">Sem tabulação no período</span><ul>`;
        naoTab.forEach(p => html += `<li><span>${p.code} ${p.desc}</span><b class="num">${fmt(p.real)} realiz. / 0 tab.</b></li>`);
        html += `</ul></div>`;
      }
      if (perdas.length) {
        html += `<div class="perda-grupo"><span class="badge badge-warn">Maior diferença Realizado - Tabulado</span><ul>`;
        perdas.forEach(p => html += `<li><span>${p.code} ${p.desc}</span><b class="num">${fmt(p.real)} - ${fmt(p.tab)} = ${fmt(p.gap)} (${pct(p.taxa)})</b></li>`);
        html += `</ul></div>`;
      }
      block.innerHTML = html;
      wrap.appendChild(block);
    });
  }

  // ========================================================
  //  Seção: QUALIDADE DOS DADOS
  // ========================================================
  function secQualidade(data) {
    const s = el('section','bloco'); s.id='qualidade';
    const itens = []; // [kind, titulo, desc, expandableHTML?]

    if (data.network.tabGtCount>0) {
      const linhas = data.network.tabGtList.map(r =>
        `<tr><td>${r.unidade}</td><td>${r.code} ${r.desc}</td><td>${MSHORT[r.mes]}</td><td class="num">${fmt(r.real)}</td><td class="num">${fmt(r.tab)}</td><td class="num qa-gap">+${fmt(r.tab-r.real)}</td></tr>`
      ).join('');
      const tabela = `
        <button class="qa-toggle" type="button">Ver os ${data.network.tabGtCount} registros</button>
        <div class="qa-detalhe" hidden>
          <div class="qa-tablewrap"><table class="qa-table">
            <thead><tr><th>Unidade</th><th>Procedimento</th><th>Mês</th><th>Realizado</th><th>Tabulado</th><th>Diferença</th></tr></thead>
            <tbody>${linhas}</tbody>
          </table></div>
        </div>`;
      itens.push(['warn', `${data.network.tabGtCount} registros com Tabulado maior que Realizado`, 'Possível diferença de competência entre produção e tabulação. A investigar.', tabela]);
    }
    data.warnings.forEach(w => itens.push(['warn', 'Valor inconsistente', w]));
    data.units.filter(u=>u.hasTab).forEach(u=>{
      if (u.naoTabulados.length)
        itens.push(['info', `${u.name}: ${u.naoTabulados.length} procedimentos sem tabulação`, 'Alguns podem não ser tabuláveis pela natureza (articulação de redes, matriciamento); outros deveriam ser registrados.']);
    });
    const semTab = data.units.filter(u=>!u.hasTab);
    if (semTab.length)
      itens.push(['info', `${semTab.length} unidades sem coluna de tabulação`, semTab.map(u=>u.name).join(', ')+'. Serão incluídas quando os dados forem informados.']);
    if (!itens.length) itens.push(['ok','Nenhuma inconsistência relevante detectada','Os dados passaram nas verificações automáticas.']);

    s.innerHTML = `
      <div class="bloco-head">
        <div class="bloco-kicker">Transparência</div>
        <div class="bloco-title">Qualidade dos dados</div>
        <p class="bloco-lede">Pontos de atenção identificados automaticamente na planilha. Não alteram os números apresentados: são sinalizados para leitura cuidadosa e para orientar a conferência com as equipes.</p>
      </div>
      <div class="qa-list">
        ${itens.map(([k,t,d,extra])=>{
          const ico = k==='ok'?ICON.check:k==='warn'?ICON.alert:ICON.info;
          const cls = k==='ok'?'qa-ok':k==='warn'?'qa-warn':'qa-info';
          return `<div class="qa-item ${cls}"><span class="qa-ico">${ico}</span><div class="qa-body"><div class="qa-t">${t}</div><div class="qa-d">${d}</div>${extra||''}</div></div>`;
        }).join('')}
      </div>`;

    // liga o botão de expandir
    setTimeout(()=>{
      const btn = s.querySelector('.qa-toggle');
      if (btn) btn.onclick = () => {
        const det = btn.nextElementSibling;
        const open = !det.hasAttribute('hidden');
        if (open) { det.setAttribute('hidden',''); btn.textContent = `Ver os ${data.network.tabGtCount} registros`; }
        else { det.removeAttribute('hidden'); btn.textContent = 'Ocultar registros'; }
      };
    },0);
    return s;
  }

  // ========================================================
  //  Seção: INSIGHTS
  // ========================================================
  function secInsights(data) {
    const s = el('section','bloco'); s.id='insights';
    const caps = data.units.filter(u=>u.hasTab);
    const cards = [];

    // 1. pior taxa de tabulação
    const pior = [...caps].filter(u=>u.taxaGeral!=null).sort((a,b)=>a.taxaGeral-b.taxaGeral)[0];
    if (pior) cards.push(['crit', ICON.percent, `Tabulação crítica no ${pior.name}`,
      `Apenas ${pct(pior.taxaGeral)} da produção chegou ao sistema nos meses com tabulação. Recomenda-se investigar o fluxo de registro desta unidade.`]);

    // 2. maior perda concentrada num procedimento
    let maiorGapItem=null;
    caps.forEach(u => u.naoTabulados.forEach(p => { if(!maiorGapItem||p.real>maiorGapItem.real) maiorGapItem={...p,unidade:u.name}; }));
    if (maiorGapItem) cards.push(['warn', ICON.target,
      `Perda concentrada: ${maiorGapItem.desc}`,
      `No ${maiorGapItem.unidade}, este procedimento teve ${fmt(maiorGapItem.real)} realizações e nenhuma tabulação no período. Verificar se é tabulável e, em caso afirmativo, corrigir o registro.`]);

    // 3. concentração do cuidado
    const top2 = data.network.mix.slice(0,2);
    const top2share = 100*top2.reduce((a,m)=>a+m.real,0)/data.network.totalReal;
    cards.push(['info', ICON.activity, 'Produção concentrada em atendimentos diretos',
      `${top2.map(m=>m.desc.toLowerCase()).join(' e ')} respondem por ${pct(top2share)} da produção da rede, refletindo o foco em atendimento individual e familiar.`]);

    // 4. tendência de produção
    const months = data.network.months;
    if (months.length>=3) {
      const first = data.units.reduce((a,u)=>a+u.procedures.reduce((x,p)=>x+(p.realizado[months[0]]||0),0),0);
      const last  = data.units.reduce((a,u)=>a+u.procedures.reduce((x,p)=>x+(p.realizado[months[months.length-1]]||0),0),0);
      if (last>first) cards.push(['ok', ICON.trending, 'Produção em crescimento',
        `A produção mensal da rede passou de ${fmt(first)} (${MSHORT[months[0]]}) para ${fmt(last)} (${MSHORT[months[months.length-1]]}), indicando ampliação da oferta de cuidado.`]);
    }

    // 5. unidade com volume muito baixo
    const baixa = data.units.filter(u=>u.totalReal>0 && u.totalReal<100).sort((a,b)=>a.totalReal-b.totalReal)[0];
    if (baixa) cards.push(['warn', ICON.alert, `Volume muito baixo: ${baixa.name}`,
      `Apenas ${fmt(baixa.totalReal)} procedimentos em todo o período. Avaliar se há subutilização do serviço ou subnotificação da produção.`]);

    s.innerHTML = `
      <div class="bloco-head">
        <div class="bloco-kicker">Leitura dos dados</div>
        <div class="bloco-title">Principais insights</div>
        <p class="bloco-lede">Sínteses geradas a partir dos números desta planilha, para orientar a conversa com a equipe e as decisões de gestão.</p>
      </div>
      <div class="insight-grid">
        ${cards.map(([k,ico,t,d])=>{
          const cls = k==='crit'?'ins-crit':k==='warn'?'ins-warn':k==='ok'?'ins-ok':'ins-info';
          return `<div class="insight ${cls}"><div class="insight-ico">${ico}</div><div class="insight-t">${t}</div><div class="insight-d">${d}</div></div>`;
        }).join('')}
      </div>
      <div class="rec">
        <div class="rec-title">Recomendações</div>
        <ul>
          <li>Priorizar a revisão do fluxo de tabulação nas unidades com taxa crítica, começando pelos procedimentos de maior volume não registrado.</li>
          <li>Esclarecer, com a equipe, os casos de tabulado maior que realizado, verificando a competência de lançamento.</li>
          <li>Confirmar se procedimentos sem tabulação são de fato não tabuláveis no sistema, separando-os dos que exigem correção de registro.</li>
          <li>Acompanhar a unidade de baixo volume para distinguir subutilização de subnotificação.</li>
        </ul>
      </div>`;
    return s;
  }

  // registro de gráficos: guarda instância + função que gera a option,
  // permitindo redesenhar ao trocar de tema.
  const REG = [];
  function regChart(dom, buildOption) {
    const ch = echarts.init(dom);
    ch.setOption(buildOption());
    REG.push({ ch, buildOption });
    return ch;
  }
  function renderAllCharts() {
    REG.forEach(r => r.ch.setOption(r.buildOption(), true));
  }
  window.addEventListener('resize', () => REG.forEach(r => r.ch.resize()));

  // ========================================================
  //  UPLOAD
  // ========================================================
  function handleFile(file) {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type:'array' });
        const data = parseWorkbook(wb);
        if (data.units.length) computeMetrics(data);
        renderValidation(data);
      } catch (err) {
        const box = $('#validacao'); box.classList.add('show'); box.innerHTML='';
        addVal(box,'err','Não foi possível ler o arquivo. Verifique se é um .xlsx válido. ('+err.message+')');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function initUpload() {
    applyTheme(getTheme());
    const tbu = $('#tema-btn-upload');
    if (tbu) { tbu.innerHTML = getTheme()==='dark'?ICON.sun:ICON.moon; tbu.onclick = () => { toggleTheme(); tbu.innerHTML = getTheme()==='dark'?ICON.sun:ICON.moon; }; }
    const dz = $('#dropzone'), input = $('#file-input');
    dz.onclick = () => input.click();
    input.onchange = () => { if (input.files[0]) handleFile(input.files[0]); };
    ['dragover','dragenter'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('drag');}));
    ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('drag');}));
    dz.addEventListener('drop', e => { const f=e.dataTransfer.files[0]; if(f) handleFile(f); });
  }

  document.addEventListener('DOMContentLoaded', initUpload);
})();
