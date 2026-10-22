

const MESES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
// Nomes sem acento (o texto da planilha é normalizado antes da comparação)
const MAPA_MESES = {"JANEIRO":1,"FEVEREIRO":2,"MARCO":3,"ABRIL":4,"MAIO":5,"JUNHO":6,"JULHO":7,"AGOSTO":8,"SETEMBRO":9,"OUTUBRO":10,"NOVEMBRO":11,"DEZEMBRO":12};

const $ = id => document.getElementById(id);
const brl = v => v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
const label = k => { const [a,m] = k.split("-"); return MESES[+m-1] + "/" + a; };

// Converte número ou texto ("R$ 1.234,56", "1234.56", "1.234,56") em número
const num = v => {
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  let s = String(v).replace(/[^\d,.\-]/g, "");
  if (!/\d/.test(s)) return 0;
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};
const pct = p => (p*100).toLocaleString("pt-BR",{maximumFractionDigits:4}) + "%";

let data = {}, warnings = [], nfiles = 0, chart = null, subs = {};
const cv = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();

const ehAnoValido = a => {
  const n = parseInt(a, 10);
  return !isNaN(n) && n >= 2022 && n <= 2026;
};

/* ============ LEITURA DO SUBSÍDIO (2022 a 2026) ============ */
const norm = s => String(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
const RE_ANO = /(?:^|\D)(202[2-6])(?!\d)/;
const anoDeTexto = t => { const m = String(t).match(RE_ANO); return m ? m[1] : null; };

// Ano contido numa célula "de cabeçalho": número inteiro 2022-2026, texto com letras
// ("SUBSÍDIO 2023", "ANO: 2024") ou texto que seja só o ano. Datas (15/03/2023) são ignoradas.
function anoDeCabecalho(cel){
  if(typeof cel === "number") return (Number.isInteger(cel) && cel >= 2022 && cel <= 2026) ? String(cel) : null;
  const t = norm(cel);
  if(!t) return null;
  const a = anoDeTexto(t);
  if(!a) return null;
  return (/[A-Z]/.test(t) || t === a) ? a : null;
}

// Percorre TODAS as abas do arquivo e devolve {"AAAA-MM": valor}
function lerSubsidio(wb, nomeArquivo){
  const achados = {}, semAno = [], conflitos = [];
  const anoArq = anoDeTexto(nomeArquivo);

  for(const sn of wb.SheetNames){
    const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header: 1, defval: ""});
    let anoAtual = anoDeTexto(sn);           // aba chamada "2023", "SUBSIDIO 2024" etc.

    for(let r = 0; r < matrix.length; r++){
      const row = matrix[r];
      const colsMensal = [];
      row.forEach((cel, c) => { if(/^MENSAL\b/.test(norm(cel))) colsMensal.push(c); });

      // Linha sem "MENSAL": se tiver um ano de cabeçalho, ele vale daqui para baixo
      if(!colsMensal.length){
        for(const cel of row){ const a = anoDeCabecalho(cel); if(a){ anoAtual = a; break; } }
        continue;
      }

      for(const c of colsMensal){
        const txt = norm(row[c]);
        const m = txt.match(/^MENSAL\s*[-:.\/]?\s*([A-Z]+)/);
        if(!m || !MAPA_MESES[m[1]]) continue;
        const mes = MAPA_MESES[m[1]];

        // Prioridade do ano: na própria célula > à esquerda na mesma linha > cabeçalho acima > nome da aba/arquivo
        let ano = anoDeTexto(txt);
        if(!ano){
          for(let i = c - 1; i >= 0 && !ano; i--) ano = anoDeCabecalho(row[i]);
        }
        if(!ano) ano = anoAtual || anoArq;
        if(!ano || !ehAnoValido(ano)){
          semAno.push(m[1] + (ano ? "/" + ano : "") + " (aba " + sn + ")");
          continue;
        }

        // Valor: primeira célula numérica à direita (ignora células vazias e anos soltos)
        let valor = 0;
        for(let off = 1; off <= 8 && c + off < row.length; off++){
          const cel = row[c + off];
          if(cel === "" || cel === null) continue;
          if(anoDeCabecalho(cel) && num(cel) >= 2022 && num(cel) <= 2026 && Number.isInteger(num(cel))) continue;
          const n = num(cel);
          if(n > 0){ valor = n; break; }
        }

        const key = ano + "-" + String(mes).padStart(2, "0");
        if(achados[key] !== undefined && achados[key] !== valor){
          conflitos.push(label(key));
        }
        achados[key] = valor;
      }
    }
  }
  return {achados, semAno, conflitos};
}
/* ============================================================ */

/* ============ SUBSÍDIO EM TABELA (data na coluna A, valor ao lado) ============ */
// Aceita: data do Excel (serial ou Date), "01/07/2022", "07/2022", "jul/22", "jul-2022", "2022-07".
function mesDeCelula(cel){
  let y, m;
  if(cel instanceof Date && !isNaN(cel)){ y = cel.getFullYear(); m = cel.getMonth() + 1; }
  else if(typeof cel === "number"){
    if(!(cel > 40000 && cel < 60000)) return null;       // serial de data (2009-2064)
    const d = XLSX.SSF.parse_date_code(cel);
    if(!d) return null; y = d.y; m = d.m;
  } else if(typeof cel === "string"){
    const t = norm(cel); let r;
    if((r = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))){ m = +r[2]; y = +r[3]; }
    else if((r = t.match(/^(\d{1,2})[\/-](\d{4})$/))){ m = +r[1]; y = +r[2]; }
    else if((r = t.match(/^(\d{4})-(\d{1,2})(?:-\d{1,2})?$/))){ y = +r[1]; m = +r[2]; }
    else if((r = t.match(/^([A-Z]{3,9})\.?\s*[\/-]\s*(\d{2}|\d{4})$/))){
      const nome = Object.keys(MAPA_MESES).find(k => k === r[1] || k.slice(0,3) === r[1].slice(0,3));
      if(!nome) return null; m = MAPA_MESES[nome]; y = +r[2]; if(y < 100) y += 2000;
    } else return null;
  } else return null;
  if(!(m >= 1 && m <= 12) || !ehAnoValido(y)) return null;
  return y + "-" + String(m).padStart(2, "0");
}

function lerSubsidioPorData(wb){
  const achados = {}, conflitos = [];
  for(const sn of wb.SheetNames){
    const M = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header: 1, defval: ""});
    for(const row of M){
      const c0 = row.findIndex(x => x !== "" && x !== null);
      if(c0 < 0) continue;
      const key = mesDeCelula(row[c0]);
      if(!key) continue;
      let valor = null;
      for(let c = c0 + 1; c < row.length; c++){
        if(row[c] === "" || row[c] === null) continue;
        if(typeof row[c] === "number" || /\d/.test(String(row[c]))){ valor = num(row[c]); break; }
      }
      if(valor === null) continue;
      if(achados[key] !== undefined && achados[key] !== valor) conflitos.push(label(key));
      achados[key] = valor;
    }
  }
  return {achados, semAno: [], conflitos};
}

/* ============ LEITURA DA ANÁLISE (planilha "Análise - Ideal", aba Receita) ============ */
// Pega a "Receita Total" da coluna RDO na linha de total de cada bloco mensal.
// Aceita títulos "Receita - Janeiro/2025", "Receita Outubro/2025" e blocos sem mês
// (o mês é deduzido como o seguinte ao bloco anterior).
function lerAnalise(wb){
  const achados = {}, avisos = [];
  const sn = wb.SheetNames.find(n => norm(n) === "RECEITA");
  if(!sn) return {achados, avisos};
  const M = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header: 1, defval: ""});

  const blocos = [];
  let ult = null;
  for(let r = 0; r < M.length; r++){
    for(let c = 0; c < Math.min(M[r].length, 4); c++){
      if(typeof M[r][c] !== "string") continue;
      const m = norm(M[r][c]).match(/^RECEITA\s*(?:-\s*)?([A-Z]+)?\s*\/?\s*(\d{4})?$/);
      if(!m) continue;
      let ano, mes;
      if(m[1]){
        mes = MAPA_MESES[m[1]];
        if(!mes) continue;                       // "Receita BUC", "Receita Total" etc.
        ano = m[2] ? +m[2] : (ult ? ult.ano : null);
      } else if(ult){                            // título sem mês: segue o bloco anterior
        mes = ult.mes + 1; ano = ult.ano;
        if(mes > 12){ mes = 1; ano++; }
        avisos.push("bloco sem mês na aba " + sn + " (linha " + (r+1) + ") tratado como " + label(ano + "-" + String(mes).padStart(2,"0")));
      } else continue;
      if(!ano) continue;
      ult = {ano, mes};
      blocos.push({r, ano, mes});
      break;
    }
  }

  blocos.forEach((b, i) => {
    const fim = i + 1 < blocos.length ? blocos[i+1].r : M.length;
    const key = b.ano + "-" + String(b.mes).padStart(2, "0");
    let col = -1, hdr = -1;
    for(let h = b.r + 1; h <= Math.min(b.r + 4, fim - 1) && col < 0; h++){
      for(let c = 0; c < M[h].length; c++){
        if(norm(M[h][c]) === "RECEITA TOTAL" && norm(M[h-1][c]) === "RDO"){ col = c; hdr = h; break; }
      }
    }
    if(col < 0){ avisos.push("coluna RDO/Receita Total não encontrada em " + label(key)); return; }
    let tot = null;
    for(let i2 = hdr + 1; i2 < fim; i2++){
      const rot = norm(M[i2][0]);
      if((rot === "TOTAL:" || rot === "TOTAL" || rot === "") && typeof M[i2][col] === "number" && M[i2][col] > 0) tot = M[i2][col];
    }
    if(tot === null){ avisos.push("linha de total não encontrada em " + label(key)); return; }
    if(achados[key] !== undefined) avisos.push("mês repetido na análise: " + label(key) + " (usado o último)");
    achados[key] = tot;
  });
  return {achados, avisos};
}
/* ============================================================ */

async function readFiles(files){
  const bad = [];
  const infoSubs = [];
  const infoAna = [];
  warnings = [];
  for(const f of files){
    try{
      const wb = XLSX.read(await f.arrayBuffer(),{type:"array"});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet,{defval:0});

      if(rows.length && "RECTOTAL" in rows[0] && "MES" in rows[0] && "ANO" in rows[0]){
        const tot = {};
        for(const r of rows){
          const anoStr = String(r.ANO).trim();
          if (!ehAnoValido(anoStr)) continue;

          const k = anoStr + "-" + String(r.MES).trim().padStart(2,"0");
          tot[k] = (tot[k] || 0) + num(r.RECTOTAL);
        }
        nfiles++;
        for(const k in tot){
          const old = data[k];
          if(old && old.file){
            warnings.push(label(k) + ": RDO de " + old.file + " e " + f.name);
            if(f.lastModified < old.mod) continue;
          }
          data[k] = {rdo:tot[k], file:f.name, mod:f.lastModified};
        }
      } 
      else if(Object.keys(lerAnalise(wb).achados).length){
        const ana = lerAnalise(wb), ka = Object.keys(ana.achados).sort();
        const fim = ka[ka.length-1], repetidos = [];
        nfiles++;
        for(const k of ka){
          const old = data[k];
          if(old && old.file){
            repetidos.push(label(k));
            // vale o arquivo de período mais recente (reapuração); empate: o salvo por último
            const manter = old.fim !== undefined
              ? (old.fim > fim || (old.fim === fim && f.lastModified < old.mod))
              : f.lastModified < old.mod;
            if(manter) continue;
          }
          data[k] = {rdo:ana.achados[k], file:f.name, mod:f.lastModified, fim:fim};
        }
        if(repetidos.length) warnings.push(f.name + ": " + repetidos.length + " mes(es) ja existiam em outro arquivo (mantido o de periodo mais recente)");
        infoAna.push(f.name + ": RDO de " + label(ka[0]) + " a " + label(ka[ka.length-1]) + " (" + ka.length + " meses)");
        ana.avisos.forEach(a => warnings.push(a));
      }
      else {
        let res = lerSubsidio(wb, f.name);
        if(!Object.keys(res.achados).length) res = lerSubsidioPorData(wb);
        const ks = Object.keys(res.achados);

        if(ks.length){
          nfiles++;
          for(const k of ks) subs[k] = res.achados[k];
          try{ localStorage.setItem("bi_rdo_subs", JSON.stringify(subs)); }catch(x){}
          const porAno = {};
          ks.forEach(k => { const y = k.slice(0,4); porAno[y] = (porAno[y] || 0) + 1; });
          infoSubs.push(f.name + ": " + Object.keys(porAno).sort().map(y => y + " (" + porAno[y] + " meses)").join(", "));
          if(res.semAno.length) warnings.push("Subsídio sem ano identificado em " + f.name + ": " + res.semAno.join(", "));
          if(res.conflitos.length) warnings.push("Subsídio repetido com valores diferentes em " + f.name + " (usado o último): " + [...new Set(res.conflitos)].join(", "));
        } else {
          bad.push(f.name);
        }
      }
    }catch(e){ bad.push(f.name); }
  }

  const allKeys = [...new Set([...Object.keys(data), ...Object.keys(subs)])]
    .filter(k => ehAnoValido(k.slice(0,4)))
    .sort();

  allKeys.forEach(k => { if(!data[k]) data[k] = {rdo:0, file:"", mod:0}; });

  const ks = keys();
  let h = "<b>" + nfiles + "</b> arquivo(s) processado(s). <b>" + ks.length + "</b> mes(es) valido(s) (2022 a 2026) localizado(s)" + (ks.length ? ", de " + label(ks[0]) + " a " + label(ks[ks.length-1]) : "") + ".";
  if(infoSubs.length) h += "<div class='ok'>Subsidio lido - " + infoSubs.join(" | ") + "</div>";
  if(infoAna.length) h += "<div class='ok'>Analise lida (Receita Total RDO) - " + infoAna.join(" | ") + "</div>";
  if(bad.length) h += "<div class='warn'>Arquivos nao reconhecidos: " + bad.join(", ") + "</div>";
  if(warnings.length) h += "<div class='warn'>Avisos: " + warnings.join("; ") + "</div>";
  $("msgs").innerHTML = h;

  const cur = $("ano").value, anos = [...new Set(ks.map(k => k.slice(0,4)))].filter(ehAnoValido);
  $("ano").innerHTML = "<option value=''>Todos os anos (2022-2026)</option>" + anos.map(a => "<option>" + a + "</option>").join("");
  $("ano").value = anos.includes(cur) ? cur : "";
  
  buildMonths();
  update();
}

const keys = () => Object.keys(data)
  .filter(k => ehAnoValido(k.slice(0,4)))
  .sort()
  .filter(k => !$("ano").value || k.startsWith($("ano").value));

function buildMonths(){
  const ks = keys();
  if(!ks.length){
    $("months").innerHTML = "<tbody><tr><td class='empty' colspan='7'>Carregue os arquivos para visualizar os relatorios (2022 a 2026).</td></tr></tbody>";
    return;
  }
  $("months").innerHTML = "<thead><tr><th>Mes</th><th>Receita RDO</th><th>Subsidio</th><th>Receita Total</th><th id='hA'></th><th id='hB'></th><th>Diferenca</th></tr></thead>" +
  "<tbody>" + ks.map(k => "<tr data-k='" + k + "'><td><b>" + label(k) + "</b></td><td>" + brl(data[k].rdo) + "</td>" +
    "<td><input type='number' step='0.01' min='0' placeholder='0,00' value='" + (subs[k] !== undefined ? subs[k] : "") + "' aria-label='Subsidio de " + label(k) + "'></td>" +
    "<td class='t'></td><td class='a'></td><td class='b'></td><td class='dif d'></td></tr>").join("") + "</tbody><tfoot id='mf'></tfoot>";
  
  document.querySelectorAll("#months tbody input").forEach(i => i.addEventListener("input", e => {
    subs[e.target.closest("tr").dataset.k] = num(e.target.value);
    try{ localStorage.setItem("bi_rdo_subs", JSON.stringify(subs)); }catch(x){}
    update();
  }));
}

function update(){
  const pa = num($("ra").value)/100, pb = num($("rb").value)/100, ks = keys();
  let R = 0, S = 0;
  document.querySelectorAll("#months tbody tr[data-k]").forEach(tr => {
    const k = tr.dataset.k, r = data[k].rdo, s = subs[k] || 0, t = r + s;
    tr.querySelector(".t").textContent = brl(t);
    tr.querySelector(".a").textContent = brl(t * pa);
    tr.querySelector(".b").textContent = brl(t * pb);
    tr.querySelector(".d").textContent = brl(t * (pb - pa));
    R += r; S += s;
  });
  const T = R + S, has = ks.length > 0, v = x => has ? brl(x) : "0,00";
  if($("hA")){ $("hA").textContent = "ISS a " + pct(pa); $("hB").textContent = "ISS a " + pct(pb); }
  if($("mf")) $("mf").innerHTML = "<tr><td>Total Consolidado</td><td>" + brl(R) + "</td><td>" + brl(S) + "</td><td>" + brl(T) + "</td><td>" + brl(T*pa) + "</td><td>" + brl(T*pb) + "</td><td>" + brl(T*(pb-pa)) + "</td></tr>";
  const pr = T ? R/T*100 : 100, ps = 100 - pr, an = $("ano").value || "2022 a 2026";
  $("kpis").innerHTML = 
   "<div class='kpi main'><small>Receita Total (" + an + ")</small><b>" + v(T) + "</b>" +
     "<div class='split'><i style='width:" + pr + "%;background:#60a5fa'></i><i style='width:" + ps + "%;background:var(--orange-primary)'></i></div>" +
     "<div class='legend'><span style='--c:#60a5fa'>RDO " + v(R) + "</span><span style='--c:var(--orange-primary)'>Subsidio " + v(S) + "</span></div></div>" +
   "<div class='kpi a'><small>ISS na Aliquota " + pct(pa) + "</small><b>" + v(T*pa) + "</b></div>" +
   "<div class='kpi b'><small>ISS na Aliquota " + pct(pb) + "</small><b>" + v(T*pb) + "</b></div>" +
   "<div class='kpi d'><small>Diferenca Total de ISS</small><b>" + v(T*(pb-pa)) + "</b></div>";
  $("sign").textContent = ks.length ? (label(ks[0]) + " a " + label(ks[ks.length-1])).toUpperCase() : "SEM DADOS";
  drawYears(pa,pb); drawChart(ks);
}

function drawYears(pa,pb){
  const ys = {};
  keys().forEach(k => {
    const y = k.slice(0,4);
    if(ehAnoValido(y)){
      ys[y] = ys[y] || {r:0, s:0, n:0};
      ys[y].r += data[k].rdo;
      ys[y].s += subs[k] || 0;
      ys[y].n++;
    }
  });
  const ay = Object.keys(ys).sort();
  if(!ay.length){
    $("years").innerHTML = "<tbody><tr><td class='empty' colspan='8'>Sem dados carregados para o periodo de 2022 a 2026.</td></tr></tbody>";
    return;
  }
  $("years").innerHTML = "<thead><tr><th>Ano</th><th>Meses</th><th>Receita RDO</th><th>Subsidio</th><th>Receita Total</th><th>ISS a " + pct(pa) + "</th><th>ISS a " + pct(pb) + "</th><th>Diferenca</th></tr></thead><tbody>" +
   ay.map(y => { const o = ys[y], t = o.r + o.s; return "<tr><td><b>" + y + "</b></td><td>" + o.n + "</td><td>" + brl(o.r) + "</td><td>" + brl(o.s) + "</td><td>" + brl(t) + "</td><td>" + brl(t*pa) + "</td><td>" + brl(t*pb) + "</td><td class='dif'>" + brl(t*(pb-pa)) + "</td></tr>"; }).join("") + "</tbody>";
}

function drawChart(ks){
  if(chart){ chart.destroy(); chart = null; }
  if(!ks.length || typeof Chart === "undefined") return;
  Chart.defaults.font.family = '"Plus Jakarta Sans", sans-serif';
  Chart.defaults.color = "#64748b";
  const compact = v => v.toLocaleString("pt-BR",{notation:"compact",style:"currency",currency:"BRL"});
  chart = new Chart($("chart"), {
    type: "bar",
    data: {
      labels: ks.map(label),
      datasets: [
        {label: "Receita RDO", data: ks.map(k => data[k].rdo), backgroundColor: cv("--blue-accent"), borderRadius: 4},
        {label: "Subsidio", data: ks.map(k => subs[k] || 0), backgroundColor: cv("--orange-primary"), borderRadius: 4}
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {mode: "index", intersect: false},
      scales: {
        x: {stacked: true, grid: {display: false}},
        y: {stacked: true, grid: {color: "#f1f5f9"}, border: {display: false}, ticks: {callback: compact}}
      },
      plugins: {
        legend: {position: "bottom", labels: {usePointStyle: true, boxWidth: 8}},
        tooltip: {callbacks: {label: c => " " + c.dataset.label + ": " + brl(c.parsed.y), footer: i => "Total: " + brl(i.reduce((s,x) => s + x.parsed.y, 0))}}
      }
    }
  });
}

function downloadCsv(){
  const pa = num($("ra").value)/100, pb = num($("rb").value)/100, f = n => n.toFixed(2).replace(".",",");
  const L = [["Mes","Receita RDO","Subsidio","Receita Total","ISS " + pct(pa),"ISS " + pct(pb),"Diferenca"].join(";")];
  keys().forEach(k => {
    const r = data[k].rdo, s = subs[k] || 0, t = r + s;
    L.push([label(k), f(r), f(s), f(t), f(t*pa), f(t*pb), f(t*(pb-pa))].join(";"));
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob(["\ufeff" + L.join("\n")], {type: "text/csv"}));
  a.download = "relatorio_receita_iss.csv";
  a.click();
}

const DEF = {c1: "#122258", c2: "#f25c05"};
let theme = Object.assign({}, DEF);
try{ Object.assign(theme, JSON.parse(localStorage.getItem("bi_rdo_theme") || "{}")); }catch(e){}

function applyTheme(){
  const r = document.documentElement.style;
  r.setProperty("--blue-primary", theme.c1);
  r.setProperty("--orange-primary", theme.c2);
  $("c1").value = theme.c1; $("c2").value = theme.c2;
  update();
}

function saveTheme(){ try{ localStorage.setItem("bi_rdo_theme", JSON.stringify(theme)); }catch(e){} }
$("c1").oninput = e => { theme.c1 = e.target.value; saveTheme(); applyTheme(); };
$("c2").oninput = e => { theme.c2 = e.target.value; saveTheme(); applyTheme(); };
$("reset").onclick = () => {
  theme = Object.assign({}, DEF);
  saveTheme();
  applyTheme();  
};

$("pick").onclick = () => $("file").click();
$("file").onchange = e => readFiles(e.target.files);
const dz = $("drop");
dz.ondragover = e => { e.preventDefault(); dz.classList.add("over"); };
dz.ondragleave = () => dz.classList.remove("over");
dz.ondrop = e => { e.preventDefault(); dz.classList.remove("over"); readFiles(e.dataTransfer.files); };
["ra","rb"].forEach(i => $(i).addEventListener("input", update));
$("ano").onchange = () => { buildMonths(); update(); };
$("csv").onclick = downloadCsv;
$("print").onclick = () => window.print();

buildMonths(); applyTheme();