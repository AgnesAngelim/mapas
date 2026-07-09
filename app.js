'use strict';

/* ================================================================
   URLs GeoJSON
================================================================ */
const GEO_BR      = 'https://raw.githubusercontent.com/codeforamerica/click_that_hood/master/public/data/brazil-states.geojson';
const GEO_USA     = 'https://raw.githubusercontent.com/PublicaMundi/MappingAPI/master/data/geojson/us-states.json';
const GEO_NAMRICA = 'https://raw.githubusercontent.com/holtzy/D3-graph-gallery/master/DATA/world.geojson';

/* ================================================================
   ESTADO GLOBAL
================================================================ */
let geoBR     = null;
let geoUSA    = null;
let geoNAm    = null;   // GeoJSON América do Norte (EUA + Canadá + México)

// Dados processados
let licenciados        = [];   // aba licenciados
let clientes           = [];   // aba clientes
let internacional      = [];   // aba internacional (clientes)
let licInternacional   = [];   // aba licenciados internacional

// Filtros ativos por tela
const F = {
  L:  { license: 'all', status: 'all' },   // tela 1
  C:  { status: 'all',  port: 'all'   },   // tela 2
  CR: { licensee: null                },   // tela 3
  US: {},                                  // tela 4
  UL: { license: 'all', status: 'all' },   // tela 5
};

/* ================================================================
   CONSTANTES DE DOMÍNIO
================================================================ */
const LICENSE_MAP = {
  telecom:  { label: 'Telecom',          cls: 'telecom',  color: '#58a6ff' },
  expert:   { label: 'Expert + Telecom', cls: 'expert',   color: '#bc8cff' },
  expansao: { label: 'Expansão USA',     cls: 'expansao', color: '#ff7b72' },
  other:    { label: 'Outros',           cls: 'other',    color: '#8b949e' },
};

// Estados sem conexão Green (energia)
const ESTADOS_SEM_GREEN = {
  AC: 'Energisa Acre',
  AM: 'AME',
  AP: 'CEA',
  DF: 'CEB',
  RO: 'Energisa Rondônia',
  RR: 'Boa Vista Energia',
  TO: 'Energisa Tocantins',
};

// Siglas dos estados dos EUA → nome
const USA_STATE_NAMES = {
  AL:'Alabama',AK:'Alaska',AZ:'Arizona',AR:'Arkansas',CA:'California',CO:'Colorado',
  CT:'Connecticut',DE:'Delaware',FL:'Florida',GA:'Georgia',HI:'Hawaii',ID:'Idaho',
  IL:'Illinois',IN:'Indiana',IA:'Iowa',KS:'Kansas',KY:'Kentucky',LA:'Louisiana',
  ME:'Maine',MD:'Maryland',MA:'Massachusetts',MI:'Michigan',MN:'Minnesota',MS:'Mississippi',
  MO:'Missouri',MT:'Montana',NE:'Nebraska',NV:'Nevada',NH:'New Hampshire',NJ:'New Jersey',
  NM:'New Mexico',NY:'New York',NC:'North Carolina',ND:'North Dakota',OH:'Ohio',
  OK:'Oklahoma',OR:'Oregon',PA:'Pennsylvania',RI:'Rhode Island',SC:'South Carolina',
  SD:'South Dakota',TN:'Tennessee',TX:'Texas',UT:'Utah',VT:'Vermont',VA:'Virginia',
  WA:'Washington',WV:'West Virginia',WI:'Wisconsin',WY:'Wyoming',DC:'D.C.'
};

const USA_NAME_TO_SIGLA = {};
Object.entries(USA_STATE_NAMES).forEach(([s, n]) => { USA_NAME_TO_SIGLA[n.toLowerCase()] = s; });

/* ================================================================
   PRÉ-CARREGA GeoJSONs
================================================================ */
Promise.all([
  fetch(GEO_BR).then(r => r.json()).catch(() => null),
  fetch(GEO_USA).then(r => r.json()).catch(() => null),
  fetch(GEO_NAMRICA).then(r => r.json()).catch(() => null),
]).then(([br, us, nam]) => {
  geoBR  = br;
  geoUSA = us;
  // Filtra apenas EUA, Canadá e México
  if (nam) {
    geoNAm = {
      type: 'FeatureCollection',
      features: nam.features.filter(f => {
        const n = (f.properties.name || '').toLowerCase();
        return n === 'united states of america' || n === 'canada' || n === 'mexico';
      })
    };
  }
});

/* ================================================================
   NAV TABS
================================================================ */
document.querySelectorAll('.nav-tab').forEach(btn => {
  btn.addEventListener('click', function () {
    document.querySelectorAll('.nav-tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    this.classList.add('active');
    document.getElementById('tab-' + this.dataset.tab).classList.add('active');
  });
});

/* ================================================================
   MODAL
================================================================ */
let modalRows = [];

document.getElementById('modalClose').addEventListener('click', closeModal);
document.getElementById('modalOverlay').addEventListener('click', function (e) {
  if (e.target === this) closeModal();
});
document.getElementById('modalSearch').addEventListener('input', function () {
  renderModalTable(modalRows, this.value);
});

function openModal(title, kpis, rows, columns) {
  modalRows = rows;
  document.getElementById('modalTitle').textContent = title;
  document.getElementById('modalKpis').innerHTML = kpis.map(k =>
    '<div class="modal-kpi"><div class="modal-kpi-val" style="color:' + (k.color || 'var(--text)') + '">' + k.val + '</div><div class="modal-kpi-label">' + k.label + '</div></div>'
  ).join('');
  document.getElementById('modalSearch').value = '';
  window._modalColumns = columns;
  renderModalTable(rows, '');
  document.getElementById('modalOverlay').style.display = 'flex';
}
function closeModal() { document.getElementById('modalOverlay').style.display = 'none'; }

function renderModalTable(rows, search) {
  const cols = window._modalColumns || [];
  const s = search.toLowerCase();
  const filtered = rows.filter(r => cols.some(c => String(r[c.key] || '').toLowerCase().includes(s)));
  document.getElementById('modalBody').innerHTML =
    '<table><thead><tr>' + cols.map(c => '<th>' + c.label + '</th>').join('') + '</tr></thead><tbody>' +
    filtered.map(r => '<tr>' + cols.map(c => '<td>' + (c.render ? c.render(r) : (r[c.key] ?? '—')) + '</td>').join('') + '</tr>').join('') +
    '</tbody></table>';
}

/* ================================================================
   FILE INPUT — lê as 2 abas automaticamente
================================================================ */
document.getElementById('fileInput').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('uploadText').textContent = '📄 ' + file.name;
  const reader = new FileReader();
  reader.onload = function (evt) {
    const wb = XLSX.read(evt.target.result, { type: 'binary' });
    parseWorkbook(wb);
  };
  reader.readAsBinaryString(file);
});

function parseWorkbook(wb) {
  // Detecta abas por nome (licenciado/cliente)
  const sheetNames = wb.SheetNames.map(n => n.toLowerCase());
  const licIdx = sheetNames.findIndex(n => n.includes('licencia') || n.includes('licença') || n.includes('licensee'));
  const cliIdx = sheetNames.findIndex(n => n.includes('cliente') || n.includes('client'));

  // Detecta abas: clientes internacionais e licenciados internacionais separados
  const intlCliIdx = sheetNames.findIndex(n => n.includes('intern') && !n.includes('licencia') && !n.includes('licença'));
  const intlLicIdx = sheetNames.findIndex(n => n.includes('intern') && (n.includes('licencia') || n.includes('licença')));
  // Fallback: se só tiver uma aba internacional, usa para clientes
  const intlIdx = intlCliIdx >= 0 ? intlCliIdx : sheetNames.findIndex(n => n.includes('intern'));

  const licSheet    = wb.Sheets[wb.SheetNames[licIdx >= 0 ? licIdx : 0]];
  const cliSheet    = cliIdx    >= 0 ? wb.Sheets[wb.SheetNames[cliIdx]]    : null;
  const intlSheet   = intlIdx   >= 0 ? wb.Sheets[wb.SheetNames[intlIdx]]   : null;
  const intlLicSheet= intlLicIdx>= 0 ? wb.Sheets[wb.SheetNames[intlLicIdx]]: null;

  const rawLic     = XLSX.utils.sheet_to_json(licSheet,     { defval: '', cellDates: true });
  const rawCli     = cliSheet     ? XLSX.utils.sheet_to_json(cliSheet,     { defval: '', cellDates: true }) : [];
  const rawIntl    = intlSheet    ? XLSX.utils.sheet_to_json(intlSheet,    { defval: '', cellDates: true }) : [];
  const rawIntlLic = intlLicSheet ? XLSX.utils.sheet_to_json(intlLicSheet, { defval: '', cellDates: true }) : [];

  // --- Processa licenciados ---
  // Colunas: Codigo, Nome, Cidade, Uf, Status, Cep, Endereco, Numero, Bairro,
  //          Complemento, Cpf, Data Ativo, Data Ativo Telecom, Idpatrocinador,
  //          Patrocinador, Tipo Licenca, Origem
  // Limite: 1 ano atrás a partir de hoje
  const umAnoAtras = new Date();
  umAnoAtras.setFullYear(umAnoAtras.getFullYear() - 1);

  licenciados = rawLic.map(r => {
    const origem     = String(r['Origem'] || r['origem'] || '').trim();
    const uf         = String(r['Uf'] || r['UF'] || r['uf'] || '').trim().toUpperCase();
    const rawLicenca = String(r['Tipo Licenca'] || r['Tipo Licença'] || r['TipoLicenca'] || '').trim();
    const tipoLicenca = normalizeLicense(rawLicenca);

    // Licenciado EUA = Origem contém EXPANSAO_EUA, EXPANSAO_USA ou variações
    const isUSA = /expan/i.test(origem) && /eua|usa/i.test(origem);

    // Estado americano: tenta ler UF como sigla americana, ou extrai da origem
    const usaState = isUSA ? normalizeUSAState(uf) || normalizeUSAState(origem) : null;

    // Status baseado em Data Ativo (coluna L — formato yyyy-mm-dd ou Date do XLSX):
    // Ativo = Data Ativo + 365 dias >= hoje
    // Inativo = Data Ativo + 365 dias < hoje, ou sem data
    const dataAtivoCell = r['Data Ativo'];
    let dataAtivoParsed = null;

    if (dataAtivoCell instanceof Date) {
      // XLSX.js com cellDates:true retorna Date object
      dataAtivoParsed = new Date(dataAtivoCell.getFullYear(), dataAtivoCell.getMonth(), dataAtivoCell.getDate());
    } else if (typeof dataAtivoCell === 'number' && dataAtivoCell > 100) {
      // Serial numérico do Excel (dias desde 30/12/1899)
      const excelEpoch = new Date(1899, 11, 30);
      const d = new Date(excelEpoch.getTime() + dataAtivoCell * 86400000);
      dataAtivoParsed = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    } else if (dataAtivoCell) {
      const s = String(dataAtivoCell).trim();
      // yyyy-mm-dd
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
        const [y, m, d] = s.substring(0, 10).split('-');
        dataAtivoParsed = new Date(+y, +m - 1, +d);
      // dd/mm/yyyy
      } else if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) {
        const [d, m, y] = s.substring(0, 10).split('/');
        dataAtivoParsed = new Date(+y, +m - 1, +d);
      // mm/dd/yyyy (formato americano que raw:false às vezes retorna)
      } else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) {
        const parts = s.split('/');
        dataAtivoParsed = new Date(+parts[2], +parts[0] - 1, +parts[1]);
      } else if (s) {
        const tmp = new Date(s);
        if (!isNaN(tmp.getTime())) dataAtivoParsed = new Date(tmp.getFullYear(), tmp.getMonth(), tmp.getDate());
      }
    }

    // Formata para exibição dd/mm/yyyy
    let dataAtivoRaw = '';
    if (dataAtivoParsed && !isNaN(dataAtivoParsed.getTime())) {
      const dd = String(dataAtivoParsed.getDate()).padStart(2, '0');
      const mm = String(dataAtivoParsed.getMonth() + 1).padStart(2, '0');
      const yy = dataAtivoParsed.getFullYear();
      dataAtivoRaw = dd + '/' + mm + '/' + yy;
    }

    // Coluna "Comprimido": S = inativo imediato, N = calcula pela Data Ativo
    const comprimido = String(r['Comprimido'] || '').trim().toUpperCase();

    let status = 'inativo';
    if (comprimido === 'S') {
      // Inativo independente da data
      status = 'inativo';
    } else {
      // N ou vazio: calcula Data Ativo + 365 dias exatos
      // Millisegundos respeitam meses de 28/29/30/31 dias automaticamente
      if (dataAtivoParsed && !isNaN(dataAtivoParsed.getTime())) {
        const MS_365_DIAS = 365 * 24 * 60 * 60 * 1000;
        const expiraMs = dataAtivoParsed.getTime() + MS_365_DIAS;
        const hoje = new Date();
        const hojeMs = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()).getTime();
        status = expiraMs >= hojeMs ? 'ativo' : 'inativo';
      }
    }

    return {
      codigo:       String(r['Codigo'] || r['Código'] || '').trim(),
      nome:         String(r['Nome'] || '').trim(),
      cidade:       String(r['Cidade'] || '').trim(),
      uf:           isUSA ? null : (uf.length >= 2 ? uf : null),
      status,
      cpf:          String(r['Cpf'] || r['CPF'] || '').trim(),
      dataAtivo:    dataAtivoRaw,
      patrocinador: String(r['Patrocinador'] || '').trim(),
      tipoLicenca,
      rawLicenca,
      origem,
      isUSA,
      usaState,
      clientCount: 0,
    };
  }).filter(r => r.nome);

  // --- Processa clientes ---
  // Colunas: Codigo Cliente, Nome, Cpf Cnpj, Data Ativo, Plano I Green,
  //          Tipo De Linha, Forma De Pagamento, ID Licenciado, Nome Licenciado,
  //          Data Cancelado, Cidade, Estado
  clientes = rawCli.map(r => {
    const cancelado  = String(r['Data Cancelado'] || '').trim();
    const status     = cancelado && cancelado !== '' && cancelado !== '0' ? 'cancelado' : 'ativo';
    const tipoLinha  = String(r['Tipo De Linha'] || r['TipoDeLinha'] || '').trim(); // Esim ou Fisico
    // Coluna "Portabilidade": S = sim, qualquer outro valor = não
    const portCol    = String(r['Portabilidade'] || '').trim().toUpperCase();
    const isPort     = portCol === 'VERDADEIRO' || portCol === 'TRUE';
    const plano      = String(r['Plano I Green'] || '').trim();
    const estado     = String(r['Estado'] || r['UF'] || r['Uf'] || '').trim().toUpperCase();

    // Cliente EUA = Plano I Green contém "Connect" ou "Connect Global"
    const isUSA = /connect/i.test(plano);

    // Estado americano: lê da coluna Estado se for sigla válida dos EUA
    const usaState = isUSA ? (normalizeUSAState(estado) || null) : null;

    return {
      codigo:         String(r['Codigo Cliente'] || r['Código Cliente'] || '').trim(),
      nome:           String(r['Nome'] || '').trim(),
      cpf:            String(r['Cpf Cnpj'] || r['CPF'] || '').trim(),
      dataAtivo:      String(r['Data Ativo'] || '').trim(),
      dataCancelado:  cancelado,
      plano,
      tipoLinha,
      isPort,
      formaPag:       String(r['Forma De Pagamento'] || '').trim(),
      idLicenciado:   String(r['ID Licenciado'] || r['Id Licenciado'] || '').trim(),
      nomeLicenciado: String(r['Nome Licenciado'] || '').trim(),
      cidade:         String(r['Cidade'] || '').trim(),
      // Estado BR apenas para não-EUA
      estado:         isUSA ? null : (estado.length >= 2 ? normalizeStateBR(estado) : null),
      status,
      isUSA,
      usaState,
    };
  }).filter(r => r.nome);

  // Join: conta clientes por licenciado
  const clientsByLic = {};
  clientes.forEach(c => {
    const id = c.idLicenciado;
    if (!id) return;
    if (!clientsByLic[id]) clientsByLic[id] = 0;
    clientsByLic[id]++;
  });
  licenciados.forEach(l => { l.clientCount = clientsByLic[l.codigo] || 0; });

  // --- Processa Internacional ---
  // Colunas: Codigo, Nome, Estado, Data Ativo, Patrocinador, Tipo Licenca, Status Licenca, Cancelado, Origem
  // Debug: loga primeiras chaves da aba Internacional
  if (rawIntl.length > 0) console.log('[Internacional] colunas:', Object.keys(rawIntl[0]));
  if (rawIntl.length > 0) console.log('[Internacional] primeiro registro:', rawIntl[0]);

  internacional = rawIntl.map(r => {
    // Busca coluna Estado ignorando espaços no nome da coluna
    const estadoKey = Object.keys(r).find(k => k.trim().toLowerCase() === 'estado') || 'Estado';
    const estadoRaw = String(r[estadoKey] || '').trim();
    // Extrai sigla do formato "Estado NJ" → "NJ"
    const estadoMatch = estadoRaw.match(/([A-Z]{2})$/);
    const estado = estadoMatch ? estadoMatch[1] : estadoRaw.toUpperCase().slice(-2);
    const usaState = (USA_STATE_NAMES[estado] ? estado : null) || normalizeUSAState(estadoRaw) || estado || null;
    const cancelado = String(r['Cancelado'] || '').trim();
    const status   = cancelado && cancelado !== '' && cancelado !== '0' && cancelado.toLowerCase() !== 'n' ? 'cancelado' : 'ativo';
    const rawLic2  = String(r['Tipo Licenca'] || r['Tipo Licença'] || '').trim();
    return {
      codigo:         String(r['Codigo Cliente'] || r['Codigo'] || r['Código'] || '').trim(),
      nome:           String(r['Nome'] || '').trim(),
      estado,
      usaState,
      dataAtivo:      String(r['Data Ativo'] || '').trim(),
      plano:          String(r['Plano I Green'] || '').trim(),
      tipoLinha:      String(r['Tipo De Linha'] || '').trim(),
      idLicenciado:   String(r['ID Licenciado'] || '').trim(),
      nomeLicenciado: String(r['Nome Licenciado'] || '').trim(),
      tipoLicenca:    normalizeLicense(rawLic2),
      rawLicenca:     rawLic2,
      status,
    };
  }).filter(r => r.nome);

  // --- Processa Licenciados Internacional ---
  // Colunas: Codigo, Nome, Celular, Estado, Data Ativo, Tipo Licenca,
  //          Status Licenca, Cancelado, Comprimido, Patrocinador, Origem
  licInternacional = rawIntlLic.map(r => {
    const estadoKey  = Object.keys(r).find(k => k.trim().toLowerCase() === 'estado') || 'Estado';
    const estadoRaw  = String(r[estadoKey] || '').trim();
    const estadoMatch = estadoRaw.match(/\b([A-Z]{2})\b$/);
    const estado      = estadoMatch ? estadoMatch[1] : estadoRaw.toUpperCase().slice(-2);
    const usaState    = (USA_STATE_NAMES[estado] ? estado : null) || normalizeUSAState(estadoRaw) || estado || null;

    const comprimido  = String(r['Comprimido'] || '').trim().toUpperCase();
    const cancelado   = String(r['Cancelado']  || '').trim();
    const dataAtivoCell = r['Data Ativo'];
    let dataAtivoParsed = null;
    if (dataAtivoCell instanceof Date) {
      dataAtivoParsed = new Date(dataAtivoCell.getFullYear(), dataAtivoCell.getMonth(), dataAtivoCell.getDate());
    } else if (typeof dataAtivoCell === 'number' && dataAtivoCell > 100) {
      const epoch = new Date(1899, 11, 30);
      const d = new Date(epoch.getTime() + dataAtivoCell * 86400000);
      dataAtivoParsed = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    } else if (dataAtivoCell) {
      const s = String(dataAtivoCell).trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) { const [y,m,d] = s.substring(0,10).split('-'); dataAtivoParsed = new Date(+y,+m-1,+d); }
      else if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) { const [d,m,y] = s.split('/'); dataAtivoParsed = new Date(+y,+m-1,+d); }
    }

    let dataAtivoRaw = '';
    if (dataAtivoParsed && !isNaN(dataAtivoParsed.getTime())) {
      dataAtivoRaw = String(dataAtivoParsed.getDate()).padStart(2,'0') + '/' +
                     String(dataAtivoParsed.getMonth()+1).padStart(2,'0') + '/' +
                     dataAtivoParsed.getFullYear();
    }

    let status = 'inativo';
    if (comprimido === 'S') {
      status = 'inativo';
    } else if (dataAtivoParsed && !isNaN(dataAtivoParsed.getTime())) {
      const expira = new Date(dataAtivoParsed.getTime() + 365 * 24 * 60 * 60 * 1000);
      status = expira >= new Date() ? 'ativo' : 'inativo';
    }

    const rawLic2 = String(r['Tipo Licenca'] || r['Tipo Licença'] || '').trim();
    return {
      codigo:       String(r['Codigo'] || r['Código'] || '').trim(),
      nome:         String(r['Nome']   || '').trim(),
      celular:      String(r['Celular']|| '').trim(),
      email:        String(r['Email']  || '').trim(),
      estado,
      usaState,
      dataAtivo:    dataAtivoRaw,
      patrocinador: String(r['Patrocinador'] || '').trim(),
      tipoLicenca:  normalizeLicense(rawLic2),
      rawLicenca:   rawLic2,
      statusLic:    String(r['Status Licenca'] || '').trim(),
      origem:       String(r['Origem'] || '').trim(),
      status,
      clientCount:  0,
    };
  }).filter(r => r.nome);

  initAllTabs();
}

/* ================================================================
   NORMALIZAÇÃO
================================================================ */
let brByName = {}, brBySigla = {};

function buildBRLookup() {
  if (!geoBR || Object.keys(brBySigla).length) return;
  geoBR.features.forEach(f => {
    const p = f.properties;
    brBySigla[p.sigla.toUpperCase()] = p.sigla;
    brByName[p.name.toLowerCase()]   = p.sigla;
  });
}

function normalizeStateBR(raw) {
  if (!raw) return null;
  buildBRLookup();
  const s  = String(raw).trim();
  if (!s || s.length < 2) return null;
  // 1. Sigla exata (ex: "SP", "RJ")
  const up = s.toUpperCase();
  if (brBySigla[up]) return brBySigla[up];
  // 2. Nome completo exato (ex: "São Paulo")
  const lo = s.toLowerCase();
  if (brByName[lo]) return brByName[lo];
  // 3. Nome completo contendo a string (só para strings longas > 3 chars, evita falsos positivos)
  if (lo.length > 3) {
    for (const [name, sigla] of Object.entries(brByName)) {
      if (name === lo) return sigla;
      if (lo.length > 5 && name.startsWith(lo)) return sigla;
    }
  }
  return null;
}

function normalizeUSAState(raw) {
  if (!raw) return null;
  const up = String(raw).trim().toUpperCase();
  if (USA_STATE_NAMES[up]) return up;
  const lo = String(raw).trim().toLowerCase();
  if (USA_NAME_TO_SIGLA[lo]) return USA_NAME_TO_SIGLA[lo];
  for (const [name, sigla] of Object.entries(USA_NAME_TO_SIGLA)) {
    if (name.includes(lo) || lo.includes(name)) return sigla;
  }
  return null;
}

function normalizeLicense(raw) {
  const s = String(raw || '').toLowerCase().trim();
  if (s.includes('expert')) return 'expert';
  if (s.includes('expansao') || s.includes('expansão') || s.includes('usa')) return 'expansao';
  if (s.includes('telecom')) return 'telecom';
  return 'other';
}

function getStateName(sigla) {
  if (!geoBR) return sigla;
  const f = geoBR.features.find(f => f.properties.sigla === sigla);
  return f ? f.properties.name : sigla;
}

/* ================================================================
   PROJEÇÃO — BR e EUA
================================================================ */
const SVG_W = 760, SVG_H = 820, PAD = 28;
const USA_W = 900, USA_H = 580, USA_PAD = 24;

function buildProjection(features, W, H, padVal) {
  let minLng = Infinity, maxLng = -Infinity, minLat = Infinity, maxLat = -Infinity;
  function scan(coord) {
    if (typeof coord[0] === 'number') {
      minLng = Math.min(minLng, coord[0]); maxLng = Math.max(maxLng, coord[0]);
      minLat = Math.min(minLat, coord[1]); maxLat = Math.max(maxLat, coord[1]);
    } else { coord.forEach(scan); }
  }
  features.forEach(f => scan(f.geometry.coordinates));

  const latMid = (minLat + maxLat) / 2;
  const cosLat = Math.cos(latMid * Math.PI / 180);
  const lngSpan = (maxLng - minLng) * cosLat;
  const latSpan = maxLat - minLat;
  const scale = Math.min((W - padVal * 2) / lngSpan, (H - padVal * 2) / latSpan);
  const offX = (W - scale * lngSpan) / 2 + padVal;
  const offY = (H - scale * latSpan) / 2 + padVal;

  return (lng, lat) => [
    offX + (lng - minLng) * cosLat * scale,
    offY + (maxLat - lat) * scale,
  ];
}

function ringToPath(ring, proj) {
  return ring.map(([lng, lat], i) => {
    const [x, y] = proj(lng, lat);
    return (i === 0 ? 'M' : 'L') + x.toFixed(2) + ',' + y.toFixed(2);
  }).join(' ') + ' Z';
}
function geomToPath(geom, proj) {
  const rings = geom.type === 'Polygon' ? [geom.coordinates] : geom.coordinates;
  return rings.map(poly => poly.map(ring => ringToPath(ring, proj)).join(' ')).join(' ');
}
function ringCentroid(ring, proj) {
  let area = 0, cx = 0, cy = 0;
  const pts = ring.map(([lng, lat]) => proj(lng, lat));
  const n = pts.length;
  for (let i = 0, j = n - 1; i < n; j = i++) {
    const [xi, yi] = pts[i], [xj, yj] = pts[j];
    const cross = xi * yj - xj * yi;
    area += cross; cx += (xi + xj) * cross; cy += (yi + yj) * cross;
  }
  area /= 2;
  if (Math.abs(area) < 1e-10) {
    let sx = 0, sy = 0; pts.forEach(([x, y]) => { sx += x; sy += y; });
    return [sx / n, sy / n];
  }
  return [cx / (6 * area), cy / (6 * area)];
}
function geomCentroid(geom, proj) {
  let ring;
  if (geom.type === 'Polygon') ring = geom.coordinates[0];
  else ring = geom.coordinates.reduce((a, b) => a[0].length > b[0].length ? a : b)[0];
  return ringCentroid(ring, proj);
}

/* ================================================================
   HEAT COLOR
================================================================ */
const HEAT_STOPS = ['#1a2a1a','#1a4731','#196f3d','#1e8449','#27ae60','#2ecc71','#58d68d','#82e0aa'];
function heatColor(v, max) {
  if (!v || !max) return HEAT_STOPS[0];
  return HEAT_STOPS[Math.min(Math.floor((v / max) * (HEAT_STOPS.length - 1)), HEAT_STOPS.length - 1)];
}

/* ================================================================
   RENDER SVG MAP (genérico)
================================================================ */
function renderSVGMap({ containerId, tooltipId, geoFeatures, W, H, padVal, byState, maxVal,
                        labelKey, nameKey, onClickState, tooltipContent, legendId }) {
  const container = document.getElementById(containerId);
  const tooltip   = document.getElementById(tooltipId || 'tooltip');
  if (!geoFeatures) { container.innerHTML = '<p style="color:var(--text-muted);padding:40px;text-align:center">⏳ Carregando mapa…</p>'; return; }

  const proj = buildProjection(geoFeatures, W, H, padVal);

  const pathEls = geoFeatures.map(f => {
    const sigla = f.properties[labelKey] || f.properties.sigla || f.properties.name || '';
    const name  = f.properties[nameKey]  || f.properties.name  || sigla;
    const st    = byState[sigla] || {};
    const color = heatColor(st.val || 0, maxVal);
    const d     = geomToPath(f.geometry, proj);
    const [cx, cy] = geomCentroid(f.geometry, proj);
    const extra = JSON.stringify(st).replace(/"/g, '&quot;');
    return '<g class="state-group" data-sigla="' + sigla + '" data-name="' + name + '" data-extra="' + extra + '">'
      + '<path class="state-path" d="' + d + '" fill="' + color + '" />'
      + '<text class="state-label" x="' + cx.toFixed(1) + '" y="' + cy.toFixed(1) + '">' + sigla + '</text>'
      + '</g>';
  }).join('');

  container.innerHTML = '<svg viewBox="0 0 ' + W + ' ' + H + '" xmlns="http://www.w3.org/2000/svg">' + pathEls + '</svg>';

  container.querySelectorAll('.state-group').forEach(g => {
    g.addEventListener('mousemove', e => {
      const st = JSON.parse(g.dataset.extra.replace(/&quot;/g, '"'));
      tooltip.style.display = 'block';
      tooltip.style.left = (e.clientX + 14) + 'px';
      tooltip.style.top  = (e.clientY - 10) + 'px';
      tooltip.innerHTML  = tooltipContent(g.dataset.name, g.dataset.sigla, st);
    });
    g.addEventListener('mouseleave', () => { tooltip.style.display = 'none'; });
    if (onClickState) g.addEventListener('click', () => {
      tooltip.style.display = 'none';
      onClickState(g.dataset.sigla, g.dataset.name);
    });
  });

  if (legendId) {
    document.getElementById(legendId).innerHTML =
      '<span class="legend-label">Menos</span><div class="legend-scale">'
      + HEAT_STOPS.map(c => '<div class="legend-box" style="background:' + c + '"></div>').join('')
      + '</div><span class="legend-label">Mais</span>';
  }
}

/* ================================================================
   HELPERS UI
================================================================ */
function renderKPIs(containerId, cards) {
  document.getElementById(containerId).innerHTML = cards.map(c =>
    '<div class="kpi-card ' + (c.cls || '') + '"><div class="kpi-value">' + c.val + '</div><div class="kpi-label">' + c.label + '</div></div>'
  ).join('');
}

function renderRanking(containerId, entries, max) {
  document.getElementById(containerId).innerHTML = entries.slice(0, 15).map(([sig, v], i) =>
    '<div class="ranking-item"><span class="rank-num">' + (i + 1) + '</span>'
    + '<span class="rank-state">' + sig + '</span>'
    + '<div class="rank-bar-wrap"><div class="rank-bar" style="width:' + (v / max * 100).toFixed(0) + '%"></div></div>'
    + '<span class="rank-count">' + v.toLocaleString('pt-BR') + '</span></div>'
  ).join('') || '<p style="color:var(--text-muted);font-size:12px">Nenhum dado</p>';
}

function makePills(containerId, pills, onSelect) {
  const el = document.getElementById(containerId);
  el.innerHTML = pills.map((p, i) =>
    '<button class="pill ' + p.cls + (i === 0 ? ' active' : '') + '" data-val="' + p.val + '">' + p.label + '</button>'
  ).join('');
  el.querySelectorAll('.pill').forEach(btn => {
    btn.addEventListener('click', function () {
      el.querySelectorAll('.pill').forEach(b => b.classList.remove('active'));
      this.classList.add('active');
      onSelect(this.dataset.val);
    });
  });
}

/* ================================================================
   TELA 1 — LICENCIADOS BR
================================================================ */
function initLicensees() {
  document.getElementById('emptyLicensees').style.display = 'none';
  document.getElementById('dashLicensees').style.display  = 'block';

  makePills('pillsLicensees', [
    { val:'all', label:'Todos', cls:'pill-all' },
    { val:'telecom', label:'Telecom', cls:'pill-telecom' },
    { val:'expert',  label:'Expert + Telecom', cls:'pill-expert' },
    { val:'expansao',label:'Expansão USA', cls:'pill-expansao' },
  ], v => { F.L.license = v; renderLicensees(); });

  makePills('pillsStatusL', [
    { val:'all',    label:'Todos',    cls:'pill-all' },
    { val:'ativo',  label:'Ativos',   cls:'pill-ativo' },
    { val:'inativo',label:'Inativos', cls:'pill-cancelado' },
  ], v => { F.L.status = v; renderLicensees(); });

  renderLicensees();
}

function renderLicensees() {
  // Filtra apenas licenciados BR (não EUA)
  let data = licenciados.filter(l => !l.isUSA);
  if (F.L.license !== 'all') data = data.filter(l => l.tipoLicenca === F.L.license);
  if (F.L.status !== 'all')  data = data.filter(l => l.status === F.L.status || (F.L.status === 'ativo' && l.status !== 'inativo'));

  // Agrega por estado
  const byState = {};
  data.forEach(l => {
    const s = l.uf || 'XX';
    if (!byState[s]) byState[s] = { val: 0, licenciados: [], inactive: 0 };
    byState[s].val++;
    byState[s].licenciados.push(l);
    if (l.clientCount === 0) byState[s].inactive++;
  });

  const maxVal = Math.max(...Object.values(byState).map(v => v.val), 1);
  const totalAtivos = data.filter(l => l.status !== 'inativo').length;
  const semClientes = data.filter(l => l.clientCount === 0).length;

  renderKPIs('kpiLicensees', [
    { val: data.length, label: 'Licenciados', cls: 'kpi-accent' },
    { val: data.filter(l=>l.tipoLicenca==='telecom').length, label: 'Telecom', cls: 'kpi-accent' },
    { val: data.filter(l=>l.tipoLicenca==='expert').length, label: 'Expert+Telecom', cls: 'kpi-purple' },
    { val: semClientes, label: 'Sem Clientes', cls: 'kpi-alert' },
  ]);

  if (!geoBR) { setTimeout(renderLicensees, 600); return; }

  renderSVGMap({
    containerId: 'mapLicensees', geoFeatures: geoBR.features,
    W: SVG_W, H: SVG_H, padVal: PAD,
    byState, maxVal,
    labelKey: 'sigla', nameKey: 'name',
    legendId: 'legendLicensees',
    tooltipContent: (name, sig, st) =>
      '<div class="tooltip-state">' + name + ' (' + sig + ')</div>'
      + '<div class="tooltip-row"><span>Licenciados:</span><strong>' + (st.val || 0) + '</strong></div>'
      + '<div class="tooltip-row"><span>Sem clientes:</span><strong>' + (st.inactive || 0) + '</strong></div>',
    onClickState: (sig, name) => openLicenseesModal(sig, name, byState[sig] || { licenciados: [] }),
  });

  // Rankinging
  const sorted = Object.entries(byState).filter(([s]) => s !== 'XX').sort((a,b) => b[1].val - a[1].val);
  renderRanking('rankLicensees', sorted.map(([s,v]) => [s, v.val]), maxVal);

  // Sem clientes
  const inactive = data.filter(l => l.clientCount === 0);
  document.getElementById('inactiveLicensees').innerHTML = inactive.length
    ? inactive.slice(0, 50).map(l =>
        '<div class="inactive-item"><span class="inactive-name">' + l.codigo + ' — ' + l.nome + '</span>'
        + '<span class="badge badge-' + (LICENSE_MAP[l.tipoLicenca]?.cls || 'other') + '">' + (LICENSE_MAP[l.tipoLicenca]?.label || 'Outros') + '</span></div>'
      ).join('')
    : '<p style="color:var(--accent-2);font-size:12px">✅ Todos têm clientes!</p>';
}

function openLicenseesModal(sig, name, stData) {
  const rows = stData.licenciados;
  const ativos  = rows.filter(l => l.status !== 'inativo').length;
  const semCli  = rows.filter(l => l.clientCount === 0).length;
  openModal(
    '🏢 ' + name + ' (' + sig + ')',
    [
      { val: rows.length, label: 'Licenciados' },
      { val: ativos, label: 'Ativos', color: 'var(--accent-2)' },
      { val: semCli, label: 'Sem clientes', color: 'var(--alert)' },
    ],
    rows,
    [
      { key: 'codigo', label: 'Código' },
      { key: 'nome', label: 'Nome' },
      { key: 'cidade', label: 'Cidade' },
      { key: 'tipoLicenca', label: 'Licença', render: r => '<span class="badge badge-' + (LICENSE_MAP[r.tipoLicenca]?.cls||'other') + '">' + (LICENSE_MAP[r.tipoLicenca]?.label||'Outros') + '</span>' },
      { key: 'status', label: 'Status', render: r => '<span class="badge badge-' + (r.status === 'inativo' ? 'cancelado' : 'ativo') + '">' + r.status + '</span>' },
      { key: 'clientCount', label: 'Clientes' },
      { key: 'dataAtivo', label: 'Data Ativo' },
    ]
  );
}

/* ================================================================
   TELA 2 — CLIENTES BR
================================================================ */
function initClients() {
  document.getElementById('emptyClients').style.display = 'none';
  document.getElementById('dashClients').style.display  = 'block';

  makePills('pillsClients', [
    { val:'all', label:'Todos', cls:'pill-all' },
    { val:'ativo', label:'Ativos', cls:'pill-ativo' },
    { val:'cancelado', label:'Cancelados', cls:'pill-cancelado' },
  ], v => { F.C.status = v; renderClients(); });

  makePills('pillsPort', [
    { val:'all',   label:'Todos',  cls:'pill-all' },
    { val:'Esim',  label:'e-SIM',  cls:'pill-port' },
    { val:'físico',label:'Físico', cls:'pill-port' },
  ], v => { F.C.port = v; renderClients(); });

  renderClients();
}

function renderClients() {
  let data = clientes.filter(c => !c.isUSA);
  if (F.C.status !== 'all') data = data.filter(c => c.status === F.C.status);
  if (F.C.port   !== 'all') data = data.filter(c => c.tipoLinha.toLowerCase() === F.C.port.toLowerCase());

  const byState = {};
  data.forEach(c => {
    const s = c.estado || 'XX';
    if (!byState[s]) byState[s] = { val: 0, ativos: 0, cancelados: 0, port: 0, clientes: [] };
    byState[s].val++;
    if (c.status === 'ativo') byState[s].ativos++;
    else byState[s].cancelados++;
    if (c.isPort) byState[s].port++;
    byState[s].clientes.push(c);
  });

  const maxVal = Math.max(...Object.values(byState).map(v => v.val), 1);
  const ativos = data.filter(c => c.status === 'ativo').length;
  const cancelados = data.length - ativos;
  const portCount = data.filter(c => c.isPort).length;

  renderKPIs('kpiClients', [
    { val: data.length, label: 'Clientes', cls: 'kpi-accent' },
    { val: ativos, label: 'Ativos', cls: 'kpi-green' },
    { val: cancelados, label: 'Cancelados', cls: 'kpi-alert' },
    { val: portCount, label: 'Portabilidade', cls: 'kpi-purple' },
  ]);

  if (!geoBR) { setTimeout(renderClients, 600); return; }

  renderSVGMap({
    containerId: 'mapClients', geoFeatures: geoBR.features,
    W: SVG_W, H: SVG_H, padVal: PAD,
    byState, maxVal,
    labelKey: 'sigla', nameKey: 'name',
    legendId: 'legendClients',
    tooltipContent: (name, sig, st) => {
      const semGreen = ESTADOS_SEM_GREEN[sig];
      return '<div class="tooltip-state">' + name + ' (' + sig + ')'
        + (semGreen ? ' <span style="color:#ff0000c8;font-size:16px">⚡ Sem Green</span>' : '') + '</div>'
        + '<div class="tooltip-row"><span>Total:</span><strong>' + (st.val||0) + '</strong></div>'
        + '<div class="tooltip-row"><span>Ativos:</span><strong>' + (st.ativos||0) + '</strong></div>'
        + '<div class="tooltip-row"><span>Cancelados:</span><strong>' + (st.cancelados||0) + '</strong></div>'
        + '<div class="tooltip-row"><span>Portabilidade:</span><strong>' + (st.port||0) + '</strong></div>'
        + (semGreen
          ? '<div style="margin-top:6px;padding-top:6px;border-top:1px solid #30363d;color:#ff0000c8;font-size:14px">⚡ Sem conexão Green<br>Distribuidora: <strong>' + semGreen + '</strong></div>'
          : '');
    },
    onClickState: (sig, name) => openClientsModal(sig, name, byState[sig] || { clientes: [] }),
  });

  // Destaca estados sem Green com borda tracejada vermelha
  setTimeout(() => {
    const container = document.getElementById('mapClients');
    if (!container) return;
    container.querySelectorAll('.state-group').forEach(g => {
      if (ESTADOS_SEM_GREEN[g.dataset.sigla]) {
        const path = g.querySelector('.state-path');
        if (path) {
          path.style.stroke = '#ff00006c';
          path.style.strokeWidth = '1.8';
        }
      }
    });
  }, 150);

  const sorted = Object.entries(byState).filter(([s]) => s !== 'XX').sort((a,b) => b[1].val - a[1].val);
  renderRanking('rankClients', sorted.map(([s,v]) => [s, v.val]), maxVal);

  // Ativos vs Cancelados breakdown
  const totalAll = data.length || 1;
  document.getElementById('statusBreakdown').innerHTML = [
    { label: 'Ativos',       val: ativos,    pct: (ativos/totalAll*100).toFixed(1),    color: 'var(--accent-2)' },
    { label: 'Cancelados',   val: cancelados, pct: (cancelados/totalAll*100).toFixed(1), color: 'var(--alert)' },
    { label: 'Portabilidade',val: portCount,  pct: (portCount/totalAll*100).toFixed(1), color: 'var(--warn)' },
  ].map(b =>
    '<div class="breakdown-item"><div class="breakdown-dot" style="background:' + b.color + '"></div>'
    + '<span class="breakdown-label">' + b.label + '</span>'
    + '<span class="breakdown-val">' + b.val.toLocaleString('pt-BR') + '</span>'
    + '<span class="breakdown-pct">(' + b.pct + '%)</span></div>'
  ).join('');
}

function openClientsModal(sig, name, stData) {
  const rows = stData.clientes || [];
  openModal(
    '👥 Clientes — ' + name + ' (' + sig + ')',
    [
      { val: rows.length, label: 'Total' },
      { val: rows.filter(c=>c.status==='ativo').length, label: 'Ativos', color: 'var(--accent-2)' },
      { val: rows.filter(c=>c.status==='cancelado').length, label: 'Cancelados', color: 'var(--alert)' },
      { val: rows.filter(c=>c.isPort).length, label: 'Portabilidade', color: 'var(--warn)' },
    ],
    rows,
    [
      { key: 'nome', label: 'Nome' },
      { key: 'nomeLicenciado', label: 'Licenciado' },
      { key: 'plano', label: 'Plano' },
      { key: 'tipoLinha', label: 'Tipo Linha' },
      { key: 'status', label: 'Status', render: r => '<span class="badge badge-' + r.status + '">' + r.status + '</span>' },
      { key: 'isPort', label: 'Portabilidade', render: r => r.isPort ? '<span class="badge badge-port">Sim</span>' : 'Não' },
      { key: 'dataAtivo', label: 'Ativo em' },
      { key: 'dataCancelado', label: 'Cancelado em', render: r => r.dataCancelado || '—' },
    ]
  );
}

/* ================================================================
   TELA 3 — ATUAÇÃO CRUZADA
================================================================ */
let allLicenseesForCross = [];
let crossSearchStr = '';

function initCross() {
  document.getElementById('emptyCross').style.display = 'none';
  document.getElementById('dashCross').style.display  = 'block';

  // Apenas licenciados BR com UF definida
  allLicenseesForCross = licenciados.filter(l => !l.isUSA && l.uf);

  // KPIs gerais de cruzamento
  const crossed = allLicenseesForCross.filter(l => {
    // tem clientes em pelo menos 1 estado diferente do seu
    const cliStates = clientes.filter(c => c.idLicenciado === l.codigo && c.estado && c.estado !== l.uf);
    return cliStates.length > 0;
  });
  renderKPIs('kpiCross', [
    { val: allLicenseesForCross.length, label: 'Licenciados BR', cls: 'kpi-accent' },
    { val: crossed.length, label: 'Atuam fora do estado', cls: 'kpi-orange' },
    { val: allLicenseesForCross.length - crossed.length, label: 'Só no estado de origem', cls: 'kpi-green' },
  ]);

  renderLicenseePicker('');

  document.getElementById('crossSearch').addEventListener('input', function () {
    renderLicenseePicker(this.value);
  });
}

function renderLicenseePicker(search) {
  crossSearchStr = search.toLowerCase();
  const picker = document.getElementById('licenseePicker');
  const filtered = allLicenseesForCross.filter(l =>
    l.nome.toLowerCase().includes(crossSearchStr) || (l.uf || '').toLowerCase().includes(crossSearchStr)
  );
  picker.innerHTML = filtered.slice(0, 150).map(l =>
    '<div class="licensee-item" data-id="' + l.codigo + '">'
    + '<div>' + l.nome + '</div>'
    + '<div class="li-state">' + l.uf + ' · ' + (LICENSE_MAP[l.tipoLicenca]?.label || 'Outros') + '</div>'
    + '</div>'
  ).join('');
  picker.querySelectorAll('.licensee-item').forEach(el => {
    el.addEventListener('click', function () {
      picker.querySelectorAll('.licensee-item').forEach(e => e.classList.remove('selected'));
      this.classList.add('selected');
      const lic = allLicenseesForCross.find(l => l.codigo === this.dataset.id);
      if (lic) renderCrossMap(lic);
    });
  });
}

function renderCrossMap(lic) {
  // Clientes deste licenciado
  const myClients = clientes.filter(c => c.idLicenciado === lic.codigo && c.estado);
  const byState = {};
  myClients.forEach(c => {
    const s = c.estado;
    if (!byState[s]) byState[s] = { val: 0, ativos: 0, cancelados: 0, isOrigin: s === lic.uf };
    byState[s].val++;
    if (c.status === 'ativo') byState[s].ativos++; else byState[s].cancelados++;
  });

  const maxVal = Math.max(...Object.values(byState).map(v => v.val), 1);
  document.getElementById('crossMapTitle').textContent = lic.nome + ' — ' + (lic.uf || '?');

  // Info de origem
  document.getElementById('crossOriginCard').style.display = 'block';
  const originSt = byState[lic.uf];
  document.getElementById('crossOriginInfo').innerHTML =
    '<div class="cross-origin-badge">' + lic.uf + ' — ' + (getStateName(lic.uf) || lic.uf) + '</div>'
    + '<div style="font-size:12px;color:var(--text-muted);margin-top:4px">Clientes no estado de origem: <strong style="color:var(--text)">' + (originSt?.val || 0) + '</strong></div>'
    + '<div style="font-size:12px;color:var(--text-muted)">Clientes fora do estado: <strong style="color:var(--warn)">' + myClients.filter(c => c.estado !== lic.uf).length + '</strong></div>'
    + '<div style="font-size:12px;color:var(--text-muted);margin-top:6px">Licença: <span class="badge badge-' + (LICENSE_MAP[lic.tipoLicenca]?.cls||'other') + '">' + (LICENSE_MAP[lic.tipoLicenca]?.label||'Outros') + '</span></div>';

  if (!geoBR) { setTimeout(() => renderCrossMap(lic), 600); return; }

  // Marca o estado de origem diferente
  renderSVGMap({
    containerId: 'mapCross', geoFeatures: geoBR.features,
    W: SVG_W, H: SVG_H, padVal: PAD,
    byState, maxVal,
    labelKey: 'sigla', nameKey: 'name',
    legendId: 'legendCross',
    tooltipContent: (name, sig, st) =>
      '<div class="tooltip-state">' + name + ' (' + sig + ')'
      + (sig === lic.uf ? ' 📍 Origem' : '') + '</div>'
      + '<div class="tooltip-row"><span>Clientes:</span><strong>' + (st.val||0) + '</strong></div>'
      + '<div class="tooltip-row"><span>Ativos:</span><strong>' + (st.ativos||0) + '</strong></div>'
      + '<div class="tooltip-row"><span>Cancelados:</span><strong>' + (st.cancelados||0) + '</strong></div>',
    onClickState: null,
  });

  // Destaca estado de origem em amarelo via DOM depois do render
  setTimeout(() => {
    const container = document.getElementById('mapCross');
    container.querySelectorAll('.state-group').forEach(g => {
      if (g.dataset.sigla === lic.uf) {
        g.querySelector('.state-path').style.stroke = 'var(--warn)';
        g.querySelector('.state-path').style.strokeWidth = '2';
      }
    });
  }, 50);

  // Rankinging
  const sorted = Object.entries(byState).sort((a,b) => b[1].val - a[1].val);
  renderRanking('rankCross', sorted.map(([s,v]) => [s + (s===lic.uf?' 📍':''), v.val]), maxVal);
}

/* ================================================================
   TELA 4 — EUA
================================================================ */
function initUSA() {
  document.getElementById('emptyUSA').style.display = 'none';
  document.getElementById('dashUSA').style.display  = 'block';
  renderUSA();
}

function renderUSA() {
  // Licenciados: aba Licenciados com Origem=EXPANSAO_EUA
  const licUSA = licenciados.filter(l => l.isUSA);
  // Clientes: todos da aba Internacional
  const cliUSA = internacional;

  // Mapa de CLIENTES — agrupado por estado da aba Internacional
  const byState = {};
  cliUSA.forEach(c => {
    const s = c.usaState || c.estado || 'XX';
    if (!byState[s]) byState[s] = { val: 0, clientList: [], ativos: 0, cancelados: 0 };
    byState[s].val++;
    byState[s].clientList.push(c);
    if (c.status === 'ativo') byState[s].ativos++;
    else byState[s].cancelados++;
  });

  const maxVal = Math.max(...Object.values(byState).map(v => v.val), 1);
  const ativos = cliUSA.filter(c => c.status === 'ativo').length;

  renderKPIs('kpiUSA', [
    { val: cliUSA.length, label: 'Clientes EUA', cls: 'kpi-accent' },
    { val: ativos, label: 'Ativos', cls: 'kpi-green' },
    { val: cliUSA.length - ativos, label: 'Cancelados', cls: 'kpi-alert' },
  ]);

  if (!geoUSA) { setTimeout(renderUSA, 600); return; }

  const featuresUSA = geoUSA.features.filter(f => f.properties.name !== 'Puerto Rico');

  featuresUSA.forEach(f => {
    const nomeFull = f.properties.name || '';
    const sigla = USA_NAME_TO_SIGLA[nomeFull.toLowerCase()];
    f.properties._sigla = sigla || nomeFull.substring(0, 2).toUpperCase();
  });

  renderSVGMap({
    containerId: 'mapUSA', geoFeatures: featuresUSA,
    W: USA_W, H: USA_H, padVal: USA_PAD,
    byState: byState, maxVal,
    labelKey: '_sigla', nameKey: 'name',
    legendId: 'legendUSA',
    tooltipContent: (name, sig, st) =>
      '<div class="tooltip-state">' + name + '</div>'
      + '<div class="tooltip-row"><span>Clientes:</span><strong>' + (st.val||0) + '</strong></div>'
      + '<div class="tooltip-row"><span>Ativos:</span><strong>' + (st.ativos||0) + '</strong></div>'
      + '<div class="tooltip-row"><span>Cancelados:</span><strong>' + (st.cancelados||0) + '</strong></div>',
    onClickState: (sig, name) => {
      const stCli = cliUSA.filter(c => (c.usaState === sig) || (c.estado === sig));
      openModal(
        '🇺🇸 ' + name,
        [
          { val: stCli.length, label: 'Clientes' },
          { val: stCli.filter(c=>c.status==='ativo').length, label: 'Ativos', color: 'var(--accent-2)' },
          { val: stCli.filter(c=>c.status==='cancelado').length, label: 'Cancelados', color: 'var(--alert)' },
        ],
        stCli,
        [
          { key: 'codigo', label: 'ID Cliente' },
          { key: 'nome', label: 'Nome' },
          { key: 'estado', label: 'Estado' },
          { key: 'plano', label: 'Plano' },
          { key: 'tipoLinha', label: 'Tipo Linha' },
          { key: 'status', label: 'Status', render: r => '<span class="badge badge-' + r.status + '">' + r.status + '</span>' },
          { key: 'idLicenciado', label: 'ID Licenciado' },
          { key: 'nomeLicenciado', label: 'Licenciado' },
        ]
      );
    },
  });

  // Rankinging por clientes
  const sorted = Object.entries(byState).filter(([s]) => s !== 'XX').sort((a,b) => b[1].val - a[1].val);
  renderRanking('rankUSA', sorted.map(([s,v]) => [s, v.val]), maxVal);

  // Breakdown por tipo de licença dos clientes
  const licTypes = Object.entries(
    cliUSA.reduce((acc, c) => { acc[c.tipoLicenca] = (acc[c.tipoLicenca]||0) + 1; return acc; }, {})
  );
  document.getElementById('usaBreakdown').innerHTML = licTypes.map(([t, n]) => {
    const lt = LICENSE_MAP[t] || LICENSE_MAP.other;
    return '<div class="breakdown-item"><div class="breakdown-dot" style="background:' + lt.color + '"></div>'
      + '<span class="breakdown-label">' + lt.label + '</span>'
      + '<span class="breakdown-val">' + n + '</span></div>';
  }).join('') || '<p style="color:var(--text-muted);font-size:12px">Nenhum dado</p>';
}

/* ================================================================
   TELA 5 — LICENCIADOS EUA
================================================================ */
function initUSALic() {
  document.getElementById('emptyUSALic').style.display = 'none';
  document.getElementById('dashUSALic').style.display  = 'block';

  makePills('pillsUSALic', [
    { val:'all',      label:'Todos',           cls:'pill-all' },
    { val:'telecom',  label:'Telecom',          cls:'pill-telecom' },
    { val:'expert',   label:'Expert + Telecom', cls:'pill-expert' },
    { val:'expansao', label:'Expansão USA',     cls:'pill-expansao' },
  ], v => { F.UL.license = v; renderUSALic(); });

  makePills('pillsStatusUSALic', [
    { val:'all',    label:'Todos',    cls:'pill-all' },
    { val:'ativo',  label:'Ativos',   cls:'pill-ativo' },
    { val:'inativo',label:'Inativos', cls:'pill-cancelado' },
  ], v => { F.UL.status = v; renderUSALic(); });

  renderUSALic();
}

function renderUSALic() {
  // Licenciados EUA: vem da aba "Licenciados internacional"
  let data = licInternacional.slice();
  if (F.UL.license !== 'all') data = data.filter(l => l.tipoLicenca === F.UL.license);
  if (F.UL.status  !== 'all') data = data.filter(l => l.status === F.UL.status);

  // Agrega por estado americano usando coluna Estado da aba Licenciados internacional
  const byState = {};
  data.forEach(l => {
    const s = l.usaState || l.estado || 'XX';
    if (!byState[s]) byState[s] = { val: 0, licenciados: [], ativos: 0, inativos: 0 };
    byState[s].val++;
    byState[s].licenciados.push(l);
    if (l.status === 'ativo') byState[s].ativos++;
    else byState[s].inativos++;
  });

  const maxVal = Math.max(...Object.values(byState).map(v => v.val), 1);
  const ativos = data.filter(l => l.status === 'ativo').length;

  renderKPIs('kpiUSALic', [
    { val: data.length, label: 'Licenciados EUA', cls: 'kpi-accent' },
    { val: data.filter(l=>l.tipoLicenca==='expert').length,  label: 'Expert+Telecom', cls: 'kpi-purple' },
  ]);

  if (!geoNAm) { setTimeout(renderUSALic, 600); return; }

  // Usa GeoJSON da América do Norte (EUA + Canadá + México)
  const featuresNAm = geoNAm.features;
  featuresNAm.forEach(f => {
    const nomeFull = (f.properties.name || '').toLowerCase();
    // Para EUA: usa sigla do estado (mapa de estados já está em geoUSA)
    // Para Canadá e México: exibe o nome do país
    const sigla = nomeFull === 'united states of america' ? 'EUA'
                : nomeFull === 'canada' ? 'CA'
                : nomeFull === 'mexico' ? 'MX'
                : f.properties.name.substring(0, 2).toUpperCase();
    f.properties._sigla = sigla;
  });

  // Para o mapa de licenciados EUA, combina:
  // - GeoJSON dos estados americanos (para detalhe por estado)
  // - GeoJSON da América do Norte (para o contorno completo)
  const featuresEstados = geoUSA ? geoUSA.features.filter(f => f.properties.name !== 'Puerto Rico') : [];
  featuresEstados.forEach(f => {
    const nomeFull = f.properties.name || '';
    const sigla = USA_NAME_TO_SIGLA[nomeFull.toLowerCase()];
    f.properties._sigla = sigla || nomeFull.substring(0, 2).toUpperCase();
  });

  // Adiciona Canadá e México como decoração (sem dados)
  const canadaMexico = featuresNAm.filter(f => {
    const n = (f.properties.name || '').toLowerCase();
    return n === 'canada' || n === 'mexico';
  });
  canadaMexico.forEach(f => {
    f.properties._sigla = (f.properties.name || '').substring(0, 2).toUpperCase();
  });

  const todasFeatures = [...featuresEstados, ...canadaMexico];

  renderSVGMap({
    containerId: 'mapUSALic', geoFeatures: todasFeatures,
    W: USA_W, H: USA_H, padVal: USA_PAD,
    byState: byState, maxVal,
    labelKey: '_sigla', nameKey: 'name',
    legendId: 'legendUSALic',
    tooltipContent: (name, sig, st) =>
      '<div class="tooltip-state">' + name + '</div>'
      + '<div class="tooltip-row"><span>Licenciados:</span><strong>' + (st.val||0) + '</strong></div>'
      + '<div class="tooltip-row"><span>Ativos:</span><strong>' + (st.ativos||0) + '</strong></div>'
      + '<div class="tooltip-row"><span>Inativos:</span><strong>' + (st.inativos||0) + '</strong></div>',
    onClickState: (sig, name) => {
      const stLic = (byState[sig] || {}).licenciados || [];
      openModal(
        '🇺🇸 Licenciados — ' + name,
        [
          { val: stLic.length, label: 'Licenciados' },
          { val: stLic.filter(l=>l.status==='ativo').length, label: 'Ativos', color: 'var(--accent-2)' },
          { val: stLic.filter(l=>l.status!=='ativo').length, label: 'Inativos', color: 'var(--alert)' },
        ],
        stLic,
        [
          { key: 'codigo',      label: 'Código' },
          { key: 'nome',        label: 'Nome' },
          { key: 'celular',     label: 'Celular' },
          { key: 'estado',      label: 'Estado' },
          { key: 'tipoLicenca', label: 'Licença', render: r => '<span class="badge badge-' + (LICENSE_MAP[r.tipoLicenca]?.cls||'other') + '">' + (LICENSE_MAP[r.tipoLicenca]?.label||'Outros') + '</span>' },
          { key: 'status',      label: 'Status', render: r => '<span class="badge badge-' + (r.status==='ativo'?'ativo':'cancelado') + '">' + r.status + '</span>' },
          { key: 'patrocinador',label: 'Patrocinador' },
          { key: 'dataAtivo',   label: 'Data Ativo' },
        ]
      );
    },
  });

  const sorted = Object.entries(byState).filter(([s]) => s !== 'XX').sort((a,b) => b[1].val - a[1].val);
  renderRanking('rankUSALic', sorted.map(([s,v]) => [s, v.val]), maxVal);

  const licTypes = Object.entries(
    data.reduce((acc, l) => { acc[l.tipoLicenca] = (acc[l.tipoLicenca]||0) + 1; return acc; }, {})
  );
  document.getElementById('usaLicBreakdown').innerHTML = licTypes.map(([t, n]) => {
    const lt = LICENSE_MAP[t] || LICENSE_MAP.other;
    return '<div class="breakdown-item"><div class="breakdown-dot" style="background:' + lt.color + '"></div>'
      + '<span class="breakdown-label">' + lt.label + '</span>'
      + '<span class="breakdown-val">' + n + '</span></div>';
  }).join('') || '<p style="color:var(--text-muted);font-size:12px">Nenhum dado</p>';
}

function initAllTabs() {
  initLicensees();
  initClients();
  initCross();
  initUSA();
  initUSALic();
}