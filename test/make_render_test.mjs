// Gera um HTML de teste visual: ECharts via CDN + app + dados reais pré-injetados.
import XLSX from '../vendor/xlsx.full.min.js';
import fs from 'node:fs';

const css   = fs.readFileSync(new URL('../src/styles.css', import.meta.url),'utf8');
const icons = fs.readFileSync(new URL('../src/icons.js', import.meta.url),'utf8');
let app     = fs.readFileSync(new URL('../src/app.js', import.meta.url),'utf8');
// placeholders leves (quadrados coloridos) só para o teste visual de layout
const brasao= 'data:image/svg+xml;base64,'+Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="80" height="80"><rect width="80" height="80" fill="#123a8f"/><text x="40" y="46" font-size="26" fill="#fff" text-anchor="middle">PET</text></svg>').toString('base64');
const saude = 'data:image/svg+xml;base64,'+Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="140" height="60"><rect width="140" height="60" fill="#2f9e5f"/><text x="70" y="36" font-size="18" fill="#fff" text-anchor="middle">SAUDE</text></svg>').toString('base64');

// auto-start: depois de carregar, parseia um workbook pré-embutido? Melhor: expõe hook e injeta dados.
app = app.replace(
  "document.addEventListener('DOMContentLoaded', initUpload);",
  "document.addEventListener('DOMContentLoaded', () => { initUpload(); window.__render = (wbData)=>{ const wb = XLSX.read(Uint8Array.from(atob(wbData), c=>c.charCodeAt(0)), {type:'array'}); const data = parseWorkbook(wb); computeMetrics(data); renderValidation(data); showApp(data); }; });"
);

const wbB64 = fs.readFileSync(new URL('../data/base_final.xlsx', import.meta.url)).toString('base64');

const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>teste</title><style>${css}</style>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.1/dist/echarts.min.js"></script>
</head><body>
<div id="tela-upload"><div class="upload-card">
<div class="upload-logos"><img src="${brasao}"><div class="sep"></div><img src="${saude}"></div>
<h1>Painel de Produção da Rede CAPS</h1><div class="sub">Secretaria de Saúde de Petrolina</div>
<label class="dropzone" id="dropzone"><div class="ico"></div><div class="t1">Clique ou arraste a planilha</div><div class="t2">.xlsx</div><input type="file" id="file-input" accept=".xlsx"></label>
<div class="privacy">Processamento 100% local.</div>
<div class="validacao" id="validacao"></div>
</div></div>
<div id="tela-app"></div>
<script>const LOGO={brasao:"${brasao}",saude:"${saude}"};</script>
<script>${icons}</script>
<script>${app}</script>
<script>window.__WB="${wbB64}";
function tryRender(){ if(window.__render && window.XLSX && window.echarts){ window.__render(window.__WB); } else { setTimeout(tryRender, 150); } }
window.addEventListener('load', ()=> setTimeout(tryRender, 300));
</script>
</body></html>`;

fs.writeFileSync(new URL('./render-test.html', import.meta.url), html);
console.log('render-test.html gerado:', (html.length/1024).toFixed(0),'KB');
