// Gera test/preview_full.html: CSS real + app real + XLSX/ECharts via CDN,
// com auto-render da planilha real embutida em base64. Serve para screenshot.
import fs from 'node:fs';
const css=fs.readFileSync(new URL('../src/styles.css',import.meta.url),'utf8');
const icons=fs.readFileSync(new URL('../src/icons.js',import.meta.url),'utf8');
let app=fs.readFileSync(new URL('../src/app.js',import.meta.url),'utf8');
const brasao='data:image/png;base64,'+fs.readFileSync(new URL('../assets/image2.png',import.meta.url)).toString('base64');
const saude='data:image/png;base64,'+fs.readFileSync(new URL('../assets/image1.png',import.meta.url)).toString('base64');
const wb=fs.readFileSync(new URL('../data/base_final.xlsx',import.meta.url)).toString('base64');

app=app.replace("document.addEventListener('DOMContentLoaded', initUpload);",
  `document.addEventListener('DOMContentLoaded', () => {
     initUpload();
     const wb = XLSX.read(Uint8Array.from(atob(window.__WB), c=>c.charCodeAt(0)), {type:'array'});
     const data = parseWorkbook(wb); computeMetrics(data); renderValidation(data); showApp(data);
   });`);

const html=`<!DOCTYPE html><html lang=pt-BR><head><meta charset=UTF-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<style>${css}</style>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.1/dist/echarts.min.js"></script>
</head><body>
<div id=tela-upload><div class=upload-card>
<div class=upload-logos><img src="${brasao}"><div class=sep></div><img src="${saude}"></div>
<h1>Painel de Produção da Rede CAPS</h1><div class=sub>Secretaria de Saúde de Petrolina</div>
<label class=dropzone id=dropzone><div class=ico></div><div class=t1>Clique ou arraste a planilha</div><div class=t2>.xlsx</div><input type=file id=file-input accept=".xlsx"></label>
<div class=privacy>Processamento local.</div><div class=validacao id=validacao></div>
</div><button class="tema-btn tema-btn-upload" id=tema-btn-upload></button></div>
<div id=tela-app></div>
<script>const LOGO={brasao:"${brasao}",saude:"${saude}"};window.__WB="${wb}";</script>
<script>${icons}</script>
<script>${app}</script>
</body></html>`;
fs.writeFileSync(new URL('./preview_full.html',import.meta.url),html);
console.log('preview_full.html',(html.length/1024).toFixed(0),'KB');
