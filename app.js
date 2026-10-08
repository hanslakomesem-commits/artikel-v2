import {readDocx,buildArticleDocx,sourceFootnotes} from './docx-engine.js';
import {analyzeThesis,generatePlan,diagnostics,helpers} from './article-ai.js';
import {translateAbstractLocal,abstractSource,localTranslatorInfo} from './local-translator.js';

const $=s=>document.querySelector(s);
let state={file:null,doc:null,analysis:null,plan:null,blob:null,englishAbstract:null,translating:false,mapping:{intro:'auto',method:'auto',discussion:'auto',conclusion:'auto'}};
const app=$('#app');
app.innerHTML=`<div class="shell">
  <div class="top"><div class="brand"><h1>ZAIN.NET — Skripsi Jadi Artikel</h1><p>Scrib Article AI Lokal V1.2 • detektor universal skripsi • tanpa API token</p></div><div class="badge">AI LOKAL • NO API TOKEN</div></div>
  <div class="grid">
   <aside class="panel"><h2>1. Upload & Pengaturan</h2><div class="pad">
    <label class="drop" id="drop"><strong>Upload Skripsi Utuh (.DOCX)</strong><small>Klik atau seret file Word ke sini</small><input id="file" type="file" accept=".docx"></label>
    <div class="field"><label>Mode Artikel</label><select id="mode"><option value="contoh">Mirip Contoh Anda (~3.300 kata)</option><option value="ringkas">Ringkas (~2.300 kata)</option><option value="lengkap">Lengkap (~5.000+ kata)</option></select></div>
    <div class="field"><label>Nama Penulis</label><input id="author" placeholder="Deteksi otomatis"></div>
    <div class="field"><label>Program Studi</label><input id="prodi" placeholder="Deteksi otomatis"></div>
    <div class="field"><label>Universitas</label><input id="univ" placeholder="Deteksi otomatis"></div>
    <div class="mapBox" id="mapBox">
      <div class="mapTitle">Pemetaan Struktur Skripsi</div>
      <div class="mapHint">V1.2 mendeteksi BAB berdasarkan urutan, isi, style Word, dan kata kunci. Jika skripsi kampus lain memakai nama BAB berbeda, pilih sumbernya manual di sini.</div>
      <div class="field"><label>Pendahuluan</label><select id="mapIntro" disabled><option value="auto">AUTO — deteksi terbaik</option></select></div>
      <div class="field"><label>Metode Penelitian</label><select id="mapMethod" disabled><option value="auto">AUTO — deteksi terbaik</option></select></div>
      <div class="field"><label>Hasil / Pembahasan</label><select id="mapDiscussion" disabled><option value="auto">AUTO — deteksi terbaik</option></select></div>
      <div class="field"><label>Kesimpulan / Saran</label><select id="mapConclusion" disabled><option value="auto">AUTO — deteksi terbaik</option></select></div>
      <button class="btn ghost" id="applyMap" disabled>Terapkan Pemetaan & Buat Ulang</button>
      <div id="mapSummary" class="mapSummary">Belum dianalisis.</div>
    </div>
    <div class="checks">
      <label><input type="checkbox" checked disabled> Abstrak Indonesia</label><label><input type="checkbox" id="useEnglish" checked> Abstract English AI Lokal</label>
      <label><input type="checkbox" checked disabled> Pendahuluan</label><label><input type="checkbox" checked disabled> Metode</label>
      <label><input type="checkbox" checked disabled> Pembahasan</label><label><input type="checkbox" checked disabled> Kesimpulan/Saran</label>
      <label><input type="checkbox" checked disabled> Daftar Pustaka</label>
    </div>
    <div class="aiBox">
      <div class="aiTitle">Abstract Bahasa Inggris</div>
      <div class="aiHint">Diterjemahkan di browser dengan <b>${localTranslatorInfo.model}</b>. Pertama kali model perlu diunduh lalu disimpan di cache browser.</div>
      <button class="btn ghost" id="translate" disabled>Terjemahkan Ulang dengan AI Lokal</button>
      <div class="field"><label>English Abstract (bisa diedit)</label><textarea id="englishAbstract" rows="8" placeholder="Akan dibuat otomatis setelah analisis..."></textarea></div>
      <div class="field"><label>Keywords</label><input id="englishKeywords" placeholder="Akan diterjemahkan otomatis"></div>
    </div>
    <div class="btnrow"><button class="btn primary" id="analyze" disabled>Analisis & Buat Draft</button><button class="btn ghost" id="regen" disabled>Buat Ulang</button></div>
    <div class="progress"><i id="bar"></i></div><div id="msg"></div>
   </div></aside>
   <main class="panel"><h2>2. Hasil Scrib Article AI Lokal</h2><div class="pad">
    <div class="stats"><div class="stat"><b id="sWords">0</b><span>Kata Skripsi</span></div><div class="stat"><b id="aWords">0</b><span>Kata Draft</span></div><div class="stat"><b id="conf">0%</b><span>Deteksi Struktur</span></div><div class="stat"><b id="refs">0</b><span>Referensi Dipilih</span></div></div>
    <div class="tabs"><button class="tab on" data-tab="review">Review Sumber</button><button class="tab" data-tab="preview">Preview Artikel</button><button class="tab" data-tab="info">Cara Kerja AI Lokal</button></div>
    <div id="review" class="tabpane"><div id="diag"></div><div id="sections" class="sections"><div class="notice">Upload skripsi untuk memulai.</div></div></div>
    <div id="preview" class="tabpane hidden"><div class="previewPaper" id="paper"></div></div>
    <div id="info" class="tabpane hidden"><div class="notice"><b>Scrib Article AI Lokal memakai dua mesin lokal.</b><br>Mesin pertama adalah Scrib Universal Structure Detector: membaca BAB, style Word, isi paragraf, urutan BAB, dan variasi nama bagian lalu memilih paragraf penting secara ekstraktif. Mesin kedua menerjemahkan Abstrak Indonesia menjadi English Abstract langsung di browser menggunakan model ONNX/Transformers.js. Tidak memakai API token. Teks skripsi tidak dikirim ke API generatif.</div></div>
    <div class="btnrow"><button class="btn ok" id="download" disabled>Download Artikel .DOCX</button><button class="btn ghost" id="report" disabled>Download Laporan .TXT</button></div>
   </div></main>
  </div><div class="footer">ZAIN.NET • Scrib Universal Structure Detector V1.2 • DOCX diproses di perangkat pengguna</div></div>`;

const fileInput=$('#file'),drop=$('#drop'),bar=$('#bar'),msg=$('#msg');
function setMsg(t,type=''){msg.innerHTML=t?`<div class="notice ${type}">${t}</div>`:'';}
function progress(n){bar.style.width=n+'%'}
function esc(s){return (s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function useFile(f){if(!f||!f.name.toLowerCase().endsWith('.docx')){setMsg('Pilih file .DOCX.', 'warn');return;}state.file=f;state.englishAbstract=null;$('#englishAbstract').value='';$('#englishKeywords').value='';$('#analyze').disabled=false;setMsg(`<b>${esc(f.name)}</b> siap dianalisis.`,'ok');}
fileInput.onchange=e=>useFile(e.target.files[0]);drop.onclick=e=>{if(e.target!==fileInput)fileInput.click()};
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.style.borderColor='#38bdf8'}));['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.style.borderColor=''}));drop.addEventListener('drop',e=>useFile(e.dataTransfer.files[0]));

function mappingNow(){
  return {intro:$('#mapIntro').value||'auto',method:$('#mapMethod').value||'auto',discussion:$('#mapDiscussion').value||'auto',conclusion:$('#mapConclusion').value||'auto'};
}
function roleLabel(r){return {intro:'Pendahuluan',method:'Metode',discussion:'Hasil/Pembahasan',conclusion:'Kesimpulan/Saran'}[r]||r;}
function populateMapping(a){
  const ids={intro:'#mapIntro',method:'#mapMethod',discussion:'#mapDiscussion',conclusion:'#mapConclusion'};
  const blocks=(a.blocks||[]).filter(b=>b.wordCount>=40);
  for(const [role,sel] of Object.entries(ids)){
    const el=$(sel);el.innerHTML='<option value="auto">AUTO — deteksi terbaik</option>'+blocks.map(b=>`<option value="${esc(b.id)}">${esc(b.label)} — ${b.wordCount.toLocaleString('id-ID')} kata${b.role===role?' ✓':''}</option>`).join('');
    el.disabled=false; el.value=state.mapping[role]||'auto';
  }
  $('#applyMap').disabled=false;
  const auto=a.roleBlocks||{};
  $('#mapSummary').innerHTML=['intro','method','discussion','conclusion'].map(r=>{
    const b=auto[r];return `<div><b>${roleLabel(r)}:</b> ${b?esc(b.label)+' <span>('+b.wordCount.toLocaleString('id-ID')+' kata)</span>':'<em>belum ditemukan</em>'}</div>`;
  }).join('');
}
function planWithCurrentMapping(){
  state.mapping=mappingNow();
  return generatePlan(state.doc,state.analysis,{mode:$('#mode').value,mapping:state.mapping});
}
function currentEnglish(){
  const body=$('#englishAbstract').value.trim(),keywords=$('#englishKeywords').value.trim();
  return body?{body,keywords,model:state.englishAbstract?.model||'manual/local'}:null;
}
async function rebuildBlob(){
  if(!state.doc||!state.plan)return;
  const english=$('#useEnglish').checked?currentEnglish():null;
  state.blob=await buildArticleDocx(state.doc,state.analysis,state.plan,{author:$('#author').value,prodi:$('#prodi').value,univ:$('#univ').value},{englishAbstract:english});
  $('#download').disabled=$('#useEnglish').checked&&!english;
}
async function doTranslate(){
  if(!state.doc||!state.plan||state.translating)return false;
  state.translating=true;$('#translate').disabled=true;$('#download').disabled=true;
  try{
    const tr=await translateAbstractLocal(state.doc,state.plan,s=>setMsg(esc(s),'warn'));
    state.englishAbstract=tr;$('#englishAbstract').value=tr.body||'';$('#englishKeywords').value=tr.keywords||'';
    setMsg('Abstract bahasa Inggris selesai dibuat oleh AI lokal. Silakan review/edit bila perlu.','ok');
    render();await rebuildBlob();return true;
  }catch(e){
    console.error(e);setMsg('AI lokal gagal membuat English Abstract: '+esc(e.message||String(e))+'. Anda dapat mencoba lagi atau isi terjemahan secara manual.','warn');
    return false;
  }finally{state.translating=false;$('#translate').disabled=false;}
}

async function run(){try{
  progress(10);setMsg('Membaca seluruh struktur DOCX…');state.doc=await readDocx(state.file);progress(30);
  state.analysis=analyzeThesis(state.doc);const a=state.analysis;$('#author').value=a.meta.author||'';$('#prodi').value=a.meta.prodi||'';$('#univ').value=a.meta.univ||'';progress(48);
  setMsg('Scrib Article AI Lokal menilai relevansi paragraf dan struktur BAB…');populateMapping(a);state.plan=planWithCurrentMapping();progress(65);render();$('#translate').disabled=false;
  if($('#useEnglish').checked){progress(72);await doTranslate();progress(92);}else{state.englishAbstract=null;$('#englishAbstract').value='';$('#englishKeywords').value='';}
  await rebuildBlob();progress(100);
  if(!$('#useEnglish').checked||currentEnglish())setMsg('Draft artikel selesai. Review hasil lalu download DOCX.','ok');
  $('#report').disabled=false;$('#regen').disabled=false;$('#download').disabled=$('#useEnglish').checked&&!currentEnglish();
}catch(e){console.error(e);setMsg('Gagal: '+esc(e.message||String(e)),'warn');progress(0)}}
$('#analyze').onclick=run;
$('#translate').onclick=doTranslate;
$('#regen').onclick=async()=>{if(!state.doc)return;state.plan=planWithCurrentMapping();render();if($('#useEnglish').checked&&!currentEnglish())await doTranslate();await rebuildBlob();setMsg('Draft dibuat ulang dengan mode baru.','ok')};
$('#applyMap').onclick=async()=>{if(!state.doc)return;state.plan=planWithCurrentMapping();render();await rebuildBlob();setMsg('Pemetaan struktur diterapkan. Periksa Review Sumber sebelum download.','ok')};
['mapIntro','mapMethod','mapDiscussion','mapConclusion'].forEach(id=>$('#'+id).addEventListener('change',()=>{state.mapping=mappingNow();}));
$('#useEnglish').onchange=async()=>{if(!state.doc)return;if($('#useEnglish').checked&&!currentEnglish())await doTranslate();else await rebuildBlob();render();};
let editTimer;['englishAbstract','englishKeywords','author','prodi','univ'].forEach(id=>$('#'+id).addEventListener('input',()=>{clearTimeout(editTimer);editTimer=setTimeout(async()=>{render();await rebuildBlob();},350)}));

function render(){
 const a=state.analysis,p=state.plan,d=state.doc;if(!a||!p||!d)return;
 $('#sWords').textContent=a.wordCount.toLocaleString('id-ID');$('#aWords').textContent=p.totalWords.toLocaleString('id-ID');$('#conf').textContent=Math.round(a.confidence*100)+'%';const b=p.sections.find(x=>x.id==='biblio');$('#refs').textContent=b?b.items.length:0;
 const notes=diagnostics(d,a,p);$('#diag').innerHTML=notes.map(x=>`<div class="notice warn">${esc(x)}</div>`).join('')||'<div class="notice ok">Struktur utama terdeteksi dengan baik.</div>';
 $('#sections').innerHTML=p.sections.map(s=>`<div class="sec"><div class="secHead"><b>${esc(s.title)}</b><span>${s.items.reduce((z,x)=>z+helpers.wc(d.items[x.index]?.text||''),0)} kata • ${s.items.length} blok</span></div><div class="paras">${s.items.map(x=>`<div class="para ${x.role==='subheading'?'sub':''}">${esc(d.items[x.index]?.text||'')}<span class="trace">Sumber blok #${x.index+1} • ${esc(x.source)}</span></div>`).join('')}</div></div>`).join('');
 const m={...a.meta,author:$('#author').value||a.meta.author,prodi:$('#prodi').value||a.meta.prodi,univ:$('#univ').value||a.meta.univ};
 let html=`<h3>${esc((m.title||'ARTIKEL ILMIAH').toUpperCase())}</h3><div class="identity author"><b>${esc(m.author||'Nama Penulis')}</b></div><div class="identity"><i>${esc(m.prodi||'')}</i></div><div class="identity"><i>${esc(m.univ||'')}</i></div>`;
 for(const s of p.sections){
   const center=s.id==='abstract'||s.id==='biblio'; html+=`<h4 class="${center?'center':''}">${esc(s.title)}</h4>`;
   for(const x of s.items){const t=d.items[x.index]?.text||'';const isAbs=s.id==='abstract';const isConc=s.id==='conclusion'&&/^(?:[A-Z]\s*[.)]\s*)?(?:KESIMPULAN|SARAN)\b/i.test(t);html+=`<p class="${x.role==='subheading'?'sub ':''}${isAbs?'abstractText ':''}${isConc?'conclusionHead ':''}" style="${x.role==='bibliography'?'text-indent:-1.27cm;margin-left:1.27cm;':''}">${esc(t)}</p>`;}
   if(s.id==='abstract'&&$('#useEnglish').checked){const eng=currentEnglish();if(eng?.body){html+=`<h4 class="center">ABSTRACT</h4><p class="abstractText">${esc(eng.body)}</p>${eng.keywords?`<p class="abstractKeywords"><b>Keywords:</b> ${esc(eng.keywords)}</p>`:''}`;}else html+=`<div class="previewPending">English Abstract belum tersedia.</div>`;}
 }
 $('#paper').innerHTML=html;
}

for(const t of document.querySelectorAll('.tab'))t.onclick=()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('on'));document.querySelectorAll('.tabpane').forEach(x=>x.classList.add('hidden'));t.classList.add('on');$('#'+t.dataset.tab).classList.remove('hidden')};
function saveBlob(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},1000)}
$('#download').onclick=async()=>{if(!state.blob)return;if($('#useEnglish').checked&&!currentEnglish()){setMsg('English Abstract belum selesai. Terjemahkan dulu atau matikan opsi Abstract English.','warn');return;}await rebuildBlob();const author=($('#author').value||state.analysis.meta.author||'MAHASISWA').replace(/[^\p{L}\p{N}]+/gu,'_').replace(/^_|_$/g,'');saveBlob(state.blob,`Artikel_${author}_ZAINNET_AI_LOKAL.docx`)};
$('#report').onclick=()=>{const a=state.analysis,p=state.plan;const fns=sourceFootnotes(state.doc,p);let txt=`ZAIN.NET — LAPORAN SCRIB ARTICLE AI LOKAL\n\nFile: ${state.file.name}\nJudul: ${a.meta.title}\nPenulis: ${$('#author').value}\nKata skripsi: ${a.wordCount}\nKata draft: ${p.totalWords}\nConfidence struktur: ${Math.round(a.confidence*100)}%\nFootnote yang ikut terpakai: ${fns.length}\nEnglish Abstract: ${currentEnglish()?'YA — '+(state.englishAbstract?.model||'manual'):'TIDAK'}\n\nBAGIAN TERPILIH:\n`;for(const s of p.sections)txt+=`- ${s.title}: ${s.items.length} blok\n`;txt+='\nPEMETAAN STRUKTUR:\n';for(const r of ['intro','method','discussion','conclusion']){const b=p.usedBlocks?.[r];txt+=`- ${roleLabel(r)}: ${b?b.label+' ('+b.wordCount+' kata)':'TIDAK DITEMUKAN'}\n`;}txt+='\nFormat khusus: nama penulis bold; ABSTRAK/ABSTRACT centered; isi abstrak single spacing; Kesimpulan dan Saran bold.\n';saveBlob(new Blob([txt],{type:'text/plain;charset=utf-8'}),'Laporan_Artikel_ZAINNET.txt')};
