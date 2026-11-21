'use strict';

/** IO Insight 2023: 기존 Excel/Python 계산 엔진을 유지한 분석 UI. */
const $ = id => document.getElementById(id);
const fmt = value => Number(value).toLocaleString('ko-KR', {maximumFractionDigits:2});
const fmtRatio = value => Number(value).toLocaleString('ko-KR', {minimumFractionDigits:2,maximumFractionDigits:4});
const fmtRaw = value => Number(value).toLocaleString('ko-KR', {maximumFractionDigits:8});
const cfg = {
  large: {file:'io-2023-large.xlsx',count:33,label:'통합대분류'},
  medium:{file:'io-2023-medium.xlsx',count:83,label:'통합중분류'},
  small: {file:'io-2023-small.xlsx',count:165,label:'통합소분류'},
  basic: {file:'io-2023-basic.xlsx',count:380,label:'기본부문'}
};
const pageNames = {dashboard:'대시보드',explore:'데이터 탐색',simulate:'유발효과 분석',compare:'산업 간 비교',rank:'산업별 순위',heatmap:'산업연관 히트맵',network:'산업연관 네트워크',decompose:'직접·간접효과 분해',policy:'정책 시나리오',report:'산업 보고서',methods:'분석 방법 안내'};
let worker=null, requests=0, activeLevel=null, sectors=[], rowChoices=[], cells=[], basket=[], result=null, comparison=null;
let currentUnit='—', exploreSerial=0, loading=false, busySim=false, busyCompare=false, generation=0;
const pending=new Map();

function el(tag,text,cls=''){
  const item=document.createElement(tag);
  if(text!==undefined) item.textContent=String(text);
  if(cls) item.className=cls;
  return item;
}
function putText(id,text){$(id).textContent=text;}
function options(select,items){
  const fragment=document.createDocumentFragment();
  for(const item of items){
    const option=el('option',item.code+' · '+item.name);
    option.value=item.code;
    fragment.appendChild(option);
  }
  select.replaceChildren(fragment);
}
function message(text,type=''){
  putText('status',text);
  $('status').className='dataset-status'+(type?' '+type:'');
}
function setReady(ready){
  for(const id of ['sheet','row','row-search','col-search','sector-search','sector','amount','add','run','cmp-a','cmp-b','cmp-amount-a','cmp-amount-b','cmp-run','rank-metric','rank-search','rank-download','heat-kind','heat-sector','heat-run','network-sector','network-limit','network-threshold','network-run','dec-sector','dec-amount','dec-run','policy-a','policy-b','policy-budget','policy-split','policy-run','report-sector','report-run']) $(id).disabled=!ready;
  const badge=$('ready-badge');
  badge.classList.toggle('live',ready);
  badge.replaceChildren(el('span',undefined,'pill-dot'),document.createTextNode(ready ? `${cfg[activeLevel].label} 연결됨` : '데이터 미연결'));
  putText('global-state',ready?'연결됨':'미연결');
  $('global-state').classList.toggle('live',ready);
  putText('dashboard-ready',ready?`${cfg[activeLevel].label} ${cfg[activeLevel].count}개 산업 데이터가 연결되었습니다. 분석 페이지에서 바로 사용할 수 있습니다.`:'현재 선택한 Excel을 먼저 불러오세요.');
  $('dashboard-ready').classList.toggle('live',ready);
  if(!ready)putText('cmp-status','분석 데이터부터 불러오세요.');
}
function setPage(page,scroll=true){
  if(!(page in pageNames))page='dashboard';
  for(const section of document.querySelectorAll('[data-page]'))section.hidden=section.dataset.page!==page;
  for(const button of document.querySelectorAll('.nav-link')){
    const selected=button.dataset.nav===page;
    button.classList.toggle('active',selected);
    if(selected)button.setAttribute('aria-current','page');else button.removeAttribute('aria-current');
  }
  putText('current-page',pageNames[page]);
  document.title=pageNames[page]+' | IO Insight 2023';
  if(scroll)window.scrollTo({top:0,behavior:'auto'});
}
function go(page){
  if(!(page in pageNames))return;
  if(location.hash!=='#'+page)location.hash=page;
  setPage(page);
}
for(const button of document.querySelectorAll('[data-nav]'))button.addEventListener('click',()=>go(button.dataset.nav));
window.addEventListener('hashchange',()=>setPage(location.hash.slice(1)));
setPage(location.hash.slice(1),false);

function bootWorker(){
  if(worker)return;
  worker=new Worker('./worker.js');
  worker.onmessage=({data})=>{
    if(data.type==='status'){
      message(data.text);
      return;
    }
    const p=pending.get(data.id);
    if(!p)return;
    pending.delete(data.id);
    data.ok?p.resolve(data.result):p.reject(Error(data.error));
  };
  worker.onerror=event=>{
    for(const p of pending.values())p.reject(Error(event.message||'Python 실행 오류'));
    pending.clear();
    message('Worker 오류가 발생했습니다. 페이지를 새로고침하세요.','bad');
  };
}
function call(type,payload){
  bootWorker();
  return new Promise((resolve,reject)=>{
    const id=++requests;
    pending.set(id,{resolve,reject});
    worker.postMessage({id,type,payload});
  });
}
function emptyTbody(id,cols,text){
  const td=el('td',text,'empty');td.colSpan=cols;
  const tr=el('tr');tr.append(td);
  $(id).replaceChildren(tr);
}
function clearResults(){
  exploreSerial++;
  result=null;comparison=null;
  v3Reset();
  $('results').hidden=true;
  $('compare-results').hidden=true;
  basket=[];renderBasket();
  cells=[];rowChoices=[];
  $('row').replaceChildren();$('sheet').replaceChildren();
  emptyTbody('cells',3,'Excel을 불러오면 데이터가 표시됩니다.');
  putText('explore-count','Excel 불러오기 필요');
  putText('explore-caption','조회 결과');
  putText('unit','단위: —');
  putText('validation','Excel을 불러오면 검증 결과가 여기에 표시됩니다.');
}
async function loadWorkbook(){
  if(loading||busyCompare||busySim)return;
  const key=$('level').value;
  const item=cfg[key];
  loading=true;
  generation++;
  activeLevel=null;
  sectors=[];
  setReady(false);
  clearResults();
  $('load').disabled=true;
  $('level').disabled=true;
  $('load').textContent='데이터 불러오는 중…';
  try{
    message(`${item.label}의 Excel 파일을 내려받는 중…`);
    // Pages 업로드 시 data/ 폴더가 보존된 형태를 우선하며, 루트 업로드도 지원.
    let buffer=null;
    const attempts=[];
    for(const url of ['./data/'+item.file,'./'+item.file]){
      try{
        const response=await fetch(url,{cache:'no-store'});
        if(!response.ok){attempts.push(url+' HTTP '+response.status);continue;}
        const candidate=await response.arrayBuffer();
        const magic=new Uint8Array(candidate,0,Math.min(candidate.byteLength,2));
        if(magic[0]!==0x50||magic[1]!==0x4b){attempts.push(url+' Excel 형식 아님');continue;}
        buffer=candidate;break;
      }catch(err){attempts.push(url+' '+err.message);}
    }
    if(!buffer)throw Error('Excel 다운로드 실패: '+attempts.join(' / '));
    const info=await call('init',{buffer,count:item.count});
    if(info.count!==item.count||!Array.isArray(info.sectors)||info.sectors.length!==item.count)throw Error('산업 수가 일치하지 않습니다.');
    activeLevel=key;
    sectors=info.sectors;
    options($('sector'),sectors);
    options($('cmp-a'),sectors);
    options($('cmp-b'),sectors);
    v3InitSectors();
    const a=sectors.find(x=>x.code==='F')||sectors[0];
    const b=sectors.find(x=>x.code==='P')||sectors[Math.min(1,sectors.length-1)];
    $('cmp-a').value=a.code;
    $('cmp-b').value=b.code;
    for(const id of ['sector-search','row-search','col-search'])$(id).value='';
    $('sheet').replaceChildren(...info.sheets.map(sheet=>{
      const option=el('option',sheet);option.value=sheet;return option;
    }));
    putText('validation',JSON.stringify(info.validation,null,2));
    setReady(true);
    await explore();
    try{await loadRankings();}catch(rankErr){message('데이터는 연결되었지만 산업별 순위 조회 실패: '+rankErr.message,'bad');}
    putText('explore-count',`${info.count}개 산업 · ${info.sheets.length}개 시트`);
    putText('cmp-status','산업 A와 B를 선택하고 비교를 시작하세요.');
    message(`${item.label} ${item.count}개 산업 / ${info.sheets.length}개 시트 · 행렬 검증 통과`,'good');
  }catch(error){
    activeLevel=null;
    sectors=[];
    setReady(false);
    message('불러오기 실패: '+error.message+' — data/ 경로, Excel 파일, Pyodide 연결을 확인하세요.','bad');
  }finally{
    loading=false;
    $('load').disabled=false;
    $('level').disabled=false;
    $('load').textContent='Excel 데이터 불러오기 ↗';
  }
}
$('load').addEventListener('click',loadWorkbook);
$('level').addEventListener('change',()=>{
  activeLevel=null;sectors=[];generation++;
  setReady(false);clearResults();
  message('산업분류가 변경되었습니다. Excel을 다시 불러오세요.');
});

function filteredRowOptions(preferred){
  const q=$('row-search').value.trim().toLocaleLowerCase();
  const subset=rowChoices.filter(item=>(item.code+' '+item.name).toLocaleLowerCase().includes(q));
  const old=preferred||$('row').value;
  options($('row'),subset);
  if(subset.some(item=>item.code===old))$('row').value=old;
  return subset;
}
async function explore(){
  if(!activeLevel)return;
  const serial=++exploreSerial;
  const sheet=$('sheet').value;
  try{
    const response=await call('explore',{sheet,row:$('row').value});
    if(serial!==exploreSerial||sheet!==$('sheet').value)return;
    rowChoices=response.rows;
    const found=filteredRowOptions(response.selected);
    if(!found.some(item=>item.code===response.selected)){
      putText('explore-caption','검색 조건에 일치하는 행이 없습니다.');
      emptyTbody('cells',3,'검색 결과가 없습니다.');
      return;
    }
    cells=response.cells;
    currentUnit=response.unit;
    putText('explore-caption',`${sheet} / ${response.selected} ${response.name}`);
    putText('unit','단위: '+currentUnit);
    renderCells();
  }catch(error){if(serial===exploreSerial)message('데이터 조회 오류: '+error.message,'bad');}
}
function renderCells(){
  const q=$('col-search').value.trim().toLocaleLowerCase();
  const filtered=cells.filter(item=>(item.code+' '+item.name).toLocaleLowerCase().includes(q));
  const fragment=document.createDocumentFragment();
  for(const value of filtered){
    const tr=el('tr');
    tr.append(el('td',value.code),el('td',value.name),el('td',value.value===null?'—':typeof value.value==='number'?fmtRaw(value.value):value.value,'number'));
    fragment.append(tr);
  }
  if(!filtered.length){emptyTbody('cells',3,'검색 결과가 없습니다.');return;}
  $('cells').replaceChildren(fragment);
}
$('sheet').addEventListener('change',()=>{
  $('row').replaceChildren();$('row-search').value='';$('col-search').value='';explore();
});
$('row').addEventListener('change',explore);
$('row-search').addEventListener('input',()=>{
  const before=$('row').value;
  const subset=filteredRowOptions();
  if(!subset.length){emptyTbody('cells',3,'검색 결과가 없습니다.');putText('explore-caption','검색 조건에 일치하는 행이 없습니다.');return;}
  if($('row').value!==before)explore();
});
$('col-search').addEventListener('input',renderCells);

function filterSectors(){
  const q=$('sector-search').value.trim().toLocaleLowerCase();
  const old=$('sector').value;
  const match=sectors.filter(item=>(item.code+' '+item.name).toLocaleLowerCase().includes(q));
  options($('sector'),match);
  if(match.some(item=>item.code===old))$('sector').value=old;
}
$('sector-search').addEventListener('input',filterSectors);
for(const button of document.querySelectorAll('[data-amount]'))button.addEventListener('click',()=>{$('amount').value=button.dataset.amount;});
function renderBasket(){
  if(!basket.length){$('basket').replaceChildren(el('li','추가한 산업이 없습니다.','empty'));putText('sim-status','먼저 산업을 추가하세요.');return;}
  const fragment=document.createDocumentFragment();
  for(const [index,item] of basket.entries()){
    const li=el('li');
    li.append(el('span',`${item.code} · ${item.name}  /  ${fmt(item.amount/100)}억원`));
    const del=el('button','삭제');del.type='button';
    del.addEventListener('click',()=>{basket.splice(index,1);result=null;$('results').hidden=true;renderBasket();});
    li.append(del);fragment.append(li);
  }
  $('basket').replaceChildren(fragment);
  putText('sim-status',`${basket.length}개 산업 시나리오가 준비되었습니다.`);
}
$('scenario').addEventListener('submit',event=>{
  event.preventDefault();
  const sector=sectors.find(item=>item.code===$('sector').value);
  const amount=Number($('amount').value);
  if(!activeLevel||!sector||!Number.isFinite(amount)||amount<=0||amount>1e12){
    putText('sim-status','올바른 산업과 양(+)의 금액(억원)을 입력하세요.');return;
  }
  basket.push({...sector,amount:amount*100});
  result=null;$('results').hidden=true;renderBasket();
});
$('clear').addEventListener('click',()=>{basket=[];result=null;$('results').hidden=true;renderBasket();});

function displayResults(){
  if(!result)return;
  const items=[['최종수요 증가액',result.demand],['생산유발액',result.production],['부가가치유발액',result.valueAdded],['수입유발액',result.imports]];
  const fragment=document.createDocumentFragment();
  for(const [title,value] of items){
    const card=el('div',undefined,'stat');
    card.append(el('p',title,'stat-title'),el('strong',fmt(value/100)),el('span','억원','stat-unit'));
    fragment.append(card);
  }
  $('cards').replaceChildren(fragment);
  putText('multiplier',`생산유발배율 ${fmtRatio(result.productionMultiplier)}배 · 최종수요 1억원 증가당 약 ${fmtRatio(result.productionMultiplier)}억원의 총생산이 유발됩니다. 중간재 거래가 포함된 총산출이며 GDP 증가액은 아닙니다.`);
  const ranked=result.rows.slice(0,12);
  const max=Math.max(...ranked.map(x=>x.production),1e-10);
  const barFragment=document.createDocumentFragment();
  for(const item of ranked){
    const line=el('div',undefined,'bar-row');
    const label=el('span',item.name,'bar-label');label.title=item.code+' '+item.name;
    const track=el('div',undefined,'bar-track');const fill=el('div',undefined,'bar-fill');
    fill.style.width=(Math.max(0,item.production)/max*100)+'%';
    track.append(fill);line.append(label,track,el('span',fmt(item.production/100),'bar-value'));barFragment.append(line);
  }
  $('bars').replaceChildren(barFragment);
  const rows=document.createDocumentFragment();
  for(const item of result.rows){
    const tr=el('tr');tr.append(el('td',item.code),el('td',item.name));
    for(const amount of [item.production,item.valueAdded,item.imports])tr.append(el('td',fmt(amount/100),'number'));
    rows.append(tr);
  }
  $('result-rows').replaceChildren(rows);
  $('results').hidden=false;
}
$('run').addEventListener('click',async()=>{
  if(busySim||loading||!activeLevel||!basket.length){putText('sim-status','분석할 산업을 먼저 추가하세요.');return;}
  const snapshot=basket.map(({code,amount})=>({code,amount}));
  const serialized=JSON.stringify(snapshot);
  const currentGeneration=generation;
  busySim=true;$('run').disabled=true;
  putText('sim-status','Python에서 계수행렬을 계산 중…');
  try{
    const answer=await call('simulate',snapshot);
    if(currentGeneration!==generation||serialized!==JSON.stringify(basket.map(({code,amount})=>({code,amount})))){
      putText('sim-status','계산 도중 조건이 변경되었습니다. 다시 실행해 주세요.');return;
    }
    result=answer;displayResults();
    putText('sim-status','계산 완료 · 아래 결과를 확인하세요.');
  }catch(error){result=null;$('results').hidden=true;putText('sim-status','계산 실패: '+error.message);}
  finally{busySim=false;$('run').disabled=!activeLevel;}
});

function escapeCsv(value){
  let s=String(value);
  if(typeof value==='string'&&/^[=+@\-\t\r]/.test(s))s="'"+s;
  return '"'+s.replaceAll('"','""')+'"';
}
function downloadCsv(name,records){
  const content='\ufeff'+records.map(row=>row.map(escapeCsv).join(',')).join('\r\n');
  const blob=new Blob([content],{type:'text/csv;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const a=el('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),4000);
}
$('download').addEventListener('click',()=>{
  if(!result||!activeLevel)return;
  const records=[['분류','산업코드','산업명','최종수요증가액_억원','생산유발액_억원','부가가치유발액_억원','수입유발액_억원']];
  for(const item of result.rows)records.push([cfg[activeLevel].label,item.code,item.name,item.finalDemand/100,item.production/100,item.valueAdded/100,item.imports/100]);
  downloadCsv('io-insight-2023-'+activeLevel+'-scenario.csv',records);
});

function makeMetricBars(name,va,vb,unit='억원'){
  const outer=el('div',undefined,'cmp-metric');
  const head=el('div',undefined,'cmp-metric-head');
  head.append(el('span',name),el('span',`${fmt(va)} / ${fmt(vb)} ${unit}`,'cmp-metric-values'));
  outer.append(head);
  const max=Math.max(va,vb,1e-10);
  const group=el('div',undefined,'cmp-two-bars');
  for(const [value,type] of [[va,'a'],[vb,'b']]){
    const track=el('div',undefined,'cmp-track');
    const fill=el('div',undefined,'cmp-fill '+type);
    fill.style.width=Math.max(0,value)/max*100+'%';
    track.append(fill);group.append(track);
  }
  outer.append(group);return outer;
}
function makeCompareKpi(label,unit,va,vb,hint=''){
  const card=el('div',undefined,'compare-kpi');
  const head=el('div',undefined,'compare-kpi-head');
  head.append(el('span',label),el('small',hint));
  const values=el('div',undefined,'compare-kpi-values');
  for(const [name,value,cls] of [['A',va,'a'],['B',vb,'b']]){
    const box=el('div');
    box.append(el('small',name),el('strong',fmt(value)+unit,cls));
    values.append(box);
  }
  card.append(head,values);return card;
}
function narrativeItem(headline,description){
  const item=el('div',undefined,'narrative-item');
  item.append(el('strong',headline),el('p',description));
  return item;
}
function compareSentence(label,aName,bName,a,b,unit){
  const tolerance=Math.max(Math.abs(a),Math.abs(b),1)*1e-9;
  if(Math.abs(a-b)<=tolerance)return `${label}는 두 산업이 거의 동일합니다.`;
  const winner=a>b?aName:bName;
  const distance=Math.abs(a-b);
  return `${label}는 ${winner}이(가) 더 크며, 100억원의 동일한 최종수요를 가정하면 차이는 약 ${fmt(distance)}${unit}입니다.`;
}
function showComparison(){
  if(!comparison)return;
  const {a,b,infoA,infoB}=comparison;
  const nameA=`${infoA.code} · ${infoA.name}`;
  const nameB=`${infoB.code} · ${infoB.name}`;
  putText('compare-result-title',`${infoA.name} vs ${infoB.name}`);
  putText('compare-name-a',nameA+' · '+fmt(infoA.demand)+'억원');
  putText('compare-name-b',nameB+' · '+fmt(infoB.demand)+'억원');
  const unitA=infoA.demand,unitB=infoB.demand;
  // 금액은 각 시나리오의 총효과. 서로 다른 수요액일 때 배율과 정규화 수치로 해석한다.
  const kpis=[
    ['총생산유발액','억원',a.production/100,b.production/100,'총산출'],
    ['총부가가치유발액','억원',a.valueAdded/100,b.valueAdded/100,'국내 부가가치'],
    ['총수입유발액','억원',a.imports/100,b.imports/100,'중간재 수입'],
    ['생산유발배율','배',a.productionMultiplier,b.productionMultiplier,'투입액 대비']
  ];
  const kpiFragment=document.createDocumentFragment();
  for(const metric of kpis)kpiFragment.append(makeCompareKpi(...metric));
  $('compare-kpis').replaceChildren(kpiFragment);
  const metrics=document.createDocumentFragment();
  for(const [label,unit,va,vb] of kpis.slice(0,3))metrics.append(makeMetricBars(label,va,vb,unit));
  metrics.append(makeMetricBars('생산유발배율',a.productionMultiplier,b.productionMultiplier,'배'));
  $('compare-bars').replaceChildren(metrics);

  const aMap=new Map(a.rows.map(x=>[x.code,x]));
  const bMap=new Map(b.rows.map(x=>[x.code,x]));
  const merged=sectors.map(sector=>{
    const xa=aMap.get(sector.code)?.production||0;
    const xb=bMap.get(sector.code)?.production||0;
    return {code:sector.code,name:sector.name,a:xa/100,b:xb/100,peak:Math.max(xa,xb)};
  }).sort((x,y)=>y.peak-x.peak).slice(0,8);
  const sectorChart=document.createDocumentFragment();
  for(const entry of merged){
    sectorChart.append(makeMetricBars(`${entry.code} ${entry.name}`,entry.a,entry.b));
  }
  $('compare-sector-bars').replaceChildren(sectorChart);

  const normalize=(raw,demand)=>raw/demand*100; // 두 수요액을 100억원으로 정규화 (raw 값은 백만원)
  const naP=normalize(a.production,a.demand),nbP=normalize(b.production,b.demand);
  const naV=normalize(a.valueAdded,a.demand),nbV=normalize(b.valueAdded,b.demand);
  const naM=normalize(a.imports,a.demand),nbM=normalize(b.imports,b.demand);
  const interpret=document.createDocumentFragment();
  interpret.append(narrativeItem('생산 파급효과 비교',compareSentence('최종수요 100억원당 생산유발액',infoA.name,infoB.name,naP,nbP,'억원')));
  interpret.append(narrativeItem('부가가치 창출 비교',compareSentence('최종수요 100억원당 부가가치유발액',infoA.name,infoB.name,naV,nbV,'억원')));
  interpret.append(narrativeItem('수입 중간재 파급효과',compareSentence('최종수요 100억원당 수입유발액',infoA.name,infoB.name,naM,nbM,'억원')));
  interpret.append(narrativeItem('해석 시 주의사항',unitA!==unitB?'두 시나리오의 입력 금액이 다르므로 총액만으로 비교하면 안 됩니다. 위 해설은 최종수요 100억원당 효과로 환산해 비교했습니다.':'같은 100억원이 아니더라도 두 입력액이 동일하여 직접 비교할 수 있습니다. 생산유발액은 GDP나 부가가치 증가액이 아닙니다.'));
  $('compare-narrative').replaceChildren(interpret);
  $('compare-results').hidden=false;
}
$('cmp-run').addEventListener('click',async()=>{
  if(busyCompare||loading||!activeLevel)return;
  const infoA=sectors.find(x=>x.code===$('cmp-a').value);
  const infoB=sectors.find(x=>x.code===$('cmp-b').value);
  const amountA=Number($('cmp-amount-a').value),amountB=Number($('cmp-amount-b').value);
  if(!infoA||!infoB||!Number.isFinite(amountA)||!Number.isFinite(amountB)||amountA<=0||amountB<=0||amountA>1e12||amountB>1e12){
    putText('cmp-status','두 산업과 양(+)의 최종수요 증가액을 입력하세요.');return;
  }
  const gen=generation;
  busyCompare=true;
  $('cmp-run').disabled=true;
  comparison=null;$('compare-results').hidden=true;
  putText('cmp-status','두 산업을 각각 분석하고 있습니다…');
  try{
    // 기존 Python 모델을 각 산업에 따로 적용하여 독립 시나리오 비교.
    const [a,b]=await Promise.all([
      call('simulate',[{code:infoA.code,amount:amountA*100}]),
      call('simulate',[{code:infoB.code,amount:amountB*100}])
    ]);
    if(gen!==generation || infoA.code!==$('cmp-a').value||infoB.code!==$('cmp-b').value||amountA!==Number($('cmp-amount-a').value)||amountB!==Number($('cmp-amount-b').value)){
      putText('cmp-status','입력값이 변경되었습니다. 비교를 다시 실행하세요.');return;
    }
    comparison={a,b,infoA:{...infoA,demand:amountA},infoB:{...infoB,demand:amountB}};
    showComparison();
    putText('cmp-status','비교 완료 · 자동 해설과 산업별 생산유발액을 확인하세요.');
  }catch(error){putText('cmp-status','비교 실패: '+error.message);}
  finally{busyCompare=false;$('cmp-run').disabled=!activeLevel;}
});
for(const id of ['cmp-a','cmp-b','cmp-amount-a','cmp-amount-b'])$(id).addEventListener('change',()=>{
  comparison=null;
  $('compare-results').hidden=true;
  if(activeLevel)putText('cmp-status','비교 조건이 변경되었습니다. 다시 계산하세요.');
});
$('cmp-download').addEventListener('click',()=>{
  if(!comparison||!activeLevel)return;
  const {a,b,infoA,infoB}=comparison;
  const rows=[['산업분류',cfg[activeLevel].label],['시나리오A',infoA.code+' '+infoA.name+' / '+infoA.demand+'억원'],['시나리오B',infoB.code+' '+infoB.name+' / '+infoB.demand+'억원'],[],['산업코드','산업명','A_생산유발액_억원','A_부가가치유발액_억원','A_수입유발액_억원','B_생산유발액_억원','B_부가가치유발액_억원','B_수입유발액_억원']];
  const aMap=new Map(a.rows.map(x=>[x.code,x]));const bMap=new Map(b.rows.map(x=>[x.code,x]));
  for(const sector of sectors){
    const xa=aMap.get(sector.code),xb=bMap.get(sector.code);
    rows.push([sector.code,sector.name,xa.production/100,xa.valueAdded/100,xa.imports/100,xb.production/100,xb.valueAdded/100,xb.imports/100]);
  }
  downloadCsv('io-insight-2023-'+activeLevel+'-comparison.csv',rows);
});

// ============================================================
// IO Insight 2023 V3: 추가 산업연관 구조/순위/시나리오 분석
// Python 엔진은 원본 Excel의 계수행렬을 사용하며, 새 통계는 예시값이 아님.
// ============================================================
let ranking=null, structureData=null, decomposition=null, policyResults=null, reportData=null;
let v3Busy={heat:false,dec:false,policy:false,report:false,network:false};
let networkResult=null;
function v3Reset(){
  ranking=null; structureData=null;decomposition=null;policyResults=null;reportData=null;networkResult=null;
  $('rank-summary').replaceChildren(el('div','Excel을 불러오면 전체 산업을 비교합니다.','stat'));
  $('rank-bars').replaceChildren(el('p','Excel 데이터를 먼저 불러오세요.','fineprint'));
  emptyTbody('rank-rows',5,'데이터 연결 후 표시됩니다.');
  putText('rank-count','미연결');
  $('heat-wrapper').replaceChildren(el('div','분류를 불러오면 히트맵을 생성할 수 있습니다.','empty'));
  $('heat-top').replaceChildren(el('p','히트맵을 생성하세요.','fineprint'));
  putText('heat-heading','데이터 불러오기 필요');putText('heat-scope','–');putText('heat-inspect','셀을 선택하면 해당 산업 간 계수가 표시됩니다.');
  $('dec-results').hidden=true;$('policy-results').hidden=true;$('print-report').hidden=true;
  $('network-results').hidden=true; $('network-canvas').replaceChildren(el('div','데이터 연결 후 네트워크를 생성하세요.','empty'));
  putText('network-status','분류별 Excel을 불러오세요.');
  $('report-print').disabled=true;
  for(const [id,msg] of [['dec-status','분석 데이터를 먼저 불러오세요.'],['policy-status','Excel을 먼저 불러오세요.'],['report-status','Excel을 먼저 불러오세요.']])putText(id,msg);
}
function v3InitSectors(){
  for(const id of ['heat-sector','network-sector','dec-sector','policy-a','policy-b','report-sector'])options($(id),sectors);
  const a=sectors.find(x=>x.code==='F')||sectors[0];
  const b=sectors.find(x=>x.code==='P')||sectors[Math.min(1,sectors.length-1)];
  for(const id of ['heat-sector','network-sector','dec-sector','report-sector','policy-a'])$(id).value=a.code;
  $('policy-b').value=b.code;
  putText('dec-status','산업과 증가액을 설정해 직접·간접효과를 계산하세요.');
  putText('policy-status','두 산업과 총수요를 선택해 시나리오를 계산하세요.');
  putText('report-status','산업 선택 후 보고서를 생성하세요.');
  putText('network-status','중심 산업과 연결 수를 선택한 후 네트워크를 그리세요.');
}
async function loadRankings(){
  ranking=await call('rank',{});
  if(!activeLevel || ranking.count!==sectors.length)throw Error('산업 순위 결과의 산업 수가 다릅니다.');
  renderRankings();
}
function rankMetricName(metric){return ({production:'생산유발액',valueAdded:'부가가치유발액',imports:'수입유발액'})[metric]||'생산유발액';}
function renderRankings(){
  if(!ranking)return;
  const metric=$('rank-metric').value;
  const needle=$('rank-search').value.trim().toLocaleLowerCase();
  const data=ranking.rows.filter(item=>(item.code+' '+item.name).toLocaleLowerCase().includes(needle));
  data.sort((a,b)=>b[metric]-a[metric]);
  const stats=document.createDocumentFragment();
  for(const [key,label] of [['production','생산유발 최고'],['valueAdded','부가가치유발 최고'],['imports','수입유발 최고']]){
    const winner=[...ranking.rows].sort((a,b)=>b[key]-a[key])[0];
    const item=el('div',undefined,'stat');
    item.append(el('p',label,'stat-title'),el('strong',fmt(winner[key])+'억원'),el('p',winner.code+' · '+winner.name,'v3-stat-desc'));
    stats.append(item);
  }
  $('rank-summary').replaceChildren(stats);
  putText('rank-chart-title',rankMetricName(metric)+' 상위 12개 산업');
  putText('rank-count',`${data.length}개 산업 · 100억원당`);
  const max=Math.max(...data.slice(0,12).map(x=>x[metric]),1e-12);
  const bars=document.createDocumentFragment();
  data.slice(0,12).forEach(item=>{
    const line=el('div',undefined,'bar-row');
    const label=el('span',item.code+' '+item.name,'bar-label');label.title=item.code+' '+item.name;
    const track=el('div',undefined,'bar-track');const fill=el('div',undefined,'bar-fill');
    fill.style.width=(Math.max(0,item[metric])/max*100)+'%';track.append(fill);
    line.append(label,track,el('span',fmt(item[metric]),'bar-value'));bars.append(line);
  });
  $('rank-bars').replaceChildren(bars);
  const rows=document.createDocumentFragment();
  for(const [i,item] of data.entries()){
    const tr=el('tr');tr.append(el('td',i+1),el('td',item.code+' · '+item.name));
    for(const k of ['production','valueAdded','imports'])tr.append(el('td',fmt(item[k]),'number'));
    rows.append(tr);
  }
  if(data.length)$('rank-rows').replaceChildren(rows);else emptyTbody('rank-rows',5,'검색 결과가 없습니다.');
}
$('rank-metric').addEventListener('change',renderRankings);
$('rank-search').addEventListener('input',renderRankings);
$('rank-download').addEventListener('click',()=>{
  if(!ranking||!activeLevel)return;
  const metric=$('rank-metric').value;
  const ranked=[...ranking.rows].sort((a,b)=>b[metric]-a[metric]);
  const rows=[['산업분류',cfg[activeLevel].label],['각 산업의 독립적인 국산품 최종수요 증가액_억원',100],['정렬지표',rankMetricName(metric)],[],['순위','산업코드','산업명','생산유발액_억원','부가가치유발액_억원','수입유발액_억원','생산유발배율']];
  ranked.forEach((x,i)=>rows.push([i+1,x.code,x.name,x.production,x.valueAdded,x.imports,x.productionMultiplier]));
  downloadCsv('io-v3-'+activeLevel+'-industry-ranking.csv',rows);
});

// 투입계수행렬의 행=공급산업/열=사용산업, 그중 선택된 열의 상위 공급산업을 강조.
$('heat-run').addEventListener('click',async()=>{
  if(!activeLevel||v3Busy.heat||loading)return;
  v3Busy.heat=true;$('heat-run').disabled=true;
  const gen=generation,code=$('heat-sector').value,kind=$('heat-kind').value;
  putText('heat-heading','원본 계수표 분석 중…');
  try{
    const data=await call('structure',{code,kind});
    if(gen!==generation||code!==$('heat-sector').value||kind!==$('heat-kind').value)return;
    structureData=data;renderHeatmap();
  }catch(err){putText('heat-heading','히트맵 오류: '+err.message);}
  finally{v3Busy.heat=false;$('heat-run').disabled=!activeLevel;}
});
function renderHeatmap(){
  const data=structureData;if(!data)return;
  putText('heat-heading',`${data.target.name} · ${data.kind==='Ad'?'국산투입계수':data.kind==='Am'?'수입투입계수':'생산유발계수'}`);
  putText('heat-scope',data.selection);
  const names=data.labels, n=names.length;
  const board=el('div',undefined,'heat-grid');board.style.gridTemplateColumns='minmax(120px,165px) repeat('+n+', 30px)';
  const empty=el('div','공급 ↓ / 사용 →','heat-axis-corner');board.append(empty);
  for(const x of names){const h=el('div',x.code,'heat-col-head');h.title=x.name;board.append(h);}
  const max=Math.max(data.maxValue,1e-10);
  data.cells.forEach((row,i)=>{
    const label=el('div',names[i].code,'heat-row-head');label.title=names[i].code+' '+names[i].name;
    board.append(label);
    row.forEach((raw,j)=>{
      const value=Number(raw),rate=Math.min(1,Math.max(0,value/max));
      const cell=el('div',undefined,'heat-cell');
      cell.style.backgroundColor=`hsl(${213-25*rate}, ${57+26*rate}%, ${96-57*rate}%)`;
      cell.title=`공급 ${names[i].code} ${names[i].name} → 사용 ${names[j].code} ${names[j].name}: ${fmtRaw(value)}`;
      cell.setAttribute('aria-label',cell.title);
      cell.tabIndex=0;cell.setAttribute('role','button');
      const showCell=()=>putText('heat-inspect',cell.title);
      cell.addEventListener('click',showCell);
      cell.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();showCell();}});
      if(i===j)cell.classList.add('heat-self');
      board.append(cell);
    });
  });
  $('heat-wrapper').replaceChildren(board);
  const strongest=data.topInputs.slice(0,10),peak=Math.max(...strongest.map(x=>x.coefficient),1e-10);
  const frag=document.createDocumentFragment();
  for(const item of strongest){
    const line=el('div',undefined,'bar-row');
    const caption=el('span',item.code+' '+item.name,'bar-label');caption.title=item.name;
    const track=el('div',undefined,'bar-track');const fill=el('div',undefined,'bar-fill');
    fill.style.width=(Math.max(0,item.coefficient)/peak*100)+'%';track.append(fill);
    line.append(caption,track,el('span',fmtRaw(item.coefficient),'bar-value'));frag.append(line);
  }
  $('heat-top').replaceChildren(frag);
  putText('heat-top-heading',`선택 산업(${data.target.code})과 연결된 상위 공급 산업`);
  putText('heat-note',data.kind==='L'?'생산유발계수는 직접·간접 총산출 유발을 포함합니다. 자기 산업 계수는 최초 수요의 직접효과도 포함합니다.': '계수는 해당 열 산업 1단위 생산에 필요한 행 산업 중간재 투입량을 뜻합니다. 값이 높다고 인과적 중요도나 정책 성과가 자동으로 커지는 것은 아닙니다.');
}

$('dec-run').addEventListener('click',async()=>{
  if(!activeLevel||v3Busy.dec||loading)return;
  const code=$('dec-sector').value,amount=Number($('dec-amount').value);
  if(!Number.isFinite(amount)||amount<=0||amount>1e12){putText('dec-status','0보다 큰 유효 금액을 입력하세요.');return;}
  const gen=generation;
  v3Busy.dec=true;$('dec-run').disabled=true;putText('dec-status','직접·간접 산출액 분해 중…');
  try{
    const data=await call('decompose',[{code,amount:amount*100}]);
    if(gen!==generation||code!==$('dec-sector').value||amount!==Number($('dec-amount').value))return;
    decomposition={...data,sector:sectors.find(x=>x.code===code)};
    renderDecomposition();putText('dec-status','분해 완료: 직접효과와 모든 간접효과를 합하면 총생산유발액입니다.');
  }catch(err){putText('dec-status','분해 실패: '+err.message);}
  finally{v3Busy.dec=false;$('dec-run').disabled=!activeLevel;}
});
function renderDecomposition(){
  const d=decomposition;if(!d)return;
  const stages=[['직접',d.direct,'direct'],['1차 중간재',d.first,'first'],['2차 이상',d.higher,'higher']];
  const cards=document.createDocumentFragment();
  for(const [label,val] of [...stages,['총생산유발액',d.production,'total']]){
    const stat=el('div',undefined,'stat');stat.append(el('p',label,'stat-title'),el('strong',fmt(val)+'억원'));cards.append(stat);
  }
  $('dec-kpis').replaceChildren(cards);
  const total=Math.max(d.production,1e-10);const wrap=el('div',undefined,'stacked-track');
  stages.forEach(([label,value,cls])=>{
    const seg=el('div',undefined,'stacked-segment '+cls);seg.style.width=Math.max(0,value/total*100)+'%';seg.title=label+': '+fmt(value)+'억원';wrap.append(seg);
  });
  const labels=el('div',undefined,'stacked-legend');
  stages.forEach(([label,value,cls])=>{const x=el('span',undefined,'stacked-legend-item');x.append(el('i',undefined,'stacked-swatch '+cls),el('span',`${label}: ${fmt(value)}억원`));labels.append(x);});
  $('dec-stacked').replaceChildren(wrap,labels);
  putText('dec-explainer',`${d.sector.name}의 국산품 최종수요 ${fmt(d.demand)}억원 증가 시 직접 산출 ${fmt(d.direct)}억원, 1차 중간재 파급 ${fmt(d.first)}억원, 2차 이상 연쇄 파급 ${fmt(d.higher)}억원으로 합계 ${fmt(d.production)}억원의 총생산이 유발됩니다. 총산출은 GDP가 아닙니다.`);
  const tr=document.createDocumentFragment();
  for(const item of d.rows){const line=el('tr');line.append(el('td',item.code+' '+item.name));for(const key of ['direct','first','higher','production'])line.append(el('td',fmt(item[key]),'number'));tr.append(line);}
  $('dec-rows').replaceChildren(tr);$('dec-results').hidden=false;
}
$('dec-download').addEventListener('click',()=>{
  const d=decomposition;if(!d||!activeLevel)return;
  const rows=[['산업분류',cfg[activeLevel].label],['대상산업',d.sector.code+' '+d.sector.name],['최종수요증가_억원',d.demand],[],['산업코드','산업명','직접효과_억원','1차간접효과_억원','2차이상간접효과_억원','총생산유발액_억원']];
  for(const x of d.rows)rows.push([x.code,x.name,x.direct,x.first,x.higher,x.production]);
  downloadCsv('io-v3-'+activeLevel+'-decomposition.csv',rows);
});

function updatePolicySplit(){putText('policy-split-label',`A ${$('policy-split').value}% · B ${100-Number($('policy-split').value)}%`);}
$('policy-split').addEventListener('input',updatePolicySplit);
$('policy-run').addEventListener('click',async()=>{
  if(!activeLevel||v3Busy.policy||loading)return;
  const a=$('policy-a').value,b=$('policy-b').value,amount=Number($('policy-budget').value),share=Number($('policy-split').value)/100;
  if(!a||!b||a===b){putText('policy-status','서로 다른 두 산업을 선택하세요.');return;}
  if(!Number.isFinite(amount)||amount<=0||amount>1e12){putText('policy-status','총수요 증가액은 0보다 커야 합니다.');return;}
  const conf=[{name:'A 중심 100%',a:1},{name:'B 중심 100%',a:0},{name:'A·B 균등 50:50',a:.5},{name:`사용자 배분 ${Math.round(share*100)}:${Math.round((1-share)*100)}`,a:share}];
  const gen=generation;
  v3Busy.policy=true;$('policy-run').disabled=true;
  $('policy-results').hidden=true;policyResults=null;
  try{
    const outputs=[];
    for(let i=0;i<conf.length;i++){
      putText('policy-status',`동일 총수요 ${fmt(amount)}억원 기준 · ${i+1}/4 시나리오 계산 중…`);
      const c=conf[i];const selections=[];
      if(c.a>0)selections.push({code:a,amount:amount*100*c.a});
      if(c.a<1)selections.push({code:b,amount:amount*100*(1-c.a)});
      const result=await call('simulate',selections);
      outputs.push({name:c.name,a:c.a,...result});
    }
    if(gen!==generation||a!==$('policy-a').value||b!==$('policy-b').value||amount!==Number($('policy-budget').value)||share!==Number($('policy-split').value)/100)return;
    policyResults={budget:amount,a:sectors.find(x=>x.code===a),b:sectors.find(x=>x.code===b),out:outputs};
    renderPolicy();putText('policy-status','4가지 정책 시나리오 계산 완료.');
  }catch(err){putText('policy-status','시나리오 계산 실패: '+err.message);}
  finally{v3Busy.policy=false;$('policy-run').disabled=!activeLevel;}
});
function renderPolicy(){
  const p=policyResults;if(!p)return;
  const max=Math.max(...p.out.map(x=>x.production),1);
  const graph=document.createDocumentFragment();const rows=document.createDocumentFragment();
  for(const x of p.out){
    const row=el('div',undefined,'policy-bar-row');
    row.append(el('strong',x.name));
    const track=el('div',undefined,'bar-track');const fill=el('div',undefined,'bar-fill');fill.style.width=x.production/max*100+'%';track.append(fill);
    row.append(track,el('span',fmt(x.production/100)+'억원','bar-value'));graph.append(row);
    const tr=el('tr');tr.append(el('td',x.name));for(const v of [x.production/100,x.valueAdded/100,x.imports/100,x.productionMultiplier])tr.append(el('td',fmt(v),'number'));rows.append(tr);
  }
  $('policy-bars').replaceChildren(graph);$('policy-rows').replaceChildren(rows);
  const prod=[...p.out].sort((a,b)=>b.production-a.production)[0];const va=[...p.out].sort((a,b)=>b.valueAdded-a.valueAdded)[0];
  putText('policy-narrative',`동일한 국산품 최종수요 ${fmt(p.budget)}억원 아래, 비교한 4개 배분안 중 총생산유발액이 가장 큰 경우는 ‘${prod.name}’(${fmt(prod.production/100)}억원), 부가가치유발액이 가장 큰 경우는 ‘${va.name}’(${fmt(va.valueAdded/100)}억원)입니다. 이는 고정계수 모형의 가상 결과이며 투자 수익률·고용효과·정책 효과의 우열 판단이 아닙니다.`);
  $('policy-results').hidden=false;
}
$('policy-download').addEventListener('click',()=>{
  const p=policyResults;if(!p||!activeLevel)return;
  const rows=[['산업분류',cfg[activeLevel].label],['산업A',p.a.code+' '+p.a.name],['산업B',p.b.code+' '+p.b.name],['총수요증가_억원',p.budget],[],['시나리오','A배분비율_%','B배분비율_%','최종수요총액_억원','생산유발액_억원','부가가치유발액_억원','수입유발액_억원','생산유발배율']];
  for(const x of p.out)rows.push([x.name,100*x.a,100*(1-x.a),x.demand/100,x.production/100,x.valueAdded/100,x.imports/100,x.productionMultiplier]);
  downloadCsv('io-v3-'+activeLevel+'-policy-scenarios.csv',rows);
});

$('report-run').addEventListener('click',async()=>{
  if(!activeLevel||v3Busy.report||loading)return;
  const code=$('report-sector').value,gen=generation;
  v3Busy.report=true;$('report-run').disabled=true;
  putText('report-status','선택 산업의 파급효과와 투입관계를 분석 중…');
  try{
    const res=await call('profile',{code});
    if(gen!==generation||code!==$('report-sector').value)return;
    reportData=res;renderReport();putText('report-status','보고서 생성 완료 · 아래 인쇄/PDF 저장을 이용하세요.');
  }catch(err){putText('report-status','보고서 생성 실패: '+err.message);}
  finally{v3Busy.report=false;$('report-run').disabled=!activeLevel;}
});
function renderReport(){
  const x=reportData;if(!x)return;
  putText('report-name',x.code+' · '+x.name+' 산업연관분석 보고서');
  putText('report-subtitle',`${cfg[activeLevel].label} · 2023년 생산자가격 기준 · 국산품 최종수요 ${fmt(x.demand)}억원 증가 시나리오`);
  const f=document.createDocumentFragment();
  for(const [label,num] of [['생산유발액',x.production],['부가가치유발액',x.valueAdded],['수입유발액',x.imports],['생산유발배율',x.productionMultiplier]]){
    const s=el('div',undefined,'stat');s.append(el('p',label,'stat-title'),el('strong',fmt(num)+(label.includes('배율')?'배':'억원')));f.append(s);
  }
  $('report-kpis').replaceChildren(f);
  const list=(id,data,key,unit)=>{
    const frag=document.createDocumentFragment();
    for(const [i,r] of data.entries()){
      const line=el('div',undefined,'report-list-row');line.append(el('span',`${i+1}. ${r.code} ${r.name}`),el('strong',fmtRaw(r[key])+unit));frag.append(line);
    }
    $(id).replaceChildren(frag);
  };
  list('report-beneficiaries',x.topBeneficiaries,'production','억원');
  list('report-suppliers',x.topSuppliers,'coefficient','');
  putText('report-interpret',`${x.name}의 국산품 최종수요 100억원 증가에 대해 직접 산출 100억원과 간접 산출 ${fmt(x.indirect)}억원 등 총생산 ${fmt(x.production)}억원이 유발됩니다. 주요 생산유발 산업은 ${x.topBeneficiaries[0].name}이며, 국산투입계수 기준 주요 공급 산업은 ${x.topSuppliers[0].name}입니다. 투입 관계는 산업 간 구조를 보여주는 것이며 실제 정책 효과를 확정하지 않습니다.`);
  $('print-report').hidden=false;$('report-print').disabled=false;
}
$('report-print').addEventListener('click',()=>{if(reportData)window.print();});
for(const id of ['dec-sector','dec-amount'])$(id).addEventListener('change',()=>{decomposition=null;$('dec-results').hidden=true;});
for(const id of ['policy-a','policy-b','policy-budget','policy-split'])$(id).addEventListener('change',()=>{policyResults=null;$('policy-results').hidden=true;});
$('report-sector').addEventListener('change',()=>{reportData=null;$('print-report').hidden=true;$('report-print').disabled=true;});


// ============================================================
// V4 · 산업연관 네트워크: 연결 방향과 실제 Ad 계수를 보존
// ============================================================
const SVG_NS='http://www.w3.org/2000/svg';
function netSvg(tag,attributes={},text){
  const node=document.createElementNS(SVG_NS,tag);
  for(const [name,value] of Object.entries(attributes))node.setAttribute(name,String(value));
  if(text!==undefined)node.textContent=text;
  return node;
}
function netShort(name,max=18){return name.length>max?name.slice(0,max-1)+'…':name;}
function networkInspect(direction,entry){
  if(!networkResult)return;
  const base=networkResult.target;
  const text=direction==='supplier'
    ? `${entry.code} ${entry.name} → ${base.code} ${base.name}: Aᵈ[${entry.code}, ${base.code}] = ${fmtRaw(entry.coefficient)}. ${base.name} 생산액 1단위에 ${entry.name} 중간재 ${fmtRaw(entry.coefficient)}단위가 투입되는 구조입니다.`
    : `${base.code} ${base.name} → ${entry.code} ${entry.name}: Aᵈ[${base.code}, ${entry.code}] = ${fmtRaw(entry.coefficient)}. ${entry.name} 생산액 1단위에 ${base.name} 중간재 ${fmtRaw(entry.coefficient)}단위가 투입되는 구조입니다.`;
  putText('network-detail',text);
}
function drawNetwork(data){
  const left=data.suppliers, right=data.users;
  const n=Math.max(left.length,right.length,1);
  const width=920, top=97, step=62, height=Math.max(500,top+n*step+64);
  const svg=netSvg('svg',{viewBox:`0 0 ${width} ${height}`,role:'img','aria-label':`${data.target.name} 관련 주요 공급 산업과 사용 산업의 국산투입계수 네트워크`,class:'network-svg'});
  svg.append(netSvg('title',{},`${data.target.name} 산업 직접 투입관계`));
  svg.append(netSvg('text',{x:143,y:42,'text-anchor':'middle',class:'network-column-title'},'주요 공급 산업'));
  svg.append(netSvg('text',{x:460,y:42,'text-anchor':'middle',class:'network-column-title'},'중심 산업'));
  svg.append(netSvg('text',{x:777,y:42,'text-anchor':'middle',class:'network-column-title'},'주요 사용 산업'));
  const nodeY=i=>top+i*step;
  const midY=top+(n-1)*step/2;
  const leftX=16,leftWidth=256,centerX=331,centerWidth=258,rightX=648,rightWidth=256;
  const leftEnd=leftX+leftWidth,rightStart=rightX;
  const maxCoefficient=Math.max(...[...left,...right].map(x=>x.coefficient),0.00001);
  const edges=netSvg('g',{class:'network-edges'});
  for(const [side,entries] of [['supplier',left],['user',right]]){
    for(let i=0;i<entries.length;i++){
      const entry=entries[i],sy=side==='supplier'?nodeY(i):midY,ey=side==='supplier'?midY:nodeY(i);
      const sx=side==='supplier'?leftEnd:centerX+centerWidth;
      const ex=side==='supplier'?centerX:rightStart;
      const dx=(ex-sx)*0.53;
      const path=`M ${sx} ${sy} C ${sx+dx} ${sy}, ${ex-dx} ${ey}, ${ex} ${ey}`;
      const widthStroke=(1.6+5.4*Math.sqrt(entry.coefficient/maxCoefficient)).toFixed(2);
      const visible=netSvg('path',{d:path,fill:'none',stroke:side==='supplier'?'#178b94':'#4470ca','stroke-width':widthStroke,'stroke-opacity':'.57','stroke-linecap':'round',class:'network-connection'});
      edges.append(visible);
      const hit=netSvg('path',{d:path,fill:'none',stroke:'transparent','stroke-width':22,role:'button',tabindex:'0','aria-label':`${entry.name} ${side==='supplier'?'공급':'사용'} 연결 계수 ${fmtRaw(entry.coefficient)}`,class:'network-hit'});
      hit.append(netSvg('title',{},`${entry.code} ${entry.name} · ${fmtRaw(entry.coefficient)}`));
      hit.addEventListener('click',()=>networkInspect(side,entry));
      hit.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();networkInspect(side,entry);}});
      edges.append(hit);
    }
  }
  svg.append(edges);
  function node(x,y,w,label,name,sub,kind,entry){
    const group=netSvg('g',{class:'network-node node-'+kind,role:entry?'button':'group',tabindex:entry?'0':'-1','aria-label':`${label} ${name} ${sub}`});
    group.append(netSvg('rect',{x,y:y-24,width:w,height:48,rx:10,class:'network-node-rect'}));
    group.append(netSvg('text',{x:x+13,y:y-4,class:'network-node-code'},label));
    group.append(netSvg('text',{x:x+13,y:y+13,class:'network-node-label'},netShort(name,24)));
    group.append(netSvg('text',{x:x+w-12,y:y+8,'text-anchor':'end',class:'network-node-value'},sub));
    group.append(netSvg('title',{},`${label} ${name} ${sub}`));
    if(entry){
      group.addEventListener('click',()=>networkInspect(kind,entry));
      group.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();networkInspect(kind,entry);}});
    }
    svg.append(group);
  }
  left.forEach((entry,i)=>node(leftX,nodeY(i),leftWidth,entry.code,entry.name,fmtRaw(entry.coefficient),'supplier',entry));
  right.forEach((entry,i)=>node(rightX,nodeY(i),rightWidth,entry.code,entry.name,fmtRaw(entry.coefficient),'user',entry));
  node(centerX,midY,centerWidth,data.target.code,data.target.name,'FOCAL','center',null);
  if(!left.length)svg.append(netSvg('text',{x:138,y:midY,'text-anchor':'middle',class:'network-no-edges'},'조건에 맞는 공급 관계 없음'));
  if(!right.length)svg.append(netSvg('text',{x:776,y:midY,'text-anchor':'middle',class:'network-no-edges'},'조건에 맞는 사용 관계 없음'));
  $('network-canvas').replaceChildren(svg);
}
function renderNetwork(){
  const d=networkResult;if(!d)return;
  putText('network-heading',`${d.target.code} · ${d.target.name}의 국산 중간재 연결`);
  putText('network-detail',`중심 산업: ${d.target.name}. 왼쪽은 주요 공급 산업, 오른쪽은 주요 사용 산업입니다. 선이나 산업명을 선택하면 투입계수의 경제적 의미를 확인할 수 있습니다.`);
  const metricItems=[
    ['표시 공급 산업',`${d.suppliers.length}개`],
    ['표시 사용 산업',`${d.users.length}개`],
    ['중심 산업 국산중간재 투입계수 합',fmtRatio(d.domesticInputRatio)],
    ['자기 산업 투입계수',fmtRatio(d.selfCoefficient)]
  ];
  const f=document.createDocumentFragment();
  for(const [label,value] of metricItems){const box=el('div',undefined,'network-stat');box.append(el('p',label),el('strong',value));f.append(box);}
  $('network-kpis').replaceChildren(f);
  drawNetwork(d);
  function tableRows(id,entries,dir){
    if(!entries.length){emptyTbody(id,2,'선택 조건에 일치하는 관계가 없습니다.');return;}
    const frag=document.createDocumentFragment();
    for(const item of entries){
      const tr=el('tr');const button=el('button',item.code+' · '+item.name,'network-table-link');button.type='button';button.addEventListener('click',()=>networkInspect(dir,item));
      const td=el('td');td.append(button);tr.append(td,el('td',fmtRaw(item.coefficient),'number'));frag.append(tr);
    }
    $(id).replaceChildren(frag);
  }
  tableRows('network-suppliers',d.suppliers,'supplier');
  tableRows('network-users',d.users,'user');
  $('network-results').hidden=false;
}
$('network-run').addEventListener('click',async()=>{
  if(!activeLevel||v3Busy.network||loading)return;
  const code=$('network-sector').value, limit=Number($('network-limit').value), threshold=Number($('network-threshold').value), gen=generation;
  v3Busy.network=true;$('network-run').disabled=true;
  putText('network-status','2023년 국산투입계수에서 직접 공급·사용 관계를 계산하고 있습니다…');
  try{
    const d=await call('network',{code,limit,threshold});
    if(gen!==generation||code!==$('network-sector').value||limit!==Number($('network-limit').value)||threshold!==Number($('network-threshold').value))return;
    networkResult=d;renderNetwork();
    putText('network-status',`연결 관계 조회 완료 · 공급 ${d.supplierMatches}개/사용 ${d.userMatches}개 중 상위 ${d.limit}개씩 표시 (최소 계수 ${d.threshold}).`);
  }catch(err){putText('network-status','네트워크 생성 실패: '+err.message);}
  finally{v3Busy.network=false;$('network-run').disabled=!activeLevel;}
});
for(const id of ['network-sector','network-limit','network-threshold'])$(id).addEventListener('change',()=>{networkResult=null;$('network-results').hidden=true;if(activeLevel)putText('network-status','표시 조건이 변경되었습니다. 네트워크를 다시 생성하세요.');});
$('network-download').addEventListener('click',()=>{
  const d=networkResult;if(!d||!activeLevel)return;
  const rows=[['산업분류',cfg[activeLevel].label],['중심 산업',d.target.code+' '+d.target.name],['최소 계수',d.threshold],['단위','국산투입계수(무차원)'],[],['방향','공급 산업코드','공급 산업명','사용 산업코드','사용 산업명','국산투입계수']];
  for(const x of d.suppliers)rows.push(['공급 → 중심',x.code,x.name,d.target.code,d.target.name,x.coefficient]);
  for(const x of d.users)rows.push(['중심 → 사용',d.target.code,d.target.name,x.code,x.name,x.coefficient]);
  rows.push(['자기 산업',d.target.code,d.target.name,d.target.code,d.target.name,d.selfCoefficient]);
  downloadCsv('io-v4-'+activeLevel+'-'+d.target.code+'-network.csv',rows);
});


setReady(false);
