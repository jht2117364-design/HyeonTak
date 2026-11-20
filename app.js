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
const pageNames = {dashboard:'대시보드',explore:'데이터 탐색',simulate:'유발효과 분석',compare:'산업 간 비교',methods:'분석 방법 안내'};
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
  for(const id of ['sheet','row','row-search','col-search','sector-search','sector','amount','add','run','cmp-a','cmp-b','cmp-amount-a','cmp-amount-b','cmp-run']) $(id).disabled=!ready;
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
    const response=await fetch('./data/'+item.file);
    if(!response.ok)throw Error('Excel 다운로드 실패 (HTTP '+response.status+')');
    const buffer=await response.arrayBuffer();
    const info=await call('init',{buffer,count:item.count});
    if(info.count!==item.count||!Array.isArray(info.sectors)||info.sectors.length!==item.count)throw Error('산업 수가 일치하지 않습니다.');
    activeLevel=key;
    sectors=info.sectors;
    options($('sector'),sectors);
    options($('cmp-a'),sectors);
    options($('cmp-b'),sectors);
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
    putText('explore-count',`${info.count}개 산업 · ${info.sheets.length}개 시트`);
    putText('cmp-status','산업 A와 B를 선택하고 비교를 시작하세요.');
    message(`${item.label} ${item.count}개 산업 / ${info.sheets.length}개 시트 · 행렬 검증 통과`,'good');
  }catch(error){
    activeLevel=null;
    sectors=[];
    setReady(false);
    message('불러오기 실패: '+error.message+' — data/ 파일과 인터넷 연결을 확인하세요.','bad');
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

setReady(false);
