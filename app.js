
(()=>{
'use strict';

const STORAGE='reports.scenarios.v2';
const SCRIPT='docbuilder/scripts/reports_executor.docbuilder';
const PROBE_TTL=60000;
const PENDING=Object.create(null);
const WATCH=Object.create(null);

const state={scenarios:[],q:'',draft:null,running:null,probe:null,toastTimer:null};

const schema={
  set_cell_value:{label:'Вставить значение',d:{sheet:'Лист1',range:'A1',mode:'text',value:'',merge:false},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Ячейка/диапазон',t:'text',r:1,p:'A1 или A1:C10'},
    {k:'mode',l:'Тип',t:'select',o:[['text','Текст'],['number','Число'],['formula','Формула'],['bool','Логическое']]},
    {k:'value',l:'Значение',t:'textarea',r:1,full:1,p:'Текст, число или формула'},
    {k:'merge',l:'Объединить после записи',t:'check',full:1}
  ]},
  clear_range:{label:'Очистить диапазон',d:{sheet:'Лист1',range:'A1',mode:'contents'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},
    {k:'mode',l:'Что очищать',t:'select',o:[['contents','Только значения'],['formats','Только формат'],['hyperlinks','Только ссылки'],['all','Все']]}
  ]},
  formula_to_value:{label:'Формулы в значения',d:{sheet:'Лист1',range:'A1:Z200'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1}
  ]},
  set_border:{label:'Границы диапазона',d:{sheet:'Лист1',range:'A1:C3',scope:'all',style:'Thin',color:'#000000'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},
    {k:'scope',l:'Где ставить',t:'select',o:[['all','Все'],['outer','Внешние'],['inner','Внутренние']]},
    {k:'style',l:'Стиль',t:'select',o:[['Thin','Тонкая'],['Medium','Средняя'],['Thick','Толстая'],['Dashed','Пунктир'],['Dotted','Точки'],['Double','Двойная'],['None','Убрать']]},
    {k:'color',l:'Цвет',t:'color'}
  ]},
  merge_range:{label:'Объединить диапазон',d:{sheet:'Лист1',range:'A1:B1',across:false},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},{k:'across',l:'По строкам (Across)',t:'check',full:1}
  ]},
  unmerge_range:{label:'Разъединить диапазон',d:{sheet:'Лист1',range:'A1:B1'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1}
  ]},
  insert_rows:{label:'Вставить строки',d:{sheet:'Лист1',start_row:'2',count:'1'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'start_row',l:'С какой строки',t:'number',r:1,min:1,step:1},{k:'count',l:'Сколько строк',t:'number',r:1,min:1,step:1}
  ]},
  delete_rows:{label:'Удалить строки',d:{sheet:'Лист1',start_row:'2',count:'1'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'start_row',l:'С какой строки',t:'number',r:1,min:1,step:1},{k:'count',l:'Сколько строк',t:'number',r:1,min:1,step:1}
  ]},
  set_rows_hidden:{label:'Скрыть/показать строки',d:{sheet:'Лист1',start_row:'2',end_row:'5',hidden:true},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'start_row',l:'Начальная строка',t:'number',r:1,min:1,step:1},{k:'end_row',l:'Конечная строка',t:'number',r:1,min:1,step:1},{k:'hidden',l:'Скрыть (иначе показать)',t:'check',full:1}
  ]},
  insert_columns:{label:'Вставить столбцы',d:{sheet:'Лист1',start_column:'B',count:'1'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'start_column',l:'С какого столбца',t:'text',r:1,p:'A, B, AA'},{k:'count',l:'Сколько столбцов',t:'number',r:1,min:1,step:1}
  ]},
  delete_columns:{label:'Удалить столбцы',d:{sheet:'Лист1',start_column:'B',count:'1'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'start_column',l:'С какого столбца',t:'text',r:1,p:'A, B, AA'},{k:'count',l:'Сколько столбцов',t:'number',r:1,min:1,step:1}
  ]},
  set_columns_hidden:{label:'Скрыть/показать столбцы',d:{sheet:'Лист1',start_column:'C',end_column:'E',hidden:true},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'start_column',l:'Начальный столбец',t:'text',r:1,p:'A, B, AA'},{k:'end_column',l:'Конечный столбец',t:'text',r:1,p:'A, B, AA'},{k:'hidden',l:'Скрыть (иначе показать)',t:'check',full:1}
  ]},
  group_rows:{label:'Сгруппировать строки',d:{sheet:'Лист1',start_row:'2',end_row:'10',collapsed:false},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'start_row',l:'Начальная строка',t:'number',r:1,min:1,step:1},{k:'end_row',l:'Конечная строка',t:'number',r:1,min:1,step:1},{k:'collapsed',l:'Свернуть после группировки',t:'check',full:1}
  ]},
  group_columns:{label:'Сгруппировать столбцы',d:{sheet:'Лист1',start_column:'C',end_column:'E',collapsed:false},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'start_column',l:'Начальный столбец',t:'text',r:1,p:'A, B, AA'},{k:'end_column',l:'Конечный столбец',t:'text',r:1,p:'A, B, AA'},{k:'collapsed',l:'Свернуть после группировки',t:'check',full:1}
  ]},
  autofit:{label:'Автоподбор размеров',d:{sheet:'Лист1',range:'A:Z',rows:true,cols:true},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},{k:'rows',l:'Автоподбор высоты строк',t:'check',full:1},{k:'cols',l:'Автоподбор ширины столбцов',t:'check',full:1}
  ]},
  add_sheet:{label:'Добавить лист',d:{sheet_name:'НовыйЛист'},f:[{k:'sheet_name',l:'Имя нового листа',t:'text',r:1}]},
  rename_sheet:{label:'Переименовать лист',d:{sheet:'Лист1',new_name:'Лист2'},f:[{k:'sheet',l:'Текущий лист',t:'text',r:1},{k:'new_name',l:'Новое имя',t:'text',r:1}]},
  delete_sheet:{label:'Удалить лист',d:{sheet:'Лист1'},f:[{k:'sheet',l:'Лист для удаления',t:'text',r:1}]},
  set_font_style:{label:'Шрифт и стиль',d:{sheet:'Лист1',range:'A1',font_name:'Arial',font_size:'11',bold:false,italic:false,underline:'none',strikeout:false,font_color:'#000000'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1,p:'A1 или A1:C10'},
    {k:'font_name',l:'Шрифт',t:'text',p:'Например: Arial'},
    {k:'font_size',l:'Размер шрифта',t:'number',min:1,step:0.5},
    {k:'underline',l:'Подчеркивание',t:'select',o:[['none','Нет'],['single','Одинарное'],['double','Двойное'],['singleAccounting','Одинарное (учетное)'],['doubleAccounting','Двойное (учетное)']]},
    {k:'font_color',l:'Цвет шрифта',t:'color'},
    {k:'bold',l:'Полужирный',t:'check'},{k:'italic',l:'Курсив',t:'check'},{k:'strikeout',l:'Зачеркнутый',t:'check'}
  ]},
  set_fill_color:{label:'Заливка диапазона',d:{sheet:'Лист1',range:'A1',color:'#FFFFFF'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},
    {k:'color',l:'Цвет заливки',t:'color',r:1,none:1}
  ]},
  set_number_format:{label:'Формат числа',d:{sheet:'Лист1',range:'A1',format:'General'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},
    {k:'format',l:'Маска формата',t:'text',r:1,p:'General, 0.00, #,##0, dd.mm.yyyy, @'}
  ]},
  set_alignment:{label:'Выравнивание и перенос',d:{sheet:'Лист1',range:'A1',horizontal:'left',vertical:'top',wrap:false,orientation:'0',reading_order:'context'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},
    {k:'horizontal',l:'Горизонталь',t:'select',o:[['left','Слева'],['center','По центру'],['right','Справа'],['justify','По ширине']]},
    {k:'vertical',l:'Вертикаль',t:'select',o:[['top','Сверху'],['center','По центру'],['bottom','Снизу'],['justify','По высоте'],['distributed','Распределить']]},
    {k:'orientation',l:'Поворот текста',t:'text',p:'-90..90, 255 или xlHorizontal/xlVertical'},
    {k:'reading_order',l:'Направление текста',t:'select',o:[['context','По контексту'],['ltr','Слева направо'],['rtl','Справа налево']]},
    {k:'wrap',l:'Переносить текст',t:'check',full:1}
  ]},
  set_row_height:{label:'Высота строк',d:{sheet:'Лист1',range:'2:2',height:'20'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Строки',t:'text',r:1,p:'2:10'},
    {k:'height',l:'Высота (pt)',t:'number',r:1,min:0.5,step:0.5}
  ]},
  set_column_width:{label:'Ширина столбцов',d:{sheet:'Лист1',range:'A:A',width:'12'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Столбцы',t:'text',r:1,p:'A:C'},
    {k:'width',l:'Ширина',t:'number',r:1,min:0.5,step:0.5}
  ]},
  copy_paste_range:{label:'Копировать/вставить диапазон',d:{from_sheet:'Лист1',from_range:'A1:C5',to_sheet:'Лист1',to_cell:'E1',paste_type:'xlPasteAll',operation:'xlPasteSpecialOperationNone',skip_blanks:false,transpose:false},f:[
    {k:'from_sheet',l:'Исходный лист',t:'text',r:1},{k:'from_range',l:'Исходный диапазон',t:'text',r:1,p:'A1:C5'},
    {k:'to_sheet',l:'Лист назначения',t:'text',r:1},{k:'to_cell',l:'Вставить с ячейки',t:'text',r:1,p:'E1'},
    {k:'paste_type',l:'Тип вставки',t:'select',o:[['xlPasteAll','Все'],['xlPasteFormats','Только формат'],['xlPasteValues','Только значения'],['xlPasteFormulas','Только формулы'],['xlPasteValuesAndNumberFormats','Значения и формат числа'],['xlPasteFormulasAndNumberFormats','Формулы и формат числа'],['xlPasteComments','Комментарии'],['xlPasteColumnWidths','Ширины столбцов'],['xlPasteAllExceptBorders','Все без границ']]},
    {k:'operation',l:'Операция',t:'select',o:[['xlPasteSpecialOperationNone','Без операции'],['xlPasteSpecialOperationAdd','Сложить'],['xlPasteSpecialOperationSubtract','Вычесть'],['xlPasteSpecialOperationMultiply','Умножить'],['xlPasteSpecialOperationDivide','Разделить']]},
    {k:'skip_blanks',l:'Пропустить пустые',t:'check'},{k:'transpose',l:'Транспонировать',t:'check'}
  ]},
  sort_range:{label:'Сортировка диапазона',d:{sheet:'Лист1',range:'A1:D20',key1:'A2:A20',order1:'asc',key2:'',order2:'asc',key3:'',order3:'asc',header:true,sort_by:'rows'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Сортируемый диапазон',t:'text',r:1,p:'A1:D20'},
    {k:'key1',l:'Ключ 1 (диапазон)',t:'text',r:1,p:'A2:A20 или B2:B20'},{k:'order1',l:'Порядок 1',t:'select',o:[['asc','По возрастанию'],['desc','По убыванию']]},
    {k:'key2',l:'Ключ 2 (опц.)',t:'text',p:'B2:B20'},{k:'order2',l:'Порядок 2',t:'select',o:[['asc','По возрастанию'],['desc','По убыванию']]},
    {k:'key3',l:'Ключ 3 (опц.)',t:'text',p:'C2:C20'},{k:'order3',l:'Порядок 3',t:'select',o:[['asc','По возрастанию'],['desc','По убыванию']]},
    {k:'sort_by',l:'Ориентация сортировки',t:'select',o:[['rows','По строкам'],['columns','По столбцам']]},
    {k:'header',l:'Первая строка - заголовки',t:'check',full:1}
  ]},
  set_autofilter:{label:'Автофильтр',d:{sheet:'Лист1',range:'A1:D20',field:'',criteria1:'',operator:'',criteria2:'',visible_drop_down:true},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон с заголовками',t:'text',r:1,p:'A1:D20'},
    {k:'field',l:'Поле (№ столбца)',t:'number',min:1,step:1,p:'1 = первый столбец'},
    {k:'criteria1',l:'Критерий 1',t:'text',p:'Например >100 или A;B;C для xlFilterValues'},
    {k:'operator',l:'Оператор',t:'select',o:[['','(не задан)'],['xlAnd','И'],['xlOr','ИЛИ'],['xlFilterValues','Список значений'],['xlTop10Items','TOP 10 (шт.)'],['xlTop10Percent','TOP 10 (%)'],['xlBottom10Items','BOTTOM 10 (шт.)'],['xlBottom10Percent','BOTTOM 10 (%)'],['xlFilterCellColor','По цвету заливки'],['xlFilterFontColor','По цвету шрифта'],['xlFilterDynamic','Динамический']]},
    {k:'criteria2',l:'Критерий 2 (для And/Or)',t:'text'},
    {k:'visible_drop_down',l:'Показывать стрелку фильтра',t:'check',full:1}
  ]},
  set_array_formula:{label:'Массивная формула',d:{sheet:'Лист1',range:'A1:C1',formula:'=A1:C1*2'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},
    {k:'formula',l:'Формула массива',t:'textarea',r:1,full:1,p:'Например: =SUM(A1:A10*B1:B10)'}
  ]},
  set_gridlines:{label:'Сетка листа',d:{sheet:'Лист1',show:true},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'show',l:'Показывать линии сетки',t:'check',full:1}
  ]}
};

const THEME_COLORS=[
  ['#FFFFFF','#000000','#1F1F1F','#EEECE1','#4F81BD','#C0504D','#9BBB59','#8064A2','#4BACC6','#F79646'],
  ['#F2F2F2','#7F7F7F','#D8D8D8','#DDD9C3','#DBE5F1','#F2DCDB','#EBF1DE','#E5E0EC','#DBEEF3','#FDEADA'],
  ['#D9D9D9','#595959','#BFBFBF','#C4BD97','#B8CCE4','#E5B9B7','#D7E3BC','#CCC1D9','#B7DEE8','#FBD5B5'],
  ['#BFBFBF','#3F3F3F','#A5A5A5','#938953','#95B3D7','#D99694','#C3D69B','#B2A2C7','#92CDDC','#FAC08F'],
  ['#A6A6A6','#262626','#7F7F7F','#494429','#366092','#953734','#76923C','#5F497A','#31859B','#E36C09'],
  ['#7F7F7F','#0C0C0C','#595959','#1D1B10','#244061','#632423','#4F6128','#3F3151','#205867','#974806']
];
const STANDARD_COLORS=['#C00000','#FF0000','#FFC000','#FFFF00','#92D050','#00B050','#00B0F0','#0070C0','#002060','#7030A0'];

const typeList=Object.keys(schema).map(k=>[k,schema[k].label]);
const els={
  list:document.getElementById('scenario-list'),empty:document.getElementById('empty-state'),search:document.getElementById('search-input'),
  btnCreate:document.getElementById('btn-create'),modal:document.getElementById('scenario-modal'),title:document.getElementById('modal-title'),
  btnClose:document.getElementById('btn-close-modal'),btnCancel:document.getElementById('btn-cancel'),btnSave:document.getElementById('btn-save'),
  name:document.getElementById('scenario-name'),tpl:document.getElementById('scenario-template'),descr:document.getElementById('scenario-description'),
  btnPick:document.getElementById('btn-pick-template'),file:document.getElementById('template-file-input'),actionList:document.getElementById('action-list'),
  btnAddAction:document.getElementById('btn-add-action'),toast:document.getElementById('toast')
};

function uid(p){return `${p}-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;}
function clone(v){return JSON.parse(JSON.stringify(v));}
function parse(v){try{return typeof v==='string'?JSON.parse(v):v;}catch(_){return null;}}
function isObj(v){return v&&typeof v==='object'&&!Array.isArray(v);}
function dt(v){if(!v)return '-';const d=new Date(v);if(Number.isNaN(d.getTime()))return String(v);return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;}
function toast(msg,type){if(!els.toast)return; if(state.toastTimer)clearTimeout(state.toastTimer); if(!msg){els.toast.className='toast hidden';els.toast.textContent='';return;} els.toast.textContent=msg;els.toast.className=`toast ${type||''}`.trim();state.toastTimer=setTimeout(()=>{els.toast.className='toast hidden';},4500);}
function normSlashes(p){return String(p||'').replace(/\//g,'\\');}
function joinPath(){const a=[...arguments].filter(Boolean).map(normSlashes);if(!a.length)return '';let out=a[0];for(let i=1;i<a.length;i++)out=`${out.replace(/\\+$/,'')}\\${a[i].replace(/^\\+/,'')}`;return out.replace(/\\{2,}/g,'\\');}
function dirName(p){const s=normSlashes(p);const i=s.lastIndexOf('\\');return i>0?s.slice(0,i):'';}
function baseName(p){const s=normSlashes(p);const i=s.lastIndexOf('\\');return i>=0?s.slice(i+1):s;}
function getRoot(){
  try{if(window.parent&&window.parent.reportsUiRoot)return normSlashes(window.parent.reportsUiRoot);}catch(_){ }
  try{
    let p=decodeURIComponent(window.location.pathname||'');
    if(p.startsWith('/')&&/^[a-zA-Z]:/.test(p.slice(1)))p=p.slice(1);
    p=p.replace(/\\/g,'/');
    if(p.includes('/reports-ui/'))return normSlashes(p.split('/reports-ui/')[0]);
  }catch(_){ }
  return '';
}
function toFsPath(p){
  let s=String(p||'').trim(); if(!s)return '';
  if(/^file:\/\//i.test(s)){s=decodeURIComponent(s.replace(/^file:\/\//i,''));if(/^\/[a-zA-Z]:/.test(s))s=s.slice(1);} 
  return normSlashes(s);
}
function resolveTemplate(raw){
  const p=toFsPath(raw); if(!p)return '';
  if(/^[a-zA-Z]:\\/.test(p)||/^\\\\/.test(p))return p;
  const rel=p.replace(/^\.\\/,'').replace(/^reports-ui\\/i,'');
  const root=getRoot();
  return root?joinPath(root,'reports-ui',rel):rel;
}
function fmtStamp(){const d=new Date();return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}-${String(d.getHours()).padStart(2,'0')}${String(d.getMinutes()).padStart(2,'0')}${String(d.getSeconds()).padStart(2,'0')}`;}
function safeFile(n){
  const raw=String(n||'scenario');
  const s=raw
    .replace(/[^\x00-\x7F]/g,'_')
    .replace(/[<>:"/\\|?*]+/g,'_')
    .replace(/\s+/g,'_')
    .replace(/_+/g,'_')
    .replace(/^_+|_+$/g,'');
  return s||'scenario';
}
function mkOutputPath(sc,probe,tpl){const runtime=probe&&probe.runtimeDir?normSlashes(probe.runtimeDir):'';const outDir=runtime?joinPath(runtime,'temp'):dirName(tpl);return joinPath(outDir,`${safeFile(sc.name)}-${fmtStamp()}.xlsx`);}
function extType(path){
  const ext=(String(path||'').split('.').pop()||'xlsx').toLowerCase();
  const map={docx:65,doc:66,odt:67,rtf:68,txt:69,pdf:513,xlsx:257,xls:258,ods:259,csv:260,xlsm:261,xltx:262,xltm:263,xlsb:264,pptx:129,ppt:130,odp:131};
  try{
    if(window.parent&&window.parent.utils&&window.parent.utils.fileExtensionToFileFormat){
      const t=Number(window.parent.utils.fileExtensionToFileFormat(ext)||0);
      if(t>0)return t;
    }
  }catch(_){ }
  return map[ext]||257;
}

function expandHex3(v){return `#${v[1]}${v[1]}${v[2]}${v[2]}${v[3]}${v[3]}`;}
function normColor(v,allowNone){
  let s=String(v==null?'':v).trim();
  if(!s)return allowNone?'none':'#000000';
  if(allowNone&&/^none$/i.test(s))return 'none';
  if(!s.startsWith('#'))s=`#${s}`;
  if(/^#[0-9a-fA-F]{3}$/.test(s))s=expandHex3(s);
  if(!/^#[0-9a-fA-F]{6}$/.test(s))return allowNone?'none':'#000000';
  return s.toUpperCase();
}
function closeColorMenus(){document.querySelectorAll('.color-menu.open').forEach(x=>x.classList.remove('open'));}
function swatchState(node,value){
  if(!node)return;
  if(String(value).toLowerCase()==='none'){node.style.background='transparent';node.classList.add('is-none');}
  else{node.style.background=value;node.classList.remove('is-none');}
}
function appendColorGrid(container,colors,onPick){
  const g=document.createElement('div');g.className='color-grid';
  colors.forEach(c=>{const b=document.createElement('button');b.type='button';b.className='color-cell';b.title=c;b.style.background=c;b.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();onPick(c);};g.appendChild(b);});
  container.appendChild(g);
}
function createColorField(action,field){
  const allowNone=!!field.none;
  const box=document.createElement('div');box.className='color-field';
  const txt=document.createElement('input');txt.type='text';txt.className='color-value';txt.value=normColor(action[field.k],allowNone);if(field.p)txt.placeholder=field.p;
  const btn=document.createElement('button');btn.type='button';btn.className='color-open';
  const dot=document.createElement('span');dot.className='color-dot';btn.appendChild(dot);
  const menu=document.createElement('div');menu.className='color-menu';
  const t1=document.createElement('div');t1.className='color-section-title';t1.textContent='Цвета темы';menu.appendChild(t1);
  THEME_COLORS.forEach(row=>appendColorGrid(menu,row,applyColor));
  const t2=document.createElement('div');t2.className='color-section-title';t2.textContent='Стандартные цвета';menu.appendChild(t2);
  appendColorGrid(menu,STANDARD_COLORS,applyColor);
  const actions=document.createElement('div');actions.className='color-actions';
  if(allowNone){const noneBtn=document.createElement('button');noneBtn.type='button';noneBtn.className='btn btn-ghost color-action';noneBtn.textContent='Без заливки';noneBtn.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();applyColor('none');};actions.appendChild(noneBtn);}
  const moreBtn=document.createElement('button');moreBtn.type='button';moreBtn.className='btn btn-ghost color-action';moreBtn.textContent='Другие цвета';
  const picker=document.createElement('input');picker.type='color';picker.className='color-native';
  moreBtn.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();picker.click();};
  picker.oninput=(ev)=>applyColor(ev.target.value);
  actions.appendChild(moreBtn);actions.appendChild(picker);menu.appendChild(actions);

  function applyColor(value){
    const n=normColor(value,allowNone);
    action[field.k]=n;
    txt.value=n;
    swatchState(dot,n);
    closeColorMenus();
  }
  txt.oninput=(ev)=>{const raw=String(ev.target.value||'');if(/^#?[0-9a-fA-F]{0,6}$/.test(raw)||/^none$/i.test(raw.trim()))action[field.k]=raw;};
  txt.onblur=()=>{applyColor(txt.value);};
  btn.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();const wasOpen=menu.classList.contains('open');closeColorMenus();if(!wasOpen)menu.classList.add('open');};
  menu.onclick=(ev)=>ev.stopPropagation();
  applyColor(txt.value);
  box.appendChild(txt);box.appendChild(btn);box.appendChild(menu);
  return box;
}

function mkAction(type){const key=schema[type]?type:'set_cell_value';const a={id:uid('action'),type:key};Object.assign(a,schema[key].d||{});return a;}
function normAction(a){if(!isObj(a))return mkAction('set_cell_value');const t=schema[a.type]?a.type:'set_cell_value';const o={id:a.id||uid('action'),type:t};Object.assign(o,schema[t].d||{},a);return o;}
function normScenario(s){if(!isObj(s))return null;return{ id:s.id||uid('scenario'), name:String(s.name||'Без названия'), description:String(s.description||s.descr||''), templatePath:String(s.templatePath||s.path||''), actions:Array.isArray(s.actions)?s.actions.map(normAction):[], createdAt:s.createdAt||new Date().toISOString(), updatedAt:s.updatedAt||s.updated||new Date().toISOString(), lastRunAt:s.lastRunAt||s.last_run||null};}
function save(){try{localStorage.setItem(STORAGE,JSON.stringify(state.scenarios));}catch(_){toast('Не удалось сохранить сценарии.','error');}}
function load(){const raw=parse(localStorage.getItem(STORAGE));state.scenarios=Array.isArray(raw)?raw.map(normScenario).filter(Boolean):[];}

function openModal(id){
  const src=state.scenarios.find(x=>x.id===id);
  state.draft=src?clone(src):{id:uid('scenario'),name:'',description:'',templatePath:'',actions:[mkAction('set_cell_value')],createdAt:new Date().toISOString(),updatedAt:new Date().toISOString(),lastRunAt:null};
  els.title.textContent=src?'Изменение сценария':'Новый сценарий';
  els.name.value=state.draft.name; els.tpl.value=state.draft.templatePath; els.descr.value=state.draft.description;
  renderActions(); els.modal.classList.remove('hidden'); els.name.focus();
}
function closeModal(){els.modal.classList.add('hidden');state.draft=null;}

function renderList(){
  const q=state.q.trim().toLowerCase();
  const list=state.scenarios.slice().sort((a,b)=>new Date(b.updatedAt)-new Date(a.updatedAt)).filter(s=>!q||(`${s.name} ${s.description} ${s.templatePath}`.toLowerCase().includes(q)));
  els.list.innerHTML='';
  els.empty.hidden=list.length>0;
  list.forEach(s=>{
    const c=document.createElement('article');c.className='scenario-card';
    c.innerHTML=`<div class='scenario-card-head'><h3 class='scenario-card-title'></h3><div class='scenario-card-template'></div></div><div class='scenario-card-body'><p class='scenario-card-desc'></p><div class='scenario-meta'></div><div class='scenario-actions'></div></div>`;
    c.querySelector('.scenario-card-title').textContent=s.name;
    c.querySelector('.scenario-card-template').textContent=s.templatePath||'Шаблон не задан';
    c.querySelector('.scenario-card-desc').textContent=s.description||'Без описания';
    const m=c.querySelector('.scenario-meta');
    [['Действий',s.actions.length],['Изменен',dt(s.updatedAt)],['Запуск',dt(s.lastRunAt)]].forEach(x=>{const ch=document.createElement('span');ch.className='meta-chip';ch.textContent=`${x[0]}: ${x[1]}`;m.appendChild(ch);});
    const a=c.querySelector('.scenario-actions');
    const bRun=document.createElement('button');bRun.className='btn btn-primary';bRun.type='button';bRun.textContent=state.running===s.id?'Выполняется...':'Заполнить';bRun.disabled=!!state.running;bRun.onclick=()=>runScenario(s.id);
    const bEdit=document.createElement('button');bEdit.className='btn btn-secondary';bEdit.type='button';bEdit.textContent='Изменить';bEdit.onclick=()=>openModal(s.id);
    const bDel=document.createElement('button');bDel.className='btn btn-danger';bDel.type='button';bDel.textContent='Удалить';bDel.onclick=()=>{if(confirm(`Удалить сценарий "${s.name}"?`)){state.scenarios=state.scenarios.filter(x=>x.id!==s.id);save();renderList();toast('Сценарий удален.','ok');}};
    [bRun,bEdit,bDel].forEach(x=>a.appendChild(x));
    els.list.appendChild(c);
  });
}

function fieldWrap(full){const w=document.createElement('label');w.className=full?'field field-full':'field';return w;}
function renderActions(){
  const root=els.actionList; root.innerHTML='';
  if(!state.draft||!state.draft.actions.length){const e=document.createElement('div');e.className='action-item';e.textContent='Добавьте хотя бы одно действие.';root.appendChild(e);return;}
  state.draft.actions.forEach((a,i)=>{
    const item=document.createElement('div');item.className='action-item';
    const head=document.createElement('div');head.className='action-item-head';
    const ord=document.createElement('span');ord.className='order';ord.textContent=String(i+1);
    const sel=document.createElement('select');typeList.forEach(t=>{const o=document.createElement('option');o.value=t[0];o.textContent=t[1];o.selected=t[0]===a.type;sel.appendChild(o);});
    sel.onchange=(ev)=>{const n=mkAction(ev.target.value);n.id=a.id;if('sheet' in n&&a.sheet)n.sheet=a.sheet;state.draft.actions[i]=n;renderActions();};
    const tools=document.createElement('div');tools.className='action-tools';
    const up=document.createElement('button');up.className='btn btn-ghost';up.type='button';up.textContent='↑';up.disabled=i===0;up.onclick=()=>{const t=state.draft.actions[i-1];state.draft.actions[i-1]=state.draft.actions[i];state.draft.actions[i]=t;renderActions();};
    const dn=document.createElement('button');dn.className='btn btn-ghost';dn.type='button';dn.textContent='↓';dn.disabled=i===state.draft.actions.length-1;dn.onclick=()=>{const t=state.draft.actions[i+1];state.draft.actions[i+1]=state.draft.actions[i];state.draft.actions[i]=t;renderActions();};
    const rm=document.createElement('button');rm.className='btn btn-danger';rm.type='button';rm.textContent='Удалить';rm.onclick=()=>{state.draft.actions.splice(i,1);renderActions();};
    [up,dn,rm].forEach(x=>tools.appendChild(x)); [ord,sel,tools].forEach(x=>head.appendChild(x)); item.appendChild(head);
    const fs=document.createElement('div');fs.className='action-fields';
    (schema[a.type].f||[]).forEach(f=>{
      const w=fieldWrap(!!f.full);
      if(f.t==='check'){
        const inl=document.createElement('span');inl.className='field-inline';const inp=document.createElement('input');inp.type='checkbox';inp.checked=!!a[f.k];inp.onchange=(ev)=>{a[f.k]=!!ev.target.checked;};const t=document.createElement('span');t.textContent=f.l;inl.appendChild(inp);inl.appendChild(t);w.appendChild(inl);fs.appendChild(w);return;
      }
      const cap=document.createElement('span');cap.textContent=f.l;w.appendChild(cap);
      let inp;
      if(f.t==='textarea'){inp=document.createElement('textarea');inp.rows=2;inp.value=a[f.k]==null?'':String(a[f.k]);}
      else if(f.t==='select'){inp=document.createElement('select');(f.o||[]).forEach(op=>{const o=document.createElement('option');o.value=op[0];o.textContent=op[1];o.selected=String(op[0])===String(a[f.k]);inp.appendChild(o);});}
      else if(f.t==='color'){inp=createColorField(a,f);}
      else {inp=document.createElement('input');inp.type=f.t==='number'?'number':'text';inp.value=a[f.k]==null?'':String(a[f.k]);if(f.min!=null)inp.min=String(f.min);if(f.step!=null)inp.step=String(f.step);}
      if(f.t!=='color'){if(f.p)inp.placeholder=f.p;inp.oninput=(ev)=>{a[f.k]=ev.target.value;};}
      w.appendChild(inp);fs.appendChild(w);
    });
    item.appendChild(fs); root.appendChild(item);
  });
}
function readDraft(){if(!state.draft)return;state.draft.name=String(els.name.value||'').trim();state.draft.templatePath=String(els.tpl.value||'').trim();state.draft.description=String(els.descr.value||'').trim();}
function validateDraft(){
  if(!state.draft)return 'Сценарий не открыт.'; readDraft();
  if(!state.draft.name)return 'Введите название сценария.';
  if(!state.draft.templatePath)return 'Выберите файл шаблона.';
  if(/fakepath/i.test(state.draft.templatePath))return 'Получен fakepath. Укажите полный путь к шаблону.';
  if(!Array.isArray(state.draft.actions)||!state.draft.actions.length)return 'Добавьте хотя бы одно действие.';
  for(let i=0;i<state.draft.actions.length;i++){
    const a=state.draft.actions[i];if(!schema[a.type])return `Шаг ${i+1}: неизвестный тип действия.`;
    for(const f of schema[a.type].f||[]){if(!f.r||f.t==='check')continue;const v=a[f.k];if(v==null||String(v).trim()==='')return `Шаг ${i+1} (${schema[a.type].label}): заполните поле «${f.l}».`;}
  }
  return '';
}
function saveDraft(){
  const err=validateDraft();if(err){toast(err,'error');return;}
  const s={id:state.draft.id,name:state.draft.name,description:state.draft.description,templatePath:state.draft.templatePath,actions:state.draft.actions.map(normAction),createdAt:state.draft.createdAt||new Date().toISOString(),updatedAt:new Date().toISOString(),lastRunAt:state.draft.lastRunAt||null};
  const idx=state.scenarios.findIndex(x=>x.id===s.id);if(idx>=0){state.scenarios[idx]=s;toast('Сценарий обновлен.','ok');}else{state.scenarios.push(s);toast('Сценарий создан.','ok');}
  save();renderList();closeModal();
}

function post(message){try{if(window.parent&&window.parent!==window)window.parent.postMessage(JSON.stringify(message),'*');}catch(_){ }}
function reqId(p){return `${p}-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;}
function parseNativeParam(param){
  const obj=parse(param);
  if(isObj(obj))return obj;
  try{
    const div=document.createElement('div');
    div.innerHTML=String(param||'');
    const decoded=div.textContent||div.innerText||'';
    const obj2=parse(decoded);
    if(isObj(obj2))return obj2;
  }catch(_){ }
  return null;
}
function boolVal(v){
  if(typeof v==='boolean')return v;
  if(typeof v==='number')return v!==0;
  if(typeof v==='string'){
    const s=v.trim().toLowerCase();
    if(s==='true'||s==='1')return true;
    if(s==='false'||s==='0'||s==='')return false;
    try{return !!JSON.parse(s);}catch(_){return false;}
  }
  return !!v;
}
function clearWatch(id){const w=WATCH[id];if(!w)return; if(w.t)clearInterval(w.t); delete WATCH[id];}
function clearPending(id){const p=PENDING[id];if(!p)return null;if(p.t)clearTimeout(p.t);delete PENDING[id];clearWatch(id);return p;}
function addPending(id,res,rej,ms){PENDING[id]={res,rej,t:setTimeout(()=>{const p=clearPending(id);if(!p)return;const e=new Error('Таймаут DocumentBuilder');e.details={requestId:id,error:'timeout'};p.rej(e);},ms)}}
function dbErr(r){const e=new Error((r&&r.error)||'Ошибка DocumentBuilder');e.details=r||{};return e;}
function handleFilesChecked(param){
  const data=parseNativeParam(param);if(!isObj(data))return;
  Object.keys(WATCH).forEach(id=>{
    const w=WATCH[id];if(!w)return;
    if(!Object.prototype.hasOwnProperty.call(data,w.k))return;
    if(!boolVal(data[w.k]))return;
    if(w.openAfterRun){clearWatch(id);return;}
    clearWatch(id);
    onDbResult({ok:true,requestId:id,exitCode:0,via:'files:checked'});
  });
}
function bindDirectNative(){
  try{
    if(window.__reportsDirectNativeBound)return true;
    const sdk=window.parent&&window.parent.sdk;
    if(!sdk||typeof sdk.on!=='function')return false;
    sdk.on('on_native_message',(cmd,param)=>{
      const c=String(cmd||'').trim();
      if(c==='docbuilder:result'||c==='docbuilder:probeResult'){const data=parseNativeParam(param);if(data)onDbResult(data);return;}
      if(c==='files:checked'){handleFilesChecked(param);}
    });
    window.__reportsDirectNativeBound=true;
    return true;
  }catch(_){return false;}
}
function startWatch(id,payload,ms){
  clearWatch(id);
  const out=isObj(payload)&&isObj(payload.argument)&&payload.argument.outputPath?String(payload.argument.outputPath):'';
  if(!out)return;
  const key=`reports_db_${String(id).replace(/[^a-zA-Z0-9_-]/g,'_')}`;
  const started=Date.now();
  const tick=()=>{
    if(!WATCH[id])return;
    if(Date.now()-started>(ms+30000)){clearWatch(id);return;}
    try{
      const sdk=window.parent&&window.parent.sdk;
      if(!sdk||typeof sdk.command!=='function')return;
      const map={}; map[key]=out; sdk.command('files:check',JSON.stringify(map));
    }catch(_){ }
  };
  WATCH[id]={k:key,t:setInterval(tick,1000),openAfterRun:!!(isObj(payload)&&payload.openAfterRun)};
  tick();
}
function requestDb(event,payload,ms){
  const d=isObj(payload)?Object.assign({},payload):{};const id=d.requestId||reqId('docbuilder');d.requestId=id;
  return new Promise((resolve,reject)=>{
    addPending(id,resolve,reject,ms);
    try{
      const sdk=window.parent&&window.parent.sdk;
      if(bindDirectNative()&&sdk&&typeof sdk.command==='function'){
        if(event==='reportsDocBuilderRun'){startWatch(id,d,ms);sdk.command('docbuilder:run',JSON.stringify(d));return;}
        if(event==='reportsDocBuilderProbe'){sdk.command('docbuilder:probe',JSON.stringify(d));return;}
      }
    }catch(_){ }
    post({event,source:'reports-ui',data:{requestId:id,payload:d}});
  });
}
function runDb(p){const d=isObj(p)?Object.assign({},p):{};if(!d.requestId)d.requestId=reqId('docbuilder-run');const n=d&&d.argument&&Array.isArray(d.argument.actions)?d.argument.actions.length:1;const t=Math.max(120000,120000+n*4500);return requestDb('reportsDocBuilderRun',d,t);}
function probeDb(p){const d=isObj(p)?Object.assign({},p):{};if(!d.requestId)d.requestId=reqId('docbuilder-probe');return requestDb('reportsDocBuilderProbe',d,30000);}
function normalizeDbResult(raw){
  if(!isObj(raw))return null;
  let out=raw;
  if(!out.requestId&&isObj(out.payload))out=Object.assign({},out.payload,out);
  if(!out.requestId&&isObj(out.data))out=Object.assign({},out.data,out);
  return out;
}
function onDbResult(raw){
  const r=normalizeDbResult(raw);if(!r)return;
  let id=r.requestId;
  if(!id){
    const ids=Object.keys(PENDING);
    if(ids.length===1)id=ids[0];
  }
  if(!id)return;
  const p=clearPending(id);if(!p)return;
  if(!r.requestId)r.requestId=id;
  if(r.ok)p.res(r);else p.rej(dbErr(r));
}
async function ensureProbe(force){if(!force&&state.probe&&Date.now()-state.probe.ts<PROBE_TTL)return state.probe.data;const r=await probeDb({});if(!r||!r.ok)throw dbErr(r||{error:'runtime_not_found'});state.probe={ts:Date.now(),data:r};return r;}

function openResult(path,scenarioId){
  const typeId=extType(path);
  post({event:'reportsOpenFile',source:'reports-ui',data:{id:null,path,typeId}});
}

async function runScenario(id){
  const sc=state.scenarios.find(x=>x.id===id);if(!sc||state.running)return;
  const tpl=resolveTemplate(sc.templatePath);if(!tpl){toast('Некорректный путь к шаблону.','error');return;}
  state.running=id;renderList();
  try{
    toast('Проверка DocumentBuilder...','ok');
    const probe=await ensureProbe(false);
    const out=mkOutputPath(sc,probe,tpl);
    toast('Выполнение сценария через DocumentBuilder...','ok');
    const runRequestId=reqId('docbuilder-run');
    const runPayload={requestId:runRequestId,script:SCRIPT,openAfterRun:false,argument:{templatePath:tpl,outputPath:out,scenarioId:sc.id,scenarioName:sc.name,stopOnError:true,actions:sc.actions}};
    const runPromise=runDb(runPayload); runPromise.catch(()=>{});
    const fallbackMs=Math.max(8000,Math.min(120000,5000+(Array.isArray(sc.actions)?sc.actions.length:1)*1500));
    const res=await Promise.race([runPromise,new Promise(resolve=>setTimeout(()=>resolve({ok:true,requestId:runRequestId,exitCode:0,fallback:true}),fallbackMs))]);
    let alreadyOpened=false;
    if(res&&res.fallback){
      clearPending(runRequestId);
      toast('Нет ответа моста. Открываю результат напрямую...','ok');
    }else{
      if(!res||!res.ok)throw dbErr(res||{error:'run_failed'});
      if(typeof res.exitCode!=='number')throw dbErr({error:'invalid_run_response',details:'DocumentBuilder returned unexpected response payload.'});
      alreadyOpened=!!res.opened;
    }
    const i=state.scenarios.findIndex(x=>x.id===id);if(i>=0){state.scenarios[i].lastRunAt=new Date().toISOString();state.scenarios[i].updatedAt=new Date().toISOString();save();}
    renderList();
    if(!alreadyOpened){
      await new Promise(r=>setTimeout(r,350));
      openResult(out,id);
    }
    toast(`Готово: ${out}`,'ok');
  }catch(err){
    const d=err&&err.details?err.details:null;
    const reason=d&&d.stderr?d.stderr:(d&&d.error?d.error:(err&&err.message?err.message:'Неизвестная ошибка'));
    toast(`Ошибка выполнения: ${reason}`,'error');
    state.probe=null;
  }finally{state.running=null;renderList();}
}

function pickTemplate(){
  try{
    const asc=window.parent&&window.parent.AscDesktopEditor;
    if(asc&&typeof asc.OpenFilenameDialog==='function'){
      asc.OpenFilenameDialog('cell',false,function(files){const f=Array.isArray(files)?files[0]:files;if(!f)return;els.tpl.value=toFsPath(String(f));if(state.draft)state.draft.templatePath=els.tpl.value;});
      return;
    }
  }catch(_){ }
  if(els.file)els.file.click();
}

function bind(){
  bindDirectNative();
  els.search&&els.search.addEventListener('input',e=>{state.q=e.target.value||'';renderList();});
  els.btnCreate&&els.btnCreate.addEventListener('click',()=>openModal(null));
  els.btnClose&&els.btnClose.addEventListener('click',closeModal);
  els.btnCancel&&els.btnCancel.addEventListener('click',closeModal);
  els.btnSave&&els.btnSave.addEventListener('click',saveDraft);
  els.btnAddAction&&els.btnAddAction.addEventListener('click',()=>{if(!state.draft)return;state.draft.actions.push(mkAction('set_cell_value'));renderActions();});
  els.btnPick&&els.btnPick.addEventListener('click',pickTemplate);
  els.file&&els.file.addEventListener('change',()=>{const f=els.file.files&&els.file.files[0];if(f&&f.path){els.tpl.value=toFsPath(f.path);if(state.draft)state.draft.templatePath=els.tpl.value;return;}const v=els.file.value||'';if(v&&!/fakepath/i.test(v)){els.tpl.value=toFsPath(v);if(state.draft)state.draft.templatePath=els.tpl.value;return;}toast('Введите полный путь к шаблону вручную.','error');});
  els.modal&&els.modal.addEventListener('click',e=>{if(e.target&&e.target.classList&&e.target.classList.contains('modal-backdrop'))closeModal();});
  document.addEventListener('click',()=>closeColorMenus());
  document.addEventListener('keydown',e=>{if(e.key==='Escape'){closeColorMenus();if(els.modal&&!els.modal.classList.contains('hidden'))closeModal();}});
  window.addEventListener('message',e=>{const m=parse(e.data);if(!isObj(m))return;if(m.event==='reportsDocBuilderResult'&&m.data)onDbResult(m.data);if(m.event==='reportsDocBuilderProbeResult'&&m.data)onDbResult(m.data);});
}

function init(){load();bind();renderList();window.ReportsDocBuilder={run:runDb,probe:probeDb};}
init();
})();
