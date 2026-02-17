
(()=>{
'use strict';

const STORAGE='reports.scenarios.v2';
const SCRIPT='docbuilder/scripts/reports_executor.docbuilder';
const PROBE_TTL=60000;
const PENDING=Object.create(null);
const WATCH=Object.create(null);
const SCENARIO_TEMPLATE_EXT='xlsx';
const THEME_IDS=['theme-light','theme-classic-light','theme-dark','theme-contrast-dark','theme-gray','theme-white','theme-night'];

const state={scenarios:[],q:'',draft:null,running:null,probe:null,toastTimer:null,runInput:null};

const schema={
  set_cell_value:{label:'Вставить значение',d:{sheet:'Лист1',range:'A1',mode:'text',value:'',merge:false,keep_template_format:true,apply_alignment:false,horizontal:'',vertical:'',wrap:false,apply_number_format:false,format_preset:'general',decimals:'2',use_thousands:false,currency_symbol:'₽',negative_red:false,custom_format:'General',format:'General',apply_font:false,font_name:'Arial',font_size:'11',bold:false,italic:false,underline:'none',strikeout:false,font_color:'auto',apply_fill:false,fill_color:'none'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Ячейка/диапазон',t:'text',r:1,p:'A1 или A1:C10'},
    {k:'mode',l:'Тип',t:'select',o:[['text','Текст'],['number','Число'],['formula','Формула'],['bool','Логическое']]},
    {k:'value',l:'Значение',t:'textarea',r:1,full:1,p:'Текст, число или формула'},
    {k:'merge',l:'Объединить после записи',t:'check',full:1},
    {k:'keep_template_format',l:'Формат текста как в шаблоне',t:'check',full:1,redraw:1},
    {k:'apply_alignment',l:'Применять выравнивание',t:'check',full:1,hideIf:'keep_template_format',redraw:1},
    {k:'horizontal',l:'Горизонталь',t:'select',showIf:'apply_alignment',hideIf:'keep_template_format',o:[['','(как в шаблоне)'],['left','Слева'],['center','По центру'],['right','Справа'],['justify','По ширине']]},
    {k:'vertical',l:'Вертикаль',t:'select',showIf:'apply_alignment',hideIf:'keep_template_format',o:[['','(как в шаблоне)'],['top','Сверху'],['center','По центру'],['bottom','Снизу'],['justify','По высоте'],['distributed','Распределить']]},
    {k:'wrap',l:'Переносить текст',t:'check',full:1,showIf:'apply_alignment',hideIf:'keep_template_format'},
    {k:'apply_number_format',l:'Применять числовой формат',t:'check',full:1,hideIf:'keep_template_format',redraw:1},
    {k:'format_builder',l:'Числовой формат',t:'number_format',full:1,showIf:'apply_number_format',hideIf:'keep_template_format'},
    {k:'apply_font',l:'Применять настройки шрифта',t:'check',full:1,hideIf:'keep_template_format',redraw:1},
    {k:'font_name',l:'Шрифт',t:'font',showIf:'apply_font',hideIf:'keep_template_format'},
    {k:'font_size',l:'Размер шрифта',t:'font_size',showIf:'apply_font',hideIf:'keep_template_format'},
    {k:'font_color',l:'Цвет шрифта',t:'color',auto:1,showIf:'apply_font',hideIf:'keep_template_format'},
    {k:'font_style',l:'Начертание',t:'font_style',full:1,showIf:'apply_font',hideIf:'keep_template_format'},
    {k:'apply_fill',l:'Применять заливку',t:'check',full:1,hideIf:'keep_template_format',redraw:1},
    {k:'fill_color',l:'Цвет заливки',t:'color',none:1,showIf:'apply_fill',hideIf:'keep_template_format'}
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
    {k:'scope',l:'Где ставить',t:'select',o:[['all','Все'],['outer','Внешние'],['inner','Внутренние'],['top','Верхняя'],['bottom','Нижняя'],['left','Левая'],['right','Правая'],['inside_h','Внутренние горизонтальные'],['inside_v','Внутренние вертикальные']]},
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
  set_font_style:{label:'Шрифт и стиль',d:{sheet:'Лист1',range:'A1',font_name:'Arial',font_size:'11',bold:false,italic:false,underline:'none',strikeout:false,font_color:'auto'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1,p:'A1 или A1:C10'},
    {k:'font_name',l:'Шрифт',t:'font'},
    {k:'font_size',l:'Размер шрифта',t:'font_size'},
    {k:'font_color',l:'Цвет шрифта',t:'color',auto:1},
    {k:'font_style',l:'Начертание',t:'font_style',full:1}
  ]},
  set_fill_color:{label:'Заливка диапазона',d:{sheet:'Лист1',range:'A1',color:'#FFFFFF'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},
    {k:'color',l:'Цвет заливки',t:'color',r:1,none:1}
  ]},
  set_number_format:{label:'Формат числа',d:{sheet:'Лист1',range:'A1',format_preset:'general',decimals:'2',use_thousands:false,currency_symbol:'₽',negative_red:false,custom_format:'General',format:'General'},f:[
    {k:'sheet',l:'Лист',t:'text',r:1},{k:'range',l:'Диапазон',t:'text',r:1},
    {k:'format_builder',l:'Формат',t:'number_format',full:1}
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
  ['#FFFFFF','#000000','#44546A','#E7E6E6','#5B9BD5','#ED7D31','#A5A5A5','#FFC000','#4472C4','#70AD47'],
  ['#F2F2F2','#7F7F7F','#D9E1F2','#F2F2F2','#DDEBF7','#FCE4D6','#EDEDED','#FFF2CC','#D9E1F2','#E2F0D9'],
  ['#D9D9D9','#595959','#B4C6E7','#D0CECE','#BDD7EE','#F8CBAD','#DBDBDB','#FFE699','#B4C6E7','#C6E0B4'],
  ['#BFBFBF','#3F3F3F','#8EA9DB','#AEAAAA','#9DC3E6','#F4B183','#C9C9C9','#FFD966','#8EA9DB','#A9D18E'],
  ['#A6A6A6','#262626','#2F5597','#7F7F7F','#2E75B6','#C55A11','#7B7B7B','#BF8F00','#2F5597','#538135'],
  ['#808080','#0D0D0D','#203864','#595959','#1F4E78','#833C0C','#525252','#7F6000','#203864','#375623']
];
const THEME_COLORS_FLAT=THEME_COLORS.reduce((acc,row)=>acc.concat(row),[]);
const STANDARD_COLORS=['#C00000','#FF0000','#FFC000','#FFFF00','#92D050','#00B050','#00B0F0','#0070C0','#002060','#7030A0'];
const FONT_FAMILIES=['Arial','Calibri','Cambria','Candara','Century Gothic','Comic Sans MS','Consolas','Courier New','Georgia','Impact','Lucida Console','Segoe UI','Tahoma','Times New Roman','Trebuchet MS','Verdana'];
const FONT_SIZES=['8','9','10','11','12','14','16','18','20','22','24','26','28','36','48','72','96'];
const NUMBER_FORMAT_PRESETS=[
  ['general','Общий'],
  ['number','Числовой'],
  ['scientific','Научный'],
  ['financial','Финансовый'],
  ['currency','Денежный'],
  ['short_date','Краткий формат даты'],
  ['long_date','Длинный формат даты'],
  ['time','Время'],
  ['percent','Процентный'],
  ['fraction','Дробный'],
  ['text','Текстовый'],
  ['custom','Другие форматы']
];
const H_ALIGN_OPTIONS=[['','(как в шаблоне)'],['left','Слева'],['center','По центру'],['right','Справа'],['justify','По ширине']];
const V_ALIGN_OPTIONS=[['','(как в шаблоне)'],['top','Сверху'],['center','По центру'],['bottom','Снизу'],['justify','По высоте'],['distributed','Распределить']];
const INPUT_TYPES=[['text','Текст'],['multiline','Многострочный текст'],['number','Число'],['date','Дата'],['select','Список'],['boolean','Да/Нет']];
const CONDITION_OPERATORS=[
  ['equals','Равно'],
  ['not_equals','Не равно'],
  ['contains','Содержит'],
  ['not_contains','Не содержит'],
  ['empty','Пусто'],
  ['not_empty','Не пусто'],
  ['gt','>'],
  ['gte','>='],
  ['lt','<'],
  ['lte','<=']
];

const typeList=Object.keys(schema).map(k=>[k,schema[k].label]);
const els={
  list:document.getElementById('scenario-list'),empty:document.getElementById('empty-state'),search:document.getElementById('search-input'),
  btnCreate:document.getElementById('btn-create'),modal:document.getElementById('scenario-modal'),title:document.getElementById('modal-title'),
  btnClose:document.getElementById('btn-close-modal'),btnCancel:document.getElementById('btn-cancel'),btnSave:document.getElementById('btn-save'),
  name:document.getElementById('scenario-name'),tpl:document.getElementById('scenario-template'),descr:document.getElementById('scenario-description'),
  inputList:document.getElementById('input-field-list'),btnAddInputField:document.getElementById('btn-add-input-field'),
  btnPick:document.getElementById('btn-pick-template'),file:document.getElementById('template-file-input'),actionList:document.getElementById('action-list'),
  btnAddAction:document.getElementById('btn-add-action'),toast:document.getElementById('toast'),
  runInputModal:document.getElementById('run-input-modal'),runInputTitle:document.getElementById('run-input-title'),
  runInputFields:document.getElementById('run-input-fields'),btnRunInputCancel:document.getElementById('btn-run-input-cancel'),
  btnRunInputSubmit:document.getElementById('btn-run-input-submit')
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
function samePath(a,b){return normSlashes(a).toLowerCase()===normSlashes(b).toLowerCase();}
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
function getReportsUiDir(){const root=getRoot();return root?joinPath(root,'reports-ui'):'reports-ui';}
function getTemplatesDir(){return joinPath(getReportsUiDir(),'templates');}
function sanitizeTemplateName(name){
  let out=String(name||'').replace(/[<>:"/\\|?*\u0000-\u001F]+/g,' ').replace(/\s+/g,' ').trim();
  out=out.replace(/\.(xlsx|xlsm|xls)$/i,'');
  out=out.replace(/[. ]+$/g,'');
  if(!out)out='Сценарий';
  if(/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i.test(out))out=`${out}_`;
  return out;
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
function getThemeTypeFromId(themeId){
  const id=String(themeId||'').toLowerCase();
  if(!id)return '';
  if(id.indexOf('dark')>=0||id.indexOf('night')>=0)return 'dark';
  if(id.indexOf('light')>=0||id.indexOf('white')>=0||id.indexOf('gray')>=0||id.indexOf('grey')>=0)return 'light';
  return '';
}
function readRendererTheme(win){
  try{
    const theme=win&&win.RendererProcessVariable&&win.RendererProcessVariable.theme;
    if(theme&&typeof theme==='object'){
      return {
        id:String(theme.id||''),
        type:String(theme.type||''),
        system:String(theme.system||'')
      };
    }
  }catch(_){ }
  return null;
}
function findThemeIdInClassName(className){
  const tokens=` ${String(className||'').toLowerCase()} `;
  for(let i=0;i<THEME_IDS.length;i++){
    const id=THEME_IDS[i];
    if(tokens.indexOf(` ${id.toLowerCase()} `)>=0)return id;
  }
  return '';
}
function findThemeTypeInClassName(className){
  const tokens=` ${String(className||'').toLowerCase()} `;
  if(tokens.indexOf(' theme-type-dark ')>=0)return 'dark';
  if(tokens.indexOf(' theme-type-light ')>=0)return 'light';
  return '';
}
function getParentThemeHint(){
  const out={id:'',type:''};
  try{
    const cls=window.parent&&window.parent.document&&window.parent.document.body&&window.parent.document.body.className;
    out.id=findThemeIdInClassName(cls);
    out.type=findThemeTypeInClassName(cls);
  }catch(_){ }
  return out;
}
function systemThemeId(themeType){
  const forced=String(themeType||'').toLowerCase();
  if(forced==='light')return 'theme-white';
  if(forced==='dark')return 'theme-night';
  try{
    const own=readRendererTheme(window);
    const parent=readRendererTheme(window.parent);
    const sys=String((own&&own.system)||(parent&&parent.system)||'').toLowerCase();
    if(sys==='light')return 'theme-white';
    if(sys==='dark')return 'theme-night';
  }catch(_){ }
  const hint=getParentThemeHint();
  if(hint.type==='light')return 'theme-white';
  return 'theme-night';
}
function normalizeThemeId(id,themeType){
  const raw=String(id||'').trim();
  if(!raw||raw==='theme-system')return systemThemeId(themeType);
  if(THEME_IDS.indexOf(raw)>=0)return raw;
  const s=raw.toLowerCase();
  if(s.indexOf('contrast')>=0&&s.indexOf('dark')>=0)return 'theme-contrast-dark';
  if(s.indexOf('night')>=0)return 'theme-night';
  if(s.indexOf('dark')>=0)return 'theme-dark';
  if(s.indexOf('gray')>=0||s.indexOf('grey')>=0)return 'theme-gray';
  if(s.indexOf('classic')>=0&&s.indexOf('light')>=0)return 'theme-classic-light';
  if(s.indexOf('light')>=0||s.indexOf('white')>=0)return 'theme-white';
  return systemThemeId(themeType);
}
function applyThemeClass(themeId,themeType){
  const resolved=normalizeThemeId(themeId,themeType);
  let resolvedType=String(themeType||'').toLowerCase();
  if(resolvedType!=='light'&&resolvedType!=='dark')resolvedType=getThemeTypeFromId(resolved)||'dark';
  const roots=[document.documentElement,document.body].filter(Boolean);
  roots.forEach(node=>{
    THEME_IDS.forEach(id=>node.classList.remove(id));
    node.classList.remove('theme-type-light','theme-type-dark');
    node.classList.add(resolved);
    node.classList.add(`theme-type-${resolvedType}`);
    node.setAttribute('data-theme',resolved);
    node.setAttribute('data-theme-type',resolvedType);
  });
}
function resolveInitialTheme(){
  const own=readRendererTheme(window);
  const parent=readRendererTheme(window.parent);
  const hint=getParentThemeHint();
  return {
    id:(own&&own.id)||(parent&&parent.id)||hint.id||'theme-system',
    type:(own&&own.type)||(parent&&parent.type)||hint.type||''
  };
}
function applyThemeFromInfo(info){
  if(isObj(info)){
    if(isObj(info.theme)){
      applyThemeClass(info.theme.id||info.theme.name||'',info.theme.type||info.type||'');
      return;
    }
    if(info.theme!=null){
      applyThemeClass(info.theme,info.type||'');
      return;
    }
    if(info.name||info.id){
      applyThemeClass(info.name||info.id,info.type||'');
      return;
    }
  }
  applyThemeClass(info,'');
}
function handleThemeBridgeMessage(data){
  const payload=isObj(data)?data:{};
  const name=payload.name||payload.id||'';
  const type=payload.type||'';
  if(name||type){
    applyThemeClass(name||'theme-system',type);
    return;
  }
  const hint=getParentThemeHint();
  applyThemeClass(hint.id||'theme-system',hint.type||'');
}
function bindThemeSync(){
  const initial=resolveInitialTheme();
  applyThemeClass(initial.id,initial.type);
  handleThemeBridgeMessage({});
  try{
    const prev=window.on_update_plugin_info;
    window.on_update_plugin_info=function(info){
      try{
        if(info)applyThemeFromInfo(info);
      }catch(_){ }
      if(typeof prev==='function'){
        try{prev(info);}catch(_){ }
      }
    };
  }catch(_){ }
  try{
    const body=window.parent&&window.parent.document&&window.parent.document.body;
    if(body&&typeof MutationObserver!=='undefined'){
      const observer=new MutationObserver(()=>handleThemeBridgeMessage({}));
      observer.observe(body,{attributes:true,attributeFilter:['class']});
      window.__reportsThemeObserver=observer;
    }
  }catch(_){ }
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
function isTrue(v){
  if(typeof v==='boolean')return v;
  if(typeof v==='number')return v!==0;
  const s=String(v==null?'':v).trim().toLowerCase();
  return s==='1'||s==='true'||s==='yes';
}
function clampInt(value,min,max,fallback){
  let n=parseInt(value,10);
  if(Number.isNaN(n))n=fallback;
  if(n<min)n=min;
  if(n>max)n=max;
  return n;
}
function repeatChar(ch,count){
  let out='';
  for(let i=0;i<count;i+=1)out+=ch;
  return out;
}
function decimalsMask(value){
  const decimals=clampInt(value,0,15,2);
  return decimals>0?`.${repeatChar('0',decimals)}`:'';
}
function normalizeCurrencySymbol(value){
  const s=String(value||'').trim();
  return s||'₽';
}
function buildNumberFormatMask(action){
  const preset=String(action&&action.format_preset||'').trim().toLowerCase();
  const decimals=clampInt(action&&action.decimals,0,15,2);
  const grouped=isTrue(action&&action.use_thousands);
  const numberBase=`${grouped?'#,##0':'0'}${decimalsMask(decimals)}`;
  const symbol=normalizeCurrencySymbol(action&&action.currency_symbol);
  const red=isTrue(action&&action.negative_red);
  if(preset==='general')return 'General';
  if(preset==='number')return red?`${numberBase};[Red]-${numberBase}`:numberBase;
  if(preset==='scientific'){
    const sci=`0${decimalsMask(decimals)}E+00`;
    return red?`${sci};[Red]-${sci}`:sci;
  }
  if(preset==='financial'){
    const fin=`${symbol} ${numberBase}`;
    return red?`${fin};[Red]-${fin}`:`${fin};-${fin}`;
  }
  if(preset==='currency'){
    const cur=`${symbol}${numberBase}`;
    return red?`${cur};[Red]-${cur}`:`${cur};-${cur}`;
  }
  if(preset==='short_date')return 'dd.mm.yyyy';
  if(preset==='long_date')return '[$-419]dd mmmm yyyy г.';
  if(preset==='time')return 'hh:mm:ss';
  if(preset==='percent')return `0${decimalsMask(decimals)}%`;
  if(preset==='fraction')return '# ?/?';
  if(preset==='text')return '@';
  if(preset==='custom'){
    const custom=String(action&&action.custom_format||action&&action.format||'').trim();
    return custom||'General';
  }
  const fallback=String(action&&action.format||'').trim();
  return fallback||'General';
}
function normalizeNumberFormatSettings(action){
  if(!action)return;
  const sourceFormat=String(action.format||'').trim();
  if(action.format_preset===undefined||action.format_preset===null||String(action.format_preset).trim()==='')
    action.format_preset=!sourceFormat||/^general$/i.test(sourceFormat)?'general':'custom';
  if(action.decimals===undefined||action.decimals===null||String(action.decimals).trim()==='')
    action.decimals='2';
  action.decimals=String(clampInt(action.decimals,0,15,2));
  if(action.use_thousands===undefined||action.use_thousands===null)
    action.use_thousands=false;
  if(action.currency_symbol===undefined||action.currency_symbol===null||String(action.currency_symbol).trim()==='')
    action.currency_symbol='₽';
  if(action.negative_red===undefined||action.negative_red===null)
    action.negative_red=false;
  if(action.custom_format===undefined||action.custom_format===null||String(action.custom_format).trim()==='')
    action.custom_format=sourceFormat||'General';
  action.format=buildNumberFormatMask(action);
}
function normalizeNumberFormatAction(action){
  if(!action||action.type!=='set_number_format')return;
  normalizeNumberFormatSettings(action);
}
function normalizeUnderlineValue(value){
  const raw=String(value==null?'none':value).trim().toLowerCase();
  if(raw==='single')return 'single';
  if(raw==='double')return 'double';
  if(raw==='singleaccounting')return 'singleAccounting';
  if(raw==='doubleaccounting')return 'doubleAccounting';
  return 'none';
}
function normalizeInsertStyleSettings(target,source){
  if(!target)return;
  const src=isObj(source)?source:target;
  const hasKeep=src.keep_template_format!==undefined&&src.keep_template_format!==null&&String(src.keep_template_format).trim()!=='';
  const hasLegacyApply=src.apply_style!==undefined&&src.apply_style!==null;
  const legacyApply=isTrue(src.apply_style);
  const hasCustomStyle=legacyApply||isTrue(src.apply_alignment)||isTrue(src.apply_number_format)||isTrue(src.apply_font)||isTrue(src.apply_fill);
  target.keep_template_format=hasKeep?isTrue(src.keep_template_format):!hasCustomStyle;
  if(src.apply_alignment===undefined||src.apply_alignment===null)
    target.apply_alignment=hasLegacyApply?legacyApply:isTrue(target.apply_alignment);
  else
    target.apply_alignment=isTrue(src.apply_alignment);
  target.horizontal=String(src.horizontal===undefined||src.horizontal===null?'':src.horizontal).trim().toLowerCase();
  target.vertical=String(src.vertical===undefined||src.vertical===null?'':src.vertical).trim().toLowerCase();
  target.wrap=isTrue(src.wrap);
  target.apply_number_format=isTrue(src.apply_number_format);
  normalizeNumberFormatSettings(target);
  target.apply_font=isTrue(src.apply_font);
  target.font_name=String(src.font_name===undefined||src.font_name===null?target.font_name:src.font_name).trim()||'Arial';
  target.font_size=String(src.font_size===undefined||src.font_size===null?target.font_size:src.font_size).trim()||'11';
  target.bold=isTrue(src.bold);
  target.italic=isTrue(src.italic);
  target.underline=normalizeUnderlineValue(src.underline);
  target.strikeout=isTrue(src.strikeout);
  target.font_color=normColor(src.font_color===undefined||src.font_color===null?target.font_color:src.font_color,false,true);
  target.apply_fill=isTrue(src.apply_fill);
  target.fill_color=normColor(src.fill_color===undefined||src.fill_color===null?target.fill_color:src.fill_color,true,false);
}
function normalizeInputType(value){
  const raw=String(value==null?'text':value).trim().toLowerCase();
  if(raw==='multiline'||raw==='number'||raw==='date'||raw==='select'||raw==='boolean')return raw;
  return 'text';
}
function parseInputOptions(raw){
  const out=[];
  const seen=Object.create(null);
  String(raw||'').split(/[\n;,]+/).forEach(part=>{
    const v=String(part||'').trim();
    if(!v)return;
    const key=v.toLowerCase();
    if(seen[key])return;
    seen[key]=true;
    out.push(v);
  });
  return out;
}
function normalizeConditionOperator(value){
  const raw=String(value==null?'equals':value).trim().toLowerCase();
  const ok=CONDITION_OPERATORS.some(item=>item[0]===raw);
  return ok?raw:'equals';
}
function conditionNeedsValue(op){
  return ['equals','not_equals','contains','not_contains','gt','gte','lt','lte'].indexOf(normalizeConditionOperator(op))>=0;
}
function normalizeActionCondition(target,source){
  if(!target)return;
  const src=isObj(source)?source:target;
  target.cond_enabled=isTrue(src.cond_enabled);
  target.cond_field=String(src.cond_field==null?'':src.cond_field).trim();
  target.cond_operator=normalizeConditionOperator(src.cond_operator);
  target.cond_value=String(src.cond_value==null?'':src.cond_value);
  target.cond_else=isTrue(src.cond_else);
}
function colLettersToNumber(raw){
  const s=String(raw||'').trim().toUpperCase().replace(/[^A-Z]/g,'');
  if(!s)return 0;
  let out=0;
  for(let i=0;i<s.length;i+=1)out=out*26+(s.charCodeAt(i)-64);
  return out;
}
function numberToColLetters(raw){
  let n=Math.floor(Number(raw)||0);
  if(n<1)n=1;
  let out='';
  while(n>0){
    const rem=(n-1)%26;
    out=String.fromCharCode(65+rem)+out;
    n=Math.floor((n-rem-1)/26);
  }
  return out;
}
function parseA1Range(raw){
  const s=String(raw||'').trim().replace(/\$/g,'').toUpperCase();
  const m=s.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if(!m)return null;
  const c1=colLettersToNumber(m[1]);
  const r1=parseInt(m[2],10);
  const c2=colLettersToNumber(m[3]||m[1]);
  const r2=parseInt(m[4]||m[2],10);
  if(!(c1>0&&c2>0&&r1>0&&r2>0))return null;
  return {
    c1:Math.min(c1,c2),
    c2:Math.max(c1,c2),
    r1:Math.min(r1,r2),
    r2:Math.max(r1,r2)
  };
}
function formatA1Range(range){
  if(!range)return '';
  const start=`${numberToColLetters(range.c1)}${range.r1}`;
  const end=`${numberToColLetters(range.c2)}${range.r2}`;
  return start===end?start:`${start}:${end}`;
}
function offsetA1Range(raw,offset,direction){
  const parsed=parseA1Range(raw);
  if(!parsed)return null;
  const step=Math.max(0,parseInt(offset,10)||0);
  if(step===0)return parsed;
  if(direction==='columns'){
    parsed.c1+=step;
    parsed.c2+=step;
  }else{
    parsed.r1+=step;
    parsed.r2+=step;
  }
  return parsed;
}
function parseNamedTargets(raw){
  const out=[];
  const seen=Object.create(null);
  String(raw||'').split(/[\n;,]+/).forEach(part=>{
    const v=String(part||'').trim();
    if(!v)return;
    const key=v.toLowerCase();
    if(seen[key])return;
    seen[key]=true;
    out.push(v);
  });
  return out;
}
function normalizeBindingMode(value){
  return String(value||'').trim().toLowerCase()==='named'?'named':'cells';
}
function isFieldValueFilled(value,inputType){
  const type=normalizeInputType(inputType);
  if(type==='boolean')return value===true||value===false;
  if(Array.isArray(value))return value.some(item=>isFieldValueFilled(item,type));
  return String(value==null?'':value).trim()!=='';
}
function normalizeNumberInput(value){
  if(value===null||value===undefined)return '';
  const text=String(value).trim().replace(',', '.');
  if(!text)return '';
  const n=Number(text);
  return Number.isFinite(n)?n:'';
}
function normalizeDateInput(value){
  const text=String(value==null?'':value).trim();
  if(!text)return '';
  if(/^\d{4}-\d{2}-\d{2}$/.test(text))return text;
  const d=new Date(text);
  if(Number.isNaN(d.getTime()))return text;
  const y=d.getFullYear();
  const m=String(d.getMonth()+1).padStart(2,'0');
  const day=String(d.getDate()).padStart(2,'0');
  return `${y}-${m}-${day}`;
}
function normalizeScalarFieldValue(value,inputType){
  const type=normalizeInputType(inputType);
  if(type==='boolean'){
    if(typeof value==='boolean')return value;
    const s=String(value==null?'':value).trim().toLowerCase();
    if(!s)return false;
    return s==='1'||s==='true'||s==='yes'||s==='on'||s==='да';
  }
  if(type==='number'){
    const n=normalizeNumberInput(value);
    return n===''?'':n;
  }
  if(type==='date')return normalizeDateInput(value);
  return String(value==null?'':value);
}
function normalizeRuntimeFieldValues(value,field){
  const f=normInputField(field);
  if(f.multiple){
    const source=Array.isArray(value)?value:[value];
    const items=source
      .map(v=>normalizeScalarFieldValue(v,f.input_type))
      .filter(v=>isFieldValueFilled(v,f.input_type));
    return items;
  }
  return [normalizeScalarFieldValue(value,f.input_type)];
}
function valueModeByInputType(inputType){
  const type=normalizeInputType(inputType);
  if(type==='number')return 'number';
  if(type==='boolean')return 'bool';
  return 'text';
}
function valueToConditionText(value){
  if(Array.isArray(value))return value.map(v=>valueToConditionText(v)).join('\n');
  if(value===true)return 'true';
  if(value===false)return 'false';
  return String(value==null?'':value);
}
function readConditionSourceValue(action,inputValues,scenario){
  const key=String(action&&action.cond_field||'').trim();
  if(!key)return '';
  if(isObj(inputValues)&&Object.prototype.hasOwnProperty.call(inputValues,key))return inputValues[key];
  const fields=Array.isArray(scenario&&scenario.inputFields)?scenario.inputFields.map(normInputField):[];
  const found=fields.find(field=>field.id===key||field.name===key);
  if(found){
    if(isObj(inputValues)&&Object.prototype.hasOwnProperty.call(inputValues,found.id))return inputValues[found.id];
    if(found.multiple)return normalizeRuntimeFieldValues(found.default_value,found);
    return normalizeScalarFieldValue(found.default_value,found.input_type);
  }
  return '';
}
function evaluateActionCondition(action,inputValues,scenario){
  if(!isObj(action)||!isTrue(action.cond_enabled))return true;
  const op=normalizeConditionOperator(action.cond_operator);
  const raw=readConditionSourceValue(action,inputValues,scenario);
  const text=valueToConditionText(raw);
  const expect=String(action.cond_value==null?'':action.cond_value);
  let ok=true;
  if(op==='equals')ok=text===expect;
  else if(op==='not_equals')ok=text!==expect;
  else if(op==='contains')ok=text.indexOf(expect)>=0;
  else if(op==='not_contains')ok=text.indexOf(expect)<0;
  else if(op==='empty')ok=String(text).trim()==='';
  else if(op==='not_empty')ok=String(text).trim()!=='';
  else{
    const left=Number(String(Array.isArray(raw)?(raw.length?raw[0]:''):raw).replace(',','.'));
    const right=Number(String(expect).replace(',','.'));
    if(!Number.isFinite(left)||!Number.isFinite(right))ok=false;
    else if(op==='gt')ok=left>right;
    else if(op==='gte')ok=left>=right;
    else if(op==='lt')ok=left<right;
    else if(op==='lte')ok=left<=right;
    else ok=false;
  }
  if(isTrue(action.cond_else))ok=!ok;
  return ok;
}
function syncNumberFormat(action){
  if(!action)return;
  action.format=buildNumberFormatMask(action);
}
function normColor(v,allowNone,allowAuto){
  let s=String(v==null?'':v).trim();
  if(!s){
    if(allowAuto)return 'auto';
    return allowNone?'none':'#000000';
  }
  if(allowAuto&&/^auto$/i.test(s))return 'auto';
  if(allowNone&&/^none$/i.test(s))return 'none';
  if(!s.startsWith('#'))s=`#${s}`;
  if(/^#[0-9a-fA-F]{3}$/.test(s))s=expandHex3(s);
  if(!/^#[0-9a-fA-F]{6}$/.test(s)){
    if(allowAuto)return 'auto';
    return allowNone?'none':'#000000';
  }
  return s.toUpperCase();
}
function closeColorMenus(){document.querySelectorAll('.color-menu.open').forEach(x=>x.classList.remove('open'));}
function swatchState(node,value){
  if(!node)return;
  const val=String(value||'').toLowerCase();
  node.classList.remove('is-none','is-auto');
  if(val==='none'){
    node.style.background='transparent';
    node.classList.add('is-none');
    return;
  }
  if(val==='auto'){
    node.style.background='#000000';
    node.classList.add('is-auto');
    return;
  }
  node.style.background=value;
}
function appendColorGrid(container,colors,onPick,kind){
  const g=document.createElement('div');g.className=kind==='standard'?'color-grid color-grid-standard':'color-grid color-grid-theme';
  const cells=[];
  colors.forEach(c=>{
    const value=normColor(c,false,false).toLowerCase();
    const b=document.createElement('button');
    b.type='button';
    b.className='color-cell';
    b.title=value.toUpperCase();
    b.dataset.color=value;
    b.style.background=value;
    b.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();onPick(value);};
    cells.push(b);
    g.appendChild(b);
  });
  container.appendChild(g);
  return cells;
}
function createColorField(action,field){
  const allowNone=!!field.none;
  const allowAuto=!!field.auto;
  const box=document.createElement('div');box.className='color-field';
  const btn=document.createElement('button');btn.type='button';btn.className='color-open';
  if(field.k==='font_color')btn.classList.add('color-open-font');
  else if(field.k==='fill_color'||(action&&action.type==='set_fill_color'))btn.classList.add('color-open-fill');
  else if(action&&action.type==='set_border')btn.classList.add('color-open-border');
  const icon=document.createElement('span');icon.className='color-open-icon';
  const dot=document.createElement('span');dot.className='color-dot color-open-indicator';
  btn.appendChild(icon);
  btn.appendChild(dot);
  const menu=document.createElement('div');menu.className='color-menu';
  const paletteCells=[];
  const hasAutoRow=allowAuto||allowNone;
  let autoBtn=null;
  if(hasAutoRow){
    autoBtn=document.createElement('button');
    autoBtn.type='button';
    autoBtn.className='color-auto';
    const autoDot=document.createElement('span');
    autoDot.className='color-auto-dot';
    if(!allowAuto)autoDot.classList.add('is-none');
    const autoText=document.createElement('span');
    autoText.textContent=allowAuto?'Автоматический':'Нет заливки';
    autoBtn.appendChild(autoDot);
    autoBtn.appendChild(autoText);
    autoBtn.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();applyColor(allowAuto?'auto':'none');};
    menu.appendChild(autoBtn);
  }
  const t1=document.createElement('div');t1.className='color-section-title';t1.textContent='Цвета темы';menu.appendChild(t1);
  paletteCells.push(...appendColorGrid(menu,THEME_COLORS_FLAT,applyColor,'theme'));
  const t2=document.createElement('div');t2.className='color-section-title';t2.textContent='Стандартные цвета';menu.appendChild(t2);
  paletteCells.push(...appendColorGrid(menu,STANDARD_COLORS,applyColor,'standard'));
  const links=document.createElement('div');links.className='color-links';
  const pipetteBtn=document.createElement('button');
  pipetteBtn.type='button';
  pipetteBtn.className='color-link color-link-pipette';
  const pipetteIcon=document.createElement('span');
  pipetteIcon.className='color-link-icon color-link-icon-pipette';
  const pipetteText=document.createElement('span');
  pipetteText.textContent='Пипетка';
  pipetteBtn.appendChild(pipetteIcon);
  pipetteBtn.appendChild(pipetteText);
  const moreBtn=document.createElement('button');
  moreBtn.type='button';
  moreBtn.className='color-link';
  moreBtn.textContent='Другие цвета';
  const picker=document.createElement('input');picker.type='color';picker.className='color-native';
  pipetteBtn.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();picker.click();};
  moreBtn.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();picker.click();};
  picker.oninput=(ev)=>applyColor(ev.target.value);
  links.appendChild(pipetteBtn);
  links.appendChild(moreBtn);
  links.appendChild(picker);
  menu.appendChild(links);

  function applyColor(value){
    const n=normColor(value,allowNone,allowAuto);
    action[field.k]=n;
    swatchState(dot,n);
    const low=String(n||'').toLowerCase();
    paletteCells.forEach(cell=>cell.classList.toggle('selected',cell.dataset.color===low));
    if(autoBtn)autoBtn.classList.toggle('selected',low==='auto'||low==='none');
    closeColorMenus();
  }
  btn.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();const wasOpen=menu.classList.contains('open');closeColorMenus();if(!wasOpen)menu.classList.add('open');};
  menu.onclick=(ev)=>ev.stopPropagation();
  applyColor(action[field.k]);
  box.appendChild(btn);box.appendChild(menu);
  return box;
}
function createFontField(action,field){
  const box=document.createElement('div');box.className='font-field';
  const select=document.createElement('select');select.className='font-value';
  const current=String(action[field.k]||'Arial').trim()||'Arial';
  const values=FONT_FAMILIES.slice();
  if(values.indexOf(current)<0)values.unshift(current);
  values.forEach(name=>{const opt=document.createElement('option');opt.value=name;opt.textContent=name;select.appendChild(opt);});
  select.value=current;
  select.onchange=(ev)=>{action[field.k]=String(ev.target.value||'').trim();};
  box.appendChild(select);
  return box;
}
function createFontSizeField(action,field){
  const box=document.createElement('div');box.className='font-size-field';
  const select=document.createElement('select');select.className='font-size-value';
  const current=String(action[field.k]||'11').trim()||'11';
  const values=FONT_SIZES.slice();
  if(values.indexOf(current)<0)values.push(current);
  values.forEach(size=>{const opt=document.createElement('option');opt.value=size;opt.textContent=size;select.appendChild(opt);});
  select.value=current;
  select.onchange=(ev)=>{action[field.k]=String(ev.target.value||'').trim();};
  box.appendChild(select);
  return box;
}
function createFontStyleField(action){
  const box=document.createElement('div');box.className='font-style-bar';
  const mkBtn=(title,text,getter,setter)=>{
    const btn=document.createElement('button');
    btn.type='button';
    btn.className='font-style-btn';
    btn.title=title;
    btn.textContent=text;
    btn.onclick=(ev)=>{ev.preventDefault();ev.stopPropagation();setter(!getter());refresh();};
    return btn;
  };
  const boldBtn=mkBtn('Полужирный','Ж',()=>isTrue(action.bold),(v)=>{action.bold=v;});
  const italicBtn=mkBtn('Курсив','К',()=>isTrue(action.italic),(v)=>{action.italic=v;});
  const strikeBtn=mkBtn('Зачеркнутый','S',()=>isTrue(action.strikeout),(v)=>{action.strikeout=v;});
  const underlineBtn=mkBtn('Подчеркивание','Ч',()=>String(action.underline||'none')!=='none',(v)=>{action.underline=v?'single':'none';});
  const underlineSelect=document.createElement('select');underlineSelect.className='font-style-underline';
  [['none','Нет подчеркивания'],['single','Одинарное'],['double','Двойное'],['singleAccounting','Одинарное (учетное)'],['doubleAccounting','Двойное (учетное)']]
    .forEach(item=>{const opt=document.createElement('option');opt.value=item[0];opt.textContent=item[1];underlineSelect.appendChild(opt);});
  underlineSelect.onchange=(ev)=>{action.underline=ev.target.value;refresh();};
  function refresh(){
    const underlineRaw=String(action.underline||'none');
    const underline=(underlineRaw==='single'||underlineRaw==='double'||underlineRaw==='singleAccounting'||underlineRaw==='doubleAccounting')?underlineRaw:'none';
    action.underline=underline;
    boldBtn.classList.toggle('active',isTrue(action.bold));
    italicBtn.classList.toggle('active',isTrue(action.italic));
    strikeBtn.classList.toggle('active',isTrue(action.strikeout));
    underlineBtn.classList.toggle('active',underline!=='none');
    underlineSelect.value=underline;
  }
  [boldBtn,italicBtn,underlineBtn,strikeBtn,underlineSelect].forEach(node=>box.appendChild(node));
  refresh();
  return box;
}
function createNumberFormatField(action){
  normalizeNumberFormatSettings(action);
  const box=document.createElement('div');box.className='numfmt-box';
  const presetWrap=document.createElement('div');presetWrap.className='numfmt-row';
  const preset=document.createElement('select');preset.className='numfmt-preset';
  NUMBER_FORMAT_PRESETS.forEach(item=>{const opt=document.createElement('option');opt.value=item[0];opt.textContent=item[1];preset.appendChild(opt);});
  preset.value=String(action.format_preset||'general');
  presetWrap.appendChild(preset);

  const options=document.createElement('div');options.className='numfmt-options';
  const decimalsWrap=document.createElement('label');decimalsWrap.className='numfmt-option';
  const decimalsTitle=document.createElement('span');decimalsTitle.textContent='Знаков после запятой';
  const decimalsInput=document.createElement('input');decimalsInput.type='number';decimalsInput.min='0';decimalsInput.max='15';decimalsInput.step='1';decimalsInput.value=String(action.decimals||'2');
  decimalsWrap.appendChild(decimalsTitle);decimalsWrap.appendChild(decimalsInput);

  const symbolWrap=document.createElement('label');symbolWrap.className='numfmt-option';
  const symbolTitle=document.createElement('span');symbolTitle.textContent='Валюта';
  const symbolSelect=document.createElement('select');
  [['₽','Рубль (₽)'],['$','Доллар ($)'],['€','Евро (€)'],['£','Фунт (£)'],['¥','Иена (¥)'],['₸','Тенге (₸)']].forEach(item=>{const opt=document.createElement('option');opt.value=item[0];opt.textContent=item[1];symbolSelect.appendChild(opt);});
  symbolSelect.value=normalizeCurrencySymbol(action.currency_symbol);
  symbolWrap.appendChild(symbolTitle);symbolWrap.appendChild(symbolSelect);

  const thousandWrap=document.createElement('label');thousandWrap.className='numfmt-option numfmt-check';
  const thousandInline=document.createElement('span');thousandInline.className='field-inline';
  const thousandCheck=document.createElement('input');thousandCheck.type='checkbox';thousandCheck.checked=isTrue(action.use_thousands);
  const thousandText=document.createElement('span');thousandText.textContent='Разделитель групп разрядов';
  thousandInline.appendChild(thousandCheck);thousandInline.appendChild(thousandText);thousandWrap.appendChild(thousandInline);

  const redWrap=document.createElement('label');redWrap.className='numfmt-option numfmt-check';
  const redInline=document.createElement('span');redInline.className='field-inline';
  const redCheck=document.createElement('input');redCheck.type='checkbox';redCheck.checked=isTrue(action.negative_red);
  const redText=document.createElement('span');redText.textContent='Отрицательные числа красным';
  redInline.appendChild(redCheck);redInline.appendChild(redText);redWrap.appendChild(redInline);

  options.appendChild(decimalsWrap);
  options.appendChild(symbolWrap);
  options.appendChild(thousandWrap);
  options.appendChild(redWrap);

  const customWrap=document.createElement('label');customWrap.className='numfmt-option numfmt-custom';
  const customTitle=document.createElement('span');customTitle.textContent='Пользовательская маска';
  const customInput=document.createElement('input');customInput.type='text';customInput.value=String(action.custom_format||action.format||'General');customInput.placeholder='Например: #,##0.00';
  customWrap.appendChild(customTitle);customWrap.appendChild(customInput);

  const previewWrap=document.createElement('label');previewWrap.className='numfmt-option numfmt-preview';
  const previewTitle=document.createElement('span');previewTitle.textContent='Итоговая маска';
  const previewInput=document.createElement('input');previewInput.type='text';previewInput.readOnly=true;
  previewWrap.appendChild(previewTitle);previewWrap.appendChild(previewInput);

  const needsDecimals=(id)=>['number','scientific','financial','currency','percent'].indexOf(id)>=0;
  const needsSymbol=(id)=>['financial','currency'].indexOf(id)>=0;
  const needsThousands=(id)=>['number','financial','currency'].indexOf(id)>=0;
  const needsNegative=(id)=>['number','scientific','financial','currency'].indexOf(id)>=0;

  function refresh(){
    const presetId=String(preset.value||'general');
    action.format_preset=presetId;
    action.decimals=String(clampInt(decimalsInput.value,0,15,2));
    action.use_thousands=!!thousandCheck.checked;
    action.currency_symbol=normalizeCurrencySymbol(symbolSelect.value);
    action.negative_red=!!redCheck.checked;
    action.custom_format=String(customInput.value||'').trim()||'General';
    decimalsInput.value=action.decimals;
    syncNumberFormat(action);
    previewInput.value=action.format;
    decimalsWrap.style.display=needsDecimals(presetId)?'flex':'none';
    symbolWrap.style.display=needsSymbol(presetId)?'flex':'none';
    thousandWrap.style.display=needsThousands(presetId)?'flex':'none';
    redWrap.style.display=needsNegative(presetId)?'flex':'none';
    customWrap.style.display=presetId==='custom'?'flex':'none';
  }

  [preset,decimalsInput,symbolSelect,thousandCheck,redCheck,customInput].forEach(ctrl=>ctrl.addEventListener('input',refresh));
  preset.addEventListener('change',refresh);
  box.appendChild(presetWrap);
  box.appendChild(options);
  box.appendChild(customWrap);
  box.appendChild(previewWrap);
  refresh();
  return box;
}

function mkInputField(){
  const out={
    id:uid('input'),
    name:'',
    placeholder:'',
    default_value:'',
    required:true,
    mode:'cells',
    binding_mode:'cells',
    targets:'Лист1:A1',
    named_targets:'',
    token:'{{key}}',
    scope:'workbook',
    sheets:'',
    input_type:'text',
    options:'',
    multiple:false,
    multiple_direction:'rows',
    multiple_insert:true,
    keep_template_format:true,
    apply_alignment:false,
    horizontal:'',
    vertical:'',
    wrap:false,
    apply_number_format:false,
    format_preset:'general',
    decimals:'2',
    use_thousands:false,
    currency_symbol:'₽',
    negative_red:false,
    custom_format:'General',
    format:'General',
    apply_font:false,
    font_name:'Arial',
    font_size:'11',
    bold:false,
    italic:false,
    underline:'none',
    strikeout:false,
    font_color:'auto',
    apply_fill:false,
    fill_color:'none'
  };
   out.input_type=normalizeInputType(out.input_type);
   out.binding_mode=normalizeBindingMode(out.binding_mode);
   normalizeInsertStyleSettings(out,out);
  return out;
}

function normInputField(field){
  const base=mkInputField();
  const src=isObj(field)?field:{};
  const out=Object.assign({},base,src);
  out.id=String(src.id||base.id);
  out.name=String(src.name||'').trim();
  out.placeholder=String(src.placeholder||'').trim();
  out.default_value=String(src.default_value==null?'':src.default_value);
  out.required=src.required===undefined?true:!!src.required;
  out.mode=String(src.mode||'cells').trim().toLowerCase()==='token'?'token':'cells';
  out.binding_mode=normalizeBindingMode(src.binding_mode||src.target_kind||'cells');
  out.targets=String(src.targets||'').trim();
  out.named_targets=String(src.named_targets||src.names||'').trim();
  out.token=String(src.token||'').trim();
  out.scope=String(src.scope||'workbook').trim().toLowerCase()==='sheets'?'sheets':'workbook';
  out.sheets=String(src.sheets||'').trim();
  out.input_type=normalizeInputType(src.input_type||src.type||'text');
  out.options=String(src.options||'').trim();
  out.multiple=!!src.multiple;
  out.multiple_direction=String(src.multiple_direction||src.repeat_direction||'rows').trim().toLowerCase()==='columns'?'columns':'rows';
  out.multiple_insert=src.multiple_insert===undefined?true:!!src.multiple_insert;
  if(out.mode!=='cells'){
    out.binding_mode='cells';
    out.multiple=false;
    out.multiple_insert=false;
  }else if(out.binding_mode==='named'){
    out.multiple_insert=false;
  }
  normalizeInsertStyleSettings(out,src);
  return out;
}

function parseSheetList(raw){
  const out=[];
  const seen=Object.create(null);
  String(raw||'').split(/[;,\n]+/).forEach(part=>{
    const name=String(part||'').trim();
    if(!name)return;
    const key=name.toLowerCase();
    if(seen[key])return;
    seen[key]=true;
    out.push(name);
  });
  return out;
}

function parseCellBindings(raw){
  const text=String(raw||'').trim();
  if(!text)return {items:[],error:'Укажите адреса в формате "Лист:A1,B2; Лист2:C3".'};
  const chunks=text.split(/[;\n]+/);
  const items=[];
  for(let i=0;i<chunks.length;i+=1){
    const block=String(chunks[i]||'').trim();
    if(!block)continue;
    let sep=block.indexOf(':');
    if(sep<=0){
      const bang=block.indexOf('!');
      if(bang>0)sep=bang;
    }
    if(sep<=0)return {items:[],error:`Блок ${i+1}: используйте формат "Лист:адрес".`};
    const sheet=block.slice(0,sep).trim();
    const rawRanges=block.slice(sep+1).trim();
    if(!sheet)return {items:[],error:`Блок ${i+1}: не указан лист.`};
    if(!rawRanges)return {items:[],error:`Блок ${i+1}: не указаны адреса.`};
    const ranges=rawRanges.split(',').map(x=>String(x||'').trim()).filter(Boolean);
    if(!ranges.length)return {items:[],error:`Блок ${i+1}: не указаны адреса.`};
    ranges.forEach(range=>items.push({sheet,range}));
  }
  if(!items.length)return {items:[],error:'Не удалось разобрать адреса.'};
  return {items,error:''};
}

function inputFieldLabel(field,index){
  const name=String(field&&field.name||'').trim();
  return name||`Поле ${index+1}`;
}

function mkAction(type){
  const key=schema[type]?type:'set_cell_value';
  const a={id:uid('action'),type:key};
  Object.assign(a,schema[key].d||{});
  if(key==='set_cell_value')normalizeInsertStyleSettings(a,a);
  else if(key==='set_number_format')normalizeNumberFormatSettings(a);
  if(key==='set_font_style'){
    a.font_color=normColor(a.font_color,false,true);
    a.underline=normalizeUnderlineValue(a.underline);
  }
  normalizeActionCondition(a,a);
  return a;
}
function normAction(a){
  if(!isObj(a))return mkAction('set_cell_value');
  const t=schema[a.type]?a.type:'set_cell_value';
  const o={id:a.id||uid('action'),type:t};
  Object.assign(o,schema[t].d||{},a);
  if(t==='set_cell_value')normalizeInsertStyleSettings(o,a);
  else if(t==='set_number_format')normalizeNumberFormatSettings(o);
  if(t==='set_font_style'){
    o.font_color=normColor(o.font_color,false,true);
    o.underline=normalizeUnderlineValue(o.underline);
  }
  normalizeActionCondition(o,a);
  return o;
}
function normScenario(s){
  if(!isObj(s))return null;
  const rawInputs=Array.isArray(s.inputFields)?s.inputFields:(Array.isArray(s.input_fields)?s.input_fields:[]);
  return{
    id:s.id||uid('scenario'),
    name:String(s.name||'Без названия'),
    description:String(s.description||s.descr||''),
    templatePath:String(s.templatePath||s.path||''),
    actions:Array.isArray(s.actions)?s.actions.map(normAction):[],
    inputFields:rawInputs.map(normInputField),
    createdAt:s.createdAt||new Date().toISOString(),
    updatedAt:s.updatedAt||s.updated||new Date().toISOString(),
    lastRunAt:s.lastRunAt||s.last_run||null
  };
}
function save(){try{localStorage.setItem(STORAGE,JSON.stringify(state.scenarios));}catch(_){toast('Не удалось сохранить сценарии.','error');}}
function load(){const raw=parse(localStorage.getItem(STORAGE));state.scenarios=Array.isArray(raw)?raw.map(normScenario).filter(Boolean):[];}

function openModal(id){
  const src=state.scenarios.find(x=>x.id===id);
  state.draft=src?clone(src):{id:uid('scenario'),name:'',description:'',templatePath:'',actions:[mkAction('set_cell_value')],inputFields:[],createdAt:new Date().toISOString(),updatedAt:new Date().toISOString(),lastRunAt:null};
  if(!Array.isArray(state.draft.inputFields))state.draft.inputFields=[];
  state.draft.inputFields=state.draft.inputFields.map(normInputField);
  els.title.textContent=src?'Изменение сценария':'Новый сценарий';
  els.name.value=state.draft.name; els.tpl.value=state.draft.templatePath; els.descr.value=state.draft.description;
  renderInputFields(); renderActions(); els.modal.classList.remove('hidden'); els.name.focus();
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
    const inputCount=Array.isArray(s.inputFields)?s.inputFields.length:0;
    [['Действий',s.actions.length],['Полей',inputCount],['Изменен',dt(s.updatedAt)],['Запуск',dt(s.lastRunAt)]].forEach(x=>{const ch=document.createElement('span');ch.className='meta-chip';ch.textContent=`${x[0]}: ${x[1]}`;m.appendChild(ch);});
    const a=c.querySelector('.scenario-actions');
    const bRun=document.createElement('button');bRun.className='btn btn-primary';bRun.type='button';bRun.textContent=state.running===s.id?'Выполняется...':'Заполнить';bRun.disabled=!!state.running||!!state.runInput;bRun.onclick=()=>runScenario(s.id);
    const bEdit=document.createElement('button');bEdit.className='btn btn-secondary';bEdit.type='button';bEdit.textContent='Изменить';bEdit.onclick=()=>openModal(s.id);
    const bTpl=document.createElement('button');bTpl.className='btn btn-ghost';bTpl.type='button';bTpl.textContent='Открыть шаблон';bTpl.disabled=!s.templatePath;bTpl.onclick=()=>openTemplateInEditor(s.id);
    const bDel=document.createElement('button');bDel.className='btn btn-danger';bDel.type='button';bDel.textContent='Удалить';bDel.onclick=()=>{if(confirm(`Удалить сценарий "${s.name}"?`)){state.scenarios=state.scenarios.filter(x=>x.id!==s.id);save();renderList();toast('Сценарий удален.','ok');}};
    [bRun,bEdit,bTpl,bDel].forEach(x=>a.appendChild(x));
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
      if(f.showIf&&!a[f.showIf])return;
      if(f.hideIf&&a[f.hideIf])return;
      const w=fieldWrap(!!f.full);
      if(f.t==='check'){
        const inl=document.createElement('span');inl.className='field-inline';const inp=document.createElement('input');inp.type='checkbox';inp.checked=!!a[f.k];inp.onchange=(ev)=>{a[f.k]=!!ev.target.checked;if(f.redraw)renderActions();};const t=document.createElement('span');t.textContent=f.l;inl.appendChild(inp);inl.appendChild(t);w.appendChild(inl);fs.appendChild(w);return;
      }
      const cap=document.createElement('span');cap.textContent=f.l;w.appendChild(cap);
      let inp;
      if(f.t==='textarea'){inp=document.createElement('textarea');inp.rows=2;inp.value=a[f.k]==null?'':String(a[f.k]);}
      else if(f.t==='select'){inp=document.createElement('select');(f.o||[]).forEach(op=>{const o=document.createElement('option');o.value=op[0];o.textContent=op[1];o.selected=String(op[0])===String(a[f.k]);inp.appendChild(o);});}
      else if(f.t==='color'){inp=createColorField(a,f);}
      else if(f.t==='font'){inp=createFontField(a,f);}
      else if(f.t==='font_size'){inp=createFontSizeField(a,f);}
      else if(f.t==='font_style'){inp=createFontStyleField(a);}
      else if(f.t==='number_format'){inp=createNumberFormatField(a);}
      else {inp=document.createElement('input');inp.type=f.t==='number'?'number':'text';inp.value=a[f.k]==null?'':String(a[f.k]);if(f.min!=null)inp.min=String(f.min);if(f.step!=null)inp.step=String(f.step);}
      if(['color','font','font_size','font_style','number_format'].indexOf(f.t)<0){
        if(f.p)inp.placeholder=f.p;
        const applyValue=(ev)=>{a[f.k]=ev.target.value;if(f.redraw)renderActions();};
        inp.oninput=applyValue;
        if(f.t==='select')inp.onchange=applyValue;
      }
      w.appendChild(inp);fs.appendChild(w);
    });
    const wCondEnable=fieldWrap(true);
    const inlCondEnable=document.createElement('span');inlCondEnable.className='field-inline';
    const inCondEnable=document.createElement('input');inCondEnable.type='checkbox';inCondEnable.checked=!!a.cond_enabled;inCondEnable.onchange=(ev)=>{a.cond_enabled=!!ev.target.checked;renderActions();};
    const txtCondEnable=document.createElement('span');txtCondEnable.textContent='Условие выполнения (if/else)';
    inlCondEnable.appendChild(inCondEnable);inlCondEnable.appendChild(txtCondEnable);wCondEnable.appendChild(inlCondEnable);fs.appendChild(wCondEnable);
    if(a.cond_enabled){
      const wCondField=fieldWrap(false);
      const capCondField=document.createElement('span');capCondField.textContent='Поле для условия';
      const inCondField=document.createElement('select');
      const inputFields=(Array.isArray(state.draft&&state.draft.inputFields)?state.draft.inputFields:[]).map(normInputField);
      const blank=document.createElement('option');blank.value='';blank.textContent='Выберите поле';inCondField.appendChild(blank);
      inputFields.forEach(field=>{
        const opt=document.createElement('option');
        opt.value=field.id;
        opt.textContent=`${field.name||field.id}`;
        if(String(a.cond_field||'')===field.id)opt.selected=true;
        inCondField.appendChild(opt);
      });
      inCondField.onchange=(ev)=>{a.cond_field=String(ev.target.value||'');};
      wCondField.appendChild(capCondField);wCondField.appendChild(inCondField);fs.appendChild(wCondField);

      const wCondOp=fieldWrap(false);
      const capCondOp=document.createElement('span');capCondOp.textContent='Оператор';
      const inCondOp=document.createElement('select');
      CONDITION_OPERATORS.forEach(op=>{
        const opt=document.createElement('option');
        opt.value=op[0];
        opt.textContent=op[1];
        if(normalizeConditionOperator(a.cond_operator)===op[0])opt.selected=true;
        inCondOp.appendChild(opt);
      });
      inCondOp.onchange=(ev)=>{a.cond_operator=normalizeConditionOperator(ev.target.value);renderActions();};
      wCondOp.appendChild(capCondOp);wCondOp.appendChild(inCondOp);fs.appendChild(wCondOp);

      if(conditionNeedsValue(a.cond_operator)){
        const wCondVal=fieldWrap(false);
        const capCondVal=document.createElement('span');capCondVal.textContent='Сравнить со значением';
        const inCondVal=document.createElement('input');
        inCondVal.type='text';
        inCondVal.value=String(a.cond_value==null?'':a.cond_value);
        inCondVal.oninput=(ev)=>{a.cond_value=ev.target.value;};
        wCondVal.appendChild(capCondVal);wCondVal.appendChild(inCondVal);fs.appendChild(wCondVal);
      }

      const wCondElse=fieldWrap(true);
      const inlCondElse=document.createElement('span');inlCondElse.className='field-inline';
      const inCondElse=document.createElement('input');inCondElse.type='checkbox';inCondElse.checked=!!a.cond_else;inCondElse.onchange=(ev)=>{a.cond_else=!!ev.target.checked;};
      const txtCondElse=document.createElement('span');txtCondElse.textContent='Ветка else (выполнять при ложном условии)';
      inlCondElse.appendChild(inCondElse);inlCondElse.appendChild(txtCondElse);wCondElse.appendChild(inlCondElse);fs.appendChild(wCondElse);
    }
    item.appendChild(fs); root.appendChild(item);
  });
}

function renderInputFields(){
  const root=els.inputList;
  if(!root)return;
  root.innerHTML='';
  if(!state.draft||!Array.isArray(state.draft.inputFields))return;
  if(!state.draft.inputFields.length){
    const empty=document.createElement('div');
    empty.className='input-field-empty';
    empty.textContent='Поля ввода не добавлены. Нажмите «Добавить поле», чтобы показывать форму перед запуском.';
    root.appendChild(empty);
    return;
  }
  state.draft.inputFields.forEach((rawField,i)=>{
    const f=normInputField(rawField);
    state.draft.inputFields[i]=f;
    const item=document.createElement('div');
    item.className='action-item input-field-item';

    const head=document.createElement('div');
    head.className='action-item-head';
    const ord=document.createElement('span');
    ord.className='order';
    ord.textContent=String(i+1);
    const title=document.createElement('div');
    title.className='input-field-title';
    title.textContent=inputFieldLabel(f,i);
    const tools=document.createElement('div');
    tools.className='action-tools';
    const up=document.createElement('button');
    up.className='btn btn-ghost';
    up.type='button';
    up.textContent='↑';
    up.disabled=i===0;
    up.onclick=()=>{const t=state.draft.inputFields[i-1];state.draft.inputFields[i-1]=state.draft.inputFields[i];state.draft.inputFields[i]=t;renderInputFields();};
    const dn=document.createElement('button');
    dn.className='btn btn-ghost';
    dn.type='button';
    dn.textContent='↓';
    dn.disabled=i===state.draft.inputFields.length-1;
    dn.onclick=()=>{const t=state.draft.inputFields[i+1];state.draft.inputFields[i+1]=state.draft.inputFields[i];state.draft.inputFields[i]=t;renderInputFields();};
    const rm=document.createElement('button');
    rm.className='btn btn-danger';
    rm.type='button';
    rm.textContent='Удалить';
    rm.onclick=()=>{state.draft.inputFields.splice(i,1);renderInputFields();};
    [up,dn,rm].forEach(x=>tools.appendChild(x));
    [ord,title,tools].forEach(x=>head.appendChild(x));
    item.appendChild(head);

    const fs=document.createElement('div');
    fs.className='action-fields';

    const wName=fieldWrap(false);
    const capName=document.createElement('span');capName.textContent='Название поля';
    const inName=document.createElement('input');inName.type='text';inName.value=f.name;inName.placeholder='Например: Название компании';inName.oninput=(ev)=>{f.name=ev.target.value;title.textContent=inputFieldLabel(f,i);};
    wName.appendChild(capName);wName.appendChild(inName);fs.appendChild(wName);

    const wMode=fieldWrap(false);
    const capMode=document.createElement('span');capMode.textContent='Тип привязки';
    const inMode=document.createElement('select');
    [['cells','Вставка в ячейки/диапазоны'],['token','Замена ключа ({{key}})']].forEach(opt=>{const o=document.createElement('option');o.value=opt[0];o.textContent=opt[1];if(f.mode===opt[0])o.selected=true;inMode.appendChild(o);});
    inMode.onchange=(ev)=>{
      f.mode=ev.target.value==='token'?'token':'cells';
      if(f.mode!=='cells'){
        f.multiple=false;
        f.multiple_insert=false;
      }
      renderInputFields();
    };
    wMode.appendChild(capMode);wMode.appendChild(inMode);fs.appendChild(wMode);

    const wPlaceholder=fieldWrap(false);
    const capPlaceholder=document.createElement('span');capPlaceholder.textContent='Подсказка для ввода';
    const inPlaceholder=document.createElement('input');inPlaceholder.type='text';inPlaceholder.value=f.placeholder;inPlaceholder.placeholder='Например: ООО Ромашка';inPlaceholder.oninput=(ev)=>{f.placeholder=ev.target.value;};
    wPlaceholder.appendChild(capPlaceholder);wPlaceholder.appendChild(inPlaceholder);fs.appendChild(wPlaceholder);

    const wDefault=fieldWrap(false);
    const capDefault=document.createElement('span');capDefault.textContent='Значение по умолчанию';
    if(f.input_type==='multiline'){
      const inDefault=document.createElement('textarea');inDefault.rows=2;inDefault.value=String(f.default_value==null?'':f.default_value);inDefault.placeholder='Можно оставить пустым';inDefault.oninput=(ev)=>{f.default_value=ev.target.value;};
      wDefault.appendChild(capDefault);wDefault.appendChild(inDefault);
    }else if(f.input_type==='number'){
      const inDefault=document.createElement('input');inDefault.type='number';inDefault.step='any';const start=normalizeNumberInput(f.default_value);inDefault.value=start===''?'':String(start);inDefault.placeholder='Например: 10';inDefault.oninput=(ev)=>{f.default_value=ev.target.value;};
      wDefault.appendChild(capDefault);wDefault.appendChild(inDefault);
    }else if(f.input_type==='date'){
      const inDefault=document.createElement('input');inDefault.type='date';inDefault.value=normalizeDateInput(f.default_value);inDefault.oninput=(ev)=>{f.default_value=ev.target.value;};
      wDefault.appendChild(capDefault);wDefault.appendChild(inDefault);
    }else if(f.input_type==='boolean'){
      const inDefault=document.createElement('select');
      [['true','Да'],['false','Нет']].forEach(item=>{const o=document.createElement('option');o.value=item[0];o.textContent=item[1];inDefault.appendChild(o);});
      inDefault.value=normalizeScalarFieldValue(f.default_value,'boolean')?'true':'false';
      inDefault.onchange=(ev)=>{f.default_value=ev.target.value==='true'?'true':'false';};
      wDefault.appendChild(capDefault);wDefault.appendChild(inDefault);
    }else if(f.input_type==='select'){
      const inDefault=document.createElement('select');
      const opts=parseInputOptions(f.options);
      const emptyOpt=document.createElement('option');emptyOpt.value='';emptyOpt.textContent='(пусто)';inDefault.appendChild(emptyOpt);
      opts.forEach(opt=>{const o=document.createElement('option');o.value=opt;o.textContent=opt;inDefault.appendChild(o);});
      const start=String(f.default_value==null?'':f.default_value);
      if(start&&opts.indexOf(start)<0){const extra=document.createElement('option');extra.value=start;extra.textContent=start;inDefault.appendChild(extra);}
      inDefault.value=start;
      inDefault.onchange=(ev)=>{f.default_value=ev.target.value;};
      wDefault.appendChild(capDefault);wDefault.appendChild(inDefault);
    }else{
      const inDefault=document.createElement('input');inDefault.type='text';inDefault.value=String(f.default_value==null?'':f.default_value);inDefault.placeholder='Можно оставить пустым';inDefault.oninput=(ev)=>{f.default_value=ev.target.value;};
      wDefault.appendChild(capDefault);wDefault.appendChild(inDefault);
    }
    fs.appendChild(wDefault);

    const wInputType=fieldWrap(false);
    const capInputType=document.createElement('span');capInputType.textContent='Тип поля ввода';
    const inInputType=document.createElement('select');
    INPUT_TYPES.forEach(item=>{const o=document.createElement('option');o.value=item[0];o.textContent=item[1];if(f.input_type===item[0])o.selected=true;inInputType.appendChild(o);});
    inInputType.onchange=(ev)=>{
      f.input_type=normalizeInputType(ev.target.value);
      if(f.input_type==='boolean')f.multiple=false;
      renderInputFields();
    };
    wInputType.appendChild(capInputType);wInputType.appendChild(inInputType);fs.appendChild(wInputType);

    if(f.input_type==='select'){
      const wOptions=fieldWrap(true);
      const capOptions=document.createElement('span');capOptions.textContent='Значения списка';
      const inOptions=document.createElement('textarea');inOptions.rows=2;inOptions.value=f.options;inOptions.placeholder='Вариант 1; Вариант 2; Вариант 3';inOptions.oninput=(ev)=>{f.options=ev.target.value;};
      const helpOptions=document.createElement('small');helpOptions.className='field-help';helpOptions.textContent='Можно разделять точкой с запятой, запятой или переносом строки.';
      wOptions.appendChild(capOptions);wOptions.appendChild(inOptions);wOptions.appendChild(helpOptions);fs.appendChild(wOptions);
    }

    const wRequired=fieldWrap(true);
    const inlRequired=document.createElement('span');inlRequired.className='field-inline';
    const inRequired=document.createElement('input');inRequired.type='checkbox';inRequired.checked=!!f.required;inRequired.onchange=(ev)=>{f.required=!!ev.target.checked;};
    const txtRequired=document.createElement('span');txtRequired.textContent='Обязательное поле';
    inlRequired.appendChild(inRequired);inlRequired.appendChild(txtRequired);wRequired.appendChild(inlRequired);fs.appendChild(wRequired);

    if(f.mode==='cells'){
      const wBinding=fieldWrap(false);
      const capBinding=document.createElement('span');capBinding.textContent='Привязка вставки';
      const inBinding=document.createElement('select');
      [['cells','По адресам ячеек'],['named','По именованным диапазонам']].forEach(item=>{const o=document.createElement('option');o.value=item[0];o.textContent=item[1];if(f.binding_mode===item[0])o.selected=true;inBinding.appendChild(o);});
      inBinding.onchange=(ev)=>{f.binding_mode=normalizeBindingMode(ev.target.value);if(f.binding_mode==='named'){f.multiple_insert=false;}renderInputFields();};
      wBinding.appendChild(capBinding);wBinding.appendChild(inBinding);fs.appendChild(wBinding);

      if(f.binding_mode==='named'){
        const wNamedTargets=fieldWrap(true);
        const capNamedTargets=document.createElement('span');capNamedTargets.textContent='Именованные диапазоны';
        const inNamedTargets=document.createElement('textarea');inNamedTargets.rows=2;inNamedTargets.value=f.named_targets;inNamedTargets.placeholder='CompanyName; ProductName';inNamedTargets.oninput=(ev)=>{f.named_targets=ev.target.value;};
        const helpNamed=document.createElement('small');helpNamed.className='field-help';helpNamed.textContent='Укажите имена диапазонов через ; , или перенос строки.';
        wNamedTargets.appendChild(capNamedTargets);wNamedTargets.appendChild(inNamedTargets);wNamedTargets.appendChild(helpNamed);fs.appendChild(wNamedTargets);
      }else{
        const wTargets=fieldWrap(true);
        const capTargets=document.createElement('span');capTargets.textContent='Ячейки/диапазоны';
        const inTargets=document.createElement('textarea');inTargets.rows=2;inTargets.value=f.targets;inTargets.placeholder='Лист1:A1,B2; Лист2:C3:D5';inTargets.oninput=(ev)=>{f.targets=ev.target.value;};
        const help=document.createElement('small');help.className='field-help';help.textContent='Формат: Лист:адрес1,адрес2; Лист2:адрес3. Адрес может быть ячейкой или диапазоном.';
        wTargets.appendChild(capTargets);wTargets.appendChild(inTargets);wTargets.appendChild(help);fs.appendChild(wTargets);

        const wMultiple=fieldWrap(true);
        const inlMultiple=document.createElement('span');inlMultiple.className='field-inline';
        const inMultiple=document.createElement('input');inMultiple.type='checkbox';inMultiple.checked=!!f.multiple;inMultiple.onchange=(ev)=>{f.multiple=!!ev.target.checked;renderInputFields();};
        const txtMultiple=document.createElement('span');txtMultiple.textContent='Можно вводить несколько значений (кнопка + в форме)';
        inlMultiple.appendChild(inMultiple);inlMultiple.appendChild(txtMultiple);wMultiple.appendChild(inlMultiple);fs.appendChild(wMultiple);

        if(f.multiple){
          const wDirection=fieldWrap(false);
          const capDirection=document.createElement('span');capDirection.textContent='Куда добавлять следующие значения';
          const inDirection=document.createElement('select');
          [['rows','По строкам вниз'],['columns','По столбцам вправо']].forEach(item=>{const o=document.createElement('option');o.value=item[0];o.textContent=item[1];if(f.multiple_direction===item[0])o.selected=true;inDirection.appendChild(o);});
          inDirection.onchange=(ev)=>{f.multiple_direction=ev.target.value==='columns'?'columns':'rows';};
          wDirection.appendChild(capDirection);wDirection.appendChild(inDirection);fs.appendChild(wDirection);

          const wInsert=fieldWrap(true);
          const inlInsert=document.createElement('span');inlInsert.className='field-inline';
          const inInsert=document.createElement('input');inInsert.type='checkbox';inInsert.checked=!!f.multiple_insert;inInsert.onchange=(ev)=>{f.multiple_insert=!!ev.target.checked;};
          const txtInsert=document.createElement('span');txtInsert.textContent='Добавлять новую строку/столбец для 2-го и следующих значений';
          inlInsert.appendChild(inInsert);inlInsert.appendChild(txtInsert);wInsert.appendChild(inlInsert);fs.appendChild(wInsert);
        }
      }

      const wKeepTpl=fieldWrap(true);
      const inlKeepTpl=document.createElement('span');inlKeepTpl.className='field-inline';
      const inKeepTpl=document.createElement('input');inKeepTpl.type='checkbox';inKeepTpl.checked=!!f.keep_template_format;inKeepTpl.onchange=(ev)=>{f.keep_template_format=!!ev.target.checked;renderInputFields();};
      const txtKeepTpl=document.createElement('span');txtKeepTpl.textContent='Формат текста как в шаблоне';
      inlKeepTpl.appendChild(inKeepTpl);inlKeepTpl.appendChild(txtKeepTpl);wKeepTpl.appendChild(inlKeepTpl);fs.appendChild(wKeepTpl);

      if(!f.keep_template_format){
        const wApplyAlign=fieldWrap(true);
        const inlApplyAlign=document.createElement('span');inlApplyAlign.className='field-inline';
        const inApplyAlign=document.createElement('input');inApplyAlign.type='checkbox';inApplyAlign.checked=!!f.apply_alignment;inApplyAlign.onchange=(ev)=>{f.apply_alignment=!!ev.target.checked;renderInputFields();};
        const txtApplyAlign=document.createElement('span');txtApplyAlign.textContent='Применять выравнивание';
        inlApplyAlign.appendChild(inApplyAlign);inlApplyAlign.appendChild(txtApplyAlign);wApplyAlign.appendChild(inlApplyAlign);fs.appendChild(wApplyAlign);

        if(f.apply_alignment){
          const wHorizontal=fieldWrap(false);
          const capHorizontal=document.createElement('span');capHorizontal.textContent='Горизонталь';
          const inHorizontal=document.createElement('select');
          H_ALIGN_OPTIONS.forEach(opt=>{const o=document.createElement('option');o.value=opt[0];o.textContent=opt[1];if(String(f.horizontal||'')===String(opt[0]))o.selected=true;inHorizontal.appendChild(o);});
          inHorizontal.onchange=(ev)=>{f.horizontal=String(ev.target.value||'');};
          wHorizontal.appendChild(capHorizontal);wHorizontal.appendChild(inHorizontal);fs.appendChild(wHorizontal);

          const wVertical=fieldWrap(false);
          const capVertical=document.createElement('span');capVertical.textContent='Вертикаль';
          const inVertical=document.createElement('select');
          V_ALIGN_OPTIONS.forEach(opt=>{const o=document.createElement('option');o.value=opt[0];o.textContent=opt[1];if(String(f.vertical||'')===String(opt[0]))o.selected=true;inVertical.appendChild(o);});
          inVertical.onchange=(ev)=>{f.vertical=String(ev.target.value||'');};
          wVertical.appendChild(capVertical);wVertical.appendChild(inVertical);fs.appendChild(wVertical);

          const wWrap=fieldWrap(true);
          const inlWrap=document.createElement('span');inlWrap.className='field-inline';
          const inWrap=document.createElement('input');inWrap.type='checkbox';inWrap.checked=!!f.wrap;inWrap.onchange=(ev)=>{f.wrap=!!ev.target.checked;};
          const txtWrap=document.createElement('span');txtWrap.textContent='Переносить текст';
          inlWrap.appendChild(inWrap);inlWrap.appendChild(txtWrap);wWrap.appendChild(inlWrap);fs.appendChild(wWrap);
        }

        const wApplyNum=fieldWrap(true);
        const inlApplyNum=document.createElement('span');inlApplyNum.className='field-inline';
        const inApplyNum=document.createElement('input');inApplyNum.type='checkbox';inApplyNum.checked=!!f.apply_number_format;inApplyNum.onchange=(ev)=>{f.apply_number_format=!!ev.target.checked;renderInputFields();};
        const txtApplyNum=document.createElement('span');txtApplyNum.textContent='Применять числовой формат';
        inlApplyNum.appendChild(inApplyNum);inlApplyNum.appendChild(txtApplyNum);wApplyNum.appendChild(inlApplyNum);fs.appendChild(wApplyNum);

        if(f.apply_number_format){
          const wNumFormat=fieldWrap(true);
          const capNumFormat=document.createElement('span');capNumFormat.textContent='Числовой формат';
          const numFormat=createNumberFormatField(f);
          wNumFormat.appendChild(capNumFormat);
          wNumFormat.appendChild(numFormat);
          fs.appendChild(wNumFormat);
        }

        const wApplyFont=fieldWrap(true);
        const inlApplyFont=document.createElement('span');inlApplyFont.className='field-inline';
        const inApplyFont=document.createElement('input');inApplyFont.type='checkbox';inApplyFont.checked=!!f.apply_font;inApplyFont.onchange=(ev)=>{f.apply_font=!!ev.target.checked;renderInputFields();};
        const txtApplyFont=document.createElement('span');txtApplyFont.textContent='Применять настройки шрифта';
        inlApplyFont.appendChild(inApplyFont);inlApplyFont.appendChild(txtApplyFont);wApplyFont.appendChild(inlApplyFont);fs.appendChild(wApplyFont);

        if(f.apply_font){
          const wFontName=fieldWrap(false);
          const capFontName=document.createElement('span');capFontName.textContent='Шрифт';
          const inFontName=createFontField(f,{k:'font_name'});
          wFontName.appendChild(capFontName);wFontName.appendChild(inFontName);fs.appendChild(wFontName);

          const wFontSize=fieldWrap(false);
          const capFontSize=document.createElement('span');capFontSize.textContent='Размер шрифта';
          const inFontSize=createFontSizeField(f,{k:'font_size'});
          wFontSize.appendChild(capFontSize);wFontSize.appendChild(inFontSize);fs.appendChild(wFontSize);

          const wFontColor=fieldWrap(false);
          const capFontColor=document.createElement('span');capFontColor.textContent='Цвет шрифта';
          const inFontColor=createColorField(f,{k:'font_color',auto:1});
          wFontColor.appendChild(capFontColor);wFontColor.appendChild(inFontColor);fs.appendChild(wFontColor);

          const wFontStyle=fieldWrap(true);
          const capFontStyle=document.createElement('span');capFontStyle.textContent='Начертание';
          const inFontStyle=createFontStyleField(f);
          wFontStyle.appendChild(capFontStyle);wFontStyle.appendChild(inFontStyle);fs.appendChild(wFontStyle);
        }

        const wApplyFill=fieldWrap(true);
        const inlApplyFill=document.createElement('span');inlApplyFill.className='field-inline';
        const inApplyFill=document.createElement('input');inApplyFill.type='checkbox';inApplyFill.checked=!!f.apply_fill;inApplyFill.onchange=(ev)=>{f.apply_fill=!!ev.target.checked;renderInputFields();};
        const txtApplyFill=document.createElement('span');txtApplyFill.textContent='Применять заливку';
        inlApplyFill.appendChild(inApplyFill);inlApplyFill.appendChild(txtApplyFill);wApplyFill.appendChild(inlApplyFill);fs.appendChild(wApplyFill);

        if(f.apply_fill){
          const wFillColor=fieldWrap(false);
          const capFillColor=document.createElement('span');capFillColor.textContent='Цвет заливки';
          const inFillColor=createColorField(f,{k:'fill_color',none:1});
          wFillColor.appendChild(capFillColor);wFillColor.appendChild(inFillColor);fs.appendChild(wFillColor);
        }
      }
    }else{
      const wToken=fieldWrap(false);
      const capToken=document.createElement('span');capToken.textContent='Ключ для замены';
      const inToken=document.createElement('input');inToken.type='text';inToken.value=f.token;inToken.placeholder='{{product}}';inToken.oninput=(ev)=>{f.token=ev.target.value;};
      wToken.appendChild(capToken);wToken.appendChild(inToken);fs.appendChild(wToken);

      const wScope=fieldWrap(false);
      const capScope=document.createElement('span');capScope.textContent='Где искать ключ';
      const inScope=document.createElement('select');
      [['workbook','Во всей книге'],['sheets','Только на выбранных листах']].forEach(opt=>{const o=document.createElement('option');o.value=opt[0];o.textContent=opt[1];if(f.scope===opt[0])o.selected=true;inScope.appendChild(o);});
      inScope.onchange=(ev)=>{f.scope=ev.target.value==='sheets'?'sheets':'workbook';renderInputFields();};
      wScope.appendChild(capScope);wScope.appendChild(inScope);fs.appendChild(wScope);

      if(f.scope==='sheets'){
        const wSheets=fieldWrap(true);
        const capSheets=document.createElement('span');capSheets.textContent='Листы для поиска';
        const inSheets=document.createElement('input');inSheets.type='text';inSheets.value=f.sheets;inSheets.placeholder='Лист1, Лист2';inSheets.oninput=(ev)=>{f.sheets=ev.target.value;};
        wSheets.appendChild(capSheets);wSheets.appendChild(inSheets);fs.appendChild(wSheets);
      }
    }

    item.appendChild(fs);
    root.appendChild(item);
  });
}
function readDraft(){if(!state.draft)return;state.draft.name=String(els.name.value||'').trim();state.draft.templatePath=String(els.tpl.value||'').trim();state.draft.description=String(els.descr.value||'').trim();}
function validateDraft(){
  if(!state.draft)return 'Сценарий не открыт.'; readDraft();
  if(!state.draft.name)return 'Введите название сценария.';
  if(!state.draft.templatePath)return 'Выберите файл шаблона.';
  if(/fakepath/i.test(state.draft.templatePath))return 'Получен fakepath. Укажите полный путь к шаблону.';
  if(!Array.isArray(state.draft.actions))state.draft.actions=[];
  if(!Array.isArray(state.draft.inputFields))state.draft.inputFields=[];
  if(!state.draft.actions.length&&!state.draft.inputFields.length)return 'Добавьте хотя бы одно действие или поле ввода.';
  const normalizedInputs=state.draft.inputFields.map(normInputField);
  state.draft.inputFields=normalizedInputs;
  const inputFieldIds=normalizedInputs.map(field=>field.id);
  for(let i=0;i<state.draft.actions.length;i++){
    const a=normAction(state.draft.actions[i]);state.draft.actions[i]=a;if(!schema[a.type])return `Шаг ${i+1}: неизвестный тип действия.`;
    for(const f of schema[a.type].f||[]){if(!f.r||f.t==='check')continue;const v=a[f.k];if(v==null||String(v).trim()==='')return `Шаг ${i+1} (${schema[a.type].label}): заполните поле «${f.l}».`;}
    if(a.cond_enabled){
      if(!a.cond_field)return `Шаг ${i+1} (${schema[a.type].label}): выберите поле в условии.`;
      if(inputFieldIds.indexOf(a.cond_field)<0)return `Шаг ${i+1} (${schema[a.type].label}): поле условия не найдено.`;
      if(conditionNeedsValue(a.cond_operator)&&String(a.cond_value==null?'':a.cond_value).trim()==='')
        return `Шаг ${i+1} (${schema[a.type].label}): укажите значение для условия.`;
    }
  }
  for(let i=0;i<normalizedInputs.length;i+=1){
    const field=normalizedInputs[i];
    const label=inputFieldLabel(field,i);
    if(!field.name)return `Поле ввода ${i+1}: укажите название.`;
    if(field.input_type==='select'&&!parseInputOptions(field.options).length)return `${label}: добавьте хотя бы один вариант списка.`;
    if(field.mode==='cells'){
      if(field.binding_mode==='named'){
        const names=parseNamedTargets(field.named_targets);
        if(!names.length)return `${label}: укажите хотя бы один именованный диапазон.`;
        if(field.multiple)return `${label}: множественный ввод доступен только для адресов ячеек.`;
      }else{
        const parsed=parseCellBindings(field.targets);
        if(parsed.error)return `${label}: ${parsed.error}`;
        if(field.multiple&&field.multiple_insert){
          for(let t=0;t<parsed.items.length;t+=1){
            if(!parseA1Range(parsed.items[t].range))return `${label}: для авто-добавления строк/столбцов используйте адреса формата A1 или A1:C3.`;
          }
        }
      }
    }else{
      if(!field.token)return `${label}: укажите ключ для замены.`;
      if(field.scope==='sheets'&&!parseSheetList(field.sheets).length)
        return `${label}: укажите хотя бы один лист для поиска ключа.`;
      if(field.multiple)return `${label}: множественный ввод сейчас доступен только для вставки в ячейки/диапазоны.`;
    }
  }
  return '';
}
function dbReason(err){
  const d=err&&err.details?err.details:null;
  return d&&d.stderr?d.stderr:(d&&d.error?d.error:(err&&err.message?err.message:'Неизвестная ошибка'));
}
async function prepareTemplatePath(draft){
  const source=resolveTemplate(draft.templatePath);
  if(!source)throw new Error('Некорректный путь к шаблону.');
  const targetName=sanitizeTemplateName(draft.name);
  const targetPath=joinPath(getTemplatesDir(),`${targetName}.${SCENARIO_TEMPLATE_EXT}`);
  if(samePath(source,targetPath))return targetPath;
  const requestId=reqId('docbuilder-template');
  const runPayload={requestId,script:SCRIPT,openAfterRun:false,argument:{templatePath:source,outputPath:targetPath,scenarioId:draft.id||'',scenarioName:draft.name||'',stopOnError:true,actions:[]}};
  const runPromise=runDb(runPayload); runPromise.catch(()=>{});
  const fallbackMs=12000;
  const res=await Promise.race([
    runPromise,
    new Promise(resolve=>setTimeout(()=>resolve({ok:true,requestId,exitCode:0,fallback:true}),fallbackMs))
  ]);
  if(res&&res.fallback){
    clearPending(requestId);
    return targetPath;
  }
  if(!res||!res.ok)throw dbErr(res||{error:'template_prepare_failed'});
  if(typeof res.exitCode!=='number'||res.exitCode!==0)throw dbErr(res||{error:'template_prepare_failed'});
  return targetPath;
}
async function saveDraft(){
  const err=validateDraft();if(err){toast(err,'error');return;}
  const draftRef=state.draft;
  const btn=els.btnSave;
  if(btn){btn.disabled=true;btn.textContent='Сохранение...';}
  try{
    toast('Подготовка шаблона...','ok');
    const preparedPath=await prepareTemplatePath(draftRef);
    if(state.draft!==draftRef||!state.draft)return;
    state.draft.templatePath=preparedPath;
    els.tpl.value=preparedPath;
    const s={id:state.draft.id,name:state.draft.name,description:state.draft.description,templatePath:state.draft.templatePath,actions:state.draft.actions.map(normAction),inputFields:(state.draft.inputFields||[]).map(normInputField),createdAt:state.draft.createdAt||new Date().toISOString(),updatedAt:new Date().toISOString(),lastRunAt:state.draft.lastRunAt||null};
    const idx=state.scenarios.findIndex(x=>x.id===s.id);if(idx>=0){state.scenarios[idx]=s;toast('Сценарий обновлен.','ok');}else{state.scenarios.push(s);toast('Сценарий создан.','ok');}
    save();renderList();closeModal();
  }catch(e){
    toast(`Не удалось подготовить шаблон: ${dbReason(e)}`,'error');
  }finally{
    if(btn){btn.disabled=false;btn.textContent='Сохранить сценарий';}
  }
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
function openTemplateInEditor(id){
  const sc=state.scenarios.find(x=>x.id===id);
  if(!sc||!sc.templatePath){toast('Шаблон не задан.','error');return;}
  const tpl=resolveTemplate(sc.templatePath);
  if(!tpl){toast('Некорректный путь к шаблону.','error');return;}
  openResult(tpl,id);
}

function closeRunInputModal(result){
  const ctx=state.runInput;
  if(!ctx)return;
  state.runInput=null;
  if(els.runInputModal)els.runInputModal.classList.add('hidden');
  if(els.runInputFields)els.runInputFields.innerHTML='';
  renderList();
  if(typeof ctx.resolve==='function')ctx.resolve(result);
}

function createRuntimeInputControl(field,initialValue){
  const type=normalizeInputType(field.input_type);
  if(type==='multiline'){
    const ta=document.createElement('textarea');
    ta.rows=3;
    ta.value=String(initialValue==null?'':initialValue);
    ta.placeholder=field.placeholder||'Введите значение';
    return {node:ta,getValue:()=>String(ta.value==null?'':ta.value),focus:()=>ta.focus()};
  }
  if(type==='select'){
    const sel=document.createElement('select');
    const options=parseInputOptions(field.options);
    if(!field.required){
      const empty=document.createElement('option');
      empty.value='';
      empty.textContent='(пусто)';
      sel.appendChild(empty);
    }
    options.forEach(opt=>{
      const o=document.createElement('option');
      o.value=opt;
      o.textContent=opt;
      sel.appendChild(o);
    });
    const start=String(initialValue==null?'':initialValue);
    if(start&&options.indexOf(start)<0){
      const extra=document.createElement('option');
      extra.value=start;
      extra.textContent=start;
      sel.insertBefore(extra,sel.firstChild);
    }
    sel.value=start;
    return {node:sel,getValue:()=>String(sel.value==null?'':sel.value),focus:()=>sel.focus()};
  }
  if(type==='boolean'){
    const wrap=document.createElement('label');
    wrap.className='field-inline runtime-bool';
    const chk=document.createElement('input');
    chk.type='checkbox';
    chk.checked=!!normalizeScalarFieldValue(initialValue,'boolean');
    const txt=document.createElement('span');
    txt.textContent=field.placeholder||'Да';
    wrap.appendChild(chk);
    wrap.appendChild(txt);
    return {node:wrap,getValue:()=>!!chk.checked,focus:()=>chk.focus()};
  }
  const inp=document.createElement('input');
  if(type==='number'){
    inp.type='number';
    inp.step='any';
    const num=normalizeNumberInput(initialValue);
    inp.value=num===''?'':String(num);
  }else if(type==='date'){
    inp.type='date';
    inp.value=normalizeDateInput(initialValue);
  }else{
    inp.type='text';
    inp.value=String(initialValue==null?'':initialValue);
  }
  inp.placeholder=field.placeholder||'Введите значение';
  return {node:inp,getValue:()=>inp.value,focus:()=>inp.focus()};
}

function createRuntimeFieldControl(field){
  if(field.multiple){
    const box=document.createElement('div');
    box.className='runtime-multi-box';
    const list=document.createElement('div');
    list.className='runtime-multi-list';
    const actions=document.createElement('div');
    actions.className='runtime-multi-actions';
    const addBtn=document.createElement('button');
    addBtn.type='button';
    addBtn.className='btn btn-ghost runtime-multi-add';
    addBtn.textContent='+ Добавить значение';
    actions.appendChild(addBtn);
    box.appendChild(list);
    box.appendChild(actions);
    const items=[];
    const makeRow=(value)=>{
      const row=document.createElement('div');
      row.className='runtime-multi-row';
      const ctrl=createRuntimeInputControl(field,value);
      const delBtn=document.createElement('button');
      delBtn.type='button';
      delBtn.className='btn btn-danger runtime-multi-remove';
      delBtn.textContent='-';
      const item={row,ctrl,delBtn};
      delBtn.onclick=(ev)=>{
        ev.preventDefault();
        if(field.required&&items.length<=1)return;
        const idx=items.indexOf(item);
        if(idx>=0)items.splice(idx,1);
        row.remove();
      };
      row.appendChild(ctrl.node);
      row.appendChild(delBtn);
      list.appendChild(row);
      items.push(item);
      return item;
    };
    const defaults=normalizeRuntimeFieldValues(field.default_value,field);
    if(defaults.length)defaults.forEach(v=>makeRow(v));
    else makeRow('');
    addBtn.onclick=(ev)=>{ev.preventDefault();makeRow('');};
    return {
      node:box,
      getValue:()=>items.map(item=>normalizeScalarFieldValue(item.ctrl.getValue(),field.input_type)),
      focus:()=>{if(items.length&&items[0].ctrl&&typeof items[0].ctrl.focus==='function')items[0].ctrl.focus();}
    };
  }
  const ctrl=createRuntimeInputControl(field,normalizeScalarFieldValue(field.default_value,field.input_type));
  return {node:ctrl.node,getValue:()=>normalizeScalarFieldValue(ctrl.getValue(),field.input_type),focus:ctrl.focus};
}

function submitRunInputModal(){
  const ctx=state.runInput;
  if(!ctx)return;
  const values={};
  for(let i=0;i<ctx.fields.length;i+=1){
    const field=ctx.fields[i];
    const input=ctx.inputs[field.id];
    const rawValue=input&&typeof input.getValue==='function'?input.getValue():'';
    let value;
    if(field.multiple){
      value=normalizeRuntimeFieldValues(rawValue,field);
      if(field.required&&value.length===0){
        toast(`Заполните поле «${inputFieldLabel(field,i)}».`,'error');
        if(input&&typeof input.focus==='function')input.focus();
        return;
      }
    }else{
      value=normalizeScalarFieldValue(rawValue,field.input_type);
      if(field.required&&!isFieldValueFilled(value,field.input_type)){
        toast(`Заполните поле «${inputFieldLabel(field,i)}».`,'error');
        if(input&&typeof input.focus==='function')input.focus();
        return;
      }
    }
    if(field.required&&field.input_type==='number'&&!field.multiple&&value===''){
      toast(`Заполните поле «${inputFieldLabel(field,i)}».`,'error');
      if(input&&typeof input.focus==='function')input.focus();
      return;
    }
    values[field.id]=value;
  }
  closeRunInputModal(values);
}

function openRunInputModal(sc){
  const fields=(Array.isArray(sc&&sc.inputFields)?sc.inputFields:[]).map(normInputField);
  if(!fields.length)return Promise.resolve({});
  return new Promise(resolve=>{
    const inputs=Object.create(null);
    state.runInput={resolve,scenarioId:sc.id,fields,inputs};
    renderList();
    if(els.runInputTitle)els.runInputTitle.textContent=`Заполнение: ${sc.name}`;
    if(els.runInputFields){
      els.runInputFields.innerHTML='';
      fields.forEach((field,i)=>{
        const wrap=fieldWrap(true);
        const cap=document.createElement('span');
        cap.textContent=field.required?`${inputFieldLabel(field,i)} *`:inputFieldLabel(field,i);
        const ctrl=createRuntimeFieldControl(field);
        inputs[field.id]=ctrl;
        wrap.appendChild(cap);
        wrap.appendChild(ctrl.node);
        if(field.mode==='cells'){
          const hint=document.createElement('small');
          hint.className='field-help';
          if(field.binding_mode==='named')hint.textContent=`Вставка в именованные диапазоны: ${field.named_targets||'-'}`;
          else if(field.multiple)hint.textContent=`Вставка в: ${field.targets||'-'} (${field.multiple_direction==='columns'?'по столбцам':'по строкам'}${field.multiple_insert?', с добавлением':''})`;
          else hint.textContent=`Вставка в: ${field.targets||'-'}`;
          wrap.appendChild(hint);
        }else{
          const hint=document.createElement('small');
          hint.className='field-help';
          hint.textContent=field.scope==='sheets'?`Ключ ${field.token||'-'} (листы: ${field.sheets||'-'})`:`Ключ ${field.token||'-'} (во всей книге)`;
          wrap.appendChild(hint);
        }
        els.runInputFields.appendChild(wrap);
      });
    }
    if(els.runInputModal)els.runInputModal.classList.remove('hidden');
    const first=fields.length&&inputs[fields[0].id];
    if(first&&typeof first.focus==='function')setTimeout(()=>first.focus(),10);
  });
}

function buildInputActions(sc,inputValues){
  const fields=(Array.isArray(sc&&sc.inputFields)?sc.inputFields:[]).map(normInputField);
  if(!fields.length)return [];
  const values=isObj(inputValues)?inputValues:{};
  const out=[];
  const makeSetCellValueAction=(field,sheet,range,value)=>{
    const preferredMode=valueModeByInputType(field.input_type);
    const mode=(preferredMode==='number'&&value==='')?'text':preferredMode;
    const action={
      id:uid('action'),
      type:'set_cell_value',
      sheet:sheet||'',
      range,
      named_range:field.binding_mode==='named',
      mode,
      value,
      merge:false,
      keep_template_format:!!field.keep_template_format,
      apply_alignment:!!field.apply_alignment,
      horizontal:String(field.horizontal||''),
      vertical:String(field.vertical||''),
      wrap:!!field.wrap,
      apply_number_format:!!field.apply_number_format,
      format_preset:String(field.format_preset||'general'),
      decimals:String(field.decimals||'2'),
      use_thousands:!!field.use_thousands,
      currency_symbol:String(field.currency_symbol||'₽'),
      negative_red:!!field.negative_red,
      custom_format:String(field.custom_format||'General'),
      format:String(field.format||'General'),
      apply_font:!!field.apply_font,
      font_name:String(field.font_name||'Arial'),
      font_size:String(field.font_size||'11'),
      bold:!!field.bold,
      italic:!!field.italic,
      underline:String(field.underline||'none'),
      strikeout:!!field.strikeout,
      font_color:String(field.font_color||'auto'),
      apply_fill:!!field.apply_fill,
      fill_color:String(field.fill_color||'none')
    };
    normalizeInsertStyleSettings(action,action);
    return action;
  };
  for(let i=0;i<fields.length;i+=1){
    const field=fields[i];
    const label=inputFieldLabel(field,i);
    const hasValue=Object.prototype.hasOwnProperty.call(values,field.id);
    const rawValue=hasValue?values[field.id]:field.default_value;
    const valueItems=field.multiple?normalizeRuntimeFieldValues(rawValue,field):normalizeRuntimeFieldValues(rawValue,field).slice(0,1);
    if(!valueItems.length)continue;
    if(field.mode==='cells'){
      if(field.binding_mode==='named'){
        const names=parseNamedTargets(field.named_targets);
        if(!names.length)throw new Error(`${label}: укажите хотя бы один именованный диапазон.`);
        valueItems.forEach(value=>{
          names.forEach(name=>out.push(makeSetCellValueAction(field,'',name,value)));
        });
        continue;
      }
      const parsed=parseCellBindings(field.targets);
      if(parsed.error)throw new Error(`${label}: ${parsed.error}`);
      const targets=parsed.items.map(item=>({
        sheet:String(item.sheet||''),
        range:String(item.range||''),
        parsedRange:parseA1Range(item.range)
      }));
      if(field.multiple&&field.multiple_insert){
        for(let t=0;t<targets.length;t+=1){
          if(!targets[t].parsedRange)throw new Error(`${label}: для добавления строк/столбцов используйте адреса A1 или A1:C3.`);
        }
      }
      for(let valueIndex=0;valueIndex<valueItems.length;valueIndex+=1){
        const value=valueItems[valueIndex];
        if(field.multiple&&field.multiple_insert&&valueIndex>0){
          const sheetStarts=Object.create(null);
          targets.forEach(target=>{
            const key=target.sheet.toLowerCase();
            const anchor=target.parsedRange;
            if(!anchor)return;
            const start=field.multiple_direction==='columns'?anchor.c1+valueIndex:anchor.r1+valueIndex;
            if(sheetStarts[key]===undefined||start<sheetStarts[key])sheetStarts[key]=start;
          });
          Object.keys(sheetStarts).forEach(key=>{
            const originalSheet=targets.find(t=>t.sheet.toLowerCase()===key);
            if(!originalSheet)return;
            if(field.multiple_direction==='columns'){
              out.push({
                id:uid('action'),
                type:'insert_columns',
                sheet:originalSheet.sheet,
                start_column:numberToColLetters(sheetStarts[key]),
                count:'1'
              });
            }else{
              out.push({
                id:uid('action'),
                type:'insert_rows',
                sheet:originalSheet.sheet,
                start_row:String(sheetStarts[key]),
                count:'1'
              });
            }
          });
        }
        targets.forEach(target=>{
          let targetRange=target.range;
          if(field.multiple&&valueIndex>0){
            const shifted=offsetA1Range(target.range,valueIndex,field.multiple_direction);
            if(!shifted)throw new Error(`${label}: диапазон ${target.range} не поддерживает смещение. Используйте A1 или A1:C3.`);
            targetRange=formatA1Range(shifted);
          }
          out.push(makeSetCellValueAction(field,target.sheet,targetRange,value));
        });
      }
      continue;
    }
    const token=String(field.token||'').trim();
    if(!token)continue;
    const scope=field.scope==='sheets'?'sheets':'workbook';
    const sheets=scope==='sheets'?parseSheetList(field.sheets):[];
    valueItems.forEach(value=>{
      out.push({
        id:uid('action'),
        type:'replace_token',
        token,
        value:String(value==null?'':value),
        scope,
        sheets
      });
    });
  }
  return out;
}

function buildRunActions(sc,inputValues){
  const base=(Array.isArray(sc&&sc.actions)?sc.actions.map(normAction):[]).filter(action=>evaluateActionCondition(action,inputValues,sc));
  const extra=buildInputActions(sc,inputValues);
  return extra.concat(base);
}

async function runScenario(id){
  const sc=state.scenarios.find(x=>x.id===id);if(!sc||state.running||state.runInput)return;
  const tpl=resolveTemplate(sc.templatePath);if(!tpl){toast('Некорректный путь к шаблону.','error');return;}
  const inputValues=await openRunInputModal(sc);
  if(inputValues===null)return;
  state.running=id;renderList();
  try{
    const runActions=buildRunActions(sc,inputValues);
    toast('Проверка DocumentBuilder...','ok');
    const probe=await ensureProbe(false);
    const out=mkOutputPath(sc,probe,tpl);
    toast('Выполнение сценария через DocumentBuilder...','ok');
    const runRequestId=reqId('docbuilder-run');
    const runPayload={requestId:runRequestId,script:SCRIPT,openAfterRun:false,argument:{templatePath:tpl,outputPath:out,scenarioId:sc.id,scenarioName:sc.name,stopOnError:true,actions:runActions}};
    const runPromise=runDb(runPayload); runPromise.catch(()=>{});
    const fallbackMs=Math.max(8000,Math.min(120000,5000+(Array.isArray(runActions)?runActions.length:1)*1500));
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
  const templatesDir=getTemplatesDir();
  try{
    const sdk=window.parent&&window.parent.sdk;
    if(sdk&&typeof sdk.command==='function'&&templatesDir)sdk.command('files:setOpenPath',templatesDir);
  }catch(_){ }
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
  els.btnAddInputField&&els.btnAddInputField.addEventListener('click',()=>{if(!state.draft)return;state.draft.inputFields.push(mkInputField());renderInputFields();});
  els.btnAddAction&&els.btnAddAction.addEventListener('click',()=>{if(!state.draft)return;state.draft.actions.push(mkAction('set_cell_value'));renderActions();});
  els.btnPick&&els.btnPick.addEventListener('click',pickTemplate);
  els.btnRunInputCancel&&els.btnRunInputCancel.addEventListener('click',()=>closeRunInputModal(null));
  els.btnRunInputSubmit&&els.btnRunInputSubmit.addEventListener('click',submitRunInputModal);
  els.file&&els.file.addEventListener('change',()=>{const f=els.file.files&&els.file.files[0];if(f&&f.path){els.tpl.value=toFsPath(f.path);if(state.draft)state.draft.templatePath=els.tpl.value;return;}const v=els.file.value||'';if(v&&!/fakepath/i.test(v)){els.tpl.value=toFsPath(v);if(state.draft)state.draft.templatePath=els.tpl.value;return;}toast('Введите полный путь к шаблону вручную.','error');});
  els.modal&&els.modal.addEventListener('click',e=>{if(e.target&&e.target.classList&&e.target.classList.contains('modal-backdrop'))closeModal();});
  els.runInputModal&&els.runInputModal.addEventListener('click',e=>{if(e.target&&e.target.classList&&e.target.classList.contains('modal-backdrop'))closeRunInputModal(null);});
  document.addEventListener('click',()=>closeColorMenus());
  document.addEventListener('keydown',e=>{
    if(e.key==='Escape'){
      closeColorMenus();
      if(els.runInputModal&&!els.runInputModal.classList.contains('hidden')){closeRunInputModal(null);return;}
      if(els.modal&&!els.modal.classList.contains('hidden'))closeModal();
      return;
    }
    if((e.key==='Enter'||e.keyCode===13)&&els.runInputModal&&!els.runInputModal.classList.contains('hidden')){
      const tag=e.target&&e.target.tagName?String(e.target.tagName).toLowerCase():'';
      if(tag!=='textarea'){
        e.preventDefault();
        submitRunInputModal();
      }
    }
  });
  window.addEventListener('message',e=>{
    const m=parse(e.data);
    if(!isObj(m))return;
    if(m.event==='uiThemeChanged'){handleThemeBridgeMessage(m.data);return;}
    if(m.event==='reportsDocBuilderResult'&&m.data)onDbResult(m.data);
    if(m.event==='reportsDocBuilderProbeResult'&&m.data)onDbResult(m.data);
  });
}

function init(){bindThemeSync();load();bind();renderList();window.ReportsDocBuilder={run:runDb,probe:probeDb};}
init();
})();
