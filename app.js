
(()=>{
'use strict';

const STORAGE='reports.scenarios.v3';
const SCENARIOS_LEGACY_FILE='data\\scenarios.json';
const SCENARIOS_DIR='data\\scenarios';
const SCENARIOS_INDEX_FILE='data\\scenarios\\index.json';
const SCRIPT='docbuilder/scripts/reports_executor.docbuilder';
const SOURCE_PROBE_SCRIPT='docbuilder/scripts/reports_source_probe.docbuilder';
const SOURCE_PROBE_PREFIX='__REPORTS_SOURCE_OK__';
const PROBE_TTL=900000;
const SOURCE_PROBE_CACHE_TTL=900000;
const PENDING=Object.create(null);
const WATCH=Object.create(null);
const PATH_CHECK_PENDING=Object.create(null);
const FILE_PENDING=Object.create(null);
const SOURCE_PROBE_CACHE=Object.create(null);
const DB_SCRIPT_CACHE=Object.create(null);
const SCENARIO_TEMPLATE_EXT='xlsx';
const GENERATED_DIR='generated';
const THEME_IDS=['theme-light','theme-classic-light','theme-dark','theme-contrast-dark','theme-gray','theme-white','theme-night'];

const state={scenarios:[],scenarioFiles:Object.create(null),persistScenariosPromise:Promise.resolve(),q:'',draft:null,running:null,probe:null,toastTimer:null,runInput:null,listEditor:null,exportPack:null,inputCollapsed:false,actionCollapsed:false,collapsedInputItems:Object.create(null),collapsedActionItems:Object.create(null),fileBridge:'unknown',fileSaveWarned:false};

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
  ]},
  insert_source_field:{label:'Вставить поле из файла',d:{source_field_id:'',targets:'Лист1:A1',keep_template_format:true},f:[
    {k:'source_field_id',l:'Поле из файла',t:'select',dynamicOptions:'source_fields',r:1},
    {k:'targets',l:'Точки вставки',t:'textarea',r:1,full:1,p:'Лист1:A1\\nЛист2:F10\\nСвод:C5,D5'},
    {k:'keep_template_format',l:'Формат текста как в шаблоне',t:'check',full:1}
  ]},
  insert_source_list:{label:'Вставить список из файла',d:{source_list_id:'',targets:'Лист1:C13|A13:K13',insert_mode:'insert_rows',keep_template_format:true},f:[
    {k:'source_list_id',l:'Список из файла',t:'select',dynamicOptions:'source_lists',r:1},
    {k:'targets',l:'Точки вставки',t:'textarea',r:1,full:1,p:'Ф1:C13|A13:K13\\nСвод:D20|A20:H20'},
    {k:'insert_mode',l:'Режим вставки',t:'select',o:[['insert_rows','Вставлять строки'],['overwrite','Записывать поверх']]},
    {k:'keep_template_format',l:'Формат строк как в шаблоне',t:'check',full:1}
  ]},
  insert_source_table:{label:'Вставить таблицу из файла',d:{source_table_id:'',targets:'Лист1:A15|A15:N15',insert_mode:'insert_rows',keep_template_format:true,mappings:''},f:[
    {k:'source_table_id',l:'Таблица из файла',t:'select',dynamicOptions:'source_tables',r:1,redraw:1},
    {k:'targets',l:'Точки вставки',t:'textarea',r:1,full:1,p:'Лист1:A15|A15:N15\\nЛист2:A40|A40:N40'},
    {k:'insert_mode',l:'Режим вставки',t:'select',o:[['insert_rows','Вставлять строки'],['overwrite','Записывать поверх']]},
    {k:'keep_template_format',l:'Формат строк как в шаблоне',t:'check',full:1},
    {k:'mappings',l:'Соответствие колонок',t:'textarea',r:1,full:1,p:'number=A\\ntitle=B\\nprice=C\\ndate=D'}
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
  btnExportPack:document.getElementById('btn-export-pack'),btnImportPack:document.getElementById('btn-import-pack'),
  importPackInput:document.getElementById('import-pack-input'),
  btnClose:document.getElementById('btn-close-modal'),btnCancel:document.getElementById('btn-cancel'),btnSave:document.getElementById('btn-save'),
  name:document.getElementById('scenario-name'),tpl:document.getElementById('scenario-template'),descr:document.getElementById('scenario-description'),
  inputList:document.getElementById('input-field-list'),btnAddInputField:document.getElementById('btn-add-input-field'),
  btnAddInputFieldBottom:document.getElementById('btn-add-input-field-bottom'),
  btnToggleInputSection:document.getElementById('btn-toggle-input-section'),
  inputSectionBody:document.getElementById('input-section-body'),
  sourceConfigRoot:document.getElementById('source-config-root'),
  btnAddSourceField:document.getElementById('btn-add-source-field'),
  btnAddSourceList:document.getElementById('btn-add-source-list'),
  btnAddSourceTable:document.getElementById('btn-add-source-table'),
  btnPick:document.getElementById('btn-pick-template'),file:document.getElementById('template-file-input'),actionList:document.getElementById('action-list'),
  btnAddAction:document.getElementById('btn-add-action'),
  btnAddActionBottom:document.getElementById('btn-add-action-bottom'),
  btnToggleActionSection:document.getElementById('btn-toggle-action-section'),
  actionSectionBody:document.getElementById('action-section-body'),
  toast:document.getElementById('toast'),
  runInputModal:document.getElementById('run-input-modal'),runInputTitle:document.getElementById('run-input-title'),
  runInputFields:document.getElementById('run-input-fields'),btnRunInputCancel:document.getElementById('btn-run-input-cancel'),
  btnRunInputSubmit:document.getElementById('btn-run-input-submit'),
  runSourceSection:document.getElementById('run-source-section'),
  runSourceSubtitle:document.getElementById('run-source-subtitle'),
  btnRunSourcePick:document.getElementById('btn-run-source-pick'),
  runSourceFileInput:document.getElementById('run-source-file-input'),
  runSourceFilePath:document.getElementById('run-source-file-path'),
  runSourceStatus:document.getElementById('run-source-status'),
  runSourceStatusText:document.getElementById('run-source-status-text'),
  runSourceError:document.getElementById('run-source-error'),
  runSourcePreview:document.getElementById('run-source-preview'),
  listEditorModal:document.getElementById('list-editor-modal'),listEditorTitle:document.getElementById('list-editor-title'),
  listEditorSubtitle:document.getElementById('list-editor-subtitle'),listEditorText:document.getElementById('list-editor-text'),
  btnListEditorClose:document.getElementById('btn-list-editor-close'),btnListEditorCancel:document.getElementById('btn-list-editor-cancel'),
  btnListEditorApply:document.getElementById('btn-list-editor-apply'),
  exportPackModal:document.getElementById('export-pack-modal'),exportPackList:document.getElementById('export-pack-list'),
  btnExportPackClose:document.getElementById('btn-export-pack-close'),btnExportPackCancel:document.getElementById('btn-export-pack-cancel'),
  btnExportPackSubmit:document.getElementById('btn-export-pack-submit')
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
async function ensureReadableTemplatePath(raw){
  const resolved=resolveTemplate(raw);
  if(!resolved)throw new Error('Некорректный путь к шаблону.');
  const ok=await canReadBinaryFile(resolved,2000);
  if(!ok)throw new Error(`Файл шаблона не найден: ${resolved}`);
  return resolved;
}
function isAbsPath(p){const s=normSlashes(p);return /^[a-zA-Z]:\\/.test(s)||/^\\\\/.test(s);}
function getReportsUiDir(){const root=getRoot();return root?joinPath(root,'reports-ui'):'reports-ui';}
function getTemplatesDir(){return joinPath(getReportsUiDir(),'templates');}
function getLegacyScenariosPath(){return joinPath(getReportsUiDir(),SCENARIOS_LEGACY_FILE);}
function getScenariosDir(){return joinPath(getReportsUiDir(),SCENARIOS_DIR);}
function getScenariosIndexPath(){return joinPath(getReportsUiDir(),SCENARIOS_INDEX_FILE);}
function getScenarioPath(fileName){return joinPath(getScenariosDir(),String(fileName||''));}
function getScenarioTemplatePath(name){return joinPath(getTemplatesDir(),`${sanitizeTemplateName(name)}.${SCENARIO_TEMPLATE_EXT}`);}
function isManagedScenarioTemplatePath(path,scenarioName){
  const resolved=resolveTemplate(path);
  if(!resolved)return false;
  return samePath(resolved,getScenarioTemplatePath(scenarioName));
}
function reportsDirFromProbe(probe){
  const runtime=probe&&probe.runtimeDir?normSlashes(probe.runtimeDir):'';
  if(!runtime||!isAbsPath(runtime))return '';
  const probeReports=dirName(dirName(runtime));
  return /\\reports-ui$/i.test(probeReports)?probeReports:'';
}
function getGeneratedDir(probe,tpl){
  const configured=getReportsUiDir();
  if(isAbsPath(configured))return joinPath(configured,GENERATED_DIR);
  const fromProbe=reportsDirFromProbe(probe);
  if(fromProbe)return joinPath(fromProbe,GENERATED_DIR);
  const base=dirName(tpl||'');
  return base?joinPath(base,GENERATED_DIR):joinPath('reports-ui',GENERATED_DIR);
}
function resolveReportsPath(raw){
  const p=toFsPath(raw);
  if(!p)return '';
  if(isAbsPath(p))return p;
  return joinPath(getReportsUiDir(),p.replace(/^\.\\/,'').replace(/^reports-ui\\/i,''));
}
function stripBom(text){
  return String(text||'').replace(/^\uFEFF/,'');
}
function toAsciiJsLiteral(value){
  const json=JSON.stringify(value);
  return String(json==null?'null':json)
    .replace(/[\u2028\u2029]/g,ch=>`\\u${ch.charCodeAt(0).toString(16).padStart(4,'0')}`)
    .replace(/[^\x20-\x7E]/g,ch=>`\\u${ch.charCodeAt(0).toString(16).padStart(4,'0')}`);
}
function toDocbuilderLiteralPath(path){
  return normSlashes(path||'').replace(/\\/g,'/');
}
function applyDocbuilderLiterals(source,literals){
  let out=String(source||'');
  const map=isObj(literals)?literals:{};
  Object.keys(map).forEach(key=>{
    const token=`\"__REPORTS_LITERAL_${String(key)}__\"`;
    out=out.split(token).join(JSON.stringify(toDocbuilderLiteralPath(String(map[key]||''))));
  });
  return out;
}
function pickDocbuilderWorkDir(probe){
  const runtimeDir=probe&&probe.runtimeDir?toFsPath(probe.runtimeDir):'';
  if(runtimeDir)return runtimeDir;
  const runtimeExe=probe&&probe.runtimeExe?toFsPath(probe.runtimeExe):'';
  return runtimeExe?dirName(runtimeExe):'';
}
async function loadDocbuilderScript(scriptPath){
  const resolved=resolveReportsPath(scriptPath);
  if(!resolved)throw new Error(`Не найден путь к скрипту DocumentBuilder: ${scriptPath}`);
  if(Object.prototype.hasOwnProperty.call(DB_SCRIPT_CACHE,resolved))return DB_SCRIPT_CACHE[resolved];
  const res=await readTextFile(resolved,5000);
  if(!res||!res.ok)throw fileErr(res||{error:'script_read_failed',path:resolved});
  const text=stripBom(res.content||'');
  DB_SCRIPT_CACHE[resolved]=text;
  return text;
}
async function materializeDocbuilderScript(scriptPath,variables,tag,literals){
  const base=applyDocbuilderLiterals(await loadDocbuilderScript(scriptPath),literals);
  return materializeDocbuilderSource(base,variables,tag);
}
async function materializeDocbuilderSource(baseSource,variables,tag,literals){
  const base=applyDocbuilderLiterals(String(baseSource||''),literals);
  const lines=[];
  const vars=isObj(variables)?variables:{};
  Object.keys(vars).forEach(key=>{
    if(vars[key]===undefined)return;
    lines.push(`var __REPORTS_${key} = ${toAsciiJsLiteral(vars[key])};`);
  });
  const wrapperDir=joinPath(getReportsUiDir(),GENERATED_DIR,'_runtime');
  const wrapperPath=joinPath(wrapperDir,`reports_${String(tag||'run')}_${Date.now()}_${Math.random().toString(36).slice(2,8)}.docbuilder`);
  await writeTextFile(wrapperPath,`${lines.join('\n')}\n${base}\n`,7000);
  return wrapperPath;
}
function makeProbeResultPath(){
  const resultDir=joinPath(getReportsUiDir(),GENERATED_DIR,'_runtime');
  return joinPath(resultDir,`reports_probe_result_${Date.now()}_${Math.random().toString(36).slice(2,8)}.json`);
}
function makeTempOutputPath(){
  const resultDir=joinPath(getReportsUiDir(),GENERATED_DIR,'_runtime');
  return joinPath(resultDir,`reports_output_${Date.now()}_${Math.random().toString(36).slice(2,8)}.xlsx`);
}
async function readProbeResultFile(resultPath,timeoutMs){
  const res=await readTextFile(resultPath,timeoutMs||5000);
  if(!res||!res.ok)throw fileErr(res||{error:'probe_result_read_failed',path:resultPath});
  const raw=stripBom(res.content||'').trim();
  if(!raw)throw new Error('Probe result file is empty.');
  const parsed=parse(raw);
  if(!isObj(parsed))throw new Error('Probe result file has invalid JSON.');
  return parsed;
}
async function waitProbeResultFile(resultPath,timeoutMs,pollMs){
  const started=Date.now();
  const limit=Math.max(1000,toInt(timeoutMs,15000)||15000);
  const step=Math.max(100,toInt(pollMs,250)||250);
  while(Date.now()-started<limit){
    try{
      return await readProbeResultFile(resultPath,600);
    }catch(err){
      const details=err&&err.details?err.details:null;
      const code=details&&details.error?String(details.error):'';
      const text=String(err&&err.message||'').toLowerCase();
      const transient=code==='not_found'||code==='read_failed'||text.indexOf('empty')>=0||text.indexOf('invalid json')>=0;
      if(!transient)throw err;
    }
    await new Promise(resolve=>setTimeout(resolve,step));
  }
  throw new Error(`Probe result timeout: ${resultPath}`);
}
function sanitizeTemplateName(name){
  let out=String(name||'').replace(/[<>:"/\\|?*\u0000-\u001F]+/g,' ').replace(/\s+/g,' ').trim();
  out=out.replace(/\.(xlsx|xlsm|xls)$/i,'');
  out=out.replace(/[. ]+$/g,'');
  if(!out)out='Сценарий';
  if(/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i.test(out))out=`${out}_`;
  return out;
}
function sanitizeScenarioFileBase(name){
  let out=String(name||'').replace(/[<>:"/\\|?*\u0000-\u001F]+/g,' ').replace(/\s+/g,' ').trim();
  out=out.replace(/\.json$/i,'');
  out=out.replace(/[. ]+$/g,'');
  if(!out)out='Сценарий';
  if(/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i.test(out))out=`${out}_`;
  return out;
}
function makeScenarioFileName(name){return `${sanitizeScenarioFileBase(name)}.json`;}
function ensureScenarioFileNameUniq(fileName,used){
  let candidate=String(fileName||'').trim();
  if(!candidate)candidate='Сценарий.json';
  const base=sanitizeScenarioFileBase(candidate);
  let next=`${base}.json`;
  let i=2;
  while(used[next.toLowerCase()]){
    next=`${base} (${i}).json`;
    i+=1;
  }
  used[next.toLowerCase()]=true;
  return next;
}
function buildScenarioFileUsedMap(excludeScenarioId){
  const used=Object.create(null);
  state.scenarios.forEach(sc=>{
    if(excludeScenarioId&&sc.id===excludeScenarioId)return;
    const file=String(state.scenarioFiles[sc.id]||'').trim();
    if(file)used[file.toLowerCase()]=true;
  });
  return used;
}
function makeScenarioFileNameForScenario(scenarioId,scenarioName){
  return ensureScenarioFileNameUniq(makeScenarioFileName(scenarioName),buildScenarioFileUsedMap(scenarioId));
}
function safeSourceKey(value,fallback){
  let out=String(value||'').toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]+/g,'_').replace(/^_+|_+$/g,'');
  if(!out)out=String(fallback||'item');
  return out;
}
function mkSourceFileConfig(){
  return {
    enabled:false,
    label:'Файл с исходными данными',
    required:true,
    accept:['xlsx','xls'],
    maxFileSizeMb:25,
    maxRowsDefault:200,
    probeTimeoutMs:45000,
    previewRows:10,
    strictMode:true
  };
}
function normSourceFileConfig(raw){
  const base=mkSourceFileConfig();
  const src=isObj(raw)?raw:{};
  const out=Object.assign({},base,src);
  out.enabled=!!src.enabled;
  out.label=String(src.label||base.label).trim()||base.label;
  out.required=src.required===undefined?true:!!src.required;
  out.maxFileSizeMb=Math.max(1,parseInt(src.maxFileSizeMb,10)||base.maxFileSizeMb);
  out.maxRowsDefault=Math.max(1,parseInt(src.maxRowsDefault,10)||base.maxRowsDefault);
  out.probeTimeoutMs=Math.max(5000,parseInt(src.probeTimeoutMs,10)||base.probeTimeoutMs);
  out.previewRows=Math.max(1,parseInt(src.previewRows,10)||base.previewRows);
  out.strictMode=src.strictMode===undefined?true:!!src.strictMode;
  out.accept=Array.isArray(src.accept)?src.accept.map(x=>String(x||'').trim().toLowerCase()).filter(Boolean):base.accept.slice();
  if(!out.accept.length)out.accept=base.accept.slice();
  return out;
}
function mkSourceField(){
  return {
    id:uid('source-field'),
    key:'field',
    label:'Новое поле',
    sheet:'Лист1',
    address:'A1',
    type:'text',
    required:true
  };
}
function normSourceField(raw,index){
  const base=mkSourceField();
  const src=isObj(raw)?raw:{};
  const out=Object.assign({},base,src);
  out.id=String(src.id||base.id);
  out.label=String(src.label||base.label).trim()||`Поле ${index+1}`;
  out.key=safeSourceKey(src.key||out.label,`field_${index+1}`);
  out.sheet=String(src.sheet||base.sheet).trim()||base.sheet;
  out.address=String(src.address||base.address).trim()||base.address;
  out.type=String(src.type||base.type).trim().toLowerCase();
  if(['text','number','date'].indexOf(out.type)<0)out.type='text';
  out.required=src.required===undefined?true:!!src.required;
  return out;
}
function mkSourceList(){
  return {
    id:uid('source-list'),
    key:'list',
    label:'Новый список',
    sheet:'Лист1',
    startAddress:'A1',
    direction:'down',
    stopMode:'empty',
    stopValue:'СТОП',
    maxItems:'200',
    type:'text',
    required:true
  };
}
function normSourceList(raw,index){
  const base=mkSourceList();
  const src=isObj(raw)?raw:{};
  const out=Object.assign({},base,src);
  out.id=String(src.id||base.id);
  out.label=String(src.label||base.label).trim()||`Список ${index+1}`;
  out.key=safeSourceKey(src.key||out.label,`list_${index+1}`);
  out.sheet=String(src.sheet||base.sheet).trim()||base.sheet;
  out.startAddress=String(src.startAddress||src.address||base.startAddress).trim()||base.startAddress;
  out.direction=String(src.direction||base.direction).trim().toLowerCase()==='right'?'right':'down';
  out.stopMode=String(src.stopMode||base.stopMode).trim().toLowerCase()==='stop_value'?'stop_value':'empty';
  out.stopValue=String(src.stopValue==null?base.stopValue:src.stopValue);
  out.maxItems=String(src.maxItems||base.maxItems).trim()||base.maxItems;
  out.type=String(src.type||base.type).trim().toLowerCase();
  if(['text','number','date'].indexOf(out.type)<0)out.type='text';
  out.required=src.required===undefined?true:!!src.required;
  return out;
}
function mkSourceColumn(){
  return {
    id:uid('source-column'),
    key:'column',
    label:'Новая колонка',
    header:'Колонка',
    type:'text',
    required:true
  };
}
function normSourceColumn(raw,index){
  const base=mkSourceColumn();
  const src=isObj(raw)?raw:{};
  const out=Object.assign({},base,src);
  out.id=String(src.id||base.id);
  out.label=String(src.label||base.label).trim()||`Колонка ${index+1}`;
  out.key=safeSourceKey(src.key||out.label,`column_${index+1}`);
  out.header=String(src.header||out.label).trim()||out.label;
  out.type=String(src.type||base.type).trim().toLowerCase();
  if(['text','number','date'].indexOf(out.type)<0)out.type='text';
  out.required=src.required===undefined?true:!!src.required;
  return out;
}
function mkSourceTable(){
  return {
    id:uid('source-table'),
    key:'table',
    label:'Новая таблица',
    sheet:'Лист1',
    headerRow:'1',
    startRowOffset:'1',
    keyHeader:'',
    emptyRowTolerance:'1',
    maxRows:'200',
    columns:[mkSourceColumn()]
  };
}
function normSourceTable(raw,index){
  const base=mkSourceTable();
  const src=isObj(raw)?raw:{};
  const out=Object.assign({},base,src);
  out.id=String(src.id||base.id);
  out.label=String(src.label||base.label).trim()||`Таблица ${index+1}`;
  out.key=safeSourceKey(src.key||out.label,`table_${index+1}`);
  out.sheet=String(src.sheet||base.sheet).trim()||base.sheet;
  out.headerRow=String(src.headerRow||base.headerRow).trim()||base.headerRow;
  out.startRowOffset=String(src.startRowOffset||base.startRowOffset).trim()||base.startRowOffset;
  out.keyHeader=String(src.keyHeader||'').trim();
  out.emptyRowTolerance=String(src.emptyRowTolerance||base.emptyRowTolerance).trim()||base.emptyRowTolerance;
  out.maxRows=String(src.maxRows||base.maxRows).trim()||base.maxRows;
  out.columns=(Array.isArray(src.columns)?src.columns:[mkSourceColumn()]).map((item,colIndex)=>normSourceColumn(item,colIndex));
  if(!out.columns.length)out.columns=[mkSourceColumn()];
  return out;
}
function mkSourceSchema(){
  return {
    fields:[],
    lists:[],
    tables:[]
  };
}
function normSourceSchema(raw){
  const src=isObj(raw)?raw:{};
  return {
    fields:(Array.isArray(src.fields)?src.fields:[]).map((item,index)=>normSourceField(item,index)),
    lists:(Array.isArray(src.lists)?src.lists:[]).map((item,index)=>normSourceList(item,index)),
    tables:(Array.isArray(src.tables)?src.tables:[]).map((item,index)=>normSourceTable(item,index))
  };
}
function sourceConfigEnabled(scenario){
  return !!(scenario&&isObj(scenario.sourceFileConfig)&&scenario.sourceFileConfig.enabled);
}
function getScenarioSourceConfig(scenario){
  return normSourceFileConfig(scenario&&scenario.sourceFileConfig);
}
function getScenarioSourceSchema(scenario){
  return normSourceSchema(scenario&&scenario.sourceSchema);
}
function scenarioUsesSourceActions(scenario){
  const actions=Array.isArray(scenario&&scenario.actions)?scenario.actions:[];
  for(let i=0;i<actions.length;i+=1){
    const type=String(actions[i]&&actions[i].type||'').trim();
    if(type==='insert_source_field'||type==='insert_source_list'||type==='insert_source_table')return true;
  }
  return false;
}
function getDraftSourceFieldOptions(){
  const schemaState=getScenarioSourceSchema(state.draft);
  const out=schemaState.fields.map(field=>[field.key,field.label]);
  return out.length?out:[['','Сначала добавьте поле файла']];
}
function getDraftSourceListOptions(){
  const schemaState=getScenarioSourceSchema(state.draft);
  const out=schemaState.lists.map(list=>[list.key,list.label]);
  return out.length?out:[['','Сначала добавьте список файла']];
}
function getDraftSourceTableOptions(){
  const schemaState=getScenarioSourceSchema(state.draft);
  const out=schemaState.tables.map(table=>[table.key,table.label]);
  return out.length?out:[['','Сначала добавьте таблицу файла']];
}
function getDynamicActionOptions(kind){
  if(kind==='source_fields')return getDraftSourceFieldOptions();
  if(kind==='source_lists')return getDraftSourceListOptions();
  if(kind==='source_tables')return getDraftSourceTableOptions();
  return [['','(нет вариантов)']];
}
function parseActionSourceMappings(raw){
  const text=String(raw||'').trim();
  if(!text)return {items:[],error:'Укажите соответствие колонок в формате key=A.'};
  const items=[];
  text.split(/\r?\n+/).forEach(line=>{
    const row=String(line||'').trim();
    if(!row)return;
    const parts=row.split('=');
    if(parts.length<2)return;
    const sourceKey=safeSourceKey(parts.shift(), '');
    const targetColumn=String(parts.join('=')||'').trim().toUpperCase();
    if(sourceKey&&/^[A-Z]+$/.test(targetColumn))items.push({sourceKey,targetColumn});
  });
  if(!items.length)return {items:[],error:'Не удалось разобрать соответствие колонок.'};
  return {items,error:''};
}
function fmtStamp(){
  const d=new Date();
  const dd=String(d.getDate()).padStart(2,'0');
  const mm=String(d.getMonth()+1).padStart(2,'0');
  const yy=String(d.getFullYear()).slice(-2);
  const hh=String(d.getHours()).padStart(2,'0');
  const mi=String(d.getMinutes()).padStart(2,'0');
  return `${dd}-${mm}-${yy}_${hh}-${mi}`;
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
function mkOutputPath(sc,probe,tpl){
  const outDir=getGeneratedDir(probe,tpl);
  const baseName=sanitizeTemplateName(sc&&sc.name?sc.name:'Сценарий').replace(/\s+/g,' ');
  return joinPath(outDir,`${baseName}_${fmtStamp()}.xlsx`);
}
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
function toInt(value,fallback){
  const n=parseInt(value,10);
  return Number.isNaN(n)?fallback:n;
}
function clampInt(value,min,max,fallback){
  let n=toInt(value,fallback);
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
  const text=String(raw||'');
  let parts=[];
  if(Array.isArray(raw)){
    parts=raw;
  }else if(/\r?\n/.test(text)){
    parts=text.split(/\r?\n/);
  }else if(text.indexOf('\t')>=0){
    parts=text.split('\t');
  }else if(text.indexOf(';')>=0){
    parts=text.split(';');
  }else if(text.indexOf(',')>=0){
    const csvParts=text.split(',');
    // Legacy support: split by comma only for simple one-word tokens like "A,B,C".
    const looksLikeSimpleCsv=csvParts.length>1&&csvParts.every(item=>{
      const v=String(item||'').trim();
      return v!==''&&v.indexOf(' ')<0;
    });
    parts=looksLikeSimpleCsv?csvParts:[text];
  }else{
    parts=[text];
  }
  parts.forEach(part=>{
    const v=String(part==null?'':part).trim();
    if(!v)return;
    const key=v.toLowerCase();
    if(seen[key])return;
    seen[key]=true;
    out.push(v);
  });
  return out;
}
function stringifyInputOptions(values){
  const out=[];
  const seen=Object.create(null);
  (Array.isArray(values)?values:[]).forEach(item=>{
    const v=String(item==null?'':item).trim();
    if(!v)return;
    const key=v.toLowerCase();
    if(seen[key])return;
    seen[key]=true;
    out.push(v);
  });
  return out.join('; ');
}
function optionsFromLines(raw){
  const out=[];
  const seen=Object.create(null);
  String(raw||'').split(/\r?\n/).forEach(line=>{
    const v=String(line||'').trim();
    if(!v)return;
    const key=v.toLowerCase();
    if(seen[key])return;
    seen[key]=true;
    out.push(v);
  });
  return out;
}
function optionsToLines(values){
  const list=Array.isArray(values)?values:parseInputOptions(values);
  return list.map(item=>String(item==null?'':item).trim()).filter(Boolean).join('\n');
}
function optionsPreview(values,max){
  const list=Array.isArray(values)?values:parseInputOptions(values);
  const clean=list.map(item=>String(item==null?'':item).trim()).filter(Boolean);
  if(!clean.length)return 'Список пуст';
  const limit=Math.max(1,parseInt(max,10)||3);
  if(clean.length<=limit)return clean.join(', ');
  return `${clean.slice(0,limit).join(', ')} ... (+${clean.length-limit})`;
}
function parseDependentRules(raw){
  const rules=[];
  const lines=String(raw||'').split(/\r?\n/);
  for(let i=0;i<lines.length;i+=1){
    const line=String(lines[i]||'').trim();
    if(!line)continue;
    const posColon=line.indexOf(':');
    const posEq=line.indexOf('=');
    let sep=-1;
    if(posColon>0&&posEq>0)sep=Math.min(posColon,posEq);
    else if(posColon>0)sep=posColon;
    else if(posEq>0)sep=posEq;
    if(sep<=0){
      return {rules,error:`строка ${i+1}: используйте формат "Значение родителя: Вариант 1; Вариант 2".`};
    }
    const key=line.slice(0,sep).trim();
    const valuesRaw=line.slice(sep+1).trim();
    if(!key)return {rules,error:`строка ${i+1}: не указано значение родителя.`};
    const options=parseInputOptions(valuesRaw);
    if(!options.length)return {rules,error:`строка ${i+1}: не указаны варианты зависимого списка.`};
    rules.push({key,options});
  }
  return {rules,error:''};
}
function stringifyDependentRules(rules){
  const lines=[];
  const seen=Object.create(null);
  (Array.isArray(rules)?rules:[]).forEach(rule=>{
    const key=String(rule&&rule.key!=null?rule.key:'').trim();
    if(!key)return;
    const options=stringifyInputOptions(rule&&rule.options);
    if(!options)return;
    const uniqKey=key.toLowerCase();
    if(seen[uniqKey])return;
    seen[uniqKey]=true;
    lines.push(`${key}: ${options}`);
  });
  return lines.join('\n');
}
function parseDependentOptionsMap(raw){
  const map=Object.create(null);
  const mapLower=Object.create(null);
  const parsed=parseDependentRules(raw);
  let valuesCount=0;
  if(parsed.error)return {map,mapLower,valuesCount,error:parsed.error};
  for(let i=0;i<parsed.rules.length;i+=1){
    const rule=parsed.rules[i];
    const key=String(rule.key||'').trim();
    const options=Array.isArray(rule.options)?rule.options:[];
    map[key]=options;
    const low=key.toLowerCase();
    if(!Object.prototype.hasOwnProperty.call(mapLower,low))mapLower[low]=options;
    valuesCount+=options.length;
  }
  return {map,mapLower,valuesCount,error:''};
}
function resolveSelectOptions(field,parentValue){
  const base=parseInputOptions(field&&field.options);
  if(!field||normalizeInputType(field.input_type)!=='select')return base;
  const parentId=String(field.depends_on||'').trim();
  if(!parentId)return base;
  const allowFallback=field.fallback_on_missing!==false;
  const parsed=parseDependentOptionsMap(field.options_map);
  if(parsed.error)return allowFallback?base:[];
  const key=String(parentValue==null?'':parentValue).trim();
  if(!key)return [];
  if(Object.prototype.hasOwnProperty.call(parsed.map,key))return parsed.map[key];
  const low=key.toLowerCase();
  if(Object.prototype.hasOwnProperty.call(parsed.mapLower,low))return parsed.mapLower[low];
  return allowFallback?base:[];
}
function findInputDependencyCycle(fields){
  const links=Object.create(null);
  const normalized=Array.isArray(fields)?fields.map(normInputField):[];
  normalized.forEach(field=>{
    if(field.input_type!=='select')return;
    const parent=String(field.depends_on||'').trim();
    if(!parent||parent===field.id)return;
    links[field.id]=parent;
  });
  const visited=Object.create(null);
  const inStack=Object.create(null);
  const visit=(id)=>{
    if(inStack[id])return id;
    if(visited[id])return '';
    visited[id]=true;
    inStack[id]=true;
    const parent=links[id];
    if(parent&&links[parent]){
      const found=visit(parent);
      if(found)return found;
    }
    delete inStack[id];
    return '';
  };
  const ids=Object.keys(links);
  for(let i=0;i<ids.length;i+=1){
    const found=visit(ids[i]);
    if(found)return found;
  }
  return '';
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
    depends_on:'',
    options_map:'',
    fallback_on_missing:true,
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
  out.depends_on=String(src.depends_on||src.dependsOn||'').trim();
  out.options_map=String(src.options_map||src.optionsMap||'').trim();
  if(src.fallback_on_missing===undefined&&src.fallbackOnMissing===undefined){
    out.fallback_on_missing=out.depends_on?parseInputOptions(out.options).length>0:true;
  }else{
    out.fallback_on_missing=src.fallback_on_missing===undefined?!!src.fallbackOnMissing:!!src.fallback_on_missing;
  }
  if(out.depends_on===out.id)out.depends_on='';
  if(!out.depends_on)out.fallback_on_missing=true;
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

function parseTableTargets(raw){
  const text=String(raw||'').trim();
  if(!text)return {items:[],error:'Укажите точки вставки в формате "Лист:A15|A15:N15".'};
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
    if(sep<=0)return {items:[],error:`Блок ${i+1}: используйте формат "Лист:A15|A15:N15".`};
    const sheet=block.slice(0,sep).trim();
    const payload=block.slice(sep+1).trim();
    const pipe=payload.indexOf('|');
    if(pipe<=0)return {items:[],error:`Блок ${i+1}: разделите первую ячейку и строку-шаблон через "|".`};
    const anchorCell=payload.slice(0,pipe).trim();
    const templateRowRange=payload.slice(pipe+1).trim();
    if(!sheet)return {items:[],error:`Блок ${i+1}: не указан лист.`};
    if(!anchorCell)return {items:[],error:`Блок ${i+1}: не указана первая ячейка.`};
    if(!templateRowRange)return {items:[],error:`Блок ${i+1}: не указана строка-шаблон.`};
    items.push({sheet,anchor_cell:anchorCell,template_row_range:templateRowRange});
  }
  if(!items.length)return {items:[],error:'Не удалось разобрать точки вставки таблицы.'};
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
  if(t==='insert_source_field'){
    const hasTargets=!(a.targets==null||String(a.targets).trim()==='');
    if(!hasTargets){
      const sheet=String(a.sheet||'Лист1').trim()||'Лист1';
      const range=String(a.range||'A1').trim()||'A1';
      o.targets=`${sheet}:${range}`;
    }
  }
  if(t==='insert_source_list'){
    const hasTargets=!(a.targets==null||String(a.targets).trim()==='');
    if(!hasTargets){
      const sheet=String(a.sheet||'Лист1').trim()||'Лист1';
      const anchorCell=String(a.anchor_cell||'A1').trim()||'A1';
      const templateRowRange=String(a.template_row_range||'A1:D1').trim()||'A1:D1';
      o.targets=`${sheet}:${anchorCell}|${templateRowRange}`;
    }
  }
  if(t==='insert_source_table'){
    const hasTargets=!(a.targets==null||String(a.targets).trim()==='');
    if(!hasTargets){
      const sheet=String(a.sheet||'Лист1').trim()||'Лист1';
      const anchorCell=String(a.anchor_cell||'A1').trim()||'A1';
      const templateRowRange=String(a.template_row_range||'A1:D1').trim()||'A1:D1';
      o.targets=`${sheet}:${anchorCell}|${templateRowRange}`;
    }
  }
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
  const sourceFileConfig=normSourceFileConfig(s.sourceFileConfig||s.source_file_config);
  const sourceSchema=normSourceSchema(s.sourceSchema||s.source_schema);
  return{
    id:s.id||uid('scenario'),
    name:String(s.name||'Без названия'),
    description:String(s.description||s.descr||''),
    templatePath:String(s.templatePath||s.path||''),
    actions:Array.isArray(s.actions)?s.actions.map(normAction):[],
    inputFields:rawInputs.map(normInputField),
    sourceFileConfig,
    sourceSchema,
    createdAt:s.createdAt||new Date().toISOString(),
    updatedAt:s.updatedAt||s.updated||new Date().toISOString(),
    lastRunAt:s.lastRunAt||s.last_run||null
  };
}
function normalizeScenarios(raw){return Array.isArray(raw)?raw.map(normScenario).filter(Boolean):[];}
function saveLocalStorageOnly(){try{localStorage.setItem(STORAGE,JSON.stringify(state.scenarios));}catch(_){toast('Не удалось сохранить сценарии.','error');}}
function loadLocalStorageOnly(){
  const raw=parse(localStorage.getItem(STORAGE));
  state.scenarios=normalizeScenarios(raw);
  state.scenarioFiles=Object.create(null);
}
function clearFilePending(id){
  const p=FILE_PENDING[id];
  if(!p)return null;
  if(p.t)clearTimeout(p.t);
  delete FILE_PENDING[id];
  return p;
}
function addFilePending(id,res,rej,ms){
  FILE_PENDING[id]={
    res,rej,
    t:setTimeout(()=>{
      const p=clearFilePending(id);
      if(!p)return;
      if(state.fileBridge==='unknown')state.fileBridge='missing';
      const e=new Error('Таймаут доступа к файлу сценариев');
      e.details={requestId:id,error:'timeout'};
      p.rej(e);
    },ms)
  };
}
function fileErr(raw){
  const e=new Error((raw&&raw.error)||'Ошибка работы с файлом сценариев');
  e.details=raw||{};
  return e;
}
function onNativeFileResult(raw){
  const data=parseNativeParam(raw);
  if(!isObj(data))return;
  state.fileBridge='ready';
  const id=data.requestId?String(data.requestId):'';
  if(!id)return;
  const p=clearFilePending(id);
  if(!p)return;
  p.res(data);
}
function canUseNativeFileBridge(){
  if(state.fileBridge==='missing')return false;
  const sdk=window.parent&&window.parent.sdk;
  if(!sdk||typeof sdk.command!=='function')return false;
  return bindDirectNative();
}
function requestNativeFile(cmd,payload,ms){
  return new Promise((resolve,reject)=>{
    if(!canUseNativeFileBridge()){
      const e=new Error('Native file bridge unavailable');
      e.details={error:'bridge_unavailable'};
      reject(e);
      return;
    }
    const sdk=window.parent&&window.parent.sdk;
    const body=isObj(payload)?Object.assign({},payload):{};
    const id=body.requestId||reqId('reports-file');
    body.requestId=id;
    addFilePending(id,resolve,reject,ms);
    try{
      sdk.command(cmd,JSON.stringify(body));
    }catch(err){
      const p=clearFilePending(id);
      if(!p)return;
      const e=new Error('Не удалось отправить команду сохранения сценариев');
      e.details={requestId:id,error:String(err||'command_failed')};
      p.rej(e);
    }
  });
}
async function readScenariosFile(){
  const r=await requestNativeFile('reports:fileRead',{path:getLegacyScenariosPath()},3000);
  return isObj(r)?r:{ok:false,error:'invalid_read_response'};
}
async function readTextFile(path,timeoutMs){
  const r=await requestNativeFile('reports:fileRead',{path},timeoutMs||3000);
  return isObj(r)?r:{ok:false,error:'invalid_read_response'};
}
async function readBinaryFileBase64(path,timeoutMs){
  const r=await requestNativeFile('reports:fileRead',{path,encoding:'base64'},timeoutMs||12000);
  if(!isObj(r)||!r.ok)throw fileErr(r||{error:'invalid_read_response',path});
  return String(r.content||'');
}
async function canReadBinaryFile(path,timeoutMs){
  try{
    await readBinaryFileBase64(path,timeoutMs||2500);
    return true;
  }catch(_){
    return false;
  }
}
async function writeTextFile(path,content,timeoutMs){
  const r=await requestNativeFile('reports:fileWrite',{path,content:String(content||'')},timeoutMs||5000);
  if(!isObj(r)||!r.ok)throw fileErr(r||{error:'write_failed'});
  return r;
}
async function writeBinaryFileBase64(path,contentBase64,timeoutMs){
  const r=await requestNativeFile('reports:fileWrite',{path,content:String(contentBase64||''),encoding:'base64'},timeoutMs||15000);
  if(!isObj(r)||!r.ok)throw fileErr(r||{error:'write_failed'});
  return r;
}
async function movePath(fromPath,toPath,timeoutMs){
  const from=String(fromPath||'').trim();
  const to=String(toPath||'').trim();
  if(!from||!to)return {ok:false,error:'invalid_path'};
  if(samePath(from,to))return {ok:true};
  const r=await requestNativeFile('reports:fileMove',{fromPath:from,toPath:to},timeoutMs||6000);
  if(!isObj(r))throw fileErr({error:'invalid_move_response'});
  if(!r.ok&&String(r.error||'')!=='not_found')throw fileErr(r);
  return r;
}
async function deletePath(path,timeoutMs){
  const p=String(path||'').trim();
  if(!p)return {ok:false,error:'invalid_path'};
  const r=await requestNativeFile('reports:fileDelete',{path:p},timeoutMs||5000);
  if(!isObj(r))throw fileErr({error:'invalid_delete_response'});
  if(!r.ok&&String(r.error||'')!=='not_found')throw fileErr(r);
  return r;
}
async function relocateGeneratedFile(fromPath,toPath){
  const from=String(fromPath||'').trim();
  const to=String(toPath||'').trim();
  if(!from||!to)throw fileErr({error:'invalid_path',fromPath:from,toPath:to});
  if(samePath(from,to))return true;
  try{
    const moved=await movePath(from,to,2500);
    if(moved&&moved.ok){
      const exists=await ensureOutputExists(to,2500);
      if(exists)return true;
    }
  }catch(_){ }
  const base64=await readBinaryFileBase64(from,20000);
  await writeBinaryFileBase64(to,base64,20000);
  const exists=await ensureOutputExists(to,4000);
  if(!exists)throw fileErr({error:'write_failed',path:to});
  await deletePath(from,5000).catch(()=>{});
  return true;
}
function warnFileOpsNeedRebuild(){
  toast('Файлы не удалось переименовать/удалить. Обновите DesktopEditors (reports:fileMove/reports:fileDelete).','error');
}
function parseScenarioIndex(rawText){
  const raw=parse(rawText);
  const src=Array.isArray(raw)?raw:(isObj(raw)&&Array.isArray(raw.items)?raw.items:[]);
  const out=[];
  src.forEach(item=>{
    if(typeof item==='string'){
      const file=String(item||'').trim();
      if(file)out.push({id:'',file});
      return;
    }
    if(!isObj(item))return;
    const file=String(item.file||item.path||'').trim();
    if(!file)return;
    out.push({id:String(item.id||'').trim(),file});
  });
  return out;
}
function buildScenarioIndex(){
  return {
    version:1,
    items:state.scenarios.map(sc=>({
      id:sc.id,
      file:String(state.scenarioFiles[sc.id]||'').trim(),
      name:String(sc.name||''),
      updatedAt:sc.updatedAt||''
    })).filter(item=>item.file)
  };
}
function ensureScenarioFilesMap(){
  const used=Object.create(null);
  const next=Object.create(null);
  state.scenarios.forEach(sc=>{
    const prev=String(state.scenarioFiles[sc.id]||'').trim();
    const candidate=prev||makeScenarioFileName(sc.name);
    next[sc.id]=ensureScenarioFileNameUniq(candidate,used);
  });
  state.scenarioFiles=next;
}
async function persistScenariosToFiles(){
  ensureScenarioFilesMap();
  for(let i=0;i<state.scenarios.length;i+=1){
    const sc=normScenario(state.scenarios[i]);
    state.scenarios[i]=sc;
    const fileName=state.scenarioFiles[sc.id];
    if(!fileName)continue;
    const path=getScenarioPath(fileName);
    const body=`${JSON.stringify(sc,null,2)}\n`;
    await writeTextFile(path,body,6000);
  }
  const indexBody=`${JSON.stringify(buildScenarioIndex(),null,2)}\n`;
  await writeTextFile(getScenariosIndexPath(),indexBody,6000);
}
function queuePersistScenarios(){
  if(state.fileBridge!=='ready')return Promise.resolve();
  state.persistScenariosPromise=state.persistScenariosPromise
    .catch(()=>{})
    .then(()=>persistScenariosToFiles());
  return state.persistScenariosPromise;
}
async function loadScenariosFromScenarioFiles(){
  const indexPath=getScenariosIndexPath();
  const indexRes=await readTextFile(indexPath,3000);
  if(!indexRes.ok){
    if(indexRes.error==='not_found'){
      const legacyRes=await readScenariosFile();
      if(legacyRes.ok){
        state.scenarios=normalizeScenarios(parse(legacyRes.content||'[]'));
        state.scenarioFiles=Object.create(null);
        ensureScenarioFilesMap();
        await persistScenariosToFiles().catch(()=>{});
        return true;
      }
      if(legacyRes.error==='not_found'){
        state.scenarios=[];
        state.scenarioFiles=Object.create(null);
        await writeTextFile(indexPath,'{\n  "version": 1,\n  "items": []\n}\n',4000).catch(()=>{});
        return true;
      }
    }
    return false;
  }

  const refs=parseScenarioIndex(indexRes.content||'');
  if(!refs.length){
    state.scenarios=[];
    state.scenarioFiles=Object.create(null);
    return true;
  }

  const list=[];
  const filesById=Object.create(null);
  const usedIds=Object.create(null);
  for(let i=0;i<refs.length;i+=1){
    const ref=refs[i];
    const scenarioPath=getScenarioPath(ref.file);
    const scenarioRes=await readTextFile(scenarioPath,3000).catch(()=>null);
    if(!scenarioRes||!scenarioRes.ok)continue;
    const sc=normScenario(parse(scenarioRes.content||''));
    if(!sc)continue;
    if(usedIds[sc.id])sc.id=uid('scenario');
    usedIds[sc.id]=true;
    list.push(sc);
    filesById[sc.id]=ref.file;
  }
  state.scenarios=list;
  state.scenarioFiles=filesById;
  return true;
}
function save(){
  saveLocalStorageOnly();
  if(state.fileBridge!=='ready')return Promise.resolve(true);
  return queuePersistScenarios()
    .then(()=>{state.fileSaveWarned=false;return true;})
    .catch(()=>{
      if(!state.fileSaveWarned){
        state.fileSaveWarned=true;
        toast('Не удалось сохранить сценарии в reports-ui/data/scenarios/.','error');
      }
      return false;
    });
}
async function load(){
  try{localStorage.removeItem('reports.scenarios.v2');}catch(_){ }
  if(canUseNativeFileBridge()){
    try{
      const loaded=await loadScenariosFromScenarioFiles();
      if(loaded){
        saveLocalStorageOnly();
        return;
      }
    }catch(_){
      state.fileBridge='missing';
    }
  }
  loadLocalStorageOnly();
}

function getItemCollapseMap(kind){
  return kind==='action'?state.collapsedActionItems:state.collapsedInputItems;
}
function isItemCollapsed(kind,id){
  const key=String(id||'').trim();
  if(!key)return false;
  return !!getItemCollapseMap(kind)[key];
}
function setItemCollapsed(kind,id,collapsed){
  const key=String(id||'').trim();
  if(!key)return;
  const map=getItemCollapseMap(kind);
  if(collapsed)map[key]=true;
  else delete map[key];
}
function toggleItemCollapsed(kind,id){
  setItemCollapsed(kind,id,!isItemCollapsed(kind,id));
}
function resetItemCollapsedState(){
  state.collapsedInputItems=Object.create(null);
  state.collapsedActionItems=Object.create(null);
}

function openModal(id){
  const src=state.scenarios.find(x=>x.id===id);
  state.draft=src?clone(src):{id:uid('scenario'),name:'',description:'',templatePath:'',actions:[],inputFields:[],sourceFileConfig:mkSourceFileConfig(),sourceSchema:mkSourceSchema(),createdAt:new Date().toISOString(),updatedAt:new Date().toISOString(),lastRunAt:null};
  resetItemCollapsedState();
  if(!Array.isArray(state.draft.actions))state.draft.actions=[];
  state.draft.actions=state.draft.actions.map(raw=>{
    const action=normAction(raw);
    if(!action.id)action.id=uid('action');
    return action;
  });
  if(!Array.isArray(state.draft.inputFields))state.draft.inputFields=[];
  state.draft.inputFields=state.draft.inputFields.map(normInputField);
  state.draft.sourceFileConfig=normSourceFileConfig(state.draft.sourceFileConfig);
  state.draft.sourceSchema=normSourceSchema(state.draft.sourceSchema);
  if(src){
    // При редактировании существующего сценария: секции открыты, элементы свернуты.
    state.inputCollapsed=false;
    state.actionCollapsed=false;
    state.draft.inputFields.forEach(field=>setItemCollapsed('input',field.id,true));
    state.draft.actions.forEach(action=>setItemCollapsed('action',action.id,true));
  }
  els.title.textContent=src?'Изменение сценария':'Новый сценарий';
  els.name.value=state.draft.name; els.tpl.value=state.draft.templatePath; els.descr.value=state.draft.description;
  renderInputFields(); renderSourceConfig(); renderActions(); syncSectionCollapseUi(); els.modal.classList.remove('hidden'); els.name.focus();
}
function closeModal(){els.modal.classList.add('hidden');state.draft=null;resetItemCollapsedState();}

function setSectionCollapsed(kind,collapsed){
  const isInput=kind==='input';
  if(isInput)state.inputCollapsed=!!collapsed;
  else state.actionCollapsed=!!collapsed;
  const body=isInput?els.inputSectionBody:els.actionSectionBody;
  const btn=isInput?els.btnToggleInputSection:els.btnToggleActionSection;
  if(body)body.classList.toggle('hidden',!!collapsed);
  if(btn){
    btn.textContent=collapsed?'Развернуть':'Свернуть';
    btn.setAttribute('aria-expanded',collapsed?'false':'true');
  }
}
function toggleSection(kind){
  if(kind==='input')setSectionCollapsed('input',!state.inputCollapsed);
  else setSectionCollapsed('action',!state.actionCollapsed);
}
function syncSectionCollapseUi(){
  setSectionCollapsed('input',state.inputCollapsed);
  setSectionCollapsed('action',state.actionCollapsed);
}

async function removeScenario(id){
  const idx=state.scenarios.findIndex(x=>x.id===id);
  if(idx<0)return;
  const scenario=state.scenarios[idx];
  const scenarioFile=String(state.scenarioFiles[scenario.id]||'').trim();
  const scenarioPath=scenarioFile?getScenarioPath(scenarioFile):'';
  const templatePath=isManagedScenarioTemplatePath(scenario.templatePath,scenario.name)?resolveTemplate(scenario.templatePath):'';
  state.scenarios.splice(idx,1);
  delete state.scenarioFiles[scenario.id];
  renderList();
  toast('Сценарий удален.','ok');
  save().catch(()=>{});
  if(state.fileBridge==='ready'){
    Promise.all([
      scenarioPath?deletePath(scenarioPath,2500):Promise.resolve({ok:true}),
      templatePath?deletePath(templatePath,2500):Promise.resolve({ok:true})
    ]).catch(()=>{warnFileOpsNeedRebuild();});
  }
}

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
    const bDel=document.createElement('button');bDel.className='btn btn-danger';bDel.type='button';bDel.textContent='Удалить';bDel.onclick=()=>{removeScenario(s.id).catch(()=>{toast('Не удалось удалить сценарий.','error');});};
    [bRun,bEdit,bTpl,bDel].forEach(x=>a.appendChild(x));
    els.list.appendChild(c);
  });
}

function fieldWrap(full){const w=document.createElement('label');w.className=full?'field field-full':'field';return w;}
function renderActions(){
  const root=els.actionList; root.innerHTML='';
  if(!state.draft||!state.draft.actions.length){const e=document.createElement('div');e.className='action-item';e.textContent='Добавьте хотя бы одно действие.';root.appendChild(e);return;}
  state.draft.actions.forEach((a,i)=>{
    if(!a.id)a.id=uid('action');
    const item=document.createElement('div');item.className='action-item';
    const collapsed=isItemCollapsed('action',a.id);
    if(collapsed)item.classList.add('item-collapsed');
    const head=document.createElement('div');head.className='action-item-head';
    const ord=document.createElement('span');ord.className='order';ord.textContent=String(i+1);
    const sel=document.createElement('select');typeList.forEach(t=>{const o=document.createElement('option');o.value=t[0];o.textContent=t[1];o.selected=t[0]===a.type;sel.appendChild(o);});
    sel.onchange=(ev)=>{const n=mkAction(ev.target.value);n.id=a.id;if('sheet' in n&&a.sheet)n.sheet=a.sheet;state.draft.actions[i]=n;renderActions();};
    const tools=document.createElement('div');tools.className='action-tools';
    const tg=document.createElement('button');tg.className='btn btn-ghost';tg.type='button';tg.textContent=collapsed?'Развернуть':'Свернуть';tg.onclick=()=>{toggleItemCollapsed('action',a.id);renderActions();};
    const up=document.createElement('button');up.className='btn btn-ghost';up.type='button';up.textContent='↑';up.disabled=i===0;up.onclick=()=>{const t=state.draft.actions[i-1];state.draft.actions[i-1]=state.draft.actions[i];state.draft.actions[i]=t;renderActions();};
    const dn=document.createElement('button');dn.className='btn btn-ghost';dn.type='button';dn.textContent='↓';dn.disabled=i===state.draft.actions.length-1;dn.onclick=()=>{const t=state.draft.actions[i+1];state.draft.actions[i+1]=state.draft.actions[i];state.draft.actions[i]=t;renderActions();};
    const rm=document.createElement('button');rm.className='btn btn-danger';rm.type='button';rm.textContent='Удалить';rm.onclick=()=>{state.draft.actions.splice(i,1);renderActions();};
    [tg,up,dn,rm].forEach(x=>tools.appendChild(x)); [ord,sel,tools].forEach(x=>head.appendChild(x)); item.appendChild(head);
    if(collapsed){
      const note=document.createElement('div');
      note.className='item-collapsed-note';
      note.textContent='Шаг свернут';
      item.appendChild(note);
      root.appendChild(item);
      return;
    }
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
      else if(f.t==='select'){
        inp=document.createElement('select');
        const options=f.dynamicOptions?getDynamicActionOptions(f.dynamicOptions):(f.o||[]);
        options.forEach(op=>{const o=document.createElement('option');o.value=op[0];o.textContent=op[1];o.selected=String(op[0])===String(a[f.k]);inp.appendChild(o);});
      }
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
    const collapsed=isItemCollapsed('input',f.id);
    if(collapsed)item.classList.add('item-collapsed');

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
    const tg=document.createElement('button');
    tg.className='btn btn-ghost';
    tg.type='button';
    tg.textContent=collapsed?'Развернуть':'Свернуть';
    tg.onclick=()=>{toggleItemCollapsed('input',f.id);renderInputFields();};
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
    [tg,up,dn,rm].forEach(x=>tools.appendChild(x));
    [ord,title,tools].forEach(x=>head.appendChild(x));
    item.appendChild(head);
    if(collapsed){
      const note=document.createElement('div');
      note.className='item-collapsed-note';
      note.textContent='Поле свернуто';
      item.appendChild(note);
      root.appendChild(item);
      return;
    }

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
      const allFields=Array.isArray(state.draft&&state.draft.inputFields)?state.draft.inputFields.map(normInputField):[];
      const parentField=allFields.find(field=>field.id===String(f.depends_on||'').trim());
      const parentDefault=parentField?normalizeScalarFieldValue(parentField.default_value,parentField.input_type):'';
      const opts=resolveSelectOptions(f,parentDefault);
      const emptyOpt=document.createElement('option');emptyOpt.value='';emptyOpt.textContent='(пусто)';inDefault.appendChild(emptyOpt);
      opts.forEach(opt=>{const o=document.createElement('option');o.value=opt;o.textContent=opt;inDefault.appendChild(o);});
      const start=String(f.default_value==null?'':f.default_value);
      if(start&&opts.indexOf(start)<0){const extra=document.createElement('option');extra.value=start;extra.textContent=start;inDefault.appendChild(extra);}
      inDefault.value=start;
      inDefault.onchange=(ev)=>{
        f.default_value=ev.target.value;
        const hasDependents=allFields.some(field=>String(field.depends_on||'').trim()===f.id);
        if(hasDependents)renderInputFields();
      };
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
      const allFields=Array.isArray(state.draft&&state.draft.inputFields)?state.draft.inputFields.map(normInputField):[];
      const hasDependency=String(f.depends_on||'').trim()!=='';

      const buildListEditorRow=(title,values,onApply,subtitle)=>{
        const wrap=fieldWrap(true);
        const cap=document.createElement('span');cap.textContent=title;
        const row=document.createElement('div');row.className='list-editor-row';
        const btn=document.createElement('button');
        btn.type='button';
        btn.className='btn btn-secondary';
        btn.textContent='Список...';
        const status=document.createElement('span');
        status.className='list-editor-status';
        const refreshStatus=(list)=>{
          const count=Array.isArray(list)?list.length:parseInputOptions(list).length;
          const preview=optionsPreview(list,2);
          status.textContent=`Пунктов: ${count}. ${preview}`;
        };
        refreshStatus(values);
        btn.onclick=(ev)=>{
          ev.preventDefault();
          openListEditor({
            title,
            subtitle:subtitle||'Каждый вариант вводится с новой строки.',
            values:Array.isArray(values)?values:parseInputOptions(values),
            onApply:(next)=>{
              onApply(Array.isArray(next)?next:[]);
              refreshStatus(next);
              renderInputFields();
            }
          });
        };
        row.appendChild(btn);
        row.appendChild(status);
        wrap.appendChild(cap);
        wrap.appendChild(row);
        return wrap;
      };

      const wDepends=fieldWrap(false);
      const capDepends=document.createElement('span');capDepends.textContent='Зависит от поля';
      const inDepends=document.createElement('select');
      const noneOpt=document.createElement('option');noneOpt.value='';noneOpt.textContent='(нет зависимости)';inDepends.appendChild(noneOpt);
      allFields.forEach((other,idx)=>{
        if(other.id===f.id)return;
        if(other.input_type!=='select')return;
        const opt=document.createElement('option');
        opt.value=other.id;
        opt.textContent=inputFieldLabel(other,idx);
        if(String(f.depends_on||'')===other.id)opt.selected=true;
        inDepends.appendChild(opt);
      });
      if(String(f.depends_on||'')&&!allFields.some(other=>other.id===String(f.depends_on||''))){
        const missed=document.createElement('option');
        missed.value=String(f.depends_on||'');
        missed.textContent='(поле не найдено)';
        missed.selected=true;
        inDepends.appendChild(missed);
      }
      inDepends.onchange=(ev)=>{
        f.depends_on=String(ev.target.value||'').trim();
        if(!f.depends_on){
          f.fallback_on_missing=true;
        }else if(!parseInputOptions(f.options).length){
          f.fallback_on_missing=false;
        }
        renderInputFields();
      };
      wDepends.appendChild(capDepends);wDepends.appendChild(inDepends);fs.appendChild(wDepends);

      if(!hasDependency){
        const wBaseList=buildListEditorRow(
          'Варианты списка',
          parseInputOptions(f.options),
          (next)=>{f.options=stringifyInputOptions(next);},
          'Базовый список: один пункт на строку.'
        );
        const helpBase=document.createElement('small');
        helpBase.className='field-help';
        helpBase.textContent='Откройте редактор списка и вставьте варианты построчно.';
        wBaseList.appendChild(helpBase);
        fs.appendChild(wBaseList);
      }else{
        const parentField=allFields.find(other=>other.id===String(f.depends_on||'').trim());
        const parentOptions=parentField?parseInputOptions(parentField.options):[];
        const dataListId=`dep_parent_${String(f.id||i).replace(/[^a-zA-Z0-9_-]/g,'_')}`;

        const wOptionsMap=fieldWrap(true);
        const capOptionsMap=document.createElement('span');capOptionsMap.textContent='Зависимые варианты';
        const parsedRules=parseDependentRules(f.options_map);
        const ruleRows=parsedRules.rules.map(rule=>({key:rule.key,options:Array.isArray(rule.options)?rule.options:[]}));
        if(!ruleRows.length)ruleRows.push({key:'',options:[]});
        const mapBox=document.createElement('div');mapBox.className='dependent-builder';
        const mapActions=document.createElement('div');mapActions.className='dependent-builder-actions';
        const btnAddRule=document.createElement('button');btnAddRule.type='button';btnAddRule.className='btn btn-ghost';btnAddRule.textContent='+ Добавить правило';
        mapActions.appendChild(btnAddRule);

        const parentList=document.createElement('datalist');
        parentList.id=dataListId;
        parentOptions.forEach(opt=>{
          const item=document.createElement('option');
          item.value=opt;
          parentList.appendChild(item);
        });

        const syncRuleRows=()=>{
          f.options_map=stringifyDependentRules(ruleRows.map(rule=>({key:String(rule.key||'').trim(),options:Array.isArray(rule.options)?rule.options:[]})));
        };

        const renderRuleRows=()=>{
          mapBox.innerHTML='';
          ruleRows.forEach((rule,rowIndex)=>{
            const row=document.createElement('div');row.className='dependent-builder-row';
            const parentValue=document.createElement('input');
            parentValue.type='text';
            parentValue.setAttribute('list',dataListId);
            parentValue.value=String(rule.key==null?'':rule.key);
            parentValue.placeholder='Значение в первом списке';
            parentValue.oninput=(ev)=>{ruleRows[rowIndex].key=ev.target.value;syncRuleRows();};

            const preview=document.createElement('div');
            preview.className='dependent-builder-preview';
            preview.textContent=optionsPreview(rule.options,2);

            const editList=document.createElement('button');
            editList.type='button';
            editList.className='btn btn-secondary';
            editList.textContent=`Список (${Array.isArray(rule.options)?rule.options.length:0})`;
            editList.onclick=(ev)=>{
              ev.preventDefault();
              openListEditor({
                title:`Список для «${String(ruleRows[rowIndex].key||'').trim()||'значения родителя'}»`,
                subtitle:'Каждый вариант вводится с новой строки.',
                values:Array.isArray(ruleRows[rowIndex].options)?ruleRows[rowIndex].options:[],
                onApply:(next)=>{
                  ruleRows[rowIndex].options=Array.isArray(next)?next:[];
                  syncRuleRows();
                  renderRuleRows();
                }
              });
            };

            const delRule=document.createElement('button');
            delRule.type='button';
            delRule.className='btn btn-danger dependent-builder-remove';
            delRule.textContent='-';
            delRule.onclick=(ev)=>{
              ev.preventDefault();
              ruleRows.splice(rowIndex,1);
              if(!ruleRows.length)ruleRows.push({key:'',options:[]});
              syncRuleRows();
              renderRuleRows();
            };
            row.appendChild(parentValue);
            row.appendChild(preview);
            row.appendChild(editList);
            row.appendChild(delRule);
            mapBox.appendChild(row);
          });
        };

        btnAddRule.onclick=(ev)=>{
          ev.preventDefault();
          ruleRows.push({key:'',options:[]});
          renderRuleRows();
        };

        renderRuleRows();
        syncRuleRows();

        const helpMap=document.createElement('small');
        helpMap.className='field-help';
        helpMap.textContent='Для каждого значения первого списка задайте свой набор второго.';

        wOptionsMap.appendChild(parentList);
        wOptionsMap.appendChild(capOptionsMap);
        wOptionsMap.appendChild(mapBox);
        wOptionsMap.appendChild(mapActions);
        if(parsedRules.error){
          const warn=document.createElement('small');
          warn.className='field-help';
          warn.textContent=`Обнаружены старые данные: ${parsedRules.error}`;
          wOptionsMap.appendChild(warn);
        }
        wOptionsMap.appendChild(helpMap);
        fs.appendChild(wOptionsMap);

        const wFallbackToggle=fieldWrap(true);
        const inlFallbackToggle=document.createElement('span');inlFallbackToggle.className='field-inline';
        const inFallbackToggle=document.createElement('input');
        inFallbackToggle.type='checkbox';
        inFallbackToggle.checked=f.fallback_on_missing!==false;
        inFallbackToggle.onchange=(ev)=>{f.fallback_on_missing=!!ev.target.checked;renderInputFields();};
        const txtFallbackToggle=document.createElement('span');
        txtFallbackToggle.textContent='Использовать запасной список, если правило не найдено';
        inlFallbackToggle.appendChild(inFallbackToggle);
        inlFallbackToggle.appendChild(txtFallbackToggle);
        wFallbackToggle.appendChild(inlFallbackToggle);
        fs.appendChild(wFallbackToggle);

        if(f.fallback_on_missing!==false){
          const wFallbackList=buildListEditorRow(
            'Запасной список',
            parseInputOptions(f.options),
            (next)=>{f.options=stringifyInputOptions(next);},
            'Этот список используется, если для значения родителя нет отдельного правила.'
          );
          fs.appendChild(wFallbackList);
        }
      }
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
function renderSourceConfig(){
  const root=els.sourceConfigRoot;
  if(!root)return;
  root.innerHTML='';
  if(!state.draft)return;
  const config=state.draft.sourceFileConfig=normSourceFileConfig(state.draft.sourceFileConfig);
  const schemaState=state.draft.sourceSchema=normSourceSchema(state.draft.sourceSchema);

  const head=document.createElement('div');
  head.className='source-config-head';
  const toggle=document.createElement('label');
  toggle.className='source-config-toggle';
  const chk=document.createElement('input');
  chk.type='checkbox';
  chk.checked=!!config.enabled;
  chk.onchange=(ev)=>{config.enabled=!!ev.target.checked;renderSourceConfig();renderActions();};
  const txt=document.createElement('span');
  txt.textContent='Использовать файл с исходными данными';
  toggle.appendChild(chk);
  toggle.appendChild(txt);
  head.appendChild(toggle);
  const meta=document.createElement('div');
  meta.className='source-config-meta';
  meta.textContent=`Полей: ${schemaState.fields.length}. Списков: ${schemaState.lists.length}. Таблиц: ${schemaState.tables.length}.`;
  head.appendChild(meta);
  root.appendChild(head);

  if(!config.enabled)return;

  const cfgGrid=document.createElement('div');
  cfgGrid.className='form-grid';
  const addConfigField=(label,value,apply,placeholder,type)=>{
    const wrap=document.createElement('label');
    wrap.className='field';
    const cap=document.createElement('span');
    cap.textContent=label;
    const inp=document.createElement('input');
    inp.type=type||'text';
    inp.value=value;
    if(placeholder)inp.placeholder=placeholder;
    inp.oninput=(ev)=>apply(ev.target.value);
    wrap.appendChild(cap);
    wrap.appendChild(inp);
    cfgGrid.appendChild(wrap);
  };
  addConfigField('Название блока',config.label,(v)=>{config.label=v;},'Файл с исходными данными');
  addConfigField('Максимум строк по умолчанию',String(config.maxRowsDefault),(v)=>{config.maxRowsDefault=v;},'', 'number');
  addConfigField('Лимит размера файла (MB)',String(config.maxFileSizeMb),(v)=>{config.maxFileSizeMb=v;},'', 'number');
  addConfigField('Строк предпросмотра',String(config.previewRows),(v)=>{config.previewRows=v;},'', 'number');

  const checks=document.createElement('div');
  checks.className='field field-full';
  const checksRow=document.createElement('div');
  checksRow.className='field-inline';
  [['required','Файл обязателен'],['strictMode','Строгая проверка структуры']].forEach(item=>{
    const label=document.createElement('label');
    label.className='field-inline';
    const input=document.createElement('input');
    input.type='checkbox';
    input.checked=!!config[item[0]];
    input.onchange=(ev)=>{config[item[0]]=!!ev.target.checked;};
    const text=document.createElement('span');
    text.textContent=item[1];
    label.appendChild(input);
    label.appendChild(text);
    checksRow.appendChild(label);
  });
  checks.appendChild(checksRow);
  cfgGrid.appendChild(checks);
  root.appendChild(cfgGrid);

  const fieldsWrap=document.createElement('div');
  fieldsWrap.className='source-section-block';
  const fieldsTitle=document.createElement('h4');
  fieldsTitle.className='source-section-title';
  fieldsTitle.textContent='Поля из файла';
  fieldsWrap.appendChild(fieldsTitle);
  if(!schemaState.fields.length){
    const empty=document.createElement('div');
    empty.className='source-empty';
    empty.textContent='Поля файла не настроены.';
    fieldsWrap.appendChild(empty);
  }
  schemaState.fields.forEach((field,index)=>{
    const item=document.createElement('div');
    item.className='source-item';
    const headRow=document.createElement('div');
    headRow.className='source-item-head';
    const title=document.createElement('div');
    title.className='source-item-title';
    title.textContent=field.label;
    const tools=document.createElement('div');
    tools.className='source-item-tools';
    const del=document.createElement('button');
    del.type='button';
    del.className='btn btn-danger';
    del.textContent='Удалить';
    del.onclick=()=>{schemaState.fields.splice(index,1);renderSourceConfig();renderActions();};
    tools.appendChild(del);
    headRow.appendChild(title);
    headRow.appendChild(tools);
    item.appendChild(headRow);

    const grid=document.createElement('div');
    grid.className='form-grid';
    const addField=(label,value,apply,placeholder,type)=>{
      const wrap=document.createElement('label');
      wrap.className='field';
      const cap=document.createElement('span');
      cap.textContent=label;
      let inp;
      if(type==='select'){
        inp=document.createElement('select');
        [['text','Текст'],['number','Число'],['date','Дата']].forEach(opt=>{
          const o=document.createElement('option');
          o.value=opt[0];
          o.textContent=opt[1];
          if(opt[0]===value)o.selected=true;
          inp.appendChild(o);
        });
        inp.onchange=(ev)=>apply(ev.target.value);
      }else if(type==='check'){
        inp=document.createElement('input');
        inp.type='checkbox';
        inp.checked=!!value;
        inp.onchange=(ev)=>apply(!!ev.target.checked);
      }else{
        inp=document.createElement('input');
        inp.type=type||'text';
        inp.value=value;
        if(placeholder)inp.placeholder=placeholder;
        inp.oninput=(ev)=>apply(ev.target.value);
      }
      wrap.appendChild(cap);
      wrap.appendChild(inp);
      grid.appendChild(wrap);
    };
    addField('Название',field.label,(v)=>{field.label=v;title.textContent=String(v||'').trim()||`Поле ${index+1}`;});
    addField('Ключ поля',field.key,(v)=>{field.key=safeSourceKey(v,`field_${index+1}`);});
    addField('Лист',field.sheet,(v)=>{field.sheet=v;},'Лист1');
    addField('Ячейка',field.address,(v)=>{field.address=v;},'B2');
    addField('Тип значения',field.type,(v)=>{field.type=v;},'', 'select');
    addField('Обязательное',field.required,(v)=>{field.required=!!v;},'', 'check');
    item.appendChild(grid);
    fieldsWrap.appendChild(item);
  });
  root.appendChild(fieldsWrap);

  const listsWrap=document.createElement('div');
  listsWrap.className='source-section-block';
  const listsTitle=document.createElement('h4');
  listsTitle.className='source-section-title';
  listsTitle.textContent='Списки из файла';
  listsWrap.appendChild(listsTitle);
  if(!schemaState.lists.length){
    const empty=document.createElement('div');
    empty.className='source-empty';
    empty.textContent='Списки файла не настроены.';
    listsWrap.appendChild(empty);
  }
  schemaState.lists.forEach((list,index)=>{
    const item=document.createElement('div');
    item.className='source-item';
    const headRow=document.createElement('div');
    headRow.className='source-item-head';
    const title=document.createElement('div');
    title.className='source-item-title';
    title.textContent=list.label;
    const tools=document.createElement('div');
    tools.className='source-item-tools';
    const del=document.createElement('button');
    del.type='button';
    del.className='btn btn-danger';
    del.textContent='Удалить';
    del.onclick=()=>{schemaState.lists.splice(index,1);renderSourceConfig();renderActions();};
    tools.appendChild(del);
    headRow.appendChild(title);
    headRow.appendChild(tools);
    item.appendChild(headRow);

    const grid=document.createElement('div');
    grid.className='form-grid';
    const addField=(label,value,apply,placeholder,type,options)=>{
      const wrap=document.createElement('label');
      wrap.className='field';
      const cap=document.createElement('span');
      cap.textContent=label;
      let inp;
      if(type==='select'){
        inp=document.createElement('select');
        (options||[]).forEach(opt=>{
          const o=document.createElement('option');
          o.value=opt[0];
          o.textContent=opt[1];
          if(opt[0]===value)o.selected=true;
          inp.appendChild(o);
        });
        inp.onchange=(ev)=>apply(ev.target.value);
      }else if(type==='check'){
        inp=document.createElement('input');
        inp.type='checkbox';
        inp.checked=!!value;
        inp.onchange=(ev)=>apply(!!ev.target.checked);
      }else{
        inp=document.createElement('input');
        inp.type=type||'text';
        inp.value=value;
        if(placeholder)inp.placeholder=placeholder;
        inp.oninput=(ev)=>apply(ev.target.value);
      }
      wrap.appendChild(cap);
      wrap.appendChild(inp);
      grid.appendChild(wrap);
      return inp;
    };
    const stopModeOptions=[['empty','До пустой ячейки'],['stop_value','До стоп-значения']];
    addField('Название списка',list.label,(v)=>{list.label=v;title.textContent=String(v||'').trim()||`Список ${index+1}`;});
    addField('Ключ списка',list.key,(v)=>{list.key=safeSourceKey(v,`list_${index+1}`);});
    addField('Лист',list.sheet,(v)=>{list.sheet=v;},'Лист1');
    addField('Стартовая ячейка',list.startAddress,(v)=>{list.startAddress=v;},'A1');
    addField('Направление чтения',list.direction,(v)=>{list.direction=v;},'', 'select', [['down','Вниз'],['right','Вправо']]);
    addField('Граница чтения',list.stopMode,(v)=>{list.stopMode=v;renderSourceConfig();},'', 'select', stopModeOptions);
    addField('Стоп-значение',list.stopValue,(v)=>{list.stopValue=v;},'СТОП');
    addField('Максимум элементов',list.maxItems,(v)=>{list.maxItems=v;},String(config.maxRowsDefault),'number');
    addField('Тип значения',list.type,(v)=>{list.type=v;},'', 'select', [['text','Текст'],['number','Число'],['date','Дата']]);
    addField('Обязательное',list.required,(v)=>{list.required=!!v;},'', 'check');
    item.appendChild(grid);
    listsWrap.appendChild(item);
  });
  root.appendChild(listsWrap);

  const tablesWrap=document.createElement('div');
  tablesWrap.className='source-section-block';
  const tablesTitle=document.createElement('h4');
  tablesTitle.className='source-section-title';
  tablesTitle.textContent='Таблицы из файла';
  tablesWrap.appendChild(tablesTitle);
  if(!schemaState.tables.length){
    const empty=document.createElement('div');
    empty.className='source-empty';
    empty.textContent='Таблицы файла не настроены.';
    tablesWrap.appendChild(empty);
  }
  schemaState.tables.forEach((table,index)=>{
    const item=document.createElement('div');
    item.className='source-item';
    const headRow=document.createElement('div');
    headRow.className='source-item-head';
    const title=document.createElement('div');
    title.className='source-item-title';
    title.textContent=table.label;
    const tools=document.createElement('div');
    tools.className='source-item-tools';
    const addColumn=document.createElement('button');
    addColumn.type='button';
    addColumn.className='btn btn-secondary';
    addColumn.textContent='Колонка';
    addColumn.onclick=()=>{table.columns.push(mkSourceColumn());renderSourceConfig();};
    const del=document.createElement('button');
    del.type='button';
    del.className='btn btn-danger';
    del.textContent='Удалить';
    del.onclick=()=>{schemaState.tables.splice(index,1);renderSourceConfig();renderActions();};
    tools.appendChild(addColumn);
    tools.appendChild(del);
    headRow.appendChild(title);
    headRow.appendChild(tools);
    item.appendChild(headRow);

    const grid=document.createElement('div');
    grid.className='form-grid';
    const addField=(label,value,apply,placeholder,type)=>{
      const wrap=document.createElement('label');
      wrap.className='field';
      const cap=document.createElement('span');
      cap.textContent=label;
      const inp=document.createElement('input');
      inp.type=type||'text';
      inp.value=value;
      if(placeholder)inp.placeholder=placeholder;
      inp.oninput=(ev)=>apply(ev.target.value);
      wrap.appendChild(cap);
      wrap.appendChild(inp);
      grid.appendChild(wrap);
    };
    addField('Название таблицы',table.label,(v)=>{table.label=v;title.textContent=String(v||'').trim()||`Таблица ${index+1}`;});
    addField('Ключ таблицы',table.key,(v)=>{table.key=safeSourceKey(v,`table_${index+1}`);});
    addField('Лист',table.sheet,(v)=>{table.sheet=v;},'Лист1');
    addField('Строка заголовков',table.headerRow,(v)=>{table.headerRow=v;},'6','number');
    addField('Смещение до данных',table.startRowOffset,(v)=>{table.startRowOffset=v;},'1','number');
    addField('Заголовок ключевой колонки',table.keyHeader,(v)=>{table.keyHeader=v;},'Номер');
    addField('Пустых строк подряд',table.emptyRowTolerance,(v)=>{table.emptyRowTolerance=v;},'1','number');
    addField('Максимум строк',table.maxRows,(v)=>{table.maxRows=v;},String(config.maxRowsDefault),'number');
    item.appendChild(grid);

    const columnsWrap=document.createElement('div');
    columnsWrap.className='source-columns-list';
    table.columns.forEach((column,colIndex)=>{
      const row=document.createElement('div');
      row.className='source-column-row';
      const makeField=(label,value,apply,placeholder,type)=>{
        const wrap=document.createElement('label');
        wrap.className='field';
        const cap=document.createElement('span');
        cap.textContent=label;
        let inp;
        if(type==='select'){
          inp=document.createElement('select');
          [['text','Текст'],['number','Число'],['date','Дата']].forEach(opt=>{
            const o=document.createElement('option');
            o.value=opt[0];
            o.textContent=opt[1];
            if(opt[0]===value)o.selected=true;
            inp.appendChild(o);
          });
          inp.onchange=(ev)=>apply(ev.target.value);
        }else if(type==='check'){
          inp=document.createElement('input');
          inp.type='checkbox';
          inp.checked=!!value;
          inp.onchange=(ev)=>apply(!!ev.target.checked);
        }else{
          inp=document.createElement('input');
          inp.type=type||'text';
          inp.value=value;
          if(placeholder)inp.placeholder=placeholder;
          inp.oninput=(ev)=>apply(ev.target.value);
        }
        wrap.appendChild(cap);
        wrap.appendChild(inp);
        row.appendChild(wrap);
      };
      makeField('Название',column.label,(v)=>{column.label=v;},'Номер');
      makeField('Ключ',column.key,(v)=>{column.key=safeSourceKey(v,`column_${colIndex+1}`);});
      makeField('Заголовок в файле',column.header,(v)=>{column.header=v;},'Номер');
      makeField('Тип',column.type,(v)=>{column.type=v;},'', 'select');
      const actionWrap=document.createElement('div');
      actionWrap.className='source-inline-actions';
      const delCol=document.createElement('button');
      delCol.type='button';
      delCol.className='btn btn-danger';
      delCol.textContent='-';
      delCol.onclick=()=>{table.columns.splice(colIndex,1);if(!table.columns.length)table.columns.push(mkSourceColumn());renderSourceConfig();};
      actionWrap.appendChild(delCol);
      row.appendChild(actionWrap);
      columnsWrap.appendChild(row);
    });
    item.appendChild(columnsWrap);
    tablesWrap.appendChild(item);
  });
  root.appendChild(tablesWrap);
}
function readDraft(){if(!state.draft)return;state.draft.name=String(els.name.value||'').trim();state.draft.templatePath=String(els.tpl.value||'').trim();state.draft.description=String(els.descr.value||'').trim();}
function validateDraft(){
  if(!state.draft)return 'Сценарий не открыт.'; readDraft();
  if(!state.draft.name)return 'Введите название сценария.';
  if(!state.draft.templatePath)return 'Выберите файл шаблона.';
  if(/fakepath/i.test(state.draft.templatePath))return 'Получен fakepath. Укажите полный путь к шаблону.';
  if(!Array.isArray(state.draft.actions))state.draft.actions=[];
  if(!Array.isArray(state.draft.inputFields))state.draft.inputFields=[];
  state.draft.sourceFileConfig=normSourceFileConfig(state.draft.sourceFileConfig);
  state.draft.sourceSchema=normSourceSchema(state.draft.sourceSchema);
  if(!state.draft.actions.length&&!state.draft.inputFields.length&&!sourceConfigEnabled(state.draft))return 'Добавьте хотя бы одно действие, поле ввода или настройте источник файла.';
  const normalizedInputs=state.draft.inputFields.map(normInputField);
  state.draft.inputFields=normalizedInputs;
  const inputFieldIds=normalizedInputs.map(field=>field.id);
  const sourceFieldKeys=state.draft.sourceSchema.fields.map(field=>field.key);
  const sourceListKeys=state.draft.sourceSchema.lists.map(list=>list.key);
  const sourceTableKeys=state.draft.sourceSchema.tables.map(table=>table.key);
  const cyclicFieldId=findInputDependencyCycle(normalizedInputs);
  if(cyclicFieldId){
    const cyclicIndex=normalizedInputs.findIndex(field=>field.id===cyclicFieldId);
    return `${inputFieldLabel(normalizedInputs[Math.max(0,cyclicIndex)],Math.max(0,cyclicIndex))}: обнаружена циклическая зависимость списков.`;
  }
  for(let i=0;i<state.draft.actions.length;i++){
    const a=normAction(state.draft.actions[i]);state.draft.actions[i]=a;if(!schema[a.type])return `Шаг ${i+1}: неизвестный тип действия.`;
    for(const f of schema[a.type].f||[]){if(!f.r||f.t==='check')continue;const v=a[f.k];if(v==null||String(v).trim()==='')return `Шаг ${i+1} (${schema[a.type].label}): заполните поле «${f.l}».`;}
    if(a.cond_enabled){
      if(!a.cond_field)return `Шаг ${i+1} (${schema[a.type].label}): выберите поле в условии.`;
      if(inputFieldIds.indexOf(a.cond_field)<0)return `Шаг ${i+1} (${schema[a.type].label}): поле условия не найдено.`;
      if(conditionNeedsValue(a.cond_operator)&&String(a.cond_value==null?'':a.cond_value).trim()==='')
        return `Шаг ${i+1} (${schema[a.type].label}): укажите значение для условия.`;
    }
    if(a.type==='insert_source_field'){
      if(!sourceConfigEnabled(state.draft))return `Шаг ${i+1} (${schema[a.type].label}): включите источник данных из файла.`;
      if(sourceFieldKeys.indexOf(String(a.source_field_id||'').trim())<0)return `Шаг ${i+1} (${schema[a.type].label}): поле файла не найдено.`;
      const parsedTargets=parseCellBindings(a.targets);
      if(parsedTargets.error)return `Шаг ${i+1} (${schema[a.type].label}): ${parsedTargets.error}`;
    }
    if(a.type==='insert_source_list'){
      if(!sourceConfigEnabled(state.draft))return `Шаг ${i+1} (${schema[a.type].label}): включите источник данных из файла.`;
      if(sourceListKeys.indexOf(String(a.source_list_id||'').trim())<0)return `Шаг ${i+1} (${schema[a.type].label}): список файла не найден.`;
      const parsedTargets=parseTableTargets(a.targets);
      if(parsedTargets.error)return `Шаг ${i+1} (${schema[a.type].label}): ${parsedTargets.error}`;
      for(let t=0;t<parsedTargets.items.length;t+=1){
        const target=parsedTargets.items[t];
        const anchorRange=parseA1Range(target.anchor_cell);
        if(!anchorRange||anchorRange.c1!==anchorRange.c2||anchorRange.r1!==anchorRange.r2){
          return `Шаг ${i+1} (${schema[a.type].label}): неверная первая ячейка в точке вставки ${t+1}.`;
        }
        const templateRange=parseA1Range(target.template_row_range);
        if(!templateRange)return `Шаг ${i+1} (${schema[a.type].label}): неверная строка-шаблон в точке вставки ${t+1}.`;
        if(templateRange.r1!==templateRange.r2)return `Шаг ${i+1} (${schema[a.type].label}): строка-шаблон в точке вставки ${t+1} должна быть одной строкой.`;
      }
    }
    if(a.type==='insert_source_table'){
      if(!sourceConfigEnabled(state.draft))return `Шаг ${i+1} (${schema[a.type].label}): включите источник данных из файла.`;
      if(sourceTableKeys.indexOf(String(a.source_table_id||'').trim())<0)return `Шаг ${i+1} (${schema[a.type].label}): таблица файла не найдена.`;
      const parsedTargets=parseTableTargets(a.targets);
      if(parsedTargets.error)return `Шаг ${i+1} (${schema[a.type].label}): ${parsedTargets.error}`;
      for(let t=0;t<parsedTargets.items.length;t+=1){
        const target=parsedTargets.items[t];
        const anchorRange=parseA1Range(target.anchor_cell);
        if(!anchorRange||anchorRange.c1!==anchorRange.c2||anchorRange.r1!==anchorRange.r2){
          return `Шаг ${i+1} (${schema[a.type].label}): неверная первая ячейка в точке вставки ${t+1}.`;
        }
        const templateRange=parseA1Range(target.template_row_range);
        if(!templateRange)return `Шаг ${i+1} (${schema[a.type].label}): неверная строка-шаблон в точке вставки ${t+1}.`;
        if(templateRange.r1!==templateRange.r2)return `Шаг ${i+1} (${schema[a.type].label}): строка-шаблон в точке вставки ${t+1} должна быть одной строкой.`;
      }
      const parsedMappings=parseActionSourceMappings(a.mappings);
      if(parsedMappings.error)return `Шаг ${i+1} (${schema[a.type].label}): ${parsedMappings.error}`;
      const table=state.draft.sourceSchema.tables.find(item=>item.key===String(a.source_table_id||'').trim());
      if(table){
        const sourceKeys=table.columns.map(column=>column.key);
        const usedTargets=Object.create(null);
        for(let m=0;m<parsedMappings.items.length;m+=1){
          const mapping=parsedMappings.items[m];
          if(sourceKeys.indexOf(mapping.sourceKey)<0)return `Шаг ${i+1} (${schema[a.type].label}): колонка источника «${mapping.sourceKey}» не найдена.`;
          const targetKey=String(mapping.targetColumn||'').toUpperCase();
          if(usedTargets[targetKey])return `Шаг ${i+1} (${schema[a.type].label}): колонка назначения «${targetKey}» указана несколько раз.`;
          usedTargets[targetKey]=true;
        }
      }
    }
  }
  for(let i=0;i<normalizedInputs.length;i+=1){
    const field=normalizedInputs[i];
    const label=inputFieldLabel(field,i);
    if(!field.name)return `Поле ввода ${i+1}: укажите название.`;
    if(field.input_type==='select'){
      const baseOptions=parseInputOptions(field.options);
      const parentId=String(field.depends_on||'').trim();
      if(parentId){
        if(parentId===field.id)return `${label}: поле не может зависеть от самого себя.`;
        if(field.multiple)return `${label}: зависимый список не поддерживает множественный ввод.`;
        const parent=normalizedInputs.find(item=>item.id===parentId);
        if(!parent)return `${label}: поле-источник списка не найдено.`;
        if(parent.input_type!=='select')return `${label}: поле-источник должно быть типа «Список».`;
        if(parent.multiple)return `${label}: поле-источник не должно быть множественным.`;
        const optionsMap=parseDependentOptionsMap(field.options_map);
        if(optionsMap.error)return `${label}: ${optionsMap.error}`;
        const hasFallback=(field.fallback_on_missing!==false)&&baseOptions.length>0;
        if(optionsMap.valuesCount===0&&!hasFallback)return `${label}: добавьте зависимые правила или включите запасной список.`;
      }else if(!baseOptions.length){
        return `${label}: добавьте хотя бы один вариант списка.`;
      }
    }
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
  if(sourceConfigEnabled(state.draft)){
    const schemaState=state.draft.sourceSchema;
    if(!schemaState.fields.length&&!schemaState.lists.length&&!schemaState.tables.length)return 'Добавьте хотя бы одно поле, список или таблицу в источнике файла.';
    const usedFieldKeys=Object.create(null);
    for(let i=0;i<schemaState.fields.length;i+=1){
      const field=normSourceField(schemaState.fields[i],i);
      schemaState.fields[i]=field;
      if(!field.label)return `Источник файла: поле ${i+1} без названия.`;
      if(!field.sheet)return `Источник файла: укажите лист для поля «${field.label}».`;
      if(!field.address)return `Источник файла: укажите ячейку для поля «${field.label}».`;
      if(usedFieldKeys[field.key])return `Источник файла: ключ поля «${field.key}» повторяется.`;
      usedFieldKeys[field.key]=true;
    }
    const usedListKeys=Object.create(null);
    for(let i=0;i<schemaState.lists.length;i+=1){
      const list=normSourceList(schemaState.lists[i],i);
      schemaState.lists[i]=list;
      if(!list.label)return `Источник файла: список ${i+1} без названия.`;
      if(!list.sheet)return `Источник файла: укажите лист для списка «${list.label}».`;
      if(!list.startAddress)return `Источник файла: укажите стартовую ячейку для списка «${list.label}».`;
      if(!(parseInt(list.maxItems,10)>0))return `Источник файла: неверный лимит элементов для списка «${list.label}».`;
      if(list.stopMode==='stop_value'&&String(list.stopValue||'').trim()==='')return `Источник файла: укажите стоп-значение для списка «${list.label}».`;
      if(usedListKeys[list.key])return `Источник файла: ключ списка «${list.key}» повторяется.`;
      usedListKeys[list.key]=true;
    }
    const usedTableKeys=Object.create(null);
    for(let i=0;i<schemaState.tables.length;i+=1){
      const table=normSourceTable(schemaState.tables[i],i);
      schemaState.tables[i]=table;
      if(!table.label)return `Источник файла: таблица ${i+1} без названия.`;
      if(!table.sheet)return `Источник файла: укажите лист для таблицы «${table.label}».`;
      if(!(parseInt(table.headerRow,10)>0))return `Источник файла: неверная строка заголовков для таблицы «${table.label}».`;
      if(!(parseInt(table.startRowOffset,10)>=0))return `Источник файла: неверное смещение данных для таблицы «${table.label}».`;
      if(!table.keyHeader)return `Источник файла: укажите ключевую колонку для таблицы «${table.label}».`;
      if(!(parseInt(table.emptyRowTolerance,10)>=0))return `Источник файла: неверный допуск пустых строк для таблицы «${table.label}».`;
      if(!(parseInt(table.maxRows,10)>0))return `Источник файла: неверный лимит строк для таблицы «${table.label}».`;
      if(usedTableKeys[table.key])return `Источник файла: ключ таблицы «${table.key}» повторяется.`;
      usedTableKeys[table.key]=true;
      const usedColumnKeys=Object.create(null);
      for(let c=0;c<table.columns.length;c+=1){
        const column=normSourceColumn(table.columns[c],c);
        table.columns[c]=column;
        if(!column.label)return `Источник файла: колонка ${c+1} в таблице «${table.label}» без названия.`;
        if(!column.header)return `Источник файла: колонка «${column.label}» в таблице «${table.label}» без заголовка.`;
        if(usedColumnKeys[column.key])return `Источник файла: ключ колонки «${column.key}» повторяется в таблице «${table.label}».`;
        usedColumnKeys[column.key]=true;
      }
    }
  }
  return '';
}
function dbReason(err){
  const d=err&&err.details?err.details:null;
  return d&&d.stderr?d.stderr:(d&&d.error?d.error:(err&&err.message?err.message:'Неизвестная ошибка'));
}
async function prepareTemplatePath(draft,previousScenario){
  const source=resolveTemplate(draft.templatePath);
  if(!source)throw new Error('Некорректный путь к шаблону.');
  const targetName=sanitizeTemplateName(draft.name);
  const targetPath=joinPath(getTemplatesDir(),`${targetName}.${SCENARIO_TEMPLATE_EXT}`);
  if(samePath(source,targetPath))return targetPath;
  if(previousScenario){
    const prevName=String(previousScenario.name||'');
    const renamed=prevName!==String(draft.name||'');
    const prevTemplatePath=resolveTemplate(previousScenario.templatePath);
    if(renamed&&prevTemplatePath&&samePath(source,prevTemplatePath)&&isManagedScenarioTemplatePath(prevTemplatePath,prevName)&&state.fileBridge==='ready'){
      const moved=await movePath(source,targetPath,4000).catch(()=>null);
      if(moved&&moved.ok)return targetPath;
    }
  }
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
  const existingIdx=state.scenarios.findIndex(x=>x.id===draftRef.id);
  const previousScenario=existingIdx>=0?state.scenarios[existingIdx]:null;
  const previousName=String(previousScenario&&previousScenario.name||'');
  const previousScenarioFile=previousScenario?String(state.scenarioFiles[draftRef.id]||'').trim():'';
  const previousTemplatePath=previousScenario?resolveTemplate(previousScenario.templatePath):'';
  const btn=els.btnSave;
  if(btn){btn.disabled=true;btn.textContent='Сохранение...';}
  try{
    toast('Подготовка шаблона...','ok');
    const preparedPath=await prepareTemplatePath(draftRef,previousScenario);
    if(state.draft!==draftRef||!state.draft)return;
    state.draft.templatePath=preparedPath;
    els.tpl.value=preparedPath;
    const s={id:state.draft.id,name:state.draft.name,description:state.draft.description,templatePath:state.draft.templatePath,actions:state.draft.actions.map(normAction),inputFields:(state.draft.inputFields||[]).map(normInputField),sourceFileConfig:normSourceFileConfig(state.draft.sourceFileConfig),sourceSchema:normSourceSchema(state.draft.sourceSchema),createdAt:state.draft.createdAt||new Date().toISOString(),updatedAt:new Date().toISOString(),lastRunAt:state.draft.lastRunAt||null};
    let staleScenarioPath='';
    let staleTemplatePath='';
    if(existingIdx>=0){
      const renamed=previousName!==String(s.name||'');
      state.scenarios[existingIdx]=s;
      if(renamed){
        const nextScenarioFile=makeScenarioFileNameForScenario(s.id,s.name);
        if(previousScenarioFile&&previousScenarioFile.toLowerCase()!==nextScenarioFile.toLowerCase()){
          const prevScenarioPath=getScenarioPath(previousScenarioFile);
          const nextScenarioPath=getScenarioPath(nextScenarioFile);
          let renamedScenarioFile=false;
          if(state.fileBridge==='ready'){
            const moved=await movePath(prevScenarioPath,nextScenarioPath,3500).catch(()=>null);
            renamedScenarioFile=!!(moved&&moved.ok);
          }
          if(!renamedScenarioFile)staleScenarioPath=prevScenarioPath;
        }
        state.scenarioFiles[s.id]=nextScenarioFile;
        if(previousTemplatePath&&isManagedScenarioTemplatePath(previousTemplatePath,previousName)&&!samePath(previousTemplatePath,preparedPath)){
          staleTemplatePath=previousTemplatePath;
        }
      }else if(previousScenarioFile){
        state.scenarioFiles[s.id]=previousScenarioFile;
      }
      toast('Сценарий обновлен.','ok');
    }else{
      state.scenarios.push(s);
      delete state.scenarioFiles[s.id];
      toast('Сценарий создан.','ok');
    }
    await save();
    if(state.fileBridge==='ready'){
      const tasks=[];
      if(staleScenarioPath)tasks.push(deletePath(staleScenarioPath,2500));
      if(staleTemplatePath)tasks.push(deletePath(staleTemplatePath,2500));
      if(tasks.length){
        Promise.all(tasks).catch(()=>{warnFileOpsNeedRebuild();});
      }
    }
    renderList();closeModal();
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
function clearPathCheck(id){const p=PATH_CHECK_PENDING[id];if(!p)return null;if(p.t)clearTimeout(p.t);if(p.i)clearInterval(p.i);delete PATH_CHECK_PENDING[id];return p;}
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
  Object.keys(PATH_CHECK_PENDING).forEach(id=>{
    const pending=PATH_CHECK_PENDING[id];if(!pending)return;
    if(!Object.prototype.hasOwnProperty.call(data,pending.k))return;
    if(!boolVal(data[pending.k]))return;
    const p=clearPathCheck(id);
    if(p&&typeof p.res==='function')p.res(true);
  });
}
function bindDirectNative(){
  try{
    if(window.__reportsDirectNativeBound)return true;
    const sdk=window.parent&&window.parent.sdk;
    if(!sdk||typeof sdk.on!=='function')return false;
    sdk.on('on_native_message',(cmd,param)=>{
      const c=String(cmd||'').trim();
      if(c==='reports:fileResult'){onNativeFileResult(param);return;}
      if(c==='docbuilder:result'||c==='docbuilder:probeResult'){const data=parseNativeParam(param);if(data)onDbResult(data);return;}
      if(c==='files:checked'){handleFilesChecked(param);}
    });
    window.__reportsDirectNativeBound=true;
    return true;
  }catch(_){return false;}
}
function startWatch(id,payload,ms){
  clearWatch(id);
  const out=isObj(payload)&&payload.outputPath?String(payload.outputPath):(isObj(payload)&&isObj(payload.argument)&&payload.argument.outputPath?String(payload.argument.outputPath):'');
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
      const sdk=(window.parent&&window.parent.sdk)||(window.sdk||null);
      if(bindDirectNative()&&sdk&&typeof sdk.command==='function'){
        if(event==='reportsDocBuilderRun'){startWatch(id,d,ms);sdk.command('docbuilder:run',JSON.stringify(d));return;}
        if(event==='reportsDocBuilderProbe'){sdk.command('docbuilder:probe',JSON.stringify(d));return;}
      }
      if(window.parent&&window.parent!==window){
        post({event,source:'reports-ui',data:{requestId:id,payload:d}});
        return;
      }
    }catch(_){ }
    post({event,source:'reports-ui',data:{requestId:id,payload:d}});
  });
}
function runDb(p){const d=isObj(p)?Object.assign({},p):{};if(!d.requestId)d.requestId=reqId('docbuilder-run');const n=d&&d.argument&&Array.isArray(d.argument.actions)?d.argument.actions.length:1;const t=Math.max(120000,120000+n*4500);return requestDb('reportsDocBuilderRun',d,t);}
function probeDb(p){const d=isObj(p)?Object.assign({},p):{};if(!d.requestId)d.requestId=reqId('docbuilder-probe');return requestDb('reportsDocBuilderProbe',d,30000);}
function waitForPathExists(path,timeoutMs,pollMs){
  const target=normSlashes(path);
  return new Promise((resolve,reject)=>{
    try{
      const sdk=(window.parent&&window.parent.sdk)||(window.sdk||null);
      if(!bindDirectNative()||!sdk||typeof sdk.command!=='function'){
        const e=new Error('Native file bridge unavailable');
        e.details={error:'bridge_unavailable',path:target};
        reject(e);
        return;
      }
      const id=reqId('reports-path-check');
      const key=`reports_path_${String(id).replace(/[^a-zA-Z0-9_-]/g,'_')}`;
      const body={};body[key]=target;
      const tick=()=>{
        try{sdk.command('files:check',JSON.stringify(body));}catch(_){ }
      };
      PATH_CHECK_PENDING[id]={
        k:key,
        res:resolve,
        i:setInterval(tick,Math.max(150,toInt(pollMs,350)||350)),
        t:setTimeout(()=>{
          const p=clearPathCheck(id);
          if(!p)return;
          const e=new Error('File check timeout');
          e.details={error:'timeout',path:target};
          reject(e);
        },Math.max(500,toInt(timeoutMs,5000)||5000))
      };
      tick();
    }catch(err){
      reject(err);
    }
  });
}
async function ensureOutputExists(path,timeoutMs){
  const total=Math.max(1000,toInt(timeoutMs,8000)||8000);
  const started=Date.now();
  while(Date.now()-started<total){
    const remaining=Math.max(250,total-(Date.now()-started));
    try{
      await waitForPathExists(path,Math.min(remaining,1200),250);
      return true;
    }catch(_){ }
    if(await canReadBinaryFile(path,Math.min(remaining,1200)))return true;
    await new Promise(resolve=>setTimeout(resolve,200));
  }
  return false;
}
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
  const fullPath=normSlashes(path);
  let sentNative=false;
  try{
    const sdk=window.parent&&window.parent.sdk;
    if(bindDirectNative()&&sdk&&typeof sdk.command==='function'){
      const payload={requestId:reqId('docbuilder-open'),path:fullPath};
      sdk.command('docbuilder:open',JSON.stringify(payload));
      sentNative=true;
    }
  }catch(_){ }
  if(sentNative){
    // Подстраховка: часть сборок не отдает openResult, поэтому дублируем старый канал.
    setTimeout(()=>{
      const typeId=extType(fullPath);
      post({event:'reportsOpenFile',source:'reports-ui',data:{id:null,path:fullPath,typeId}});
    },220);
    return;
  }
  const typeId=extType(fullPath);
  post({event:'reportsOpenFile',source:'reports-ui',data:{id:null,path:fullPath,typeId}});
}
async function openTemplateInEditor(id){
  const sc=state.scenarios.find(x=>x.id===id);
  if(!sc||!sc.templatePath){toast('Шаблон не задан.','error');return;}
  try{
    const tpl=await ensureReadableTemplatePath(sc.templatePath);
    openResult(tpl,id);
  }catch(err){
    toast(err&&err.message?err.message:'Некорректный путь к шаблону.','error');
  }
}

function extractSourceProbePayload(errorLike){
  const texts=[];
  if(errorLike&&errorLike.details){
    if(errorLike.details.stderr)texts.push(String(errorLike.details.stderr));
    if(errorLike.details.stdout)texts.push(String(errorLike.details.stdout));
    if(errorLike.details.details)texts.push(String(errorLike.details.details));
  }
  if(errorLike&&errorLike.message)texts.push(String(errorLike.message));
  for(let i=0;i<texts.length;i+=1){
    const text=String(texts[i]||'');
    const idx=text.lastIndexOf(SOURCE_PROBE_PREFIX);
    if(idx<0)continue;
    const chunk=text.slice(idx+SOURCE_PROBE_PREFIX.length).split(/\r?\n/)[0].trim();
    if(!chunk)continue;
    try{return JSON.parse(chunk);}catch(_){ }
  }
  return null;
}
function sourceProbeErrors(payload){
  if(!payload)return [];
  const errors=Array.isArray(payload.errors)?payload.errors:[];
  return errors
    .map(item=>{
      if(isObj(item)&&item.message)return String(item.message).trim();
      return String(item||'').trim();
    })
    .filter(Boolean);
}
function sourceProbeWarnings(payload){
  if(!payload)return [];
  const warnings=Array.isArray(payload.warnings)?payload.warnings:[];
  return warnings
    .map(item=>{
      if(isObj(item)&&item.message)return String(item.message).trim();
      return String(item||'').trim();
    })
    .filter(Boolean);
}
function formatSourceProbeSummary(payload){
  if(!payload)return 'Проверка файла не выполнена.';
  if(payload.ok){
    const fields=Array.isArray(payload.fields)?payload.fields.length:0;
    const lists=Array.isArray(payload.lists)?payload.lists.length:0;
    const tables=Array.isArray(payload.tables)?payload.tables.length:0;
    const listItems=(Array.isArray(payload.lists)?payload.lists:[]).reduce((sum,list)=>sum+(parseInt(list&&list.itemsCount,10)||0),0);
    const rows=(Array.isArray(payload.tables)?payload.tables:[]).reduce((sum,table)=>sum+(parseInt(table&&table.rowsCount,10)||0),0);
    const parts=[`Поля: ${fields}`,`Списки: ${lists}`,`Таблицы: ${tables}`];
    if(listItems>0)parts.push(`Элементы списков: ${listItems}`);
    if(rows>0)parts.push(`Строки: ${rows}`);
    return `Файл проверен. ${parts.join('. ')}.`;
  }
  const errors=sourceProbeErrors(payload);
  return errors.length?errors[0]:'Файл не прошел проверку.';
}
function sourceProbeSchemaKey(scenario){
  return JSON.stringify({
    config:getScenarioSourceConfig(scenario)||{},
    schema:getScenarioSourceSchema(scenario)||{}
  });
}
function sourceProbeCacheKey(scenario,filePath,fileObject){
  const normalized=normSlashes(toFsPath(filePath)||'');
  const fileMeta=fileObject?{
    size:typeof fileObject.size==='number'?fileObject.size:'',
    lastModified:typeof fileObject.lastModified==='number'?fileObject.lastModified:''
  }:{size:'',lastModified:''};
  return JSON.stringify({
    scenarioId:String(scenario&&scenario.id||''),
    path:normalized,
    file:fileMeta,
    schema:sourceProbeSchemaKey(scenario)
  });
}
function getCachedSourceProbe(scenario,filePath,fileObject){
  const key=sourceProbeCacheKey(scenario,filePath,fileObject);
  const cached=SOURCE_PROBE_CACHE[key];
  if(!cached)return null;
  if(Date.now()-cached.ts>SOURCE_PROBE_CACHE_TTL){
    delete SOURCE_PROBE_CACHE[key];
    return null;
  }
  return cached.payload||null;
}
function setCachedSourceProbe(scenario,filePath,fileObject,payload){
  SOURCE_PROBE_CACHE[sourceProbeCacheKey(scenario,filePath,fileObject)]={ts:Date.now(),payload};
}
function renderRunSourcePreview(payload){
  if(!els.runSourcePreview)return;
  const root=els.runSourcePreview;
  root.innerHTML='';
  if(!payload||(!Array.isArray(payload.fields)&&!Array.isArray(payload.lists)&&!Array.isArray(payload.tables))){
    root.classList.add('hidden');
    return;
  }
  const warnings=sourceProbeWarnings(payload);
  if(warnings.length){
    const warningBox=document.createElement('div');
    warningBox.className='run-source-error';
    warningBox.textContent=warnings.join('\n');
    root.appendChild(warningBox);
  }
  const fields=Array.isArray(payload.fields)?payload.fields:[];
  if(fields.length){
    const grid=document.createElement('div');
    grid.className='run-source-preview-grid';
    fields.forEach(field=>{
      const card=document.createElement('div');
      card.className='run-source-preview-card';
      const label=document.createElement('div');
      label.className='run-source-preview-label';
      label.textContent=String(field&&field.label||field&&field.key||'Поле');
      const value=document.createElement('div');
      value.className='run-source-preview-value';
      const display=field&&field.displayValue!==undefined&&field.displayValue!==null?field.displayValue:field&&field.value;
      value.textContent=String(display==null?'':display) || '—';
      card.appendChild(label);
      card.appendChild(value);
      grid.appendChild(card);
    });
    root.appendChild(grid);
  }
  const lists=Array.isArray(payload.lists)?payload.lists:[];
  lists.forEach(list=>{
    const box=document.createElement('div');
    box.className='run-source-table';
    const head=document.createElement('div');
    head.className='run-source-table-head';
    const title=document.createElement('div');
    title.className='run-source-preview-label';
    title.textContent=String(list&&list.label||list&&list.key||'Список');
    const meta=document.createElement('div');
    meta.className='run-source-table-meta';
    const count=Math.max(0,parseInt(list&&list.itemsCount,10)||0);
    const shown=Array.isArray(list&&list.previewItems)?list.previewItems.length:0;
    meta.textContent=count>shown&&shown>0?`Найдено элементов: ${count}. Показано: ${shown}.`:`Найдено элементов: ${count}.`;
    head.appendChild(title);
    head.appendChild(meta);
    box.appendChild(head);
    const items=Array.isArray(list&&list.previewItems)?list.previewItems:[];
    if(items.length){
      const wrap=document.createElement('div');
      wrap.className='run-source-table-wrap';
      const tbl=document.createElement('table');
      const thead=document.createElement('thead');
      const hrow=document.createElement('tr');
      const th=document.createElement('th');
      th.textContent='Значение';
      hrow.appendChild(th);
      thead.appendChild(hrow);
      tbl.appendChild(thead);
      const tbody=document.createElement('tbody');
      items.forEach(item=>{
        const tr=document.createElement('tr');
        const td=document.createElement('td');
        td.textContent=String(item==null?'':item);
        tr.appendChild(td);
        tbody.appendChild(tr);
      });
      tbl.appendChild(tbody);
      wrap.appendChild(tbl);
      box.appendChild(wrap);
    }
    root.appendChild(box);
  });
  const tables=Array.isArray(payload.tables)?payload.tables:[];
  tables.forEach(table=>{
    const box=document.createElement('div');
    box.className='run-source-table';
    const head=document.createElement('div');
    head.className='run-source-table-head';
    const title=document.createElement('div');
    title.className='run-source-preview-label';
    title.textContent=String(table&&table.label||table&&table.key||'Таблица');
    const meta=document.createElement('div');
    meta.className='run-source-table-meta';
    const count=Math.max(0,parseInt(table&&table.rowsCount,10)||0);
    const shown=Array.isArray(table&&table.previewRows)?table.previewRows.length:0;
    meta.textContent=count>shown&&shown>0?`Найдено строк: ${count}. Показано: ${shown}.`:`Найдено строк: ${count}.`;
    head.appendChild(title);
    head.appendChild(meta);
    box.appendChild(head);
    const columns=Array.isArray(table&&table.columns)?table.columns:[];
    const rows=Array.isArray(table&&table.previewRows)?table.previewRows:[];
    if(columns.length&&rows.length){
      const wrap=document.createElement('div');
      wrap.className='run-source-table-wrap';
      const tbl=document.createElement('table');
      const thead=document.createElement('thead');
      const hrow=document.createElement('tr');
      columns.forEach(column=>{
        const th=document.createElement('th');
        th.textContent=String(column&&column.label||column&&column.key||'');
        hrow.appendChild(th);
      });
      thead.appendChild(hrow);
      tbl.appendChild(thead);
      const tbody=document.createElement('tbody');
      rows.forEach(row=>{
        const tr=document.createElement('tr');
        columns.forEach(column=>{
          const td=document.createElement('td');
          const key=String(column&&column.key||'').trim();
          const value=isObj(row)&&Object.prototype.hasOwnProperty.call(row,key)?row[key]:'';
          td.textContent=String(value==null?'':value);
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      tbl.appendChild(tbody);
      wrap.appendChild(tbl);
      box.appendChild(wrap);
    }
    root.appendChild(box);
  });
  root.classList.remove('hidden');
}
function syncRunSubmitState(){
  const ctx=state.runInput;
  if(!els.btnRunInputSubmit)return;
  if(!ctx){
    els.btnRunInputSubmit.disabled=false;
    return;
  }
  const sourceEnabled=!!ctx.sourceEnabled;
  const needsSource=sourceEnabled&&!!ctx.sourceRequired;
  const hasValidSource=!sourceEnabled||(!ctx.sourceBusy&&(!ctx.sourceFilePath?!needsSource:!!(ctx.sourceProbeResult&&ctx.sourceProbeResult.ok)));
  els.btnRunInputSubmit.disabled=!!ctx.sourceBusy||!hasValidSource;
}
function refreshRunSourceUi(){
  const ctx=state.runInput;
  if(!els.runSourceSection)return;
  if(!ctx||!ctx.sourceEnabled){
    els.runSourceSection.classList.add('hidden');
    if(els.runSourceFilePath)els.runSourceFilePath.value='';
    if(els.runSourceStatus){
      els.runSourceStatus.className='run-source-status is-idle';
    }
    if(els.runSourceStatusText)els.runSourceStatusText.textContent='Файл не выбран';
    if(els.runSourceError){
      els.runSourceError.textContent='';
      els.runSourceError.classList.add('hidden');
    }
    if(els.runSourcePreview){
      els.runSourcePreview.innerHTML='';
      els.runSourcePreview.classList.add('hidden');
    }
    syncRunSubmitState();
    return;
  }
  els.runSourceSection.classList.remove('hidden');
  if(els.runSourceSubtitle){
    const exts=(Array.isArray(ctx.sourceConfig&&ctx.sourceConfig.accept)?ctx.sourceConfig.accept:[]).map(ext=>`.${String(ext||'').replace(/^\./,'')}`);
    els.runSourceSubtitle.textContent=`Выберите файл ${exts.length?exts.join(', '):'.xlsx'} и дождитесь проверки структуры.`;
  }
  if(els.runSourceFilePath)els.runSourceFilePath.value=ctx.sourceFilePath||'';
  if(els.btnRunSourcePick)els.btnRunSourcePick.disabled=!!ctx.sourceBusy;
  if(els.runSourceStatus){
    els.runSourceStatus.className=`run-source-status ${ctx.sourceBusy?'is-busy':(ctx.sourceProbeResult&&ctx.sourceProbeResult.ok?'is-ok':(ctx.sourceError||ctx.sourceFilePath?'is-error':'is-idle'))}`.trim();
  }
  if(els.runSourceStatusText){
    if(ctx.sourceBusy)els.runSourceStatusText.textContent='Проверка структуры файла через DocumentBuilder...';
    else if(ctx.sourceProbeResult)els.runSourceStatusText.textContent=formatSourceProbeSummary(ctx.sourceProbeResult);
    else if(ctx.sourceError)els.runSourceStatusText.textContent=ctx.sourceError;
    else els.runSourceStatusText.textContent='Файл не выбран';
  }
  if(els.runSourceError){
    const errorText=ctx.sourceError||'';
    els.runSourceError.textContent=errorText;
    els.runSourceError.classList.toggle('hidden',!errorText);
  }
  renderRunSourcePreview(ctx.sourceProbeResult&&ctx.sourceProbeResult.ok?ctx.sourceProbeResult:null);
  syncRunSubmitState();
}
function isAcceptedSourceFile(path,config){
  const exts=Array.isArray(config&&config.accept)?config.accept:[];
  if(!exts.length)return true;
  const match=String(path||'').trim().toLowerCase().match(/\.([a-z0-9]+)$/);
  if(!match)return false;
  return exts.indexOf(match[1])>=0;
}
async function runSourceProbe(scenario,filePath){
  const sourceConfig=getScenarioSourceConfig(scenario);
  const probe=await ensureProbe(false);
  const timeoutMs=Math.max(5000,toInt(sourceConfig&&sourceConfig.probeTimeoutMs,45000)||45000);
  const resultPath=makeProbeResultPath();
  const wrapperPath=await materializeDocbuilderScript(SOURCE_PROBE_SCRIPT,{
    sourceFilePath:filePath,
    resultPath,
    sourceFileConfigJson:JSON.stringify(sourceConfig||{}),
    sourceSchemaJson:JSON.stringify(getScenarioSourceSchema(scenario)||{})
  },'probe',{
    SOURCE_FILE_PATH:filePath,
    RESULT_PATH:resultPath
  });
  const payload={
    requestId:reqId('docbuilder-source-probe'),
    script:wrapperPath,
    workDir:pickDocbuilderWorkDir(probe),
    outputPath:resultPath,
    openAfterRun:false,
    timeoutMs
  };
  try{
    const runPromise=runDb(payload);
    runPromise.catch(()=>{});
    const fileReadyPromise=waitProbeResultFile(resultPath,timeoutMs,200);
    return await Promise.race([
      fileReadyPromise,
      (async()=>{
        await runPromise;
        return await readProbeResultFile(resultPath,1500);
      })()
    ]);
  }catch(err){
    try{return await readProbeResultFile(resultPath,2500);}catch(_){ }
    const parsed=extractSourceProbePayload(err);
    if(parsed)return parsed;
    throw err;
  }finally{
    deletePath(resultPath,2500).catch(()=>{});
    deletePath(wrapperPath,2500).catch(()=>{});
  }
}
async function beginRunSourceProbe(filePath,fileObject){
  const ctx=state.runInput;
  if(!ctx||!ctx.sourceEnabled)return;
  const normalized=toFsPath(filePath);
  if(!normalized){
    ctx.sourceFilePath='';
    ctx.sourceProbeResult=null;
    ctx.sourceError='Не удалось определить путь к выбранному файлу.';
    ctx.sourceBusy=false;
    refreshRunSourceUi();
    return;
  }
  if(fileObject&&ctx.sourceConfig&&ctx.sourceConfig.maxFileSizeMb){
    const maxBytes=ctx.sourceConfig.maxFileSizeMb*1024*1024;
    if(fileObject.size>maxBytes){
      ctx.sourceFilePath=normalized;
      ctx.sourceProbeResult=null;
      ctx.sourceError=`Файл превышает лимит ${ctx.sourceConfig.maxFileSizeMb} MB.`;
      ctx.sourceBusy=false;
      refreshRunSourceUi();
      return;
    }
  }
  if(!isAcceptedSourceFile(normalized,ctx.sourceConfig)){
    ctx.sourceFilePath=normalized;
    ctx.sourceProbeResult=null;
    ctx.sourceError='Неподдерживаемый формат файла.';
    ctx.sourceBusy=false;
    refreshRunSourceUi();
    return;
  }
  ctx.sourceFilePath=normalized;
  const cached=getCachedSourceProbe(ctx.scenario,normalized,fileObject);
  if(cached){
    ctx.sourceProbeResult=cached;
    ctx.sourceError=cached&&cached.ok?'':(sourceProbeErrors(cached).join('\n')||formatSourceProbeSummary(cached));
    ctx.sourceBusy=false;
    refreshRunSourceUi();
    return;
  }
  ctx.sourceProbeResult=null;
  ctx.sourceError='';
  ctx.sourceBusy=true;
  ctx.sourceSeq=(ctx.sourceSeq||0)+1;
  const seq=ctx.sourceSeq;
  refreshRunSourceUi();
  try{
    const payload=await runSourceProbe(ctx.scenario,normalized);
    if(state.runInput!==ctx||ctx.sourceSeq!==seq)return;
    ctx.sourceProbeResult=payload;
    ctx.sourceError=payload&&payload.ok?'':(sourceProbeErrors(payload).join('\n')||formatSourceProbeSummary(payload));
    ctx.sourceBusy=false;
    if(payload&&payload.ok)setCachedSourceProbe(ctx.scenario,normalized,fileObject,payload);
    refreshRunSourceUi();
  }catch(err){
    if(state.runInput!==ctx||ctx.sourceSeq!==seq)return;
    ctx.sourceProbeResult=null;
    ctx.sourceBusy=false;
    ctx.sourceError=`Ошибка проверки: ${dbReason(err)}`;
    refreshRunSourceUi();
  }
}
function pickRunSourceFile(){
  const ctx=state.runInput;
  if(!ctx||!ctx.sourceEnabled)return;
  try{
    const sdk=window.parent&&window.parent.sdk;
    const sourcePath=ctx.sourceFilePath?dirName(ctx.sourceFilePath):getReportsUiDir();
    if(sdk&&typeof sdk.command==='function'&&sourcePath)sdk.command('files:setOpenPath',sourcePath);
  }catch(_){ }
  try{
    const asc=window.parent&&window.parent.AscDesktopEditor;
    if(asc&&typeof asc.OpenFilenameDialog==='function'){
      asc.OpenFilenameDialog('cell',false,function(files){
        const file=Array.isArray(files)?files[0]:files;
        if(!file)return;
        beginRunSourceProbe(String(file),null);
      });
      return;
    }
  }catch(_){ }
  if(els.runSourceFileInput){
    els.runSourceFileInput.value='';
    els.runSourceFileInput.click();
  }
}

function closeRunInputModal(result){
  const ctx=state.runInput;
  if(!ctx)return;
  state.runInput=null;
  if(els.runInputModal)els.runInputModal.classList.add('hidden');
  if(els.runInputFields)els.runInputFields.innerHTML='';
  if(els.runSourceFileInput)els.runSourceFileInput.value='';
  refreshRunSourceUi();
  renderList();
  if(typeof ctx.resolve==='function')ctx.resolve(result);
}

function closeListEditorModal(apply){
  const ctx=state.listEditor;
  if(!ctx)return;
  const values=optionsFromLines(els.listEditorText?els.listEditorText.value:'');
  state.listEditor=null;
  if(els.listEditorModal)els.listEditorModal.classList.add('hidden');
  if(els.listEditorText)els.listEditorText.value='';
  if(apply&&typeof ctx.onApply==='function')ctx.onApply(values);
}

function openListEditor(cfg){
  const settings=isObj(cfg)?cfg:{};
  const values=Array.isArray(settings.values)?settings.values:parseInputOptions(settings.values);
  state.listEditor={onApply:typeof settings.onApply==='function'?settings.onApply:null};
  if(els.listEditorTitle)els.listEditorTitle.textContent=String(settings.title||'Редактор списка');
  if(els.listEditorSubtitle)els.listEditorSubtitle.textContent=String(settings.subtitle||'Каждый вариант вводится с новой строки.');
  if(els.listEditorText)els.listEditorText.value=optionsToLines(values);
  if(els.listEditorModal)els.listEditorModal.classList.remove('hidden');
  if(els.listEditorText)setTimeout(()=>els.listEditorText.focus(),20);
}

function closeExportPackModal(){
  state.exportPack=null;
  if(els.exportPackModal)els.exportPackModal.classList.add('hidden');
  if(els.exportPackList)els.exportPackList.innerHTML='';
}
function renderExportPackList(){
  if(!els.exportPackList||!state.exportPack)return;
  els.exportPackList.innerHTML='';
  state.scenarios.forEach(sc=>{
    const row=document.createElement('div');
    row.className='export-pack-item';
    const label=document.createElement('label');
    const checkbox=document.createElement('input');
    checkbox.type='checkbox';
    checkbox.checked=!!state.exportPack.selected[sc.id];
    checkbox.onchange=(ev)=>{state.exportPack.selected[sc.id]=!!ev.target.checked;};
    const main=document.createElement('div');
    main.className='export-pack-item-main';
    const name=document.createElement('div');
    name.className='export-pack-item-name';
    name.textContent=sc.name||'Без названия';
    const meta=document.createElement('div');
    meta.className='export-pack-item-meta';
    meta.textContent=`Шаблон: ${baseName(resolveTemplate(sc.templatePath)||'')||'-'}`;
    main.appendChild(name);
    main.appendChild(meta);
    label.appendChild(checkbox);
    label.appendChild(main);
    row.appendChild(label);
    els.exportPackList.appendChild(row);
  });
}
function openExportPackModal(){
  if(!state.scenarios.length){
    toast('Нет сценариев для экспорта.','error');
    return;
  }
  const selected=Object.create(null);
  state.scenarios.forEach(sc=>{selected[sc.id]=true;});
  state.exportPack={selected};
  renderExportPackList();
  if(els.exportPackModal)els.exportPackModal.classList.remove('hidden');
}
function filePathToUrl(path){
  const p=normSlashes(path||'');
  if(!p)return '';
  let out=p.replace(/\\/g,'/');
  if(!/^\//.test(out))out=`/${out}`;
  return `file://${encodeURI(out)}`;
}
function bytesToBase64(bytes){
  const chunk=0x8000;
  let out='';
  for(let i=0;i<bytes.length;i+=chunk){
    const slice=bytes.subarray(i,Math.min(i+chunk,bytes.length));
    out+=String.fromCharCode.apply(null,slice);
  }
  return btoa(out);
}
async function readBase64ViaFetch(path){
  const url=filePathToUrl(path);
  if(!url)throw new Error('invalid_path');
  const res=await fetch(url);
  if(!res||!res.ok)throw new Error('fetch_failed');
  const ab=await res.arrayBuffer();
  return bytesToBase64(new Uint8Array(ab));
}
async function readFileAsBase64(path){
  const req={path:String(path||''),encoding:'base64'};
  const res=await requestNativeFile('reports:fileRead',req,8000).catch(()=>null);
  if(res&&res.ok&&String(res.encoding||'').toLowerCase()==='base64')return String(res.content||'');
  return readBase64ViaFetch(path);
}
async function writeFileFromBase64(path,base64){
  const req={path:String(path||''),encoding:'base64',content:String(base64||'')};
  const res=await requestNativeFile('reports:fileWrite',req,10000);
  if(res&&res.ok&&String(res.encoding||'').toLowerCase()==='base64')return true;
  throw new Error('binary_write_not_supported');
}
function downloadTextFile(fileName,content){
  const blob=new Blob([String(content||'')],{type:'application/json;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;
  a.download=fileName;
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},500);
}
function ensureTemplateFileName(name,used){
  let base=sanitizeTemplateName(name||'Шаблон');
  let candidate=`${base}.xlsx`;
  let i=2;
  while(used[candidate.toLowerCase()]){
    candidate=`${base} (${i}).xlsx`;
    i+=1;
  }
  used[candidate.toLowerCase()]=true;
  return candidate;
}
async function submitExportPack(){
  if(!state.exportPack)return;
  const selectedIds=Object.keys(state.exportPack.selected).filter(id=>state.exportPack.selected[id]);
  if(!selectedIds.length){
    toast('Выберите хотя бы один сценарий.','error');
    return;
  }
  try{
    if(els.btnExportPackSubmit){els.btnExportPackSubmit.disabled=true;els.btnExportPackSubmit.textContent='Экспорт...';}
    const usedTemplateNames=Object.create(null);
    const items=[];
    for(let i=0;i<selectedIds.length;i+=1){
      const sc=state.scenarios.find(x=>x.id===selectedIds[i]);
      if(!sc)continue;
      const templatePath=resolveTemplate(sc.templatePath);
      if(!templatePath)throw new Error(`Сценарий «${sc.name}»: не найден путь к шаблону.`);
      const templateName=ensureTemplateFileName(baseName(templatePath)||sc.name,usedTemplateNames);
      const templateBase64=await readFileAsBase64(templatePath);
      const scenarioFile=state.scenarioFiles[sc.id]||makeScenarioFileName(sc.name);
      items.push({
        scenarioFile,
        templateFile:templateName,
        scenario:clone(sc),
        templateBase64
      });
    }
    const pack={
      kind:'reports-ui-pack',
      version:1,
      exportedAt:new Date().toISOString(),
      items
    };
    downloadTextFile(`reports_pack_${fmtStamp()}.reportspack.json`,JSON.stringify(pack,null,2));
    closeExportPackModal();
    toast('Экспорт выполнен.','ok');
  }catch(err){
    const msg=err&&err.message?err.message:'Ошибка экспорта';
    toast(`Экспорт не выполнен: ${msg}`,'error');
  }finally{
    if(els.btnExportPackSubmit){els.btnExportPackSubmit.disabled=false;els.btnExportPackSubmit.textContent='Готово';}
  }
}
async function importPackFromText(rawText){
  const pack=parse(rawText);
  if(!isObj(pack)||String(pack.kind||'')!=='reports-ui-pack'||!Array.isArray(pack.items))throw new Error('Неверный формат архива.');
  const usedScenarioFileNames=Object.create(null);
  Object.keys(state.scenarioFiles).forEach(id=>{const file=String(state.scenarioFiles[id]||'').trim();if(file)usedScenarioFileNames[file.toLowerCase()]=true;});
  const usedTemplateNames=Object.create(null);
  state.scenarios.forEach(sc=>{const tpl=baseName(resolveTemplate(sc.templatePath)||'').toLowerCase();if(tpl)usedTemplateNames[tpl]=true;});
  const pendingScenarios=[];
  for(let i=0;i<pack.items.length;i+=1){
    const item=pack.items[i];
    if(!isObj(item)||!isObj(item.scenario))continue;
    const sc=normScenario(item.scenario);
    if(!sc)continue;
    const templateFile=ensureTemplateFileName(item.templateFile||sc.name,usedTemplateNames);
    const templatePath=joinPath(getTemplatesDir(),templateFile);
    await writeFileFromBase64(templatePath,String(item.templateBase64||''));
    sc.templatePath=templatePath;
    const scenarioFile=ensureScenarioFileNameUniq(item.scenarioFile||makeScenarioFileName(sc.name),usedScenarioFileNames);
    pendingScenarios.push({scenario:sc,scenarioFile});
  }
  pendingScenarios.forEach(entry=>{
    const idx=state.scenarios.findIndex(x=>x.id===entry.scenario.id);
    if(idx>=0)state.scenarios[idx]=entry.scenario;
    else state.scenarios.push(entry.scenario);
    state.scenarioFiles[entry.scenario.id]=entry.scenarioFile;
  });
  save();
  renderList();
}
async function onImportPackPicked(file){
  if(!file)return;
  const text=await file.text();
  await importPackFromText(text);
}

function createRuntimeInputControl(field,initialValue,selectOptions){
  const type=normalizeInputType(field.input_type);
  if(type==='multiline'){
    const ta=document.createElement('textarea');
    ta.rows=3;
    ta.value=String(initialValue==null?'':initialValue);
    ta.placeholder=field.placeholder||'Введите значение';
    return {
      node:ta,
      getValue:()=>String(ta.value==null?'':ta.value),
      focus:()=>ta.focus(),
      onChange:(fn)=>{if(typeof fn==='function')ta.addEventListener('input',()=>fn(String(ta.value==null?'':ta.value)));},
      setDisabled:(flag)=>{ta.disabled=!!flag;}
    };
  }
  if(type==='select'){
    const sel=document.createElement('select');
    const fillOptions=(options,preferValue,keepCurrent)=>{
      const current=String(sel.value==null?'':sel.value);
      const list=[];
      const seen=Object.create(null);
      (Array.isArray(options)?options:[]).forEach(item=>{
        const value=String(item==null?'':item).trim();
        if(!value)return;
        const key=value.toLowerCase();
        if(seen[key])return;
        seen[key]=true;
        list.push(value);
      });
      while(sel.firstChild)sel.removeChild(sel.firstChild);
      if(!field.required){
        const empty=document.createElement('option');
        empty.value='';
        empty.textContent='(пусто)';
        sel.appendChild(empty);
      }
      list.forEach(opt=>{
        const o=document.createElement('option');
        o.value=opt;
        o.textContent=opt;
        sel.appendChild(o);
      });
      const wanted=String(preferValue==null?'':preferValue).trim();
      if(wanted&&list.indexOf(wanted)<0){
        const extra=document.createElement('option');
        extra.value=wanted;
        extra.textContent=wanted;
        sel.insertBefore(extra,sel.firstChild);
      }
      const nextCurrent=keepCurrent?current:'';
      if(wanted){
        sel.value=wanted;
      }else if(nextCurrent&&list.indexOf(nextCurrent)>=0){
        sel.value=nextCurrent;
      }else if(!field.required){
        sel.value='';
      }else if(list.length){
        sel.value=list[0];
      }else{
        sel.value='';
      }
    };
    const initialOptions=Array.isArray(selectOptions)?selectOptions:resolveSelectOptions(field,'');
    fillOptions(initialOptions,initialValue,false);
    return {
      node:sel,
      getValue:()=>String(sel.value==null?'':sel.value),
      focus:()=>sel.focus(),
      setOptions:(options)=>fillOptions(options,'',true),
      setDisabled:(flag)=>{sel.disabled=!!flag;},
      onChange:(fn)=>{if(typeof fn==='function')sel.addEventListener('change',()=>fn(String(sel.value==null?'':sel.value)));}
    };
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
    return {
      node:wrap,
      getValue:()=>!!chk.checked,
      focus:()=>chk.focus(),
      onChange:(fn)=>{if(typeof fn==='function')chk.addEventListener('change',()=>fn(!!chk.checked));},
      setDisabled:(flag)=>{chk.disabled=!!flag;}
    };
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
  return {
    node:inp,
    getValue:()=>inp.value,
    focus:()=>inp.focus(),
    onChange:(fn)=>{if(typeof fn==='function')inp.addEventListener('input',()=>fn(inp.value));},
    setDisabled:(flag)=>{inp.disabled=!!flag;}
  };
}

function createRuntimeFieldControl(field,selectOptions){
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
      const ctrl=createRuntimeInputControl(field,value,selectOptions);
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
      focus:()=>{if(items.length&&items[0].ctrl&&typeof items[0].ctrl.focus==='function')items[0].ctrl.focus();},
      setDisabled:(flag)=>{
        addBtn.disabled=!!flag;
        items.forEach(item=>{
          if(item&&item.ctrl&&typeof item.ctrl.setDisabled==='function')item.ctrl.setDisabled(!!flag);
          if(item&&item.delBtn)item.delBtn.disabled=!!flag;
        });
      }
    };
  }
  const ctrl=createRuntimeInputControl(field,normalizeScalarFieldValue(field.default_value,field.input_type),selectOptions);
  return {
    node:ctrl.node,
    getValue:()=>normalizeScalarFieldValue(ctrl.getValue(),field.input_type),
    focus:ctrl.focus,
    onChange:ctrl.onChange,
    setOptions:ctrl.setOptions,
    setDisabled:ctrl.setDisabled
  };
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
  if(ctx.sourceEnabled){
    if(ctx.sourceBusy){
      toast('Дождитесь завершения проверки файла.','error');
      return;
    }
    if(ctx.sourceRequired&&!ctx.sourceFilePath){
      toast('Выберите файл с исходными данными.','error');
      return;
    }
    if(ctx.sourceFilePath&&(!ctx.sourceProbeResult||!ctx.sourceProbeResult.ok)){
      const errors=sourceProbeErrors(ctx.sourceProbeResult);
      toast(errors[0]||ctx.sourceError||'Файл с исходными данными не прошел проверку.','error');
      return;
    }
  }
  closeRunInputModal({inputValues:values,sourceFilePath:ctx.sourceFilePath||'',sourceProbeResult:ctx.sourceProbeResult||null});
}

function openRunInputModal(sc){
  const fields=(Array.isArray(sc&&sc.inputFields)?sc.inputFields:[]).map(normInputField);
  const sourceConfig=getScenarioSourceConfig(sc);
  const sourceEnabled=!!sourceConfig.enabled;
  if(!fields.length&&!sourceEnabled)return Promise.resolve({inputValues:{},sourceFilePath:'',sourceProbeResult:null});
  return new Promise(resolve=>{
    const inputs=Object.create(null);
    const fieldsById=Object.create(null);
    const childrenByParent=Object.create(null);
    const sourceRequired=sourceEnabled&&(!!sourceConfig.required||scenarioUsesSourceActions(sc));
    fields.forEach(field=>{
      fieldsById[field.id]=field;
      const parentId=String(field.depends_on||'').trim();
      if(field.input_type==='select'&&parentId){
        if(!Array.isArray(childrenByParent[parentId]))childrenByParent[parentId]=[];
        childrenByParent[parentId].push(field.id);
      }
    });
    state.runInput={
      resolve,
      scenarioId:sc.id,
      scenario:sc,
      fields,
      inputs,
      sourceEnabled,
      sourceRequired,
      sourceConfig,
      sourceFilePath:'',
      sourceProbeResult:null,
      sourceError:'',
      sourceBusy:false,
      sourceSeq:0
    };
    renderList();
    if(els.runInputTitle)els.runInputTitle.textContent=`Заполнение: ${sc.name}`;
    if(els.runInputFields){
      els.runInputFields.innerHTML='';
      fields.forEach((field,i)=>{
        const wrap=document.createElement('div');
        wrap.className='run-input-row';
        const cap=document.createElement('label');
        cap.className='run-input-label';
        cap.textContent=`${inputFieldLabel(field,i)}:`;
        const parentId=String(field.depends_on||'').trim();
        let selectOptions;
        if(field.input_type==='select'&&parentId){
          const parentField=fieldsById[parentId];
          const parentInitial=parentField?normalizeScalarFieldValue(parentField.default_value,parentField.input_type):'';
          selectOptions=resolveSelectOptions(field,parentInitial);
        }
        const ctrl=createRuntimeFieldControl(field,selectOptions);
        inputs[field.id]=ctrl;
        const ctrlWrap=document.createElement('div');
        ctrlWrap.className='run-input-control';
        wrap.appendChild(cap);
        ctrlWrap.appendChild(ctrl.node);
        wrap.appendChild(ctrlWrap);
        els.runInputFields.appendChild(wrap);
      });
      if(!fields.length){
        const empty=document.createElement('div');
        empty.className='source-empty';
        empty.textContent='Ручные поля для ввода не настроены.';
        els.runInputFields.appendChild(empty);
      }
    }

    const refreshDependentField=(fieldId,stack)=>{
      const field=fieldsById[fieldId];
      if(!field)return;
      const parentId=String(field.depends_on||'').trim();
      if(!parentId)return;
      if(stack&&stack[fieldId])return;
      const nextStack=stack||Object.create(null);
      nextStack[fieldId]=true;
      const parentCtrl=inputs[parentId];
      const ctrl=inputs[fieldId];
      if(parentCtrl&&ctrl&&typeof ctrl.setOptions==='function'){
        const parentValue=parentCtrl.getValue();
        const options=resolveSelectOptions(field,parentValue);
        ctrl.setOptions(options);
        if(typeof ctrl.setDisabled==='function')ctrl.setDisabled(options.length===0);
      }
      const children=childrenByParent[fieldId];
      if(Array.isArray(children)){
        children.forEach(childId=>refreshDependentField(childId,nextStack));
      }
      delete nextStack[fieldId];
    };
    Object.keys(childrenByParent).forEach(parentId=>{
      const parentCtrl=inputs[parentId];
      if(!parentCtrl||typeof parentCtrl.onChange!=='function')return;
      parentCtrl.onChange(()=>{
        const children=childrenByParent[parentId];
        if(!Array.isArray(children))return;
        children.forEach(childId=>refreshDependentField(childId,Object.create(null)));
      });
      const children=childrenByParent[parentId];
      if(Array.isArray(children)){
        children.forEach(childId=>refreshDependentField(childId,Object.create(null)));
      }
    });

    refreshRunSourceUi();
    if(els.runInputModal)els.runInputModal.classList.remove('hidden');
    if(sourceEnabled){
      ensureProbe(false).catch(()=>{});
    }
    const first=fields.length&&inputs[fields[0].id];
    if(first&&typeof first.focus==='function')setTimeout(()=>first.focus(),10);
    else if(sourceEnabled&&els.btnRunSourcePick)setTimeout(()=>els.btnRunSourcePick.focus(),10);
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

function sourceProbeLookup(items,key){
  const list=Array.isArray(items)?items:[];
  const wanted=String(key||'').trim().toLowerCase();
  for(let i=0;i<list.length;i+=1){
    const item=list[i];
    const itemKey=String(item&&item.key||'').trim().toLowerCase();
    if(itemKey&&itemKey===wanted)return item;
  }
  return null;
}
function sourceValueModeJs(type,value){
  const normalized=String(type||'text').trim().toLowerCase();
  if(normalized==='number')return (value===''||value===null||value===undefined||Number.isNaN(Number(value)))?'text':'number';
  if(normalized==='date'&&typeof value==='number')return 'number';
  return 'text';
}
function makeRuntimeSetCellAction(base,sheet,range,value,mode){
  const action={
    id:uid('action'),
    type:'set_cell_value',
    sheet:String(sheet||'').trim(),
    range:String(range||'').trim(),
    named_range:false,
    mode:String(mode||'text'),
    value,
    merge:false,
    keep_template_format:!!(base&&base.keep_template_format),
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
  normalizeInsertStyleSettings(action,action);
  return action;
}
function makeRuntimeInsertRowsAction(sheet,row){
  return {
    id:uid('action'),
    type:'insert_rows',
    sheet:String(sheet||'').trim(),
    start_row:String(row),
    count:'1'
  };
}
function makeRuntimeCopyTemplateRowAction(sheet,templateRange,targetRow){
  const parsed=parseA1Range(templateRange);
  if(!parsed)return null;
  return {
    id:uid('action'),
    type:'copy_template_row',
    from_sheet:String(sheet||'').trim(),
    from_range:String(templateRange||'').trim(),
    to_sheet:String(sheet||'').trim(),
    to_cell:`${numberToColLetters(parsed.c1)}${targetRow}`,
    paste_type:'xlPasteAll',
    operation:'xlPasteSpecialOperationNone',
    skip_blanks:false,
    transpose:false
  };
}
function normalizeSheetInsertKey(sheet){
  return String(sheet||'').trim().toLowerCase();
}
function addSheetInsertEvent(events,sheet,row,count){
  const key=normalizeSheetInsertKey(sheet);
  if(!key)return;
  if(!Array.isArray(events[key]))events[key]=[];
  events[key].push({
    row:Math.max(1,parseInt(row,10)||1),
    count:Math.max(1,parseInt(count,10)||1)
  });
  events[key].sort((a,b)=>a.row-b.row);
}
function shiftRowByInsertEvents(events,sheet,row){
  let current=Math.max(1,parseInt(row,10)||1);
  const list=events[normalizeSheetInsertKey(sheet)];
  if(!Array.isArray(list)||!list.length)return current;
  for(let i=0;i<list.length;i+=1){
    if(current>=list[i].row)current+=list[i].count;
  }
  return current;
}
function shiftRangeByInsertEvents(events,sheet,raw){
  const parsed=parseA1Range(raw);
  if(!parsed)return null;
  return {
    c1:parsed.c1,
    c2:parsed.c2,
    r1:shiftRowByInsertEvents(events,sheet,parsed.r1),
    r2:shiftRowByInsertEvents(events,sheet,parsed.r2)
  };
}
function expandSourceActions(actions,sourceProbeResult){
  const inputActions=Array.isArray(actions)?actions:[];
  const payload=isObj(sourceProbeResult)?sourceProbeResult:null;
  const fields=Array.isArray(payload&&payload.fields)?payload.fields:[];
  const lists=Array.isArray(payload&&payload.lists)?payload.lists:[];
  const tables=Array.isArray(payload&&payload.tables)?payload.tables:[];
  const out=[];
  const insertEvents=Object.create(null);
  for(let i=0;i<inputActions.length;i+=1){
    const action=normAction(inputActions[i]);
    if(action.type==='insert_source_field'){
      const sourceField=sourceProbeLookup(fields,action.source_field_id);
      if(!sourceField)throw new Error(`Источник не содержит поле "${String(action.source_field_id||'').trim()}".`);
      const parsedTargets=parseCellBindings(action.targets);
      const targets=parsedTargets.error?[]:parsedTargets.items;
      const bindings=targets.length?targets:[{sheet:String(action.sheet||'Лист1').trim(),range:String(action.range||'A1').trim()}];
      bindings.forEach(target=>{
        const shiftedRange=shiftRangeByInsertEvents(insertEvents,target.sheet,target.range);
        const range=shiftedRange?formatA1Range(shiftedRange):String(target.range||'').trim();
        out.push(makeRuntimeSetCellAction(action,target.sheet,range,sourceField.value,sourceValueModeJs(sourceField.type,sourceField.value)));
      });
      continue;
    }
    if(action.type==='insert_source_list'){
      const sourceList=sourceProbeLookup(lists,action.source_list_id);
      if(!sourceList)throw new Error(`Источник не содержит список "${String(action.source_list_id||'').trim()}".`);
      const parsedTargets=parseTableTargets(action.targets);
      const targets=parsedTargets.error?[]:parsedTargets.items;
      const bindings=targets.length?targets:[{
        sheet:String(action.sheet||'Лист1').trim(),
        anchor_cell:String(action.anchor_cell||'A1').trim(),
        template_row_range:String(action.template_row_range||'A1:D1').trim()
      }];
      const items=Array.isArray(sourceList.items)?sourceList.items:[];
      bindings.forEach(binding=>{
        const anchor=shiftRangeByInsertEvents(insertEvents,binding.sheet,binding.anchor_cell);
        if(!anchor)throw new Error(`Неверная точка вставки списка: ${binding.anchor_cell}`);
        const insertRows=String(action.insert_mode||'insert_rows').trim().toLowerCase()==='insert_rows';
        const shiftedTemplateRange=shiftRangeByInsertEvents(insertEvents,binding.sheet,binding.template_row_range);
        const templateRowRange=shiftedTemplateRange?formatA1Range(shiftedTemplateRange):String(binding.template_row_range||'').trim();
        for(let itemIndex=0;itemIndex<items.length;itemIndex+=1){
          const targetRow=anchor.r1+itemIndex;
          if(insertRows&&itemIndex>0){
            out.push(makeRuntimeInsertRowsAction(binding.sheet,targetRow));
            addSheetInsertEvent(insertEvents,binding.sheet,targetRow,1);
          }
          if(action.keep_template_format&&(itemIndex>0||!insertRows)){
            const copyAction=makeRuntimeCopyTemplateRowAction(binding.sheet,templateRowRange,targetRow);
            if(copyAction)out.push(copyAction);
          }
          out.push(makeRuntimeSetCellAction(action,binding.sheet,`${numberToColLetters(anchor.c1)}${targetRow}`,items[itemIndex],sourceValueModeJs(sourceList.type,items[itemIndex])));
        }
      });
      continue;
    }
    if(action.type==='insert_source_table'){
      const sourceTable=sourceProbeLookup(tables,action.source_table_id);
      if(!sourceTable)throw new Error(`Источник не содержит таблицу "${String(action.source_table_id||'').trim()}".`);
      const parsedTargets=parseTableTargets(action.targets);
      const targets=parsedTargets.error?[]:parsedTargets.items;
      const bindings=targets.length?targets:[{
        sheet:String(action.sheet||'Лист1').trim(),
        anchor_cell:String(action.anchor_cell||'A1').trim(),
        template_row_range:String(action.template_row_range||'A1:D1').trim()
      }];
      const parsedMappings=parseActionSourceMappings(action.mappings);
      if(parsedMappings.error)throw new Error(parsedMappings.error);
      const columns=Array.isArray(sourceTable.columns)?sourceTable.columns:[];
      const columnTypes=Object.create(null);
      columns.forEach(column=>{columnTypes[String(column&&column.key||'').trim()]=String(column&&column.type||'text').trim().toLowerCase()||'text';});
      const rows=Array.isArray(sourceTable.rows)?sourceTable.rows:[];
      bindings.forEach(binding=>{
        const anchor=shiftRangeByInsertEvents(insertEvents,binding.sheet,binding.anchor_cell);
        if(!anchor)throw new Error(`Неверная точка вставки таблицы: ${binding.anchor_cell}`);
        const insertRows=String(action.insert_mode||'insert_rows').trim().toLowerCase()==='insert_rows';
        const shiftedTemplateRange=shiftRangeByInsertEvents(insertEvents,binding.sheet,binding.template_row_range);
        const templateRowRange=shiftedTemplateRange?formatA1Range(shiftedTemplateRange):String(binding.template_row_range||'').trim();
        for(let rowIndex=0;rowIndex<rows.length;rowIndex+=1){
          const targetRow=anchor.r1+rowIndex;
          if(insertRows&&rowIndex>0){
            out.push(makeRuntimeInsertRowsAction(binding.sheet,targetRow));
            addSheetInsertEvent(insertEvents,binding.sheet,targetRow,1);
          }
          if(action.keep_template_format&&(rowIndex>0||!insertRows)){
            const copyAction=makeRuntimeCopyTemplateRowAction(binding.sheet,templateRowRange,targetRow);
            if(copyAction)out.push(copyAction);
          }
          const row=rows[rowIndex]&&typeof rows[rowIndex]==='object'?rows[rowIndex]:{};
          parsedMappings.items.forEach(mapping=>{
            const cellValue=row[mapping.sourceKey];
            out.push(makeRuntimeSetCellAction(action,binding.sheet,`${mapping.targetColumn}${targetRow}`,cellValue,sourceValueModeJs(columnTypes[mapping.sourceKey],cellValue)));
          });
        }
      });
      continue;
    }
    out.push(action);
  }
  return out;
}

function buildRunActions(sc,inputValues,sourceProbeResult){
  const base=(Array.isArray(sc&&sc.actions)?sc.actions.map(normAction):[]).filter(action=>evaluateActionCondition(action,inputValues,sc));
  const extra=buildInputActions(sc,inputValues);
  return expandSourceActions(extra.concat(base),sourceProbeResult);
}

async function runScenario(id){
  const sc=state.scenarios.find(x=>x.id===id);if(!sc||state.running||state.runInput)return;
  let tpl='';
  try{tpl=await ensureReadableTemplatePath(sc.templatePath);}catch(err){toast(err&&err.message?err.message:'Некорректный путь к шаблону.','error');return;}
  const runInput=await openRunInputModal(sc);
  if(runInput===null)return;
  const inputValues=isObj(runInput)&&isObj(runInput.inputValues)?runInput.inputValues:{};
  const sourceProbeResult=isObj(runInput)&&isObj(runInput.sourceProbeResult)?runInput.sourceProbeResult:null;
  let wrapperPath='';
  state.running=id;renderList();
  try{
    const runActions=buildRunActions(sc,inputValues,sourceProbeResult);
    toast('Проверка DocumentBuilder...','ok');
    const probe=await ensureProbe(false);
    const out=mkOutputPath(sc,probe,tpl);
    toast('Выполнение сценария через DocumentBuilder...','ok');
    const runRequestId=reqId('docbuilder-run');
    wrapperPath=await materializeDocbuilderScript(SCRIPT,{
      templatePath:tpl,
      outputPath:out,
      scenarioId:sc.id,
      scenarioName:sc.name,
      stopOnError:true,
      actions:runActions
    },'run',{
      TEMPLATE_PATH:tpl,
      OUTPUT_PATH:out
    });
    const runPayload={requestId:runRequestId,script:wrapperPath,workDir:pickDocbuilderWorkDir(probe),outputPath:out,openAfterRun:false};
    const runPromise=runDb(runPayload);
    runPromise.catch(()=>{});
    let alreadyOpened=false;
    const settled=await Promise.race([
      runPromise.then(res=>({kind:'bridge',res})).catch(err=>({kind:'bridge_error',err})),
      ensureOutputExists(out,90000).then(ok=>ok?{kind:'file'}:{kind:'run_timeout'})
    ]);
    if(settled.kind==='bridge'){
      const res=settled.res;
      if(!res||!res.ok){
        const outputReadyAfterBridgeFail=await ensureOutputExists(out,3000);
        if(!outputReadyAfterBridgeFail)throw dbErr(res||{error:'run_failed'});
        clearPending(runRequestId);
      }else{
        if(typeof res.exitCode!=='number')throw dbErr({error:'invalid_run_response',details:'DocumentBuilder returned unexpected response payload.'});
        alreadyOpened=!!res.opened;
      }
    }else if(settled.kind==='bridge_error'){
      const outputReadyAfterError=await ensureOutputExists(out,3000);
      if(!outputReadyAfterError)throw settled.err;
      clearPending(runRequestId);
    }else if(settled.kind==='file'){
      clearPending(runRequestId);
    }else{
      const outputReadyAfterTimeout=await ensureOutputExists(out,5000);
      if(!outputReadyAfterTimeout)throw dbErr({error:'run_timeout',details:`DocumentBuilder не завершил генерацию за разумное время: ${out}`});
      clearPending(runRequestId);
    }
    const outputReady=await ensureOutputExists(out,4000);
    if(!outputReady)throw dbErr({error:'output_not_created',details:`DocumentBuilder did not create the output file: ${out}`});
    const finalReady=await ensureOutputExists(out,5000);
    if(!finalReady)throw dbErr({error:'output_not_created',details:`Итоговый файл не найден: ${out}`});
    const i=state.scenarios.findIndex(x=>x.id===id);if(i>=0){state.scenarios[i].lastRunAt=new Date().toISOString();state.scenarios[i].updatedAt=new Date().toISOString();save();}
    state.running=null;
    renderList();
    if(!alreadyOpened){
      setTimeout(()=>{try{openResult(out,id);}catch(_){ }},800);
    }
    toast(`Готово: ${out}`,'ok');
  }catch(err){
    const d=err&&err.details?err.details:null;
    const reason=d&&d.stderr?d.stderr:(d&&d.error?d.error:(err&&err.message?err.message:'Неизвестная ошибка'));
    toast(`Ошибка выполнения: ${reason}`,'error');
    state.probe=null;
  }finally{
    if(wrapperPath)deletePath(wrapperPath,2500).catch(()=>{});
    state.running=null;renderList();
  }
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
  els.btnExportPack&&els.btnExportPack.addEventListener('click',openExportPackModal);
  els.btnImportPack&&els.btnImportPack.addEventListener('click',()=>{if(els.importPackInput){els.importPackInput.value='';els.importPackInput.click();}});
  els.importPackInput&&els.importPackInput.addEventListener('change',async()=>{
    const f=els.importPackInput.files&&els.importPackInput.files[0];
    if(!f)return;
    try{
      toast('Импорт сценариев...','ok');
      await onImportPackPicked(f);
      toast('Импорт выполнен.','ok');
    }catch(err){
      const msg=err&&err.message?err.message:'Ошибка импорта';
      toast(`Импорт не выполнен: ${msg}`,'error');
    }finally{
      if(els.importPackInput)els.importPackInput.value='';
    }
  });
  els.btnClose&&els.btnClose.addEventListener('click',closeModal);
  els.btnCancel&&els.btnCancel.addEventListener('click',closeModal);
  els.btnSave&&els.btnSave.addEventListener('click',saveDraft);
  const addInputField=()=>{if(!state.draft)return;state.draft.inputFields.push(mkInputField());renderInputFields();};
  const addAction=()=>{if(!state.draft)return;state.draft.actions.push(mkAction('set_cell_value'));renderActions();};
  const addSourceField=()=>{if(!state.draft)return;state.draft.sourceSchema.fields.push(mkSourceField());renderSourceConfig();renderActions();};
  const addSourceList=()=>{if(!state.draft)return;state.draft.sourceSchema.lists.push(mkSourceList());renderSourceConfig();renderActions();};
  const addSourceTable=()=>{if(!state.draft)return;state.draft.sourceSchema.tables.push(mkSourceTable());renderSourceConfig();renderActions();};
  els.btnAddInputField&&els.btnAddInputField.addEventListener('click',addInputField);
  els.btnAddInputFieldBottom&&els.btnAddInputFieldBottom.addEventListener('click',addInputField);
  els.btnAddAction&&els.btnAddAction.addEventListener('click',addAction);
  els.btnAddActionBottom&&els.btnAddActionBottom.addEventListener('click',addAction);
  els.btnAddSourceField&&els.btnAddSourceField.addEventListener('click',addSourceField);
  els.btnAddSourceList&&els.btnAddSourceList.addEventListener('click',addSourceList);
  els.btnAddSourceTable&&els.btnAddSourceTable.addEventListener('click',addSourceTable);
  els.btnToggleInputSection&&els.btnToggleInputSection.addEventListener('click',()=>toggleSection('input'));
  els.btnToggleActionSection&&els.btnToggleActionSection.addEventListener('click',()=>toggleSection('action'));
  els.btnPick&&els.btnPick.addEventListener('click',pickTemplate);
  els.btnRunInputCancel&&els.btnRunInputCancel.addEventListener('click',()=>closeRunInputModal(null));
  els.btnRunInputSubmit&&els.btnRunInputSubmit.addEventListener('click',submitRunInputModal);
  els.btnRunSourcePick&&els.btnRunSourcePick.addEventListener('click',pickRunSourceFile);
  els.runSourceFileInput&&els.runSourceFileInput.addEventListener('change',()=>{
    const file=els.runSourceFileInput.files&&els.runSourceFileInput.files[0];
    if(!file)return;
    if(file.path){
      beginRunSourceProbe(file.path,file);
      return;
    }
    toast('Не удалось получить путь к выбранному файлу.','error');
  });
  els.btnListEditorClose&&els.btnListEditorClose.addEventListener('click',()=>closeListEditorModal(false));
  els.btnListEditorCancel&&els.btnListEditorCancel.addEventListener('click',()=>closeListEditorModal(false));
  els.btnListEditorApply&&els.btnListEditorApply.addEventListener('click',()=>closeListEditorModal(true));
  els.btnExportPackClose&&els.btnExportPackClose.addEventListener('click',closeExportPackModal);
  els.btnExportPackCancel&&els.btnExportPackCancel.addEventListener('click',closeExportPackModal);
  els.btnExportPackSubmit&&els.btnExportPackSubmit.addEventListener('click',submitExportPack);
  els.file&&els.file.addEventListener('change',()=>{const f=els.file.files&&els.file.files[0];if(f&&f.path){els.tpl.value=toFsPath(f.path);if(state.draft)state.draft.templatePath=els.tpl.value;return;}const v=els.file.value||'';if(v&&!/fakepath/i.test(v)){els.tpl.value=toFsPath(v);if(state.draft)state.draft.templatePath=els.tpl.value;return;}toast('Введите полный путь к шаблону вручную.','error');});
  els.modal&&els.modal.addEventListener('click',e=>{if(e.target&&e.target.classList&&e.target.classList.contains('modal-backdrop'))closeModal();});
  els.runInputModal&&els.runInputModal.addEventListener('click',e=>{if(e.target&&e.target.classList&&e.target.classList.contains('modal-backdrop'))closeRunInputModal(null);});
  els.listEditorModal&&els.listEditorModal.addEventListener('click',e=>{if(e.target&&e.target.classList&&e.target.classList.contains('modal-backdrop'))closeListEditorModal(false);});
  els.exportPackModal&&els.exportPackModal.addEventListener('click',e=>{if(e.target&&e.target.classList&&e.target.classList.contains('modal-backdrop'))closeExportPackModal();});
  document.addEventListener('click',()=>closeColorMenus());
  document.addEventListener('keydown',e=>{
    if(e.key==='Escape'){
      if(els.listEditorModal&&!els.listEditorModal.classList.contains('hidden')){closeListEditorModal(false);return;}
      if(els.exportPackModal&&!els.exportPackModal.classList.contains('hidden')){closeExportPackModal();return;}
      closeColorMenus();
      if(els.runInputModal&&!els.runInputModal.classList.contains('hidden')){closeRunInputModal(null);return;}
      if(els.modal&&!els.modal.classList.contains('hidden'))closeModal();
      return;
    }
    if((e.key==='Enter'||e.keyCode===13)&&els.listEditorModal&&!els.listEditorModal.classList.contains('hidden')){
      if(e.ctrlKey||e.metaKey){
        e.preventDefault();
        closeListEditorModal(true);
      }
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

async function init(){
  bindThemeSync();
  bindDirectNative();
  await load();
  bind();
  syncSectionCollapseUi();
  renderList();
  ensureProbe(false).catch(()=>{});
  window.ReportsDocBuilder={run:runDb,probe:probeDb};
}
init().catch(()=>{bind();syncSectionCollapseUi();renderList();ensureProbe(false).catch(()=>{});});
})();
