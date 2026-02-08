const REPORTS_DEBUG_DEFAULT = true;
if (typeof window !== 'undefined' && typeof window.REPORTS_DEBUG === 'undefined') {
  window.REPORTS_DEBUG = REPORTS_DEBUG_DEFAULT;
}

const state = {
  lang: 'en',
  dict: {},
  templates: [],
  selectedId: null,
  actions: [],
  debugLog: [],
  parentLogIndex: 0
};

const DEFAULT_I18N = {
  en: {
    title: 'Reports',
    create: 'Create',
    settings: 'Settings',
    search: 'Search reports',
    empty_title: 'No templates yet',
    empty_text: 'Add a template in settings to start generating reports.',
    fill: 'Fill',
    details: 'Details',
    untitled: 'Untitled',
    fields: 'Fields',
    sheets: 'Sheets',
    category: 'Category',
    owner: 'Owner',
    updated: 'Updated',
    last_run: 'Last run',
    status_active: 'Active',
    status_draft: 'Draft',
    settings_title: 'Template settings',
    save: 'Save',
    close: 'Close',
    settings_general: 'General',
    settings_fields: 'Fields',
    settings_rules: 'Rules',
    settings_actions: 'Actions',
    settings_test: 'Test',
    template_name: 'Template name',
    template_file: 'Template file',
    template_file_help: 'Files will be stored in reports-ui/templates.',
    choose: 'Choose',
    default_sheet: 'Default sheet',
    key_scope: 'Key search scope',
    scope_sheet_label: 'Sheet to search',
    key_format: 'Key format',
    sheet_any: 'Any sheet',
    scope_workbook: 'Whole workbook',
    scope_sheet: 'Single sheet',
    scope_sheets: 'Sheet list',
    scope_named: 'Named ranges only',
    scope_all_hint: '',
    sheet_list_placeholder: 'Sheet1,Sheet2',
    key_scope_help: 'Choose whether keys are searched across the whole workbook or a specific sheet.',
    add_field: 'Add field',
    field_name: 'Name',
    field_key: 'Key',
    field_type: 'Type',
    field_required: 'Required',
    field_scope: 'Scope',
    field_sheet: 'Sheet',
    field_target: 'Target',
    type_text: 'Text',
    type_date: 'Date',
    add_rule: 'Add rule',
    rule_if: 'IF',
    rule_then: 'THEN',
    cond_empty: 'is empty',
    cond_not_empty: 'is not empty',
    action_hide_row: 'Hide row',
    action_show_row: 'Show row',
    action_set_value: 'Set value',
    add_action: 'Add action',
    action_set_border: 'Set border',
    action_clear_border: 'Clear border',
    border_all: 'All borders',
    border_outer: 'Outer borders',
    border_inner: 'Inner borders',
    run_test: 'Run test',
    preview_placeholder: 'Preview will appear here',
    form_preview: 'Form preview',
    form_preview_hint: 'Fields required before execution',
    form_preview_empty: 'No manual fields required for these actions.',
    step: 'Step',
    action_set_text: 'Set text',
    action_group_cols: 'Group columns',
    action_delete_row: 'Delete row',
    input_fixed: 'Fixed text',
    input_manual: 'Manual input',
    input_select: 'Select from list',
    input_value: 'Value',
    field_label: 'Field label',
    field: 'Field',
    field_options: 'Options (comma-separated)',
    columns_range: 'Columns range (C:E)',
    group_expanded: 'Keep expanded',
    row_number: 'Row number',
    target_cell: 'Target cell',
    target_key: 'By key',
    target_key_placeholder: 'Key name',
    target_cell_range: 'Cell or range',
    merge_cells: 'Merge cells'
  },
  ru: {
    title: 'Отчеты',
    create: 'Создать',
    settings: 'Настройки',
    search: 'Поиск отчетов',
    empty_title: 'Шаблонов пока нет',
    empty_text: 'Добавьте шаблон в настройках, чтобы начать создавать отчеты.',
    fill: 'Заполнить',
    details: 'Детали',
    untitled: 'Без названия',
    fields: 'Поля',
    sheets: 'Листы',
    category: 'Категория',
    owner: 'Автор',
    updated: 'Обновление',
    last_run: 'Запуск',
    status_active: 'Активный',
    status_draft: 'Черновик',
    settings_title: 'Настройки шаблона',
    save: 'Сохранить',
    close: 'Закрыть',
    settings_general: 'Основное',
    settings_fields: 'Поля',
    settings_rules: 'Правила',
    settings_actions: 'Действия',
    settings_test: 'Тест',
    template_name: 'Название шаблона',
    template_file: 'Файл шаблона',
    template_file_help: 'Файлы будут храниться в reports-ui/templates.',
    choose: 'Выбрать',
    default_sheet: 'Лист по умолчанию',
    key_scope: 'Область поиска ключей',
    scope_sheet_label: 'Лист для поиска',
    key_format: 'Формат ключа',
    sheet_any: 'Любой лист',
    scope_workbook: 'Вся книга',
    scope_sheet: 'Один лист',
    scope_sheets: 'Список листов',
    scope_named: 'Только именованные диапазоны',
    scope_all_hint: '',
    sheet_list_placeholder: 'Лист1,Лист2',
    key_scope_help: 'Укажите, где искать ключи: по всей книге или на конкретном листе.',
    add_field: 'Добавить поле',
    field_name: 'Название',
    field_key: 'Ключ',
    field_type: 'Тип',
    field_required: 'Обяз.',
    field_scope: 'Область',
    field_sheet: 'Лист',
    field_target: 'Цель',
    type_text: 'Текст',
    type_date: 'Дата',
    add_rule: 'Добавить правило',
    rule_if: 'ЕСЛИ',
    rule_then: 'ТОГДА',
    cond_empty: 'пусто',
    cond_not_empty: 'не пусто',
    action_hide_row: 'Скрыть строку',
    action_show_row: 'Показать строку',
    action_set_value: 'Задать значение',
    add_action: 'Добавить действие',
    action_set_border: 'Поставить границы',
    action_clear_border: 'Убрать границы',
    border_all: 'Все границы',
    border_outer: 'Внешние границы',
    border_inner: 'Внутренние границы',
    run_test: 'Запуск теста',
    preview_placeholder: 'Здесь появится предпросмотр',
    form_preview: 'Форма ввода',
    form_preview_hint: 'Поля, которые нужно заполнить перед выполнением',
    form_preview_empty: 'Для этих действий ввод не требуется.',
    step: 'Шаг',
    action_set_text: 'Вставить текст',
    action_group_cols: 'Группировать столбцы',
    action_delete_row: 'Удалить строку',
    input_fixed: 'Готовый текст',
    input_manual: 'Ввод вручную',
    input_select: 'Выбор из списка',
    input_value: 'Значение',
    field_label: 'Подпись поля',
    field: 'Поле',
    field_options: 'Список вариантов (через ;)',
    columns_range: 'Диапазон столбцов (C:E)',
    group_expanded: 'Оставить раскрытой',
    row_number: 'Номер строки',
    target_cell: 'Целевая ячейка',
    target_key: 'По ключу',
    target_key_placeholder: 'Имя ключа',
    target_cell_range: 'Ячейка или диапазон',
    merge_cells: 'Объединить ячейки'
  }
};

const els = {
  list: () => document.getElementById('reports-list'),
  empty: () => document.getElementById('reports-empty'),
  search: () => document.getElementById('search'),
  btnCreate: () => document.getElementById('btn-create'),
  btnSettings: () => document.getElementById('btn-settings'),
  viewReports: () => document.getElementById('reports-view'),
  viewSettings: () => document.getElementById('settings-view'),
  btnBack: () => document.getElementById('btn-back'),
  btnClose: () => document.getElementById('btn-close'),
  navItems: () => document.querySelectorAll('.settings-nav-item'),
  actionsList: () => document.getElementById('actions-list'),
  formPreview: () => document.getElementById('form-preview-body'),
  templateFileInput: () => document.getElementById('template-file-input'),
  templateFileBtn: () => document.getElementById('template-file-btn'),
  templateFilePath: () => document.getElementById('template-file-path')
};

function normalizeLang(lang) {
  if (!lang) return 'en';
  return lang.replace('_', '-');
}

function pickLang(raw) {
  if (!raw) return 'en';
  const n = normalizeLang(raw);
  if (n.length >= 2) return n;
  return 'en';
}

function getDefaultDict(lang) {
  if (!lang) return DEFAULT_I18N.en;
  const key = lang.toLowerCase();
  if (key.startsWith('ru')) return DEFAULT_I18N.ru;
  return DEFAULT_I18N.en;
}

function applyTranslations() {
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const key = el.getAttribute('data-i18n');
    if (state.dict[key]) el.textContent = state.dict[key];
  });
  document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
    const key = el.getAttribute('data-i18n-placeholder');
    if (state.dict[key]) el.setAttribute('placeholder', state.dict[key]);
  });
}

function getActionLabel(type) {
  const map = {
    setText: state.dict.action_set_text || 'Set text',
    groupCols: state.dict.action_group_cols || 'Group columns',
    deleteRow: state.dict.action_delete_row || 'Delete row'
  };
  return map[type] || type;
}

function initActions() {
  state.actions = [
    {
      id: 'act-1',
      type: 'setText',
      sheet: 'Лист1',
      scope: 'sheet',
      sheetList: '',
      targetMode: 'cell',
      target: 'A1',
      inputMode: 'fixed',
      value: 'Привет',
      label: 'Текст для A1',
      options: ['Привет', 'Добрый день']
    },
    {
      id: 'act-2',
      type: 'groupCols',
      sheet: 'Лист2',
      scope: 'sheet',
      sheetList: '',
      range: 'C:E',
      expanded: true
    },
    {
      id: 'act-3',
      type: 'deleteRow',
      sheet: 'Лист1',
      scope: 'sheet',
      sheetList: '',
      row: '8'
    }
  ];
}

function createActionRow() {
  const row = document.createElement('div');
  row.className = 'action-grid';
  return row;
}

function renderActions() {
  const root = els.actionsList();
  if (!root) return;
  root.innerHTML = '';

  const sheets = ['Лист1', 'Лист2'];
  state.actions.forEach((action, index) => {
    const item = document.createElement('div');
    item.className = 'action-item';

    const header = document.createElement('div');
    header.className = 'action-header';

    const step = document.createElement('div');
    step.className = 'action-step';
    step.textContent = `${state.dict.step || 'Step'} ${index + 1}`;

    const controls = document.createElement('div');
    controls.className = 'action-controls';

    const btnUp = document.createElement('button');
    btnUp.textContent = '↑';
    btnUp.disabled = index === 0;
    btnUp.addEventListener('click', () => {
      if (index === 0) return;
      const tmp = state.actions[index - 1];
      state.actions[index - 1] = action;
      state.actions[index] = tmp;
      renderActions();
      renderFormPreview();
    });

    const btnDown = document.createElement('button');
    btnDown.textContent = '↓';
    btnDown.disabled = index === state.actions.length - 1;
    btnDown.addEventListener('click', () => {
      if (index === state.actions.length - 1) return;
      const tmp = state.actions[index + 1];
      state.actions[index + 1] = action;
      state.actions[index] = tmp;
      renderActions();
      renderFormPreview();
    });

    controls.appendChild(btnUp);
    controls.appendChild(btnDown);

    header.appendChild(step);
    header.appendChild(controls);
    item.appendChild(header);

    const grid = document.createElement('div');
    grid.className = 'action-grid';
    // main row: type + scope + sheet(s)
    const typeSelect = document.createElement('select');
    ['setText', 'groupCols', 'deleteRow'].forEach(type => {
      const opt = document.createElement('option');
      opt.value = type;
      opt.textContent = getActionLabel(type);
      if (type === action.type) opt.selected = true;
      typeSelect.appendChild(opt);
    });
    typeSelect.addEventListener('change', e => {
      action.type = e.target.value;
      renderActions();
      renderFormPreview();
    });

    const scopeSelect = document.createElement('select');
    const scopes = [
      { id: 'workbook', label: state.dict.scope_workbook || 'Whole workbook' },
      { id: 'sheet', label: state.dict.scope_sheet || 'Single sheet' },
      { id: 'sheets', label: state.dict.scope_sheets || 'Sheet list' },
      { id: 'named', label: state.dict.scope_named || 'Named ranges only' }
    ];
    scopes.forEach(scope => {
      const opt = document.createElement('option');
      opt.value = scope.id;
      opt.textContent = scope.label;
      if (scope.id === action.scope) opt.selected = true;
      scopeSelect.appendChild(opt);
    });
    scopeSelect.addEventListener('change', e => {
      action.scope = e.target.value;
      renderActions();
    });

    const sheetSelect = document.createElement('select');
    sheets.forEach(sheet => {
      const opt = document.createElement('option');
      opt.value = sheet;
      opt.textContent = sheet;
      if (sheet === action.sheet) opt.selected = true;
      sheetSelect.appendChild(opt);
    });
    sheetSelect.addEventListener('change', e => {
      action.sheet = e.target.value;
      renderFormPreview();
    });

    const sheetListInput = document.createElement('input');
    sheetListInput.className = 'span-2';
    sheetListInput.value = action.sheetList || '';
    sheetListInput.placeholder = state.dict.sheet_list_placeholder || 'Sheet1,Sheet2';
    sheetListInput.addEventListener('input', e => {
      action.sheetList = e.target.value;
    });

    grid.appendChild(typeSelect);
    grid.appendChild(scopeSelect);
    if (action.scope === 'sheet') {
      grid.appendChild(sheetSelect);
      const spacer = document.createElement('div');
      spacer.className = 'span-2';
      grid.appendChild(spacer);
    } else if (action.scope === 'sheets') {
      grid.appendChild(sheetListInput);
    } else {
      const span = document.createElement('div');
      span.className = 'span-2';
      span.textContent = state.dict.scope_all_hint || '';
      grid.appendChild(span);
    }

    if (action.type === 'setText') {
      const rowTarget = createActionRow();
      const targetMode = document.createElement('select');
      const targetModes = [
        { id: 'key', label: state.dict.target_key || 'By key' },
        { id: 'cell', label: state.dict.target_cell_range || 'Cell or range' }
      ];
      targetModes.forEach(mode => {
        const opt = document.createElement('option');
        opt.value = mode.id;
        opt.textContent = mode.label;
        if (mode.id === action.targetMode) opt.selected = true;
        targetMode.appendChild(opt);
      });
      targetMode.addEventListener('change', e => {
        action.targetMode = e.target.value;
        renderActions();
      });

      const targetInput = document.createElement('input');
      targetInput.value = action.target || '';
      targetInput.placeholder = action.targetMode === 'key'
        ? (state.dict.target_key_placeholder || 'Key name')
        : (state.dict.target_cell_range || 'Cell or range');
      targetInput.addEventListener('input', e => {
        action.target = e.target.value;
      });

      const mergeWrap = document.createElement('label');
      mergeWrap.className = 'action-toggle';
      const mergeToggle = document.createElement('input');
      mergeToggle.type = 'checkbox';
      mergeToggle.checked = !!action.merge;
      mergeToggle.addEventListener('change', e => {
        action.merge = e.target.checked;
      });
      mergeWrap.appendChild(mergeToggle);
      const mergeText = document.createElement('span');
      mergeText.textContent = state.dict.merge_cells || 'Merge cells';
      mergeWrap.appendChild(mergeText);

      rowTarget.appendChild(targetMode);
      rowTarget.appendChild(targetInput);
      const mergeSlot = document.createElement('div');
      mergeSlot.className = 'span-2';
      if (action.targetMode === 'cell') {
        mergeSlot.appendChild(mergeWrap);
      }
      rowTarget.appendChild(mergeSlot);

      const modeSelect = document.createElement('select');
      const modeItems = [
        { id: 'fixed', label: state.dict.input_fixed || 'Fixed text' },
        { id: 'manual', label: state.dict.input_manual || 'Manual input' },
        { id: 'select', label: state.dict.input_select || 'Select from list' }
      ];
      modeItems.forEach(mode => {
        const opt = document.createElement('option');
        opt.value = mode.id;
        opt.textContent = mode.label;
        if (mode.id === action.inputMode) opt.selected = true;
        modeSelect.appendChild(opt);
      });
      modeSelect.addEventListener('change', e => {
        action.inputMode = e.target.value;
        renderActions();
        renderFormPreview();
      });

      const rowValue = createActionRow();
      const valueInput = document.createElement('input');
      valueInput.className = 'span-2';
      valueInput.value = action.value || '';
      valueInput.placeholder = state.dict.input_value || 'Value';
      valueInput.addEventListener('input', e => {
        action.value = e.target.value;
      });

      const labelInput = document.createElement('input');
      labelInput.className = 'span-2';
      labelInput.value = action.label || '';
      labelInput.placeholder = state.dict.field_label || 'Field label';
      labelInput.addEventListener('input', e => {
        action.label = e.target.value;
      });

      const rowOptions = createActionRow();
      const optionsInput = document.createElement('input');
      optionsInput.className = 'span-4';
      optionsInput.value = (action.options || []).join('; ');
      optionsInput.placeholder = state.dict.field_options || 'Options (comma-separated)';
      optionsInput.addEventListener('input', e => {
        action.options = e.target.value.split(';').map(v => v.trim()).filter(Boolean);
        renderFormPreview();
      });

      rowValue.appendChild(modeSelect);
      if (action.inputMode === 'fixed') {
        rowValue.appendChild(valueInput);
      } else if (action.inputMode === 'manual') {
        rowValue.appendChild(labelInput);
      } else if (action.inputMode === 'select') {
        rowValue.appendChild(labelInput);
        rowOptions.appendChild(optionsInput);
      }
      item.appendChild(grid);
      item.appendChild(rowTarget);
      item.appendChild(rowValue);
      if (action.inputMode === 'select') {
        item.appendChild(rowOptions);
      }
    } else if (action.type === 'groupCols') {
      const row = createActionRow();
      const rangeInput = document.createElement('input');
      rangeInput.value = action.range || '';
      rangeInput.placeholder = state.dict.columns_range || 'Columns range (C:E)';
      rangeInput.addEventListener('input', e => {
        action.range = e.target.value;
      });

      const toggleWrap = document.createElement('label');
      toggleWrap.className = 'action-toggle';
      const toggle = document.createElement('input');
      toggle.type = 'checkbox';
      toggle.checked = !!action.expanded;
      toggle.addEventListener('change', e => {
        action.expanded = e.target.checked;
      });
      toggleWrap.appendChild(toggle);
      const toggleText = document.createElement('span');
      toggleText.textContent = state.dict.group_expanded || 'Keep expanded';
      toggleWrap.appendChild(toggleText);

      const span = document.createElement('div');
      span.className = 'span-2';
      span.appendChild(toggleWrap);
      row.appendChild(rangeInput);
      row.appendChild(span);
      item.appendChild(grid);
      item.appendChild(row);
    } else if (action.type === 'deleteRow') {
      const row = createActionRow();
      const rowInput = document.createElement('input');
      rowInput.value = action.row || '';
      rowInput.placeholder = state.dict.row_number || 'Row number';
      rowInput.addEventListener('input', e => {
        action.row = e.target.value;
      });
      const empty = document.createElement('div');
      empty.className = 'span-2';
      row.appendChild(rowInput);
      row.appendChild(empty);
      item.appendChild(grid);
      item.appendChild(row);
    }

    root.appendChild(item);
  });
}

function renderFormPreview() {
  const body = els.formPreview();
  if (!body) return;
  body.innerHTML = '';
  const fields = [];
  state.actions.forEach((action, index) => {
    if (action.type !== 'setText') return;
    if (action.inputMode === 'manual' || action.inputMode === 'select') {
      fields.push({
        id: action.id,
        label: action.label || `${state.dict.field || 'Field'} ${index + 1}`,
        type: action.inputMode,
        options: action.options || []
      });
    }
  });
  if (!fields.length) {
    const empty = document.createElement('div');
    empty.className = 'form-preview-empty';
    empty.textContent = state.dict.form_preview_empty || 'No manual fields required for these actions.';
    body.appendChild(empty);
    return;
  }
  fields.forEach(field => {
    const wrap = document.createElement('div');
    wrap.className = 'form-preview-field';
    const label = document.createElement('label');
    label.textContent = field.label;
    wrap.appendChild(label);
    if (field.type === 'select') {
      const select = document.createElement('select');
      field.options.forEach(optVal => {
        const opt = document.createElement('option');
        opt.value = optVal;
        opt.textContent = optVal;
        select.appendChild(opt);
      });
      wrap.appendChild(select);
    } else {
      const input = document.createElement('input');
      input.type = 'text';
      wrap.appendChild(input);
    }
    body.appendChild(wrap);
  });
}

function applyLocaleDict(dict, lang) {
  if (!dict || typeof dict !== 'object') return;
  const base = getDefaultDict(lang || state.lang);
  state.dict = Object.assign({}, base, state.dict, dict);
  if (lang) {
    state.lang = pickLang(lang);
  }
  applyTranslations();
}

async function loadLocale(lang) {
  const baseLang = pickLang(lang);
  const base = baseLang.toLowerCase();
  const candidates = [base, base.split('-')[0], 'en'];
  const baseDict = Object.assign({}, getDefaultDict(baseLang), state.dict);
  for (const c of candidates) {
    try {
      const res = await fetch(`locales/${c}.json`);
      if (res.ok) {
        const dict = await res.json();
        state.dict = Object.assign({}, baseDict, dict);
        state.lang = c;
        applyTranslations();
        return;
      }
    } catch (e) {
      // ignore
    }
  }
}

async function loadTemplates() {
  try {
    const res = await fetch('data/templates.json');
    if (res.ok) {
      state.templates = await res.json();
      return;
    }
  } catch (e) {
    // ignore
  }
  state.templates = [];
}

function renderTemplates() {
  const list = els.list();
  const empty = els.empty();
  list.innerHTML = '';

  const term = (els.search().value || '').trim().toLowerCase();
  const items = state.templates.filter(t => {
    const name = (t.name || '').toLowerCase();
    const descr = (t.descr || '').toLowerCase();
    return !term || name.includes(term) || descr.includes(term);
  });

  if (!items.length) {
    empty.hidden = false;
    return;
  }
  empty.hidden = true;

  const labelFields = state.dict.fields || 'Fields';
  const labelSheets = state.dict.sheets || 'Sheets';
  const labelCategory = state.dict.category || 'Category';
  const labelOwner = state.dict.owner || 'Owner';
  const labelUpdated = state.dict.updated || 'Updated';
  const labelLastRun = state.dict.last_run || 'Last run';

  const makeCover = (seed) => {
    let hash = 0;
    const str = seed || Math.random().toString(36);
    for (let i = 0; i < str.length; i += 1) {
      hash = (hash * 31 + str.charCodeAt(i)) >>> 0;
    }
    const hue = hash % 360;
    const hue2 = (hue + 28) % 360;
    return `linear-gradient(135deg, hsl(${hue}, 40%, 38%), hsl(${hue2}, 55%, 45%))`;
  };

  for (const t of items) {
    const card = document.createElement('div');
    card.className = 'report-card' + (state.selectedId === t.id ? ' selected' : '');
    card.dataset.id = t.id;

    const cover = document.createElement('div');
    cover.className = 'report-cover';
    cover.style.background = t.cover || makeCover(t.id || t.name);

    const coverType = document.createElement('span');
    coverType.className = 'report-cover-type';
    coverType.textContent = (t.type || 'XLSX').toUpperCase();

    const status = (t.status || 'active').toLowerCase();
    const statusLabel = state.dict[`status_${status}`] || status;
    const coverStatus = document.createElement('span');
    coverStatus.className = `report-cover-status ${status}`;
    coverStatus.textContent = statusLabel;

    cover.appendChild(coverType);
    cover.appendChild(coverStatus);

    const body = document.createElement('div');
    body.className = 'report-body';

    const title = document.createElement('div');
    title.className = 'report-title';
    title.textContent = t.name || state.dict.untitled || 'Untitled';

    const descr = document.createElement('div');
    descr.className = 'report-descr';
    descr.textContent = t.descr || '';

    const metaRow = document.createElement('div');
    metaRow.className = 'report-meta-row';

    const metaPairs = [];
    if (t.category) metaPairs.push(`${labelCategory}: ${t.category}`);
    if (t.owner) metaPairs.push(`${labelOwner}: ${t.owner}`);
    if (t.updated) metaPairs.push(`${labelUpdated}: ${t.updated}`);
    if (t.last_run) metaPairs.push(`${labelLastRun}: ${t.last_run}`);
    metaPairs.forEach(text => {
      const meta = document.createElement('span');
      meta.className = 'report-meta-item';
      meta.textContent = text;
      metaRow.appendChild(meta);
    });

    const tags = document.createElement('div');
    tags.className = 'report-tags';
    if (Array.isArray(t.tags)) {
      t.tags.forEach(tag => {
        const chip = document.createElement('span');
        chip.className = 'report-chip';
        chip.textContent = tag;
        tags.appendChild(chip);
      });
    }

    const stats = document.createElement('div');
    stats.className = 'report-stats';
    if (t.fields) {
      const stat = document.createElement('div');
      stat.className = 'report-stat';
      stat.innerHTML = `<span class="stat-value">${t.fields}</span><span class="stat-label">${labelFields}</span>`;
      stats.appendChild(stat);
    }
    if (t.sheets) {
      const stat = document.createElement('div');
      stat.className = 'report-stat';
      stat.innerHTML = `<span class="stat-value">${t.sheets}</span><span class="stat-label">${labelSheets}</span>`;
      stats.appendChild(stat);
    }

    const actions = document.createElement('div');
    actions.className = 'report-actions';

    const fill = document.createElement('button');
    fill.className = 'btn btn-primary btn-small';
    fill.textContent = state.dict.fill || 'Fill';
    fill.addEventListener('click', e => {
      e.stopPropagation();
      state.selectedId = t.id;
      runReport(t);
    });

    const details = document.createElement('button');
    details.className = 'btn btn-small btn-ghost';
    details.textContent = state.dict.details || 'Details';
    details.disabled = true;

    actions.appendChild(fill);
    actions.appendChild(details);

    body.appendChild(title);
    body.appendChild(descr);
    if (metaRow.childNodes.length) body.appendChild(metaRow);
    if (tags.childNodes.length) body.appendChild(tags);
    if (stats.childNodes.length) body.appendChild(stats);
    body.appendChild(actions);

    card.appendChild(cover);
    card.appendChild(body);

    card.addEventListener('click', () => {
      state.selectedId = t.id;
      renderTemplates();
    });

    list.appendChild(card);
  }
}

function sendToHost(cmd, data) {
  try {
    const payload = {
      type: 'plugin',
      data: { 0: cmd, 1: data ? JSON.stringify(data) : '' }
    };
    window.parent.postMessage(JSON.stringify(payload), '*');
  } catch (e) {
    // ignore
  }
}

function postToParent(message) {
  try {
    if (window.parent && window.parent !== window) {
      logDebug(`postToParent ${message && message.event ? message.event : 'message'}`);
      window.parent.postMessage(JSON.stringify(message), '*');
    }
  } catch (e) {
    // ignore
  }
}

function getReportsRoot() {
  try {
    if (window.parent && window.parent.reportsUiRoot) {
      return window.parent.reportsUiRoot;
    }
  } catch (e) {
    // ignore
  }
  try {
    if (window.location && window.location.pathname) {
      let p = decodeURIComponent(window.location.pathname);
      if (p.startsWith('/') && /^[a-zA-Z]:/.test(p.slice(1))) {
        p = p.slice(1);
      }
      p = p.replace(/\\/g, '/');
      if (p.includes('/reports-ui/')) {
        return p.split('/reports-ui/')[0];
      }
      return p.replace(/\/[^\/]*$/, '');
    }
  } catch (e) {
    // ignore
  }
  return '';
}

function normalizePath(path) {
  if (!path) return '';
  return String(path).replace(/\\/g, '/');
}

function resolveTemplatePath(tpl) {
  if (!tpl) return '';
  const direct = tpl.path || tpl.file || '';
  let path = normalizePath(direct);
  if (!path) return '';
  if (/^https?:\/\//i.test(path) || /^file:\/\//i.test(path)) return path;
  if (/^[a-zA-Z]:\//.test(path)) return path.replace(/\//g, '\\');
  const root = normalizePath(getReportsRoot());
  if (!root) return path.replace(/\//g, '\\');
  const full = `${root}/${path}`.replace(/\/+/g, '/');
  return full.replace(/\//g, '\\');
}

function resolveTemplateType(path) {
  try {
    const ext = (path || '').split('.').pop() || 'xlsx';
    if (window.parent && window.parent.utils && window.parent.utils.fileExtensionToFileFormat) {
      return window.parent.utils.fileExtensionToFileFormat(ext);
    }
  } catch (e) {
    // ignore
  }
  return 0;
}

function getFileName(path) {
  if (!path) return '';
  const clean = String(path).replace(/\\/g, '/');
  const parts = clean.split('/');
  return parts[parts.length - 1] || '';
}

function buildJob(template) {
  const actions = (state.actions || []).map(action => {
    const payload = Object.assign({}, action);
    if (payload.type === 'setText') {
      if (payload.inputMode === 'select' && (!payload.value || payload.value === '')) {
        payload.value = (payload.options && payload.options[0]) ? payload.options[0] : '';
      }
    }
    return payload;
  });
  return {
    id: `job-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    templateId: template ? template.id : null,
    actions,
    debug: isDebugEnabled()
  };
}

function isDebugEnabled() {
  try {
    if (window.REPORTS_DEBUG === true) return true;
    if (window.parent && window.parent.REPORTS_DEBUG === true) return true;
    const stored = localStorage.getItem('reports.debug');
    if (!stored) return false;
    return stored === '1' || stored.toLowerCase() === 'true';
  } catch (e) {
    return false;
  }
}

function logDebug(message) {
  if (!isDebugEnabled()) return;
  try {
    const stamp = new Date().toLocaleTimeString();
    const entry = `[reports-ui ${stamp}] ${message}`;
    console.log(entry);
    appendDebug(entry);
  } catch (e) {
    // ignore
  }
}

function appendDebug(message) {
  try {
    if (!isDebugEnabled()) return;
    if (!message) return;
    state.debugLog.push(message);
    if (state.debugLog.length > 300) {
      state.debugLog = state.debugLog.slice(state.debugLog.length - 300);
    }
    const panel = document.getElementById('reports-debug-log');
    if (!panel) return;
    panel.textContent = state.debugLog.join('\n');
    panel.scrollTop = panel.scrollHeight;
  } catch (e) {
    // ignore
  }
}

function ensureDebugPanel() {
  if (!isDebugEnabled()) return;
  if (document.getElementById('reports-debug')) return;
  const panel = document.createElement('div');
  panel.id = 'reports-debug';
  panel.className = 'reports-debug';
  panel.innerHTML = `
    <div class="reports-debug-header">
      <span>Reports debug</span>
      <div class="reports-debug-actions">
        <button id="reports-debug-clear" class="btn btn-small btn-ghost">Clear</button>
      </div>
    </div>
    <pre id="reports-debug-log" class="reports-debug-log"></pre>
  `;
  document.body.appendChild(panel);
  const clearBtn = document.getElementById('reports-debug-clear');
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      state.debugLog = [];
      const panelLog = document.getElementById('reports-debug-log');
      if (panelLog) panelLog.textContent = '';
    });
  }
}

function syncParentDebugLog() {
  if (!isDebugEnabled()) return;
  try {
    const parentLog = window.parent && window.parent.__reportsLog;
    if (!parentLog || !Array.isArray(parentLog)) return;
    for (let i = state.parentLogIndex; i < parentLog.length; i += 1) {
      const item = parentLog[i];
      if (!item) continue;
      const stamp = item.time ? new Date(item.time).toLocaleTimeString() : '';
      const line = `[bridge ${stamp}] ${item.message}${item.data ? ' ' + JSON.stringify(item.data) : ''}`;
      appendDebug(line);
    }
    state.parentLogIndex = parentLog.length;
  } catch (e) {
    // ignore
  }
}

function sendExternalJob(job) {
  try {
    if (!window.parent || !window.parent.AscDesktopEditor) return false;
    const message = JSON.stringify({
      type: 'onExternalPluginMessage',
      data: { type: 'reports:run', job }
    });
    const script = `function(){try{var m=${JSON.stringify(message)};if(window.g_asc_plugins&&window.g_asc_plugins.runAllSystem){try{window.g_asc_plugins.runAllSystem();}catch(e){}}if(window.g_asc_plugins&&window.g_asc_plugins.sendToAllPlugins){window.g_asc_plugins.sendToAllPlugins(m);}else{window.postMessage(m,"*");}}catch(e){}}`;
    window.parent.AscDesktopEditor.CallInAllWindows(script);
    return true;
  } catch (e) {
    return false;
  }
}

function sendDirectJob(job, expectedName) {
  try {
    if (!window.parent || !window.parent.AscDesktopEditor) return false;
    const payload = JSON.stringify({ job, expectedName, debug: isDebugEnabled(), strictName: false });
    const script = `function(){try{(function(){var p=${JSON.stringify(payload)};p=JSON.parse(p);var api=(window.Asc&&Asc.editor)?Asc.editor:window.Asc; if(!api||!api.asc_getDocumentName)return; var st=window.__reportsState||(window.__reportsState={}); st.job=p.job||{}; st.debug=!!p.debug; st.expected=p.expectedName||''; st.strict=!!p.strictName; st.startedAt=st.startedAt||Date.now(); function ready(){return !!(api.isLoadFullApi&&api.isDocumentLoadComplete);} function setRange(sheet,addr){var a=String(addr||'').trim();if(!a)return;var s=String(sheet||'').trim();var full=s? (s+'!'+a):a; api.asc_setWorksheetRange(full);} function runOnce(){if(!ready())return false; var name=api.asc_getDocumentName?api.asc_getDocumentName():''; if(st.strict&&st.expected&&name&&String(name).toLowerCase()!==String(st.expected).toLowerCase())return false; var job=st.job||{}; if(st.debug){try{setRange('', 'Z1'); api.asc_insertInCell('REPORTS DEBUG '+(job.id||''));}catch(e){}} function runAction(action){if(!action||!action.type)return; if(action.type==='setText'){var target=String(action.target||'').trim(); if(!target)return; setRange(action.sheet,target); api.asc_insertInCell(String(action.value||'')); if(action.merge){try{api.asc_mergeCells();}catch(e){}} } else if(action.type==='groupCols'){ if(!action.range)return; setRange(action.sheet, action.range); api.asc_group(false); if(typeof action.expanded==='boolean'){try{api.asc_changeGroupDetails(!!action.expanded);}catch(e){}} } else if(action.type==='deleteRow'){ if(!action.row)return; var row=String(action.row).trim(); if(!row)return; setRange(action.sheet, row+':'+row); api.asc_deleteCells(Asc.c_oAscDeleteOptions.DeleteRows); } } var acts=job.actions||[]; for(var i=0;i<acts.length;i++){runAction(acts[i]);} st.job=null; return true;} if(runOnce())return; if(!st.timer){st.timer=setInterval(function(){try{if(runOnce()){clearInterval(st.timer);st.timer=null;}else if(Date.now()-st.startedAt>60000){clearInterval(st.timer);st.timer=null;}}catch(e){}},500);} })();}catch(e){}}`;
    window.parent.AscDesktopEditor.CallInAllWindows(script);
    return true;
  } catch (e) {
    return false;
  }
}

function runReport(template) {
  const path = resolveTemplatePath(template);
  if (!path) return;

  const typeId = resolveTemplateType(path);
  const job = buildJob(template);
  logDebug(`runReport path=${path} typeId=${typeId} actions=${job.actions ? job.actions.length : 0}`);
  postToParent({
    event: 'reportsRun',
    source: 'reports-ui',
    data: {
      path,
      typeId,
      job,
      debug: isDebugEnabled()
    }
  });
}

function openSettings() {
  const view = els.viewSettings();
  const reports = els.viewReports();
  if (view && reports) {
    reports.classList.add('hidden');
    view.classList.remove('hidden');
  }
}

function closeSettings() {
  const view = els.viewSettings();
  const reports = els.viewReports();
  if (view && reports) {
    view.classList.add('hidden');
    reports.classList.remove('hidden');
  }
}

function setActiveSection(sectionId) {
  const items = els.navItems();
  items.forEach(item => {
    item.classList.toggle('active', item.dataset.section === sectionId);
  });
  const section = document.getElementById(`section-${sectionId}`);
  if (section) {
    section.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function syncThemeVars() {
  try {
    const parentDoc = window.parent.document;
    const src = parentDoc.body || parentDoc.documentElement;
    const style = window.parent.getComputedStyle(src);
    if (parentDoc.body && parentDoc.body.classList) {
      const themeClasses = Array.from(parentDoc.body.classList).filter(c => c.startsWith('theme-'));
      document.documentElement.classList.forEach(c => {
        if (c.startsWith('theme-')) {
          document.documentElement.classList.remove(c);
        }
      });
      themeClasses.forEach(c => document.documentElement.classList.add(c));
    }
    const vars = [
      '--background-action-panel',
      '--background-normal',
      '--background-icon-normal',
      '--border-divider',
      '--border-regular-control',
      '--border-control-focus',
      '--text-normal',
      '--text-secondary',
      '--text-tertiary',
      '--text-link',
      '--icon-normal',
      '--highlight-button-hover',
      '--highlight-button-pressed',
      '--background-primary-button',
      '--background-accent-button',
      '--highlight-primary-button-hover',
      '--highlight-primary-button-pressed',
      '--highlight-accent-button-hover',
      '--highlight-accent-button-pressed',
      '--shadow-sidebar-item-pressed'
    ];
    const root = document.documentElement.style;
    vars.forEach(v => {
      const val = style.getPropertyValue(v);
      if (val) root.setProperty(v, val.trim());
    });
    // map to local vars
    root.setProperty('--bg', style.getPropertyValue('--background-normal').trim() || '#2f2f2f');
    root.setProperty('--panel', style.getPropertyValue('--background-action-panel').trim() || '#3a3a3a');
    root.setProperty('--card', style.getPropertyValue('--background-action-panel').trim() || '#3c3c3c');
    root.setProperty('--border', style.getPropertyValue('--border-divider').trim() || '#4a4a4a');
    root.setProperty('--text', style.getPropertyValue('--text-normal').trim() || 'rgba(255,255,255,0.9)');
    root.setProperty('--text-secondary', style.getPropertyValue('--text-secondary').trim() || 'rgba(255,255,255,0.65)');
    root.setProperty('--text-tertiary', style.getPropertyValue('--text-tertiary').trim() || 'rgba(255,255,255,0.45)');
    root.setProperty('--accent', style.getPropertyValue('--background-accent-button').trim() || '#4a87e7');
  } catch (e) {
    // ignore
  }
}

function initEvents() {
  els.search().addEventListener('input', () => renderTemplates());
  els.btnCreate().addEventListener('click', () => {
    if (!state.selectedId) {
      openSettings();
      return;
    }
    const tpl = state.templates.find(t => t.id === state.selectedId);
    if (tpl) {
      runReport(tpl);
    }
  });
  els.btnSettings().addEventListener('click', openSettings);
  els.btnBack().addEventListener('click', closeSettings);
  els.btnClose().addEventListener('click', closeSettings);

  els.navItems().forEach(item => {
    item.addEventListener('click', () => {
      setActiveSection(item.dataset.section);
    });
  });

  const fileBtn = els.templateFileBtn();
  const fileInput = els.templateFileInput();
  const filePath = els.templateFilePath();
  if (fileBtn && fileInput && filePath) {
    fileBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
      const file = fileInput.files && fileInput.files[0];
      if (file) {
        filePath.value = file.name;
        localStorage.setItem('reports.templateFile', file.name);
      }
    });
    const stored = localStorage.getItem('reports.templateFile');
    if (stored) filePath.value = stored;
  }

  window.addEventListener('message', e => {
    try {
      const msg = JSON.parse(e.data);
      if (msg && msg.event === 'uiLangChanged' && msg.data && msg.data.new) {
        loadLocale(msg.data.new);
      }
      if (msg && msg.event === 'uiLocaleChanged' && msg.data) {
        applyLocaleDict(msg.data.dict, msg.data.lang);
      }
      if (msg && msg.event === 'uiThemeChanged') {
        syncThemeVars();
      }
    } catch (err) {
      // ignore
    }
  });
}

async function init() {
  try {
    const parentLang = window.parent && window.parent.utils && window.parent.utils.Lang;
    if (parentLang) {
      applyLocaleDict({
        title: parentLang.actReports || 'Reports',
        create: parentLang.reportsCreate || parentLang.actCreateNew || 'Create',
        settings: parentLang.reportsSettings || parentLang.actSettings || 'Settings',
        search: parentLang.reportsSearch || 'Search reports',
        empty_title: parentLang.reportsEmptyTitle || 'No templates yet',
        empty_text: parentLang.reportsEmptyText || 'Add a template in settings to start generating reports.',
        fill: parentLang.reportsFill || 'Fill',
        untitled: parentLang.reportsUnnamed || 'Untitled'
      }, parentLang.id);
    }
  } catch (e) {
    // ignore
  }
  const params = new URLSearchParams(window.location.search);
  const lang = pickLang(params.get('lang') || navigator.language || 'en');
  await loadLocale(lang);
  await loadTemplates();
  syncThemeVars();
  initEvents();
  initActions();
  renderActions();
  renderFormPreview();
  renderTemplates();
  closeSettings();
  ensureDebugPanel();
  logDebug('reports-ui initialized');
  setInterval(syncParentDebugLog, 1000);
}

init();
