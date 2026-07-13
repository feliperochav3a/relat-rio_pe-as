/* ── Estado ──────────────────────────────────────────────────────────────── */
const state = {
  pieces:  [],   // [{file, title, previewUrl}]
  dragSrc: null,
};

/* ── Elementos ───────────────────────────────────────────────────────────── */
const $ = id => document.getElementById(id);
const dropZone    = $('dropZone');
const fileInput   = $('fileInput');
const piecesList  = $('piecesList');
const btnGenerate = $('btnGenerate');
const agentStatus = $('agentStatus');
const resultArea  = $('resultArea');

/* ── Título a partir do nome do arquivo ──────────────────────────────────── */
function filenameToTitle(filename) {
  const stem = filename.replace(/\.[^.]+$/, '');
  return stem.replace(/[_\-]+/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
}

/* ── Adicionar peças ─────────────────────────────────────────────────────── */
function addFiles(files) {
  for (const file of files) {
    if (!file.type.startsWith('image/')) continue;
    state.pieces.push({ file, title: filenameToTitle(file.name), previewUrl: URL.createObjectURL(file) });
  }
  renderList();
  updateButton();
}

/* ── Renderiza lista ─────────────────────────────────────────────────────── */
function renderList() {
  piecesList.innerHTML = '';

  state.pieces.forEach((piece, idx) => {
    const num = String(idx + 1).padStart(2, '0');
    const li  = document.createElement('li');
    li.className  = 'piece-item';
    li.draggable  = true;
    li.dataset.idx = idx;
    li.innerHTML = `
      <span class="piece-handle" title="Arrastar para reordenar">⠿</span>
      <span class="piece-num">${num}</span>
      <img class="piece-thumb" src="${piece.previewUrl}" alt="${piece.title}" loading="lazy" />
      <div class="piece-info">
        <div class="piece-title">${piece.title}</div>
        <div class="piece-subtitle">${piece.file.name}</div>
      </div>
      <button class="piece-remove" title="Remover" type="button">×</button>
    `;

    /* drag-and-drop para reordenar */
    li.addEventListener('dragstart', e => {
      state.dragSrc = idx;
      li.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
    });
    li.addEventListener('dragend', () => {
      document.querySelectorAll('.piece-item').forEach(el =>
        el.classList.remove('dragging', 'drag-target')
      );
    });
    li.addEventListener('dragover', e => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      document.querySelectorAll('.piece-item').forEach(el => el.classList.remove('drag-target'));
      li.classList.add('drag-target');
    });
    li.addEventListener('drop', e => {
      e.preventDefault();
      if (state.dragSrc === null || state.dragSrc === idx) return;
      const [moved] = state.pieces.splice(state.dragSrc, 1);
      state.pieces.splice(idx, 0, moved);
      renderList();
    });

    /* remover */
    li.querySelector('.piece-remove').addEventListener('click', () => {
      URL.revokeObjectURL(piece.previewUrl);
      state.pieces.splice(idx, 1);
      renderList();
      updateButton();
    });

    piecesList.appendChild(li);
  });
}

/* ── Botão e contador ────────────────────────────────────────────────────── */
function updateButton() {
  const subtitle = $('subtitle').value.trim();
  const n        = state.pieces.length;
  btnGenerate.disabled = !subtitle || n === 0;

  const badge = $('piecesCount');
  if (n > 0) {
    badge.textContent = `${n} ${n === 1 ? 'peça' : 'peças'}`;
    badge.classList.remove('hidden');
  } else {
    badge.classList.add('hidden');
  }
}

/* ── Drop zone ───────────────────────────────────────────────────────────── */
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', e => {
  if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over');
});
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  addFiles(Array.from(e.dataTransfer.files));
});
fileInput.addEventListener('change', () => {
  addFiles(Array.from(fileInput.files));
  fileInput.value = '';
});
$('subtitle').addEventListener('input', updateButton);

/* ── Timeline ────────────────────────────────────────────────────────────── */
const STEP_ICONS = { pending: '', active: '', done: '✓', error: '✗' };

function setStep(id, status, label, msg) {
  const step    = $(id);
  step.className = `tl-step ${status}`;

  const dot = step.querySelector('.tl-dot');
  dot.className = `tl-dot ${status}`;
  dot.querySelector('.tl-icon').textContent = STEP_ICONS[status] ?? '';

  const line = step.querySelector('.tl-line');
  if (line) line.className = `tl-line${status === 'done' ? ' done' : ''}`;

  if (label !== undefined) step.querySelector('.tl-label').textContent = label;
  if (msg   !== undefined) step.querySelector('.tl-msg').textContent   = msg;
}

function resetTimeline() {
  setStep('stepOrchestrator', 'pending', 'Orquestrador',  'Aguardando...');
  setStep('stepBuilder',      'pending', 'Builder',        'Aguardando...');
  setStep('stepArtDirector',  'pending', 'Diretor de Arte','Aguardando...');
}

/* ── Submit ──────────────────────────────────────────────────────────────── */
$('reportForm').addEventListener('submit', async e => {
  e.preventDefault();

  agentStatus.classList.remove('hidden');
  resultArea.classList.add('hidden');
  $('btnReset').classList.add('hidden');
  btnGenerate.disabled = true;
  btnGenerate.classList.add('loading');
  resetTimeline();

  const subtitle = $('subtitle').value.trim();

  /* passo 1 — Orquestrador */
  setStep('stepOrchestrator', 'active', 'Orquestrador', 'Validando input...');
  await tick();

  const fd = new FormData();
  fd.append('subtitle', subtitle);
  for (const piece of state.pieces) fd.append('files', piece.file, piece.file.name);

  setStep('stepOrchestrator', 'done',   'Orquestrador', 'Input validado');
  setStep('stepBuilder',      'active', 'Builder',       'Gerando PPTX...');
  await tick();

  /* passo 2 — Builder (fetch real) */
  let data;
  try {
    const res = await fetch('/api/generate', { method: 'POST', body: fd });
    data = await res.json();
    if (!res.ok) throw new Error(data.detail || 'Erro desconhecido');
  } catch (err) {
    setStep('stepBuilder',     'error', 'Builder',        err.message);
    setStep('stepArtDirector', 'error', 'Diretor de Arte','Abortado');
    showError(err.message);
    endLoading();
    return;
  }

  /* passo 3 — Diretor de Arte */
  setStep('stepBuilder',     'done',   'Builder',        'PPTX gerado');
  setStep('stepArtDirector', 'active', 'Diretor de Arte','Revisando consistência...');
  await tick(700);

  const hasIssues = data.issues && data.issues.length > 0;
  setStep('stepArtDirector', 'done', 'Diretor de Arte',
    hasIssues ? 'Entregue com ressalvas' : 'Aprovado'
  );

  showResult(data);
  endLoading();
});

function endLoading() {
  btnGenerate.classList.remove('loading');
  btnGenerate.disabled = false;
  updateButton();
}

/* ── Resultado ───────────────────────────────────────────────────────────── */
function showResult(data) {
  const hasIssues = data.issues && data.issues.length > 0;

  $('resultIcon').textContent  = hasIssues ? '⚠️' : '✅';
  $('resultTitle').textContent = hasIssues ? 'Gerado com ressalvas' : 'PPTX Gerado!';
  $('artDirectorMessage').textContent = data.art_director_message || '';

  const list = $('issuesList');
  if (hasIssues) {
    list.innerHTML = data.issues.map(i => `<li>${i}</li>`).join('');
    list.classList.remove('hidden');
  } else {
    list.classList.add('hidden');
  }

  const btn = $('btnDownload');
  if (data.download_url) {
    btn.href = data.download_url;
    btn.classList.remove('hidden');
  } else {
    btn.classList.add('hidden');
  }

  $('btnReset').classList.remove('hidden');
  resultArea.classList.remove('hidden');
  resultArea.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function showError(msg) {
  $('resultIcon').textContent  = '❌';
  $('resultTitle').textContent = 'Erro na geração';
  $('artDirectorMessage').textContent = msg;
  $('issuesList').classList.add('hidden');
  $('btnDownload').classList.add('hidden');
  $('btnReset').classList.remove('hidden');
  resultArea.classList.remove('hidden');
}

/* ── Reset / Nova geração ────────────────────────────────────────────────── */
$('btnReset').addEventListener('click', () => {
  state.pieces.forEach(p => URL.revokeObjectURL(p.previewUrl));
  state.pieces = [];
  $('subtitle').value = '';
  renderList();
  updateButton();
  agentStatus.classList.add('hidden');
  resultArea.classList.add('hidden');
  $('btnReset').classList.add('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

/* ── Tick helper ─────────────────────────────────────────────────────────── */
const tick = (ms = 120) => new Promise(r => setTimeout(r, ms));
