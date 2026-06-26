/* ============================================================
   Pomodoro Timer App
   ============================================================ */

// ── Constants ─────────────────────────────────────────────
const WORK_DURATION = 25 * 60;   // 25 minutes in seconds
const BREAK_DURATION = 5 * 60;   // 5 minutes in seconds
const MAX_POMODORO_DISPLAY = 8;
const POMODORO_EMOJI = '🍅';
const LS_KEY_TASKS = 'pomodoro_tasks';
const LS_KEY_DAILY = 'pomodoro_daily';

// ── Mode config (single source of truth for mode data) ────
const MODE_CONFIG = {
  work:  { duration: WORK_DURATION, nextMode: 'break' },
  break: { duration: BREAK_DURATION, nextMode: 'work'  },
};

const STATUS_LABEL = {
  idle:    { work: '准备开始', break: '准备休息' },
  running: { work: '专注中...', break: '休息中...' },
  paused:  { work: '已暂停',   break: '已暂停'   },
};

// ── State ─────────────────────────────────────────────────
const State = {
  mode: 'work',
  phase: 'idle',
  remainingSeconds: WORK_DURATION,
  timerInterval: null,
  todayCount: 0,

  // Tasks
  tasks: [],
  activeTaskId: null,
};

// ── DOM Refs ──────────────────────────────────────────────
const $ = (sel) => document.querySelector(sel);
const timerDisplay = $('#timerDisplay');
const timerStatus = $('#timerStatus');
const ringProgress = $('#ringProgress');
const btnStart = $('#btnStart');
const btnPause = $('#btnPause');
const btnReset = $('#btnReset');
const todayCount = $('#todayCount');
const taskList = $('#taskList');
const taskInput = $('#taskInput');
const btnAddTask = $('#btnAddTask');

// SVG ring circumference — read from DOM to stay in sync with markup
const ringCircumference = 2 * Math.PI * ringProgress.r.baseVal.value;

// ── Shared AudioContext (reused, not created per beep) ────
let audioCtx = null;
function getAudioCtx() {
  if (!audioCtx) {
    audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  }
  if (audioCtx.state === 'suspended') audioCtx.resume();
  return audioCtx;
}

// ── Persistence helpers ───────────────────────────────────
function safeGetJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) {
    return fallback;
  }
}

function safeSetJSON(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function getTodayKey() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function loadData() {
  State.tasks = safeGetJSON(LS_KEY_TASKS, []);
  // Load today's count into cached state
  const daily = safeGetJSON(LS_KEY_DAILY, {});
  const key = getTodayKey();
  State.todayCount = (daily[key] && daily[key].count) || 0;
}

function saveTasks() {
  safeSetJSON(LS_KEY_TASKS, State.tasks);
}

// Increment today's count, persist, return new value
function incrementTodayPomodoros() {
  State.todayCount += 1;
  const key = getTodayKey();
  const daily = safeGetJSON(LS_KEY_DAILY, {});
  if (!daily[key]) daily[key] = { count: 0 };
  daily[key].count = State.todayCount;
  safeSetJSON(LS_KEY_DAILY, daily);
  todayCount.textContent = State.todayCount;
}

// ── Notifications ─────────────────────────────────────────
function notify(title, body) {
  if (window.pomodoroAPI && window.pomodoroAPI.showNotification) {
    window.pomodoroAPI.showNotification(title, body);
  }
}

function playBeepSequence() {
  try {
    const ctx = getAudioCtx();
    const notes = [880, 660, 880];
    const noteDuration = 0.15;
    const gap = 0.1;

    notes.forEach((freq, i) => {
      const startTime = ctx.currentTime + i * (noteDuration + gap);
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      osc.frequency.value = freq;
      gain.gain.setValueAtTime(0.3, startTime);
      gain.gain.exponentialRampToValueAtTime(0.001, startTime + noteDuration);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start(startTime);
      osc.stop(startTime + noteDuration + 0.05);
    });
  } catch (e) {
    // Audio not available — silently skip
  }
}

// ── Transition handlers ───────────────────────────────────
function onPomodoroComplete() {
  incrementTodayPomodoros();

  if (State.activeTaskId) {
    const task = State.tasks.find(t => t.id === State.activeTaskId);
    if (task) {
      task.pomodoros += 1;
      saveTasks();
      renderTasks();
    }
  }

  notify('番茄钟完成！ ' + POMODORO_EMOJI, '一个工作时段结束，休息一下吧~');
  playBeepSequence();
}

function onBreakComplete() {
  notify('休息结束！ ☕', '休息时间到，开始新的番茄钟吧~');
  playBeepSequence();
}

// ── Timer Logic ───────────────────────────────────────────
function formatTime(seconds) {
  const m = Math.floor(seconds / 60);
  const s = seconds % 60;
  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}

function totalSeconds() {
  return MODE_CONFIG[State.mode].duration;
}

function refreshUI() {
  updateTimerDisplay();
  updateStatusText();
  updateButtonStates();
}

function updateTimerDisplay() {
  timerDisplay.textContent = formatTime(State.remainingSeconds);
  const fraction = 1 - State.remainingSeconds / totalSeconds();
  ringProgress.style.strokeDashoffset = ringCircumference * fraction;
}

function updateStatusText() {
  timerStatus.textContent = STATUS_LABEL[State.phase][State.mode];
}

function updateButtonStates() {
  const phase = State.phase;
  btnStart.disabled = (phase === 'running');
  btnStart.textContent = (phase === 'paused') ? '▶ 继续' : '▶ 开始';
  btnPause.disabled = (phase !== 'running');
}

function updateRingModeClass() {
  ringProgress.classList.toggle('break-mode', State.mode === 'break');
}

function clearTimer() {
  if (State.timerInterval) {
    clearInterval(State.timerInterval);
    State.timerInterval = null;
  }
}

function tick() {
  if (State.remainingSeconds <= 0) {
    clearTimer();
    timerFinished();
    return;
  }
  State.remainingSeconds -= 1;
  updateTimerDisplay();
}

function switchMode(newMode) {
  State.mode = newMode;
  State.remainingSeconds = MODE_CONFIG[newMode].duration;
  State.phase = 'idle';
  updateRingModeClass();
  refreshUI();
}

function timerFinished() {
  if (State.mode === 'work') {
    onPomodoroComplete();
    switchMode('break');
    startTimer(); // auto-start break
  } else {
    onBreakComplete();
    switchMode('work');
  }
}

function startTimer() {
  if (State.timerInterval) return;
  State.phase = 'running';
  updateStatusText();
  updateButtonStates();
  State.timerInterval = setInterval(tick, 1000);
}

function pauseTimer() {
  clearTimer();
  State.phase = 'paused';
  updateStatusText();
  updateButtonStates();
}

function resetTimer() {
  clearTimer();
  State.remainingSeconds = WORK_DURATION;
  State.phase = 'idle';
  if (State.mode !== 'work') {
    State.mode = 'work';
    updateRingModeClass();
  }
  refreshUI();
}

// ── Button Handlers ───────────────────────────────────────
btnStart.addEventListener('click', () => {
  if (State.phase === 'idle' || State.phase === 'paused') startTimer();
});

btnPause.addEventListener('click', () => {
  if (State.phase === 'running') pauseTimer();
});

btnReset.addEventListener('click', resetTimer);

// ── Task Management ───────────────────────────────────────
function generateId() {
  return crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function addTask(title) {
  State.tasks.push({
    id: generateId(),
    title: title.trim(),
    pomodoros: 0,
    completed: false,
    createdAt: new Date().toISOString(),
  });
  saveTasks();
  renderTasks();
}

function deleteTask(id) {
  State.tasks = State.tasks.filter(t => t.id !== id);
  if (State.activeTaskId === id) State.activeTaskId = null;
  saveTasks();
  renderTasks();
}

function toggleTaskComplete(id) {
  const task = State.tasks.find(t => t.id === id);
  if (task) {
    task.completed = !task.completed;
    saveTasks();
    renderTasks();
  }
}

function setActiveTask(id) {
  const task = State.tasks.find(t => t.id === id);
  State.activeTaskId = (task && !task.completed) ? id : null;
  renderTasks();
}

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function renderTasks() {
  if (State.tasks.length === 0) {
    taskList.innerHTML = '<div class="tasks-empty">暂无任务，添加一个开始吧 ✨</div>';
    return;
  }

  taskList.innerHTML = State.tasks
    .map((task) => {
      const isActive = task.id === State.activeTaskId;
      const completedClass = task.completed ? 'completed' : '';
      const activeDot = isActive && !task.completed
        ? '<span class="task-active-dot" title="当前选中">●</span>'
        : '';
      const pomodoroStr = task.pomodoros > 0
        ? `<span class="task-pomodoros">${POMODORO_EMOJI.repeat(Math.min(task.pomodoros, MAX_POMODORO_DISPLAY))}`
          + `<span class="task-pomodoro-count">${task.pomodoros}</span></span>`
        : '';

      return `
        <div class="task-item ${completedClass}">
          <input
            type="checkbox"
            class="task-checkbox"
            ${task.completed ? 'checked' : ''}
            data-action="toggle"
            data-id="${task.id}"
          />
          <span
            class="task-title"
            data-action="select"
            data-id="${task.id}"
            title="${isActive ? '当前选中任务' : '点击选中此任务'}">${escapeHtml(task.title)}${activeDot}</span>
          ${pomodoroStr}
          <button
            class="task-delete"
            data-action="delete"
            data-id="${task.id}"
            title="删除任务">✕</button>
        </div>
      `;
    })
    .join('');
}

// ── Single delegated click handler for task list ─────────
taskList.addEventListener('click', (e) => {
  const el = e.target.closest('[data-action]');
  if (!el) return;
  const action = el.dataset.action;
  const id = el.dataset.id;
  if (action === 'toggle')  toggleTaskComplete(id);
  else if (action === 'delete') deleteTask(id);
  else if (action === 'select') setActiveTask(id);
});

// ── Add Task Handler ──────────────────────────────────────
btnAddTask.addEventListener('click', () => {
  const title = taskInput.value.trim();
  if (!title) return;
  addTask(title);
  taskInput.value = '';
  taskInput.focus();
});

taskInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') btnAddTask.click();
});

// ── Init ──────────────────────────────────────────────────
function init() {
  loadData();
  todayCount.textContent = State.todayCount;
  refreshUI();
  renderTasks();
}

init();
