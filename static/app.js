/* ─────────────────────────────────────────────────────
   RESERVA DE SALAS — Reserva vía Outlook Calendar
   El deeplink se arma en el navegador para que el botón
   sea un <a> real y los bloqueadores de pop-ups no lo corten.
   ───────────────────────────────────────────────────── */

const OUTLOOK_COMPOSE_URL = 'https://outlook.office.com/calendar/action/compose';

const config = {
  rooms: [
    { name: 'Sala Colibrí',  email: 'SRR270201@gm.com' },
    { name: 'Sala Maya',     email: 'SRR271627@gm.com' },
    { name: 'Sala Alebrije', email: 'SRR240332@gm.com' },
    { name: 'Sala Ajolote',  email: 'SRR265183@gm.com' },
    { name: 'Sala Alfeñique', email: 'SRR210381@gm.com' },
  ],
  allowedDomain: 'gm.com',
  defaultDurationMin: 60,
};

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const ROOM_STORAGE_KEY = 'salas.lastRoomEmail';

let guests = []; // { email, name }
let selectedRoom = config.rooms[0];
let lastAddedGuest = null;
let hasRendered = false;
let directory = [];
let directoryPromise = null;
let suggestions = [];
let activeSuggestion = -1;

/* ── DOM refs ───────────────────────────────────────── */
const form         = document.getElementById('booking-form');
const roomTitle    = document.getElementById('room-title');
const roomSelect   = document.getElementById('room-select');
const roomNote     = document.getElementById('room-note');
const subjectInput = document.getElementById('subject-input');
const pillRow      = document.getElementById('pill-row');
const emailInput     = document.getElementById('email-input');
const suggestionsEl  = document.getElementById('guest-suggestions');
const addBtn         = document.getElementById('add-btn');
const errorMsg     = document.getElementById('error-msg');
const openBtn      = document.getElementById('open-btn');
const copyBtn      = document.getElementById('copy-btn');
const copyLabel    = document.getElementById('copy-label');
const countNote    = document.getElementById('count-note');

/* ── Animación ──────────────────────────────────────── */
const prefersReducedMotion = () =>
  window.matchMedia('(prefers-reduced-motion: reduce)').matches;

/* Quitar la clase y forzar un reflow es lo que reinicia una animación CSS. */
function replayAnimation(el, className) {
  el.classList.remove(className);
  void el.offsetWidth;
  el.classList.add(className);
}

/* ── Salas ──────────────────────────────────────────── */
function findRoom(email) {
  const needle = (email || '').trim().toLowerCase();
  return config.rooms.find((room) => room.email.toLowerCase() === needle) || null;
}

function isRoomEmail(email) {
  return Boolean(findRoom(email));
}

/* El navegador recuerda la última sala para no volver a elegirla cada vez. */
function readStoredRoom() {
  try {
    return localStorage.getItem(ROOM_STORAGE_KEY);
  } catch {
    return null;
  }
}

function storeRoom(email) {
  try {
    localStorage.setItem(ROOM_STORAGE_KEY, email);
  } catch {
    // Modo privado o almacenamiento bloqueado: la selección solo dura la sesión.
  }
}

function renderRoomOptions() {
  roomSelect.replaceChildren();

  config.rooms.forEach((room) => {
    const option = document.createElement('option');
    option.value = room.email;
    option.textContent = room.name;
    roomSelect.appendChild(option);
  });

  selectedRoom =
    findRoom(readStoredRoom()) || findRoom(selectedRoom && selectedRoom.email) || config.rooms[0];
  roomSelect.value = selectedRoom.email;
}

function selectRoom(email) {
  const room = findRoom(email);
  if (!room) return;

  selectedRoom = room;
  storeRoom(room.email);
  showError('');
  render();
}

/* ── Directorio ─────────────────────────────────────── */
const MIN_QUERY = 2;
const SUGGESTION_LIMIT = 8;

function fold(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function guestEmails() {
  return guests.map((guest) => guest.email);
}

function lookupPerson(email) {
  const needle = (email || '').trim().toLowerCase();
  return directory.find((person) => person.email === needle) || null;
}

async function fetchFirstJson(paths) {
  for (const path of paths) {
    try {
      const response = await fetch(path);
      if (!response.ok) continue;
      return await response.json();
    } catch {
      // Probar la siguiente ruta (API local o archivos estáticos en hosting).
    }
  }
  return null;
}

async function loadDirectory() {
  if (directoryPromise) return directoryPromise;

  directoryPromise = fetchFirstJson(['/api/directory', 'directory.json'])
    .then((data) => {
      directory = data && Array.isArray(data.people) ? data.people : [];
      return directory;
    })
    .catch(() => {
      directory = [];
      directoryPromise = null;
      return directory;
    });

  return directoryPromise;
}

function searchDirectory(query) {
  const q = fold(query);
  if (q.length < MIN_QUERY) return [];

  const taken = new Set([...guestEmails(), selectedRoom.email.toLowerCase()]);
  const scored = [];

  for (const person of directory) {
    if (taken.has(person.email)) continue;

    const name = fold(person.name);
    const email = fold(person.email);
    const title = fold(person.title);
    const words = name.split(/\s+/);
    let score = 0;

    if (name.startsWith(q)) score = 40;
    else if (words.some((word) => word.startsWith(q))) score = 30;
    else if (name.includes(q)) score = 20;
    else if (email.startsWith(q) || email.includes(q)) score = 10;
    else if (title.includes(q)) score = 5;
    else continue;

    scored.push({ person, score });
  }

  scored.sort(
    (a, b) => b.score - a.score || a.person.name.localeCompare(b.person.name, 'es')
  );
  return scored.slice(0, SUGGESTION_LIMIT).map((item) => item.person);
}

function hideSuggestions() {
  suggestions = [];
  activeSuggestion = -1;
  suggestionsEl.replaceChildren();
  suggestionsEl.hidden = true;
  emailInput.setAttribute('aria-expanded', 'false');
  emailInput.removeAttribute('aria-activedescendant');
}

function setActiveSuggestion(index) {
  const items = [...suggestionsEl.querySelectorAll('[role="option"]')];
  if (!items.length) {
    activeSuggestion = -1;
    return;
  }

  activeSuggestion = (index + items.length) % items.length;
  items.forEach((item, i) => {
    item.classList.toggle('is-active', i === activeSuggestion);
    if (i === activeSuggestion) {
      emailInput.setAttribute('aria-activedescendant', item.id);
      item.scrollIntoView({ block: 'nearest' });
    }
  });
}

function renderSuggestions(query) {
  suggestions = searchDirectory(query);
  suggestionsEl.replaceChildren();

  if (fold(query).length < MIN_QUERY) {
    hideSuggestions();
    return;
  }

  if (!suggestions.length) {
    const empty = document.createElement('li');
    empty.className = 'guest-suggestion guest-suggestion-empty';
    empty.textContent = directory.length
      ? 'Sin coincidencias en el directorio. También puedes escribir un correo @gm.com.'
      : 'Escribe un correo @gm.com para agregar al invitado.';
    suggestionsEl.appendChild(empty);
    suggestionsEl.hidden = false;
    emailInput.setAttribute('aria-expanded', 'true');
    activeSuggestion = -1;
    return;
  }

  suggestions.forEach((person, index) => {
    const item = document.createElement('li');
    item.id = `guest-option-${index}`;
    item.className = 'guest-suggestion';
    item.setAttribute('role', 'option');

    const nameEl = document.createElement('span');
    nameEl.className = 'guest-suggestion-name';
    nameEl.textContent = person.name;

    const meta = document.createElement('span');
    meta.className = 'guest-suggestion-meta';
    meta.textContent = [person.email, person.title || person.category]
      .filter(Boolean)
      .join(' · ');

    item.append(nameEl, meta);
    item.addEventListener('mousedown', (event) => {
      event.preventDefault();
      addPerson(person);
    });
    suggestionsEl.appendChild(item);
  });

  suggestionsEl.hidden = false;
  emailInput.setAttribute('aria-expanded', 'true');
  setActiveSuggestion(0);
}

async function onGuestInput() {
  showError('');
  const query = emailInput.value;
  await loadDirectory();
  renderSuggestions(query);
}

function onGuestKeydown(event) {
  if (event.key === 'Enter') {
    event.preventDefault();
    if (!suggestionsEl.hidden && suggestions[activeSuggestion]) {
      addPerson(suggestions[activeSuggestion]);
      return;
    }
    addGuest();
    return;
  }

  if (suggestionsEl.hidden || !suggestions.length) return;

  if (event.key === 'ArrowDown') {
    event.preventDefault();
    setActiveSuggestion(activeSuggestion + 1);
  } else if (event.key === 'ArrowUp') {
    event.preventDefault();
    setActiveSuggestion(activeSuggestion - 1);
  } else if (event.key === 'Escape') {
    event.preventDefault();
    hideSuggestions();
  }
}

/* ── Invitados ──────────────────────────────────────── */
function addPerson(person) {
  if (!person || !person.email) return;
  addGuestEmail(person.email, person.name);
}

function addGuest() {
  const value = emailInput.value.trim();
  showError('');

  if (!value) return;

  if (!suggestionsEl.hidden && suggestions[activeSuggestion]) {
    addPerson(suggestions[activeSuggestion]);
    return;
  }

  const matches = searchDirectory(value);
  if (!EMAIL_RE.test(value.toLowerCase())) {
    if (matches.length === 1) {
      addPerson(matches[0]);
      return;
    }
    if (matches.length > 1) {
      showError('Hay varias coincidencias. Elige a la persona de la lista.');
      renderSuggestions(value);
      return;
    }
  }

  addGuestEmail(value);
}

function addGuestEmail(rawEmail, name = '') {
  const value = rawEmail.trim().toLowerCase();
  showError('');

  if (!value) return;

  if (!EMAIL_RE.test(value)) {
    showError('Busca un nombre del directorio o ingresa un correo válido.');
    return;
  }
  if (!value.endsWith(`@${config.allowedDomain}`)) {
    showError(`Solo se permiten correos @${config.allowedDomain}.`);
    return;
  }
  if (value === selectedRoom.email.toLowerCase()) {
    showError('La sala ya está incluida en la invitación.');
    return;
  }
  if (isRoomEmail(value)) {
    showError('Ese correo es de otra sala; elígela en la lista de arriba.');
    return;
  }
  if (guestEmails().includes(value)) {
    showError('Este invitado ya fue agregado.');
    return;
  }

  const person = lookupPerson(value);
  guests.push({
    email: value,
    name: name || (person && person.name) || '',
  });
  lastAddedGuest = value;
  emailInput.value = '';
  hideSuggestions();
  render();
}

function removeGuest(email) {
  guests = guests.filter((guest) => guest.email !== email);
  showError('');
  render();
}

/* ── Render ─────────────────────────────────────────── */
function render() {
  const title = selectedRoom.name.toUpperCase();
  const titleChanged = roomTitle.textContent !== title;

  roomTitle.textContent = title;
  if (titleChanged && hasRendered) replayAnimation(roomTitle, 'is-swapping');

  roomNote.textContent = `Buzón de la sala: ${selectedRoom.email}`;
  document.title = `Reservar ${selectedRoom.name} — GM`;

  pillRow.replaceChildren();
  pillRow.appendChild(buildPill(selectedRoom.email, { locked: true }));
  guests.forEach((guest) => pillRow.appendChild(buildPill(guest)));
  lastAddedGuest = null;

  const total = guests.length + 1;
  countNote.textContent = `${total} destinatario${total === 1 ? '' : 's'} (sala incluida)`;

  updateOutlookLink();
  hasRendered = true;
}

function buildPill(guest, { locked = false } = {}) {
  const email = typeof guest === 'string' ? guest : guest.email;
  const name = typeof guest === 'string' ? '' : guest.name;
  const pill = document.createElement('div');
  pill.className = locked ? 'pill locked' : 'pill';
  if (email === lastAddedGuest) pill.classList.add('is-new');

  const icon = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  icon.setAttribute('class', 'icon');
  icon.setAttribute('aria-hidden', 'true');
  const use = document.createElementNS('http://www.w3.org/2000/svg', 'use');
  use.setAttribute('href', locked ? '#i-door' : '#i-user');
  icon.appendChild(use);
  pill.appendChild(icon);

  if (name) {
    const text = document.createElement('span');
    text.className = 'pill-text';
    const nameEl = document.createElement('span');
    nameEl.className = 'pill-name';
    nameEl.textContent = name;
    const mailEl = document.createElement('span');
    mailEl.className = 'pill-email';
    mailEl.textContent = email;
    text.append(nameEl, mailEl);
    pill.appendChild(text);
    pill.title = email;
  } else {
    pill.appendChild(document.createTextNode(email));
  }

  if (!locked) {
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'pill-remove';
    remove.setAttribute('aria-label', `Quitar ${name || email}`);
    remove.textContent = '×';
    remove.addEventListener('click', () => dismissPill(pill, email));
    pill.appendChild(remove);
  }

  return pill;
}

/* Deja correr la animación de salida antes de rearmar la lista. */
function dismissPill(pill, email) {
  if (prefersReducedMotion()) {
    removeGuest(email);
    return;
  }

  let settled = false;
  const finish = () => {
    if (settled) return;
    settled = true;
    removeGuest(email);
  };

  pill.classList.add('is-leaving');
  pill.addEventListener('animationend', finish, { once: true });
  setTimeout(finish, 300);
}

function showError(msg) {
  setMessage(errorMsg, msg);
}

function setMessage(el, msg) {
  if (el.textContent === msg) return;

  el.textContent = msg;
  if (msg) replayAnimation(el, 'is-shown');
  else el.classList.remove('is-shown');
}

/* ── Deeplink de Outlook ────────────────────────────── */
function buildOutlookUrl() {
  const subject = subjectInput.value.trim() || `Reunión — ${selectedRoom.name}`;
  const attendees = [selectedRoom.email, ...guestEmails()].join(',');

  const params = new URLSearchParams({
    path: '/calendar/action/compose',
    rru: 'addevent',
    subject,
    location: selectedRoom.name,
    to: attendees,
    allday: 'false',
  });

  return `${OUTLOOK_COMPOSE_URL}?${params.toString()}`;
}

function updateOutlookLink() {
  const url = buildOutlookUrl();
  openBtn.href = url;
  openBtn.removeAttribute('aria-disabled');
  copyBtn.removeAttribute('aria-disabled');
  resetCopyLabel();
  return url;
}

async function copyLink() {
  const url = buildOutlookUrl();
  if (!url) return;

  try {
    await navigator.clipboard.writeText(url);
  } catch {
    // clipboard API requiere HTTPS o localhost; en LAN por HTTP cae aquí
    const helper = document.createElement('textarea');
    helper.value = url;
    helper.setAttribute('readonly', '');
    helper.style.position = 'fixed';
    helper.style.opacity = '0';
    document.body.appendChild(helper);
    helper.select();
    document.execCommand('copy');
    helper.remove();
  }

  copyLabel.textContent = '¡Liga copiada!';
  copyBtn.classList.add('is-done');
  setTimeout(resetCopyLabel, 2000);
}

function resetCopyLabel() {
  copyLabel.textContent = 'Copiar liga de la reunión';
  copyBtn.classList.remove('is-done');
}

/* ── Configuración del servidor ─────────────────────── */
async function loadConfig() {
  try {
    const data = await fetchFirstJson(['/api/config', 'config.json']);
    if (!data) return;
    if (!Array.isArray(data.rooms) || data.rooms.length === 0) delete data.rooms;
    Object.assign(config, data);

    renderRoomOptions();
    emailInput.placeholder = `Nombre o usuario@${config.allowedDomain}`;
  } catch {
    // La página funciona con los valores por defecto si el API no responde.
  }
}

/* ── Listeners ──────────────────────────────────────── */
form.addEventListener('submit', (e) => e.preventDefault());

roomSelect.addEventListener('change', () => selectRoom(roomSelect.value));

addBtn.addEventListener('click', addGuest);

emailInput.addEventListener('input', onGuestInput);
emailInput.addEventListener('focus', onGuestInput);
emailInput.addEventListener('keydown', onGuestKeydown);
emailInput.addEventListener('blur', () => {
  setTimeout(hideSuggestions, 120);
});

subjectInput.addEventListener('input', updateOutlookLink);
subjectInput.addEventListener('change', updateOutlookLink);

copyBtn.addEventListener('click', copyLink);

/* ── Init ───────────────────────────────────────────── */
renderRoomOptions();
render();
loadDirectory();
loadConfig().then(render);
