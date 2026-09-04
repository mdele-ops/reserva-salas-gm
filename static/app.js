/* ─────────────────────────────────────────────────────
   RESERVA DE SALAS — Reserva vía Outlook Calendar
   En computadora el botón es un <a> real a Outlook Web
   (los bloqueadores de pop-ups no lo cortan). En el
   teléfono se abre la app de Outlook con el borrador.
   ───────────────────────────────────────────────────── */

const OUTLOOK_WEB_COMPOSE = 'https://outlook.office.com/calendar/action/compose';
const OUTLOOK_MOBILE_WEB = 'https://outlook.cloud.microsoft/owa';
const OUTLOOK_APP_COMPOSE = 'ms-outlook://events/new';

const RULESET_JUNTAS = 'juntas';

// Solo estas 5 llevan el reglamento actual. Salas nuevas: omitir `rules`.
const config = {
  rooms: [
    { name: 'Sala Colibrí',  email: 'SRR270201@gm.com', rules: RULESET_JUNTAS },
    { name: 'Sala Maya',     email: 'SRR271627@gm.com', rules: RULESET_JUNTAS },
    { name: 'Sala Alebrije', email: 'SRR240332@gm.com', rules: RULESET_JUNTAS },
    { name: 'Sala Ajolote',  email: 'SRR265183@gm.com', rules: RULESET_JUNTAS },
    { name: 'Sala Alfeñique', email: 'SRR210381@gm.com', rules: RULESET_JUNTAS },
  ],
  allowedDomain: 'gm.com',
  defaultDurationMin: 60,
};

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const ROOM_STORAGE_KEY = 'salas.lastRoomEmail';
const MAYA_EMAIL = 'SRR271627@gm.com';
const JUNTAS_ROOM_EMAILS = new Set([
  'srr270201@gm.com',
  'srr271627@gm.com',
  'srr240332@gm.com',
  'srr265183@gm.com',
  'srr210381@gm.com',
]);

const ROOM_RULES_BODY = [
  'REGLAMENTO DE USO DE SALAS DE JUNTAS',
  '',
  '1. Orden y limpieza',
  '• Configuración original: acomode sillas y mesas como se encontraron.',
  '• Cero basura: llévese envases y envolturas. No deje derrames.',
  '• Consumo: solo en Sala Maya se permiten snacks; en las demás, únicamente agua.',
  '',
  '2. Cuidado del mobiliario y equipo',
  '• Uso adecuado: evite sentarse en descansabrazos y no raye mesas.',
  '• Equipo de TI: apague pantallas/proyectores y ordene cables.',
  '• Reporte de daños: infórmelos de inmediato para evitar que se le responsabilice.',
  '',
  '3. Gestión de horarios',
  '• Puntualidad de salida: desocupe un par de minutos antes de finalizar.',
  '• Libere si no usa: cancele la reserva de inmediato si no se ocupará.',
  '• Salas "fantasma": pasados 10 minutos de tolerancia, la sala queda libre.',
  '',
  '4. Incumplimiento',
  'En caso de daño, desperfecto o mal uso reiterado se aplicarán las medidas del Reglamento Interno de Trabajo.',
].join('\n');

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
const rulesGroup   = document.getElementById('rules-group');
const headerSub    = document.getElementById('header-sub');
const outlookHint  = document.getElementById('outlook-hint');

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

function roomSlug(room) {
  return fold(room && room.name)
    .replace(/^sala\s+/, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function findRoomBySlug(slug) {
  const needle = fold(slug)
    .replace(/^sala\s+/, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
  if (!needle) return null;
  return config.rooms.find((room) => roomSlug(room) === needle) || null;
}

function findRoomByIdentifier(value) {
  const raw = String(value || '').trim();
  if (!raw) return null;
  return (
    findRoom(raw) ||
    findRoomBySlug(raw) ||
    config.rooms.find((room) => fold(room.name) === fold(raw)) ||
    null
  );
}

function readRoomFromUrl() {
  try {
    const params = new URLSearchParams(window.location.search);
    const fromQuery = findRoomByIdentifier(params.get('sala') || params.get('room'));
    if (fromQuery) return fromQuery;

    const hash = decodeURIComponent((window.location.hash || '').replace(/^#/, '').trim());
    return findRoomByIdentifier(hash);
  } catch {
    return null;
  }
}

function syncRoomUrl(room) {
  if (!room || !window.history || typeof window.history.replaceState !== 'function') return;
  try {
    const url = new URL(window.location.href);
    const slug = roomSlug(room);
    if (!slug) return;
    if (url.searchParams.get('sala') === slug && !url.searchParams.has('room')) return;

    url.searchParams.set('sala', slug);
    url.searchParams.delete('room');
    const hashValue = (url.hash || '').replace(/^#/, '');
    if (hashValue && findRoomByIdentifier(hashValue)) url.hash = '';

    window.history.replaceState({ sala: slug }, '', `${url.pathname}${url.search}${url.hash}`);
  } catch {
    // El identificador es informativo; si el historial no se puede tocar, no bloquea la reserva.
  }
}

function roomRules(room = selectedRoom) {
  if (!room) return '';
  if (room.rules) return room.rules;
  return JUNTAS_ROOM_EMAILS.has((room.email || '').toLowerCase()) ? RULESET_JUNTAS : '';
}

function roomHasJuntasRules(room = selectedRoom) {
  return roomRules(room) === RULESET_JUNTAS;
}

function applyKnownRules() {
  config.rooms.forEach((room) => {
    if (!room.rules && JUNTAS_ROOM_EMAILS.has((room.email || '').toLowerCase())) {
      room.rules = RULESET_JUNTAS;
    }
  });
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
    readRoomFromUrl() ||
    findRoom(readStoredRoom()) ||
    findRoom(selectedRoom && selectedRoom.email) ||
    config.rooms[0];
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
      ? 'Sin coincidencias en el directorio. También puede ingresar un correo @gm.com.'
      : 'Ingrese un correo @gm.com para agregar al invitado.';
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
      showError('Hay varias coincidencias. Seleccione a la persona de la lista.');
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
    showError('Busque un nombre del directorio o ingrese un correo válido.');
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
    showError('Ese correo corresponde a otra sala; selecciónela en la lista superior.');
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
  document.title = `Reservar ${selectedRoom.name} — Complejo Toluca`;
  syncRoomUrl(selectedRoom);

  const hasJuntasRules = roomHasJuntasRules();
  if (rulesGroup) {
    rulesGroup.hidden = !hasJuntasRules;
    const isMaya =
      hasJuntasRules &&
      (selectedRoom.email.toLowerCase() === MAYA_EMAIL ||
        fold(selectedRoom.name).includes('maya'));
    rulesGroup.classList.toggle('is-maya', isMaya);
  }
  if (headerSub) {
    headerSub.textContent = hasJuntasRules
      ? 'Consulte el reglamento, seleccione la sala e invite a los participantes'
      : 'Seleccione la sala e invite a los participantes';
  }
  if (outlookHint) {
    outlookHint.textContent = hasJuntasRules
      ? 'En dispositivos móviles se abre la aplicación con el borrador. Fecha y horario se definen ahí. El reglamento se incluye en el cuerpo de la invitación.'
      : 'En dispositivos móviles se abre la aplicación con el borrador. Fecha y horario se definen ahí.';
  }

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
function isPhone() {
  return (
    /Android|iPhone|iPad|iPod/i.test(navigator.userAgent) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1)
  );
}

function isAndroid() {
  return /Android/i.test(navigator.userAgent);
}

function meetingPayload() {
  const subject = subjectInput.value.trim() || `Reunión — ${selectedRoom.name}`;
  const body = roomHasJuntasRules() ? ROOM_RULES_BODY : '';
  return {
    subject,
    location: selectedRoom.name,
    attendees: [selectedRoom.email, ...guestEmails()].join(','),
    body,
    bodyHtml: body ? body.replace(/\n/g, '<br>') : '',
  };
}

function buildWebOutlookUrl() {
  const meeting = meetingPayload();
  const params = new URLSearchParams({
    path: '/calendar/action/compose',
    rru: 'addevent',
    subject: meeting.subject,
    location: meeting.location,
    to: meeting.attendees,
    allday: 'false',
  });
  if (meeting.bodyHtml) params.set('body', meeting.bodyHtml);
  return `${OUTLOOK_WEB_COMPOSE}?${params.toString()}`;
}

function buildMobileWebOutlookUrl() {
  const meeting = meetingPayload();
  const params = new URLSearchParams({
    path: 'calendar/action/compose',
    rru: 'addevent',
    subject: meeting.subject,
    location: meeting.location,
    to: meeting.attendees,
    allday: 'false',
  });
  if (meeting.bodyHtml) params.set('body', meeting.bodyHtml);
  return `${OUTLOOK_MOBILE_WEB}?${params.toString()}`;
}

function encodeAppQuery(params) {
  return Object.entries(params)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
    .join('&');
}

function buildAppQuery() {
  const meeting = meetingPayload();
  const params = {
    title: meeting.subject,
    location: meeting.location,
    attendees: meeting.attendees,
  };
  if (meeting.body) params.description = meeting.body;
  return encodeAppQuery(params);
}

function buildAppOutlookUrl() {
  return `${OUTLOOK_APP_COMPOSE}?${buildAppQuery()}`;
}

function buildAndroidIntentUrl() {
  const fallback = encodeURIComponent(buildMobileWebOutlookUrl());
  return `intent://events/new?${buildAppQuery()}#Intent;scheme=ms-outlook;package=com.microsoft.office.outlook;S.browser_fallback_url=${fallback};end`;
}

function buildOutlookUrl() {
  if (isAndroid()) return buildAndroidIntentUrl();
  if (isPhone()) return buildAppOutlookUrl();
  return buildWebOutlookUrl();
}

function updateOutlookLink() {
  const url = buildOutlookUrl();
  openBtn.href = url;
  if (isPhone()) {
    openBtn.removeAttribute('target');
    openBtn.removeAttribute('rel');
  } else {
    openBtn.target = '_blank';
    openBtn.rel = 'noopener noreferrer';
  }
  openBtn.removeAttribute('aria-disabled');
  copyBtn.removeAttribute('aria-disabled');
  resetCopyLabel();
  return url;
}

function openOutlook(event) {
  if (!isPhone()) return;

  event.preventDefault();

  if (isAndroid()) {
    window.location.href = buildAndroidIntentUrl();
    return;
  }

  const appUrl = buildAppOutlookUrl();
  const webUrl = buildMobileWebOutlookUrl();
  let handedOff = false;
  const markHandedOff = () => {
    handedOff = true;
  };

  window.addEventListener('pagehide', markHandedOff, { once: true });
  window.addEventListener('blur', markHandedOff, { once: true });
  document.addEventListener(
    'visibilitychange',
    () => {
      if (document.hidden) markHandedOff();
    },
    { once: true }
  );

  window.location.href = appUrl;

  window.setTimeout(() => {
    if (!handedOff && document.visibilityState === 'visible') {
      window.location.href = webUrl;
    }
  }, 1400);
}

async function copyLink() {
  const url = buildWebOutlookUrl();
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

  copyLabel.textContent = 'Liga copiada';
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
    applyKnownRules();

    renderRoomOptions();
    emailInput.placeholder = `Nombre o usuario@${config.allowedDomain}`;
  } catch {
    // La página funciona con los valores por defecto si el API no responde.
  }
}

/* ── Listeners ──────────────────────────────────────── */
form.addEventListener('submit', (e) => e.preventDefault());

roomSelect.addEventListener('change', () => selectRoom(roomSelect.value));

window.addEventListener('popstate', () => {
  const room = readRoomFromUrl();
  if (!room || room.email === selectedRoom.email) return;
  selectedRoom = room;
  storeRoom(room.email);
  roomSelect.value = room.email;
  showError('');
  render();
});

addBtn.addEventListener('click', addGuest);

emailInput.addEventListener('input', onGuestInput);
emailInput.addEventListener('focus', onGuestInput);
emailInput.addEventListener('keydown', onGuestKeydown);
emailInput.addEventListener('blur', () => {
  setTimeout(hideSuggestions, 120);
});

subjectInput.addEventListener('input', updateOutlookLink);
subjectInput.addEventListener('change', updateOutlookLink);

openBtn.addEventListener('click', openOutlook);
copyBtn.addEventListener('click', copyLink);

/* ── Init ───────────────────────────────────────────── */
applyKnownRules();
renderRoomOptions();
render();
loadDirectory();
loadConfig().then(render);
