const courses = ['1er año', '2do año', '3er año', '4to año', '5to año', '6to año'];
const galaActivityNames = ['Dubsmash', 'Videoclip', 'Baile Grupal', 'Baile en Pareja'];
function normalizeGalaActivities(storedActivities = {}) {
	return Object.fromEntries(courses.map((course) => [course, galaActivityNames.map((activity) => {
		const stored = storedActivities[course]?.find((entry) => entry.activity === activity);
		return { activity, topic: stored?.topic || '', url: stored?.url || '' };
	})]));
}
const data = {
	standings: courses.map((curso) => ({ curso, pts: 0, details: {} })),
	calendar: [],
	results: [],
	rules: [],
	galaActivities: Object.fromEntries(courses.map((course) => [course, []]))
};
const today = '2026-08-26';
const tomorrow = '2026-08-27';
let selectedCourse = localStorage.getItem('copa-course');
const defaultCourseKeys = Object.fromEntries(courses.map((course, index) => [course, `curso${index + 1}2026`]));
let courseKeys = defaultCourseKeys;
try {
	const storedCourseKeys = JSON.parse(localStorage.getItem('copa-course-keys') || 'null');
	if (storedCourseKeys && typeof storedCourseKeys === 'object') courseKeys = { ...defaultCourseKeys, ...storedCourseKeys };
} catch { }
try {
	const storedRegulation = JSON.parse(localStorage.getItem('copa-regulation') || 'null');
	if (storedRegulation?.calendar) { data.calendar = storedRegulation.calendar; data.rules = storedRegulation.rules || []; data.galaActivities = normalizeGalaActivities(storedRegulation.galaActivities || {}); }
} catch { }
try {
	const storedScores = JSON.parse(localStorage.getItem('copa-scores') || 'null');
	if (storedScores?.standings) data.standings = storedScores.standings;
	if (storedScores?.results) data.results = storedScores.results;
} catch { }
data.standings.sort((left, right) => Number(right.pts) - Number(left.pts) || courses.indexOf(left.curso) - courses.indexOf(right.curso));
if (!localStorage.getItem('copa-dubsmash-topics-cleared')) {
	courses.forEach((course) => { const dubsmash = data.galaActivities[course]?.find((entry) => entry.activity === 'Dubsmash'); if (dubsmash) dubsmash.topic = ''; });
	localStorage.removeItem('copa-gala-topics');
	localStorage.setItem('copa-regulation', JSON.stringify({ calendar: data.calendar, rules: data.rules, galaActivities: data.galaActivities }));
	localStorage.setItem('copa-dubsmash-topics-cleared', 'true');
}
let moderatorKey = localStorage.getItem('copa-moderator-key') || 'IPDVS2026';
const esc = (value) => String(value).replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[char]));
function enhanceGalaEditor() {
	const course = document.getElementById('gala-course');
	const activity = document.getElementById('gala-activity');
	const link = document.getElementById('gala-link');
	if (!course || !activity || !link || document.getElementById('gala-topic')) return;
	link.parentElement.insertAdjacentHTML('beforebegin', '<label>Tópico asignado<input id="gala-topic" type="text" placeholder="Ej: Escena los Argento"></label>');
	const topic = document.getElementById('gala-topic');
	const sync = () => { const entry = data.galaActivities[course.value]?.[activity.value]; if (!entry) { topic.value = ''; link.value = ''; return; } topic.value = entry.topic; link.value = entry.url; };
	course.addEventListener('change', sync);
	activity.addEventListener('change', sync);
	document.getElementById('save-gala-link').addEventListener('click', () => {
		const entry = data.galaActivities[course.value][activity.value];
		const topicValue = topic.value.trim();
		if (!topicValue) return showToast('Escribí un tópico');
		entry.topic = topicValue;
		const topics = Object.fromEntries(courses.map((entryCourse) => [entryCourse, data.galaActivities[entryCourse].map((activityEntry) => activityEntry.topic)]));
		localStorage.setItem('copa-gala-topics', JSON.stringify(topics));
		localStorage.setItem('copa-regulation', JSON.stringify({ calendar: data.calendar, rules: data.rules, galaActivities: data.galaActivities }));
		showToast('Tópico guardado para este curso');
	});
	sync();
}
document.addEventListener('DOMContentLoaded', () => setTimeout(enhanceGalaEditor, 0));
async function readRegulation(file) {
	if (file.name.toLowerCase().endsWith('.docx')) return (await window.mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() })).value;
	const pdf = await window.pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
	const pages = [];
	for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) pages.push((await (await pdf.getPage(pageNumber)).getTextContent()).items.map((item) => item.str).join(' '));
	return pages.join('\n');
}
function applyRegulation(text, fileName) {
	const activities = [];
	text.split(/\n+/).map((line) => line.trim()).filter(Boolean).forEach((line) => {
		const match = line.match(/^(.+?)\s*(?:\||[-–])\s*(\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?)\s*(?:\||[-–])\s*(\d+)\s*(?:puntos?)?$/i);
		if (match) { const [day, month, year = '2026'] = match[2].replaceAll('/', '-').split('-'); activities.push({ title: match[1].trim(), date: `${year.length === 2 ? `20${year}` : year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`, time: '00:00', points: `${match[3]}`, type: 'Reglamento' }); }
	});
	if (!activities.length) throw new Error('No se encontraron líneas con el formato actividad | fecha | puntos.');
	data.calendar = activities;
	data.rules = activities.map((activity) => ({ title: activity.title, points: `${activity.points} puntos`, text: `Actividad publicada en ${fileName}.` }));
	data.results = [];
	data.galaActivities = normalizeGalaActivities();
	localStorage.setItem('copa-regulation', JSON.stringify({ calendar: data.calendar, rules: data.rules, galaActivities: data.galaActivities, fileName }));
	renderAll();
	openScreen('admin');
}
function enhanceRegulationUploader() {
	const fields = document.getElementById('admin-fields');
	if (!fields || document.getElementById('activity-title')) return;
	fields.insertAdjacentHTML('afterbegin', '<p class="section-heading">REGLAMENTO OFICIAL</p><label>Actividad<input id="activity-title" type="text" placeholder="Ej: Dubsmash"></label><label>Fecha y hora<input id="activity-date" type="date"><input id="activity-time" type="time"></label><label>Puntos<input id="activity-points" type="number" min="0" step="1" inputmode="numeric" placeholder="Solo números"></label><button type="button" class="primary" id="regulation-upload">Subir</button><small id="regulation-status" class="form-hint">Cada vez que subas una actividad se agrega al reglamento y al calendario.</small>');
	fields.insertAdjacentHTML('beforeend', `<p class="section-heading">ELIMINAR ACTIVIDADES</p><div class="activity-delete-list">${data.calendar.length ? data.calendar.map((activity, index) => `<div class="activity-delete-row"><span><strong>${esc(activity.title)}</strong><small>${activity.date} · ${activity.points} puntos</small></span><button type="button" class="outline delete-activity" data-activity-index="${index}">Eliminar</button></div>`).join('') : '<p class="empty">No hay actividades cargadas.</p>'}</div>`);
	fields.insertAdjacentHTML('beforeend', `<p class="section-heading">BORRAR TÓPICOS DE GALA</p><label>Curso<select id="clear-topic-course">${courses.map((course) => `<option>${course}</option>`).join('')}</select></label><label>Actividad<select id="clear-topic-activity">${galaActivityNames.map((activity) => `<option>${activity}</option>`).join('')}</select></label><button type="button" class="outline" id="clear-gala-topic">Borrar tópico</button>`);
	fields.insertAdjacentHTML('beforeend', '<p class="section-heading">LIMPIAR RESULTADOS</p><p class="form-hint">Una vez finalizada la Copa Vélez, presioná este botón para eliminar todos los resultados de este año.</p><button type="button" class="outline" id="clear-results">Eliminar resultados</button>');
	fields.insertAdjacentHTML('beforeend', `<p class="section-heading">RESULTADOS</p><label>Actividad<select id="result-activity">${data.calendar.map((activity, index) => `<option value="${index}">${esc(activity.title)}</option>`).join('')}</select></label><label>Curso ganador<select id="result-course">${courses.map((course) => `<option>${course}</option>`).join('')}</select></label><output id="result-points-display" class="points-from-rules">Puntos según reglamento: ${data.calendar[0]?.points || 0}</output><button class="primary" id="save-result">Guardar resultado</button>`);
	document.querySelectorAll('.delete-activity').forEach((button) => button.addEventListener('click', () => {
		const index = Number(button.dataset.activityIndex);
		const activity = data.calendar[index];
		if (!activity || !window.confirm(`¿Eliminar la actividad “${activity.title}”?`)) return;
		data.calendar.splice(index, 1);
		data.rules.splice(index, 1);
		courses.forEach((course) => data.galaActivities[course]?.splice(index, 1));
		localStorage.setItem('copa-regulation', JSON.stringify({ calendar: data.calendar, rules: data.rules, galaActivities: data.galaActivities }));
		renderAll();
		openScreen('admin');
		document.getElementById('admin-fields').hidden = false;
		enhanceRegulationUploader();
		showToast('Actividad eliminada');
	}));
	document.getElementById('clear-gala-topic').addEventListener('click', () => {
		const course = document.getElementById('clear-topic-course').value;
		const activity = document.getElementById('clear-topic-activity').value;
		const entry = data.galaActivities[course]?.find((galaActivity) => galaActivity.activity === activity);
		if (!entry) return;
		entry.topic = '';
		localStorage.setItem('copa-regulation', JSON.stringify({ calendar: data.calendar, rules: data.rules, galaActivities: data.galaActivities }));
		renderAll();
		openScreen('admin');
		document.getElementById('admin-fields').hidden = false;
		enhanceRegulationUploader();
		showToast('Tópico borrado');
	});
	document.getElementById('clear-results').addEventListener('click', () => {
		if (!window.confirm('¿Eliminar todos los resultados y reiniciar los puntos de la tabla?')) return;
		data.results = [];
		data.standings.forEach((standing) => { standing.pts = 0; standing.details = {}; });
		localStorage.setItem('copa-scores', JSON.stringify({ standings: data.standings, results: data.results }));
		renderAll();
		openScreen('admin');
		document.getElementById('admin-fields').hidden = false;
		enhanceRegulationUploader();
		showToast('Todos los resultados fueron eliminados');
	});
	document.getElementById('result-activity')?.addEventListener('change', (event) => { document.getElementById('result-points-display').textContent = `Puntos según reglamento: ${data.calendar[Number(event.target.value)]?.points || 0}`; });
	document.getElementById('save-result')?.addEventListener('click', () => {
		const activity = data.calendar[Number(document.getElementById('result-activity').value)];
		const course = document.getElementById('result-course').value;
		const points = Number(activity?.points);
		if (!activity || !points) return showToast('Elegí una actividad y puntos válidos');
		const standing = data.standings.find((entry) => entry.curso === course);
		standing.pts += points;
		standing.details[activity.title] = (standing.details[activity.title] || 0) + points;
		data.standings.sort((left, right) => Number(right.pts) - Number(left.pts) || courses.indexOf(left.curso) - courses.indexOf(right.curso));
		data.results.push({ title: activity.title, result: `${course} — ganador`, detail: 'Resultado oficial', points: `+${points} para ${course}` });
		localStorage.setItem('copa-scores', JSON.stringify({ standings: data.standings, results: data.results }));
		renderAll();
		openScreen('admin');
		document.getElementById('admin-fields').hidden = false;
		enhanceRegulationUploader();
		showToast('Resultado guardado y puntos sumados');
	});
	document.getElementById('regulation-upload').addEventListener('click', () => {
		const title = document.getElementById('activity-title').value.trim();
		const date = document.getElementById('activity-date').value;
		const time = document.getElementById('activity-time').value;
		const points = document.getElementById('activity-points').value;
		if (!title || !date || !time || !/^\d+$/.test(points)) return showToast('Completá actividad, fecha, hora y puntos');
		const activity = { title, date, time, points, type: 'Reglamento' };
		data.calendar.push(activity);
		data.rules.push({ title, points: `${points} puntos`, text: 'Actividad publicada en el reglamento.' });
		localStorage.setItem('copa-regulation', JSON.stringify({ calendar: data.calendar, rules: data.rules, galaActivities: data.galaActivities }));
		renderAll();
		openScreen('admin');
		const adminFields = document.getElementById('admin-fields');
		adminFields.hidden = false;
		enhanceRegulationUploader();
		showToast('Actividad agregada al reglamento');
	});
	setTimeout(() => { enhanceGalaEditor(); enhanceRegulationUploader(); }, 0);
}
document.addEventListener('DOMContentLoaded', () => setTimeout(enhanceRegulationUploader, 0));
const dateLabel = (date) => Object.fromEntries(new Intl.DateTimeFormat('es-AR', { day: '2-digit', month: 'short' }).formatToParts(new Date(`${date}T12:00:00`)).filter((part) => part.type === 'day' || part.type === 'month').map((part) => [part.type, part.value.toUpperCase()]));
function eventRows(events) { return events.length ? events.map((event) => { const { day, month } = dateLabel(event.date); return `<article class="event-row"><time datetime="${event.date}"><b>${day}</b><small>${month}</small></time><div><strong>${esc(event.title)}</strong><p>${event.time} · ${event.points} puntos</p></div><span class="event-type">${esc(event.type)}</span></article>`; }).join('') : '<p class="empty">No hay actividades cargadas para este día.</p>'; }
function screen(title, id, body) { return `<section class="screen" id="screen-${id}"><header class="screen-header"><button class="back" data-close aria-label="Volver">←</button><h2>${title}</h2></header><div class="screen-body">${body}</div></section>`; }
function rankedStandings() { return [...data.standings].sort((left, right) => Number(right.pts) - Number(left.pts) || courses.indexOf(left.curso) - courses.indexOf(right.curso)); }
function renderStandings() { const standings = rankedStandings(); const leader = standings[0]?.pts || 1; return screen('Tabla de posiciones', 'standings', standings.map((item, index) => `<article class="rank ${index === 0 ? 'leader' : ''}"><span class="rank-number">${index + 1}°</span><div class="rank-main"><strong>${esc(item.curso)}</strong><div class="progress"><i style="width:${item.pts / leader * 100}%"></i></div></div><b class="points">${item.pts}<small>pts</small></b></article>`).join('')); }
function renderCalendar() { return screen('Calendario', 'calendar', `<div class="today-heading"><p class="eyebrow">ACTIVIDADES DE LA COPA</p><h3>Fechas y puntajes</h3></div><div class="event-list">${eventRows(data.calendar)}</div>`); }
function renderResults() { return screen('Resultados', 'results', data.results.length ? data.results.map((item) => `<article class="result-card"><div class="result-top"><span>${esc(item.title)}</span><b>${esc(item.points)}</b></div><h3>${esc(item.result)}</h3><p>${esc(item.detail)}</p></article>`).join('') : '<p class="empty">Todavía no hay resultados cargados.</p>'); }
function renderCourse() { if (!selectedCourse) return screen('Mi curso', 'course', `<div class="login-box"><span class="big-icon">♙</span><p class="eyebrow">ACCESO DE CURSO</p><h3>Elegí tu curso</h3><p>Ingresá la clave especial que te dio la directora.</p><label>Curso<select id="course-select">${courses.map((course) => `<option>${course}</option>`).join('')}</select></label><label>Clave<input id="course-key" type="password" placeholder="Clave del curso"></label><button class="primary" id="course-login">Ingresar</button><small class="form-hint">La clave la define y modifica la moderación.</small></div>`); const item = data.standings.find((entry) => entry.curso === selectedCourse) || data.standings[0]; const position = data.standings.indexOf(item) + 1; const above = data.standings[position - 2]; const pointsNeeded = above ? above.pts - item.pts + 1 : 0; const gala = data.galaActivities[item.curso] || []; return screen('Mi curso', 'course', `<div class="course-hero"><p>MI CURSO</p><h3>${esc(item.curso)}</h3><strong>${item.pts} PUNTOS</strong><span>${position}° PUESTO GENERAL</span></div><button class="text-button" id="change-course">Cambiar de curso</button>${above ? `<div class="climb-note"><strong>Necesitás ${pointsNeeded} puntos</strong><span>para superar a ${esc(above.curso)}, que está ${position - 1}°</span></div>` : '<div class="climb-note"><strong>Estás en el 1° puesto</strong><span>No tenés ningún curso por superar.</span></div>'}<p class="section-heading">DESGLOSE DE PUNTOS</p>${Object.entries(item.details).map(([name, value]) => `<div class="stat"><span>${name}</span><i><b style="width:${value / 320 * 100}%"></b></i><strong>${value}</strong></div>`).join('')}<p class="section-heading">TÓPICOS DE NOCHE DE GALA</p><div class="gala-list">${gala.map((entry) => `<article class="gala-row"><div><small>${esc(entry.activity)}</small><strong>${esc(entry.topic)}</strong></div><a href="${esc(entry.url)}" target="_blank" rel="noopener">Ver referencia ↗</a></article>`).join('')}</div>`); }
function renderRules() { return screen('Reglamento oficial', 'rules', `<div class="rules-intro"><p class="eyebrow">COPA VÉLEZ 2026 · IPDVS</p><h3>Reglamento y puntajes</h3><p>Acá se publica la versión oficial de las actividades y el sistema de puntos.</p></div>${data.rules.map((rule) => `<details class="rule"><summary><span>${esc(rule.title)}</span><b>${esc(rule.points)}</b></summary><p>${esc(rule.text)}</p></details>`).join('')}`); }
function renderAdmin() { return screen('Moderación', 'admin', `<div class="login-box"><span class="big-icon">⚙</span><p class="eyebrow">PANEL LOCAL</p><h3>Administrar claves y referencias</h3><p>Este panel funciona en este dispositivo. Para compartir cambios entre dispositivos se necesita una base de datos o servidor.</p><label>Clave de moderación<input id="moderator-login" type="password"></label><button class="primary" id="moderator-enter">Acceder</button><div id="admin-fields" hidden><p class="section-heading">CLAVES DE CURSO</p><label>Curso<select id="course-key-course">${courses.map((course) => `<option>${course}</option>`).join('')}</select></label><label>Nueva clave para cursos<input id="new-course-key" type="text" placeholder="Ej: gala2026"></label><button class="primary" id="save-key">Guardar clave</button><p class="section-heading">REFERENCIAS DE GALA</p><label>Curso<select id="gala-course">${courses.map((course) => `<option>${course}</option>`).join('')}</select></label><label>Actividad<select id="gala-activity">${data.galaActivities[courses[0]].map((entry, index) => `<option value="${index}">${esc(entry.activity)}</option>`).join('')}</select></label><label>Link del video<input id="gala-link" type="url" placeholder="https://www.youtube.com/watch?v=..."></label><button class="primary" id="save-gala-link">Guardar referencia</button></div></div>`); }
function renderAll() { document.getElementById('screens').innerHTML = [renderStandings(), renderCalendar(), renderResults(), renderCourse(), renderRules(), renderAdmin()].join(''); document.getElementById('podium')?.replaceChildren(); bind(); }
function openScreen(id) { document.getElementById(`screen-${id}`)?.classList.add('active'); document.body.classList.add('locked'); if (id === 'course') document.getElementById('course-select')?.focus(); }
function closeScreens() { document.querySelectorAll('.screen.active').forEach((screenEl) => screenEl.classList.remove('active')); document.body.classList.remove('locked'); }
function bind() { document.querySelectorAll('[data-screen]').forEach((button) => button.addEventListener('click', () => openScreen(button.dataset.screen))); document.querySelectorAll('[data-close]').forEach((button) => button.addEventListener('click', closeScreens)); document.getElementById('course-login')?.addEventListener('click', () => { const course = document.getElementById('course-select').value; const key = document.getElementById('course-key').value; if (key !== courseKeys[course]) return showToast('La clave no es correcta'); selectedCourse = course; localStorage.setItem('copa-course', course); renderAll(); openScreen('course'); }); document.getElementById('change-course')?.addEventListener('click', () => { selectedCourse = null; localStorage.removeItem('copa-course'); renderAll(); openScreen('course'); }); document.getElementById('moderator-enter')?.addEventListener('click', () => { if (document.getElementById('moderator-login').value !== 'MOD2026') return showToast('Clave de moderación incorrecta'); document.getElementById('admin-fields').hidden = false; }); document.getElementById('save-key')?.addEventListener('click', () => { const course = document.getElementById('course-key-course').value; const key = document.getElementById('new-course-key').value.trim(); if (!key) return showToast('Escribí una clave'); courseKeys[course] = key; localStorage.setItem('copa-course-keys', JSON.stringify(courseKeys)); showToast(`Clave actualizada para ${course}`); }); const galaCourse = document.getElementById('gala-course'); const galaActivity = document.getElementById('gala-activity'); const galaLink = document.getElementById('gala-link'); const syncGalaLink = () => { galaLink.value = data.galaActivities[galaCourse.value][galaActivity.value].url; }; galaCourse?.addEventListener('change', syncGalaLink); galaActivity?.addEventListener('change', syncGalaLink); document.getElementById('save-gala-link')?.addEventListener('click', () => { const course = galaCourse.value; const activityIndex = Number(galaActivity.value); const url = galaLink.value.trim(); if (!/^https?:\/\//i.test(url)) return showToast('Ingresá un link válido'); data.galaActivities[course][activityIndex].url = url; const links = Object.fromEntries(courses.map((entry) => [entry, data.galaActivities[entry].map((activity) => activity.url)])); localStorage.setItem('copa-gala-links', JSON.stringify(links)); showToast('Referencia guardada para este curso'); }); document.getElementById('download-rules')?.addEventListener('click', () => showToast('El PDF oficial se puede vincular desde app.js')); }
function showToast(message) { const toast = document.getElementById('toast'); toast.textContent = message; toast.classList.add('visible'); setTimeout(() => toast.classList.remove('visible'), 2600); }
document.addEventListener('DOMContentLoaded', () => { renderAll(); setTimeout(() => { document.getElementById('splash').classList.add('hidden'); document.querySelector('.app-shell').classList.add('ready'); }, 1900); });
