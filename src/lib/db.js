import { supabase } from './supabase';
import { DEFAULT_QUESTIONNAIRE, DEFAULT_ESTIMATOR, DEFAULT_TASKS_RAW, DEFAULT_RAID_ROWS } from './defaults';

// ── Helpers ──────────────────────────────────────────────────
export function uid() { return 'p_' + Math.random().toString(36).slice(2, 10); }
export function tid() { return 't_' + Math.random().toString(36).slice(2, 10); }

function taskRowToObj(r) {
  return {
    id: r.id, sortOrder: r.sort_order, phase: r.phase, item: r.item,
    task: r.task, taskType: r.task_type, tags: r.tags,
    responsible: r.responsible, owner: r.owner, reviewer: r.reviewer,
    ownerStatus: r.owner_status, reviewerStatus: r.reviewer_status,
    comments: r.comments, actualStart: r.actual_start, actualEnd: r.actual_end,
    expectedStart: r.expected_start, expectedEnd: r.expected_end,
    actualEffort: r.actual_effort, expectedEffort: r.expected_effort,
    actualElapsed: r.actual_elapsed, expectedElapsed: r.expected_elapsed,
  };
}

function taskObjToRow(t, projectId) {
  return {
    project_id: projectId, sort_order: t.sortOrder || 0,
    phase: t.phase || '', item: t.item || '', task: t.task || '',
    task_type: t.taskType || '', tags: t.tags || '',
    responsible: t.responsible || '', owner: t.owner || '', reviewer: t.reviewer || '',
    owner_status: t.ownerStatus || 'Not Started',
    reviewer_status: t.reviewerStatus || 'Not Started',
    comments: t.comments || '', actual_start: t.actualStart || '',
    actual_end: t.actualEnd || '', expected_start: t.expectedStart || '',
    expected_end: t.expectedEnd || '', actual_effort: t.actualEffort || '',
    expected_effort: t.expectedEffort || '', actual_elapsed: t.actualElapsed || '',
    expected_elapsed: t.expectedElapsed || '',
  };
}

function raidRowToObj(r) {
  return {
    id: r.id, classification: r.classification, item: r.item,
    details: r.details, raisedOn: r.raised_on, pendingWith: r.pending_with,
    updates: r.updates, status: r.status, eta: r.eta, comments: r.comments,
  };
}

function raidObjToRow(r, projectId) {
  return {
    project_id: projectId, classification: r.classification || 'Risk',
    item: r.item || '', details: r.details || '', raised_on: r.raisedOn || '',
    pending_with: r.pendingWith || '', updates: r.updates || '',
    status: r.status || 'Not Initiated', eta: r.eta || '', comments: r.comments || '',
  };
}

// ── Templates ─────────────────────────────────────────────────
export async function loadTemplates() {
  const { data } = await supabase.from('templates').select('*');
  const map = {};
  (data || []).forEach(t => { map[t.id] = t; });

  return {
    questionnaire: map['tpl-questionnaire']
      ? { ...map['tpl-questionnaire'].data, id: 'tpl-questionnaire', name: map['tpl-questionnaire'].name, version: map['tpl-questionnaire'].version, updated: map['tpl-questionnaire'].updated }
      : JSON.parse(JSON.stringify(DEFAULT_QUESTIONNAIRE)),
    estimator: map['tpl-estimator']
      ? { ...map['tpl-estimator'].data, id: 'tpl-estimator', name: map['tpl-estimator'].name, version: map['tpl-estimator'].version, updated: map['tpl-estimator'].updated }
      : JSON.parse(JSON.stringify(DEFAULT_ESTIMATOR)),
    tasks: map['tpl-tasks']
      ? { id: 'tpl-tasks', name: map['tpl-tasks'].name, version: map['tpl-tasks'].version, updated: map['tpl-tasks'].updated, rows: map['tpl-tasks'].data.rows || [] }
      : { id: 'tpl-tasks', name: 'Tasks & Checklist', version: '1.0', updated: '2026-05-05', rows: DEFAULT_TASKS_RAW.map(r => ({ phase: r[0], item: r[1], task: r[2], taskType: r[3], tags: r[4], responsible: r[5], owner: r[6], reviewer: r[7] })) },
    raidlog: map['tpl-raidlog']
      ? { id: 'tpl-raidlog', name: map['tpl-raidlog'].name, version: map['tpl-raidlog'].version, updated: map['tpl-raidlog'].updated, rows: map['tpl-raidlog'].data.rows || [] }
      : { id: 'tpl-raidlog', name: 'RAID Log', version: '1.0', updated: '2026-05-05', rows: DEFAULT_RAID_ROWS },
  };
}

export async function saveTemplate(id, name, version, updated, data) {
  await supabase.from('templates').upsert({ id, name, version, updated, data }, { onConflict: 'id' });
}

// ── Projects ──────────────────────────────────────────────────
export async function loadProjects() {
  const { data } = await supabase.from('projects').select('*').order('created_at');
  return (data || []).map(p => ({
    id: p.id, name: p.name, createdAt: p.created_at,
    metadata: p.metadata || {},
    questionnaire: p.questionnaire || {},
    estimator: p.estimator || {},
  }));
}

export async function saveProject(proj) {
  await supabase.from('projects').upsert({
    id: proj.id, name: proj.name, created_at: proj.createdAt,
    metadata: proj.metadata || {},
    questionnaire: proj.questionnaire || {},
    estimator: proj.estimator || {},
  }, { onConflict: 'id' });
}

export async function deleteProject(projectId) {
  await supabase.from('projects').delete().eq('id', projectId);
}

// ── Tasks ─────────────────────────────────────────────────────
export async function loadTasks(projectId) {
  const { data } = await supabase.from('tasks').select('*').eq('project_id', projectId).order('sort_order');
  return (data || []).map(taskRowToObj);
}

export async function saveTasks(projectId, tasks) {
  // Delete all existing tasks for this project, then re-insert
  await supabase.from('tasks').delete().eq('project_id', projectId);
  if (!tasks.length) return;
  const rows = tasks.map((t, i) => ({ ...taskObjToRow(t, projectId), sort_order: i }));
  await supabase.from('tasks').insert(rows);
}

export async function updateTask(projectId, task, sortOrder) {
  const row = { ...taskObjToRow(task, projectId), sort_order: sortOrder };
  // task.id is a Supabase SERIAL integer for rows loaded from DB.
  // If it exists, do a targeted UPDATE — never touch other rows.
  if (task.id != null) {
    const { error } = await supabase.from('tasks').update(row).eq('id', task.id);
    if (error) throw error;
  } else {
    // Brand-new task with no DB id yet — insert it
    const { error } = await supabase.from('tasks').insert(row);
    if (error) throw error;
  }
}

// ── RAID Items ────────────────────────────────────────────────
export async function loadRaidItems(projectId) {
  const { data } = await supabase.from('raid_items').select('*').eq('project_id', projectId).order('id');
  return (data || []).map(raidRowToObj);
}

export async function saveRaidItem(projectId, item) {
  const row = raidObjToRow(item, projectId);
  if (item.id) {
    await supabase.from('raid_items').update(row).eq('id', item.id);
  } else {
    await supabase.from('raid_items').insert(row);
  }
}

export async function deleteRaidItem(id) {
  await supabase.from('raid_items').delete().eq('id', id);
}

// ── App State ─────────────────────────────────────────────────
export async function getAppState(key) {
  const { data } = await supabase.from('app_state').select('value').eq('key', key).single();
  return data?.value || null;
}

export async function setAppState(key, value) {
  await supabase.from('app_state').upsert({ key, value }, { onConflict: 'key' });
}

// ── Create Project from Template ──────────────────────────────
export async function createProjectFromTemplate(name, templates) {
  const newProj = {
    id: uid(),
    name,
    createdAt: new Date().toISOString().slice(0, 10),
    metadata: { kickoffDate: '', targetGoLive: '', leadBA: '', engineers: '', currentPhase: 'Requirements' },
    questionnaire: {
      sections: templates.questionnaire.sections.map(s => ({
        name: s.name, collapsed: false,
        questions: s.questions.map(q => ({ text: q, answer: '', id: tid() })),
      })),
    },
    estimator: {
      inputs: { ...(templates.estimator.inputs || {}) },
      risks: { ...(templates.estimator.risks || {}) },
      stepsBase: (templates.estimator.steps || []).map(s => ({ ...s })),
    },
  };

  await saveProject(newProj);

  const tasks = templates.tasks.rows.map((r, i) => ({
    sortOrder: i, phase: r.phase || '', item: r.item || '', task: r.task || '',
    taskType: r.taskType || '', tags: r.tags || '',
    responsible: r.responsible || '', owner: r.owner || '', reviewer: r.reviewer || '',
    ownerStatus: 'Not Started', reviewerStatus: 'Not Started', comments: '',
    actualStart: '', actualEnd: '', expectedStart: '', expectedEnd: '',
    actualEffort: '', expectedEffort: '', actualElapsed: '', expectedElapsed: '',
  }));
  await saveTasks(newProj.id, tasks);

  for (const r of templates.raidlog.rows) {
    await saveRaidItem(newProj.id, r);
  }

  return newProj;
}
