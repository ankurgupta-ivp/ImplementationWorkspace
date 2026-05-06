import React, { createContext, useContext, useEffect, useState, useCallback } from 'react';
import {
  loadTemplates, loadProjects, loadTasks, loadRaidItems,
  saveProject, saveTasks, saveRaidItem, deleteRaidItem,
  saveTemplate, getAppState, setAppState,
  createProjectFromTemplate, deleteProject, uid,
} from '../lib/db';

const AppContext = createContext(null);

export function AppProvider({ children }) {
  const [loading, setLoading] = useState(true);
  const [templates, setTemplates] = useState(null);
  const [projects, setProjects] = useState([]);
  const [activeProjectId, setActiveProjectIdState] = useState(null);
  const [tasks, setTasks] = useState([]);
  const [raidItems, setRaidItems] = useState([]);
  const [currentPage, setCurrentPage] = useState('overview');
  const [toast, setToast] = useState(null);

  const activeProject = projects.find(p => p.id === activeProjectId) || null;

  // ── Toast ────────────────────────────────────────────────────
  const showToast = useCallback((msg, type = 'success') => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 2500);
  }, []);

  // ── Initial Load ─────────────────────────────────────────────
  useEffect(() => {
    async function init() {
      setLoading(true);
      try {
        const [tpls, projs, savedActiveId] = await Promise.all([
          loadTemplates(),
          loadProjects(),
          getAppState('activeProjectId'),
        ]);
        setTemplates(tpls);

        if (projs.length === 0) {
          // First run — create demo project
          const demo = await createProjectFromTemplate('Acme Capital — Multi-Asset', tpls);
          setProjects([demo]);
          setActiveProjectIdState(demo.id);
          await setAppState('activeProjectId', demo.id);
          const [t, r] = await Promise.all([loadTasks(demo.id), loadRaidItems(demo.id)]);
          setTasks(t); setRaidItems(r);
        } else {
          setProjects(projs);
          const activeId = (savedActiveId && projs.find(p => p.id === savedActiveId))
            ? savedActiveId : projs[0].id;
          setActiveProjectIdState(activeId);
          const [t, r] = await Promise.all([loadTasks(activeId), loadRaidItems(activeId)]);
          setTasks(t); setRaidItems(r);
        }
      } catch (err) {
        console.error('Init error:', err);
        showToast('Failed to load data. Check Supabase config.', 'error');
      }
      setLoading(false);
    }
    init();
  }, []); // eslint-disable-line

  // ── Switch Active Project ────────────────────────────────────
  const switchProject = useCallback(async (id) => {
    setLoading(true);
    setActiveProjectIdState(id);
    await setAppState('activeProjectId', id);
    const [t, r] = await Promise.all([loadTasks(id), loadRaidItems(id)]);
    setTasks(t); setRaidItems(r);
    setCurrentPage('overview');
    setLoading(false);
  }, []);

  // ── Create New Project ────────────────────────────────────────
  const addProject = useCallback(async (name) => {
    const proj = await createProjectFromTemplate(name, templates);
    setProjects(prev => [...prev, proj]);
    await switchProject(proj.id);
    showToast('Project created');
    return proj;
  }, [templates, switchProject, showToast]);

  // ── Remove Project ────────────────────────────────────────────
  const removeProject = useCallback(async (id) => {
    await deleteProject(id);
    const remaining = projects.filter(p => p.id !== id);
    setProjects(remaining);
    if (activeProjectId === id) {
      if (remaining.length > 0) await switchProject(remaining[0].id);
      else { setActiveProjectIdState(null); setTasks([]); setRaidItems([]); }
    }
    showToast('Project deleted');
  }, [projects, activeProjectId, switchProject, showToast]);

  // ── Update Project Metadata ───────────────────────────────────
  const updateProject = useCallback(async (updated) => {
    await saveProject(updated);
    setProjects(prev => prev.map(p => p.id === updated.id ? updated : p));
    showToast('Project saved');
  }, [showToast]);

  // ── Update Tasks ──────────────────────────────────────────────
  const updateTasks = useCallback(async (newTasks) => {
    setTasks(newTasks);
    await saveTasks(activeProjectId, newTasks);
  }, [activeProjectId]);

  const updateSingleTask = useCallback(async (taskId, changes) => {
    const updated = tasks.map((t, i) => t.id === taskId ? { ...t, ...changes } : t);
    setTasks(updated);
    const idx = updated.findIndex(t => t.id === taskId);
    if (idx >= 0) await saveTasks(activeProjectId, updated);
  }, [tasks, activeProjectId]);

  // ── RAID Items ────────────────────────────────────────────────
  const addRaidItem = useCallback(async (item) => {
    await saveRaidItem(activeProjectId, item);
    const r = await loadRaidItems(activeProjectId);
    setRaidItems(r);
    showToast('RAID item added');
  }, [activeProjectId, showToast]);

  const updateRaidItem = useCallback(async (item) => {
    await saveRaidItem(activeProjectId, item);
    const r = await loadRaidItems(activeProjectId);
    setRaidItems(r);
    showToast('RAID item updated');
  }, [activeProjectId, showToast]);

  const removeRaidItem = useCallback(async (id) => {
    await deleteRaidItem(id);
    setRaidItems(prev => prev.filter(r => r.id !== id));
    showToast('RAID item deleted');
  }, [showToast]);

  // ── Templates ─────────────────────────────────────────────────
  const updateTemplate = useCallback(async (id, name, version, updated, data) => {
    await saveTemplate(id, name, version, updated, data);
    setTemplates(prev => ({ ...prev, [id.replace('tpl-', '')]: { id, name, version, updated, ...data } }));
    showToast('Template updated');
  }, [showToast]);

  // ── Questionnaire save ────────────────────────────────────────
  const updateQuestionnaire = useCallback(async (sections) => {
    const updated = { ...activeProject, questionnaire: { sections } };
    await saveProject(updated);
    setProjects(prev => prev.map(p => p.id === updated.id ? updated : p));
  }, [activeProject]);

  // ── Estimator save ────────────────────────────────────────────
  const updateEstimator = useCallback(async (estimatorData) => {
    const updated = { ...activeProject, estimator: estimatorData };
    await saveProject(updated);
    setProjects(prev => prev.map(p => p.id === updated.id ? updated : p));
  }, [activeProject]);

  return (
    <AppContext.Provider value={{
      loading, templates, projects, activeProject, activeProjectId,
      tasks, raidItems, currentPage, toast,
      setCurrentPage, showToast,
      switchProject, addProject, removeProject, updateProject,
      updateTasks, updateSingleTask,
      addRaidItem, updateRaidItem, removeRaidItem,
      updateTemplate, updateQuestionnaire, updateEstimator,
    }}>
      {children}
    </AppContext.Provider>
  );
}

export function useApp() {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error('useApp must be used inside AppProvider');
  return ctx;
}
