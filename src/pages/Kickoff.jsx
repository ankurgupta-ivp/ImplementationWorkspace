import React, { useState, useCallback, useRef } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, EmptyState } from '../components/UI';
import * as XLSX from 'xlsx';

// ── Isolated textarea — local state, saves only on blur ───────
function QuestionInput({ value, onChange, onBlur }) {
  const [local, setLocal] = useState(value);
  const dirty = useRef(false);
  const prevValue = useRef(value);

  // Sync when parent changes value externally (import / delete) but not while typing
  if (!dirty.current && prevValue.current !== value) {
    prevValue.current = value;
    setLocal(value);
  }

  return (
    <textarea
      value={local}
      onChange={e => { dirty.current = true; setLocal(e.target.value); onChange(e.target.value); }}
      onBlur={() => { dirty.current = false; onBlur(local); }}
      onFocus={e => { e.target.style.borderColor = '#404789'; }}
      placeholder="Enter answer…"
      style={{
        width: '100%', border: '1px solid #c4c4c4', borderRadius: 4,
        padding: '7px 9px', fontFamily: 'Roboto,sans-serif', fontSize: 13,
        background: local?.trim() ? '#fafffe' : '#fff',
        minHeight: 60, resize: 'vertical', outline: 'none',
      }}
    />
  );
}

// ── Inline "add question" row inside a section ────────────────
function AddQuestionRow({ onAdd }) {
  const [text, setText] = useState('');
  const inputRef = useRef(null);

  const handleAdd = () => {
    const trimmed = text.trim();
    if (!trimmed) return;
    onAdd(trimmed);
    setText('');
    inputRef.current?.focus();
  };

  return (
    <div style={{ display: 'flex', gap: 8, marginTop: 10, alignItems: 'flex-start' }}>
      <textarea
        ref={inputRef}
        value={text}
        onChange={e => setText(e.target.value)}
        onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleAdd(); } }}
        placeholder="Type a new question and press Enter or click Add…"
        style={{
          flex: 1, border: '1px dashed #9799b1', borderRadius: 4, padding: '6px 9px',
          fontFamily: 'Roboto,sans-serif', fontSize: 12, resize: 'none', outline: 'none',
          minHeight: 38, background: '#f9f9ff',
        }}
        onFocus={e => { e.target.style.borderColor = '#404789'; }}
        onBlur={e => { e.target.style.borderColor = '#9799b1'; }}
        rows={1}
      />
      <button
        className="btn btn-outlined btn-sm"
        onClick={handleAdd}
        disabled={!text.trim()}
        style={{ whiteSpace: 'nowrap', marginTop: 2 }}
      >
        + Add
      </button>
    </div>
  );
}

export default function Kickoff() {
  const { activeProject, updateQuestionnaire, showToast } = useApp();

  const [sections, setSections] = useState(
    () => JSON.parse(JSON.stringify(activeProject?.questionnaire?.sections || []))
  );
  const [saving, setSaving] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null); // { si, qi }
  const projectId = activeProject?.id;

  // Sync when project switches
  const prevProjectId = useRef(projectId);
  if (prevProjectId.current !== projectId) {
    prevProjectId.current = projectId;
    setSections(JSON.parse(JSON.stringify(activeProject?.questionnaire?.sections || [])));
  }

  // ── Hooks (all before early return) ──────────────────────────

  const setAnswer = useCallback((secIdx, qIdx, value) => {
    setSections(prev => prev.map((s, si) => si !== secIdx ? s : {
      ...s,
      questions: s.questions.map((q, qi) => qi !== qIdx ? q : { ...q, answer: value }),
    }));
  }, []);

  const persistAnswer = useCallback((secIdx, qIdx, value) => {
    setSections(prev => {
      const updated = prev.map((s, si) => si !== secIdx ? s : {
        ...s,
        questions: s.questions.map((q, qi) => qi !== qIdx ? q : { ...q, answer: value }),
      });
      updateQuestionnaire(updated);
      return updated;
    });
  }, [updateQuestionnaire]);

  const toggleCollapse = useCallback((secIdx) => {
    setSections(prev => {
      const updated = prev.map((s, si) => si !== secIdx ? s : { ...s, collapsed: !s.collapsed });
      updateQuestionnaire(updated);
      return updated;
    });
  }, [updateQuestionnaire]);

  // ── Add a custom question to a section ───────────────────────
  const addQuestion = useCallback((secIdx, text) => {
    setSections(prev => {
      const updated = prev.map((s, si) => si !== secIdx ? s : {
        ...s,
        questions: [...s.questions, { text, answer: '', id: `cq_${Date.now()}`, custom: true }],
      });
      updateQuestionnaire(updated);
      return updated;
    });
    showToast('Question added');
  }, [updateQuestionnaire, showToast]);

  // ── Delete a question ─────────────────────────────────────────
  const deleteQuestion = useCallback((secIdx, qIdx) => {
    setSections(prev => {
      const updated = prev.map((s, si) => si !== secIdx ? s : {
        ...s,
        questions: s.questions.filter((_, qi) => qi !== qIdx),
      });
      updateQuestionnaire(updated);
      return updated;
    });
    setConfirmDelete(null);
    showToast('Question deleted');
  }, [updateQuestionnaire, showToast]);

  if (!activeProject) return <EmptyState message="No project selected." />;

  const handleSave = async () => {
    setSaving(true);
    await updateQuestionnaire(sections);
    setSaving(false);
    showToast('Questionnaire saved');
  };

  // ── Export ────────────────────────────────────────────────────
  const exportXlsx = () => {
    const rows = [['Section', 'Question', 'Answer']];
    sections.forEach(sec =>
      (sec.questions || []).forEach(q => rows.push([sec.name, q.text, q.answer || '']))
    );
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 22 }, { wch: 70 }, { wch: 60 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Questionnaire');
    XLSX.writeFile(wb, `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_Kickoff_Questionnaire.xlsx`);
    showToast('Exported to Excel');
  };

  // ── Import ─────────────────────────────────────────────────────
  // Rules:
  //   1. Match existing questions by text → update answer
  //   2. New questions (not found in any section) → add to their section (creating section if needed)
  //   3. Existing questions not in the file → keep untouched
  const importXlsx = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        // Parse rows: [section, question, answer]
        const importRows = data.slice(1)
          .filter(r => r[1] && String(r[1]).trim())
          .map(r => ({
            section: String(r[0] || '').trim(),
            text:    String(r[1] || '').trim(),
            answer:  String(r[2] || '').trim(),
          }));

        // Build a set of all existing question texts (for dedup check)
        const existingTexts = new Set();
        sections.forEach(sec => (sec.questions || []).forEach(q => existingTexts.add(q.text)));

        // Step 1: Update answers for existing questions
        let updated = sections.map(sec => ({
          ...sec,
          questions: (sec.questions || []).map(q => {
            const match = importRows.find(r => r.text === q.text);
            return match ? { ...q, answer: match.answer } : q;
          }),
        }));

        // Step 2: Add new questions that aren't already present
        const newRows = importRows.filter(r => !existingTexts.has(r.text));
        let addedCount = 0;

        newRows.forEach(r => {
          // Find existing section by name (case-insensitive)
          const secIdx = updated.findIndex(s => s.name.toLowerCase() === r.section.toLowerCase());
          if (secIdx >= 0) {
            updated = updated.map((s, si) => si !== secIdx ? s : {
              ...s,
              questions: [...s.questions, { text: r.text, answer: r.answer, id: `cq_${Date.now()}_${addedCount}`, custom: true }],
            });
          } else if (r.section) {
            // Create a new section if it doesn't exist
            updated = [...updated, {
              name: r.section, collapsed: false,
              questions: [{ text: r.text, answer: r.answer, id: `cq_${Date.now()}_${addedCount}`, custom: true }],
            }];
          }
          addedCount++;
        });

        const matchedCount = importRows.length - newRows.length;
        setSections(updated);
        updateQuestionnaire(updated);
        showToast(`Updated ${matchedCount} answers · Added ${addedCount} new question${addedCount !== 1 ? 's' : ''}`);
      } catch {
        showToast('Import failed — check file format', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const answered = sections.reduce((s, sec) => s + (sec.questions || []).filter(q => q.answer?.trim()).length, 0);
  const total    = sections.reduce((s, sec) => s + (sec.questions || []).length, 0);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        breadcrumb={`Implementation Hub › ${activeProject.name}`}
        title="Kickoff Questionnaire"
        subtitle={`${answered} of ${total} questions answered`}
        actions={
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <label className="btn btn-ghost" style={{ cursor: 'pointer', margin: 0 }}>
              ↑ Import Excel
              <input type="file" accept=".xlsx,.xls" onChange={importXlsx} style={{ display: 'none' }} />
            </label>
            <button className="btn btn-ghost" onClick={exportXlsx}>↓ Export Excel</button>
            <button className="btn btn-primary" onClick={handleSave} disabled={saving}>
              {saving ? 'Saving…' : '💾 Save All'}
            </button>
          </div>
        }
      />

      <div style={{ background: '#e3f2fd', borderBottom: '1px solid #bbdefb', padding: '7px 20px', fontSize: 12, color: '#1565c0', flexShrink: 0 }}>
        <strong>Tips:</strong> Answers auto-save on blur. Use <strong>+ Add</strong> inside any section to add client-specific questions. Click <strong>✕</strong> next to any question to delete it.
        Import merges answers for existing questions and <strong>automatically adds new questions</strong> from the file.
      </div>

      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        {sections.map((sec, si) => (
          <div key={si} style={{ background: '#fff', borderRadius: 8, boxShadow: '0 1px 4px rgba(0,0,0,.08)', marginBottom: 14, overflow: 'hidden' }}>

            {/* Section header */}
            <div
              onClick={() => toggleCollapse(si)}
              style={{ background: '#f4f4f9', padding: '10px 16px', borderBottom: sec.collapsed ? 'none' : '1px solid #f1f2fc', display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer' }}
            >
              <div style={{ fontSize: 14, fontWeight: 500, color: '#343c86', display: 'flex', alignItems: 'center', gap: 8 }}>
                {sec.name}
                <span style={{ fontSize: 11, color: '#888', fontWeight: 400 }}>
                  {(sec.questions || []).filter(q => q.answer?.trim()).length}/{(sec.questions || []).length} answered
                </span>
              </div>
              <span style={{ color: '#888', fontSize: 12 }}>{sec.collapsed ? '▼' : '▲'}</span>
            </div>

            {/* Section body */}
            {!sec.collapsed && (
              <div style={{ padding: '14px 16px' }}>
                {(sec.questions || []).map((q, qi) => (
                  <div key={q.id || qi} style={{ padding: '10px 0', borderBottom: qi < sec.questions.length - 1 ? '1px solid #f5f5f5' : 'none' }}>

                    {/* Question text row */}
                    <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 8, marginBottom: 6 }}>
                      <div style={{ fontSize: 13, color: '#404041', fontWeight: 500, flex: 1, lineHeight: 1.5 }}>
                        {q.text}
                        {q.custom && (
                          <span style={{ marginLeft: 7, fontSize: 10, background: '#fff3e0', color: '#e65100', padding: '1px 6px', borderRadius: 10, fontWeight: 500 }}>
                            Custom
                          </span>
                        )}
                        {q.answer?.trim() && (
                          <span style={{ marginLeft: 7, fontSize: 10, background: '#e8f5e9', color: '#2e7d32', padding: '1px 6px', borderRadius: 10, fontWeight: 400 }}>
                            ✓ answered
                          </span>
                        )}
                      </div>

                      {/* Delete button */}
                      {confirmDelete && confirmDelete.si === si && confirmDelete.qi === qi ? (
                        <div style={{ display: 'flex', alignItems: 'center', gap: 5, flexShrink: 0 }}>
                          <span style={{ fontSize: 11, color: '#c62828' }}>Delete?</span>
                          <button className="btn btn-danger btn-xs" onClick={() => deleteQuestion(si, qi)}>Yes</button>
                          <button className="btn btn-ghost btn-xs" onClick={() => setConfirmDelete(null)}>No</button>
                        </div>
                      ) : (
                        <button
                          onClick={e => { e.stopPropagation(); setConfirmDelete({ si, qi }); }}
                          title="Delete this question"
                          style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#ccc', fontSize: 14, padding: '0 4px', lineHeight: 1, flexShrink: 0 }}
                          onMouseEnter={e => { e.target.style.color = '#d32f2f'; }}
                          onMouseLeave={e => { e.target.style.color = '#ccc'; }}
                        >
                          ✕
                        </button>
                      )}
                    </div>

                    {/* Answer textarea */}
                    <QuestionInput
                      value={q.answer || ''}
                      onChange={(val) => setAnswer(si, qi, val)}
                      onBlur={(val) => persistAnswer(si, qi, val)}
                    />
                  </div>
                ))}

                {/* Add question row */}
                <AddQuestionRow onAdd={(text) => addQuestion(si, text)} />
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}
