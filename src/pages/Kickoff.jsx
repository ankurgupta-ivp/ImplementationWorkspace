import React, { useState, useCallback, useRef } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, EmptyState } from '../components/UI';
import * as XLSX from 'xlsx';

// ── Single question textarea — fully controlled locally ───────
// Never re-renders from parent; calls onBlur to persist.
function QuestionInput({ value, onChange, onBlur }) {
  const [local, setLocal] = useState(value);
  const ref = useRef(false);

  // Sync if parent value changes externally (e.g. import), but not while typing
  const prevValue = useRef(value);
  if (!ref.current && prevValue.current !== value) {
    prevValue.current = value;
    setLocal(value);
  }

  const handleChange = (e) => {
    ref.current = true; // mark as dirty
    setLocal(e.target.value);
    onChange(e.target.value);
  };

  const handleBlur = () => {
    ref.current = false;
    onBlur(local);
  };

  return (
    <textarea
      value={local}
      onChange={handleChange}
      onBlur={handleBlur}
      placeholder="Enter answer…"
      style={{
        width: '100%', border: '1px solid #c4c4c4', borderRadius: 4,
        padding: '7px 9px', fontFamily: 'Roboto,sans-serif', fontSize: 13,
        background: local?.trim() ? '#fafffe' : '#fff',
        minHeight: 60, resize: 'vertical', outline: 'none',
        transition: 'border-color .15s',
      }}
      onFocus={e => { e.target.style.borderColor = '#404789'; }}
    />
  );
}

export default function Kickoff() {
  const { activeProject, updateQuestionnaire, showToast } = useApp();

  // ── Local sections state to avoid parent re-renders while typing
  const [sections, setSections] = useState(
    () => JSON.parse(JSON.stringify(activeProject?.questionnaire?.sections || []))
  );
  const [saving, setSaving] = useState(false);
  const projectId = activeProject?.id;

  // Sync if project switches
  const prevProjectId = useRef(projectId);
  if (prevProjectId.current !== projectId) {
    prevProjectId.current = projectId;
    setSections(JSON.parse(JSON.stringify(activeProject?.questionnaire?.sections || [])));
  }

  // Update local state only (no Supabase call) — fast, no re-render cascade
  const setAnswer = useCallback((secIdx, qIdx, value) => {
    setSections(prev => prev.map((s, si) => si !== secIdx ? s : {
      ...s,
      questions: s.questions.map((q, qi) => qi !== qIdx ? q : { ...q, answer: value }),
    }));
  }, []);

  // Called on blur — saves to Supabase
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

  // ── Import ────────────────────────────────────────────────────
  const importXlsx = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const answerMap = {};
        data.slice(1).forEach(row => {
          const qText = String(row[1] || '').trim();
          const answer = String(row[2] || '').trim();
          if (qText) answerMap[qText] = answer;
        });
        let matched = 0;
        const updated = sections.map(sec => ({
          ...sec,
          questions: (sec.questions || []).map(q => {
            if (answerMap.hasOwnProperty(q.text)) { matched++; return { ...q, answer: answerMap[q.text] }; }
            return q;
          }),
        }));
        setSections(updated);
        updateQuestionnaire(updated);
        showToast(`Imported ${matched} answers`);
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
        <strong>Tip:</strong> Answers auto-save when you click out of a field. Use <strong>Save All</strong> to force-save everything at once. Export → fill Answer column → Import to bulk-fill answers.
      </div>
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        {sections.map((sec, si) => (
          <div key={si} style={{ background: '#fff', borderRadius: 8, boxShadow: '0 1px 4px rgba(0,0,0,.08)', marginBottom: 14, overflow: 'hidden' }}>
            <div onClick={() => toggleCollapse(si)}
              style={{ background: '#f4f4f9', padding: '10px 16px', borderBottom: sec.collapsed ? 'none' : '1px solid #f1f2fc', display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer' }}>
              <div style={{ fontSize: 14, fontWeight: 500, color: '#343c86', display: 'flex', alignItems: 'center', gap: 8 }}>
                {sec.name}
                <span style={{ fontSize: 11, color: '#888', fontWeight: 400 }}>
                  {(sec.questions || []).filter(q => q.answer?.trim()).length}/{(sec.questions || []).length} answered
                </span>
              </div>
              <span style={{ color: '#888', fontSize: 12 }}>{sec.collapsed ? '▼' : '▲'}</span>
            </div>
            {!sec.collapsed && (
              <div style={{ padding: '14px 16px' }}>
                {(sec.questions || []).map((q, qi) => (
                  <div key={qi} style={{ padding: '10px 0', borderBottom: qi < sec.questions.length - 1 ? '1px solid #f5f5f5' : 'none' }}>
                    <div style={{ fontSize: 13, color: '#404041', fontWeight: 500, marginBottom: 6 }}>
                      {q.text}
                      {q.answer?.trim() && (
                        <span style={{ marginLeft: 8, fontSize: 10, background: '#e8f5e9', color: '#2e7d32', padding: '1px 6px', borderRadius: 10, fontWeight: 400 }}>✓ answered</span>
                      )}
                    </div>
                    <QuestionInput
                      value={q.answer || ''}
                      onChange={(val) => setAnswer(si, qi, val)}
                      onBlur={(val) => persistAnswer(si, qi, val)}
                    />
                  </div>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}
