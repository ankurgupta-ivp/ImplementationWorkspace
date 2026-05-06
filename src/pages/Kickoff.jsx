import React, { useState } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, EmptyState } from '../components/UI';

export default function Kickoff() {
  const { activeProject, updateQuestionnaire, showToast } = useApp();
  const [saving, setSaving] = useState(false);

  if (!activeProject) return <EmptyState message="No project selected." />;

  const sections = activeProject.questionnaire?.sections || [];

  const setAnswer = (secIdx, qIdx, value) => {
    const updated = sections.map((s, si) => si !== secIdx ? s : {
      ...s,
      questions: s.questions.map((q, qi) => qi !== qIdx ? q : { ...q, answer: value }),
    });
    // Optimistic update via parent
    updateQuestionnaire(updated);
  };

  const toggleCollapse = (secIdx) => {
    const updated = sections.map((s, si) => si !== secIdx ? s : { ...s, collapsed: !s.collapsed });
    updateQuestionnaire(updated);
  };

  const answered = sections.reduce((s, sec) => s + (sec.questions || []).filter(q => q.answer?.trim()).length, 0);
  const total = sections.reduce((s, sec) => s + (sec.questions || []).length, 0);

  const handleSave = async () => {
    setSaving(true);
    await updateQuestionnaire(sections);
    setSaving(false);
    showToast('Questionnaire saved');
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        breadcrumb={`Implementation Hub › ${activeProject.name}`}
        title="Kickoff Questionnaire"
        subtitle={`${answered} of ${total} questions answered`}
        actions={
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>
            {saving ? 'Saving…' : '💾 Save All'}
          </button>
        }
      />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        {sections.map((sec, si) => (
          <div key={si} style={{ background: '#fff', borderRadius: 8, boxShadow: '0 1px 4px rgba(0,0,0,.08)', marginBottom: 14, overflow: 'hidden' }}>
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
            {!sec.collapsed && (
              <div style={{ padding: '14px 16px' }}>
                {(sec.questions || []).map((q, qi) => (
                  <div key={q.id || qi} style={{ padding: '10px 0', borderBottom: qi < sec.questions.length - 1 ? '1px solid #f5f5f5' : 'none' }}>
                    <div style={{ fontSize: 13, color: '#404041', fontWeight: 500, marginBottom: 6 }}>{q.text}</div>
                    <textarea
                      value={q.answer || ''}
                      onChange={e => setAnswer(si, qi, e.target.value)}
                      placeholder="Enter answer…"
                      style={{
                        width: '100%', border: '1px solid #c4c4c4', borderRadius: 4,
                        padding: '7px 9px', fontFamily: 'Roboto,sans-serif', fontSize: 13,
                        background: '#fff', minHeight: 60, resize: 'vertical', outline: 'none',
                      }}
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
