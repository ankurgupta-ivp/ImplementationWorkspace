import React, { useState, useMemo } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, Modal, Field, Input, Select, Textarea, StatusBadge, EmptyState, useConfirm } from '../components/UI';
import { RAID_CLASSIFICATIONS, RAID_STATUSES, RAID_PENDING } from '../lib/defaults';
import * as XLSX from 'xlsx';

const CLASS_STYLES = {
  Risk:       { bg: '#ffebee', color: '#c62828' },
  Action:     { bg: '#e8f5e9', color: '#2e7d32' },
  Issue:      { bg: '#fff3e0', color: '#e65100' },
  Dependency: { bg: '#e3f2fd', color: '#1565c0' },
  Decision:   { bg: '#f3e5f5', color: '#6a1b9a' },
  Assumption: { bg: '#fffde7', color: '#f57f17' },
};

const EMPTY_FORM = { classification: 'Risk', item: '', details: '', raisedOn: '', pendingWith: '', updates: '', status: 'Not Initiated', eta: '', comments: '' };

export default function RaidLog() {
  const { activeProject, raidItems, addRaidItem, updateRaidItem, removeRaidItem, showToast } = useApp();
  const { ask, dialog } = useConfirm();
  const [filterClass, setFilterClass] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [editItem, setEditItem] = useState(null);
  const [form, setForm] = useState(EMPTY_FORM);
  const [isNew, setIsNew] = useState(false);

  const filtered = useMemo(() => raidItems.filter(r =>
    (!filterClass || r.classification === filterClass) &&
    (!filterStatus || r.status === filterStatus)
  ), [raidItems, filterClass, filterStatus]);

  if (!activeProject) return <EmptyState message="No project selected." />;

  const openNew = () => { setForm({ ...EMPTY_FORM, raisedOn: new Date().toISOString().slice(0, 10) }); setIsNew(true); setEditItem({}); };
  const openEdit = (item) => { setForm({ ...item }); setIsNew(false); setEditItem(item); };

  const handleSave = async () => {
    if (!form.item.trim()) { showToast('Item name is required', 'error'); return; }
    if (isNew) await addRaidItem(form);
    else await updateRaidItem({ ...form, id: editItem.id });
    setEditItem(null);
  };

  const handleDelete = (item) => {
    ask(`Delete "${item.item}" from RAID Log?`, () => removeRaidItem(item.id));
  };

  const exportXlsx = () => {
    const rows = [['#', 'Classification', 'Item', 'Details', 'Raised On', 'Pending With', 'Updates', 'Status', 'ETA', 'Comments']];
    raidItems.forEach((r, i) => rows.push([i + 1, r.classification, r.item, r.details, r.raisedOn, r.pendingWith, r.updates, r.status, r.eta, r.comments]));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'RAID Log');
    XLSX.writeFile(wb, `${activeProject.name.replace(/[^a-z0-9]/gi, '_')}_RAID_Log.xlsx`);
    showToast('Exported to Excel');
  };

  const counts = RAID_CLASSIFICATIONS.reduce((m, c) => ({ ...m, [c]: raidItems.filter(r => r.classification === c).length }), {});

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      {dialog}
      <PageHeader
        breadcrumb={`Implementation Hub › ${activeProject.name}`}
        title="RAID Log"
        subtitle={`${raidItems.length} items tracked`}
        actions={<>
          <button className="btn btn-ghost" onClick={exportXlsx}>↓ Export Excel</button>
          <button className="btn btn-primary" onClick={openNew}>+ Add Item</button>
        </>}
      />

      {/* Toolbar */}
      <div style={{ background: '#fff', borderBottom: '1px solid #e0e0e0', padding: '8px 20px', display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', flexShrink: 0 }}>
        <select value={filterClass} onChange={e => setFilterClass(e.target.value)}
          style={{ border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 9px', fontSize: 12, fontFamily: 'Roboto,sans-serif' }}>
          <option value="">All Classifications</option>
          {RAID_CLASSIFICATIONS.map(c => <option key={c}>{c} ({counts[c] || 0})</option>)}
        </select>
        <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)}
          style={{ border: '1px solid #c4c4c4', borderRadius: 4, padding: '5px 9px', fontSize: 12, fontFamily: 'Roboto,sans-serif' }}>
          <option value="">All Statuses</option>
          {RAID_STATUSES.map(s => <option key={s}>{s}</option>)}
        </select>
        {/* KPI Chips */}
        <div style={{ display: 'flex', gap: 6, marginLeft: 'auto' }}>
          {RAID_CLASSIFICATIONS.map(c => {
            const s = CLASS_STYLES[c] || {};
            return <span key={c} style={{ background: '#fff', border: '1px solid #e0e0e0', borderRadius: 4, padding: '3px 9px', fontSize: 11, color: '#444', display: 'flex', alignItems: 'center', gap: 4 }}>
              <span style={{ width: 7, height: 7, borderRadius: '50%', background: s.color, display: 'inline-block' }} />
              {c}: {counts[c] || 0}
            </span>;
          })}
        </div>
      </div>

      {/* Table */}
      <div style={{ flex: 1, overflow: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead>
            <tr style={{ background: '#f5f5f5', borderBottom: '1px solid #d0d0d0' }}>
              {['#', 'Classification', 'Item', 'Details', 'Raised On', 'Pending With', 'Status', 'ETA', ''].map(h => (
                <th key={h} style={{ padding: '9px 8px', textAlign: 'left', fontSize: 12, fontWeight: 500, color: '#444', whiteSpace: 'nowrap' }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((r, i) => {
              const cs = CLASS_STYLES[r.classification] || {};
              return (
                <tr key={r.id} style={{ borderBottom: '1px solid #f0f0f0', background: i % 2 === 0 ? '#fafafe' : '#fff' }}
                  onMouseEnter={e => e.currentTarget.style.background = '#f0f4ff'}
                  onMouseLeave={e => e.currentTarget.style.background = i % 2 === 0 ? '#fafafe' : '#fff'}
                >
                  <td style={{ padding: '6px 8px', fontSize: 11, color: '#888' }}>{i + 1}</td>
                  <td style={{ padding: '6px 8px' }}>
                    <span style={{ display: 'inline-block', fontSize: 10, fontWeight: 500, padding: '2px 7px', borderRadius: 3, textTransform: 'uppercase', background: cs.bg, color: cs.color }}>{r.classification}</span>
                  </td>
                  <td style={{ padding: '6px 8px', fontWeight: 500, color: '#404041' }}>{r.item}</td>
                  <td style={{ padding: '6px 8px', color: '#555', maxWidth: 300 }}>{r.details}</td>
                  <td style={{ padding: '6px 8px', color: '#666', whiteSpace: 'nowrap' }}>{r.raisedOn}</td>
                  <td style={{ padding: '6px 8px', color: '#666' }}>{r.pendingWith}</td>
                  <td style={{ padding: '6px 8px' }}><StatusBadge status={r.status} /></td>
                  <td style={{ padding: '6px 8px', color: '#666', whiteSpace: 'nowrap' }}>{r.eta || '—'}</td>
                  <td style={{ padding: '6px 8px' }}>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button className="btn btn-ghost btn-xs" onClick={() => openEdit(r)}>✎</button>
                      <button className="btn btn-danger btn-xs" onClick={() => handleDelete(r)}>✕</button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {filtered.length === 0 && <div style={{ textAlign: 'center', padding: 40, color: '#999' }}>No RAID items yet. Click + Add Item to start.</div>}
      </div>

      <div style={{ height: 26, background: '#f5f5f5', borderTop: '1px solid #e0e0e0', display: 'flex', alignItems: 'center', padding: '0 14px', fontSize: 11, color: '#888', flexShrink: 0 }}>
        {filtered.length} of {raidItems.length} items shown
      </div>

      {/* Add/Edit Modal */}
      {editItem !== null && (
        <Modal title={isNew ? 'Add RAID Item' : 'Edit RAID Item'} onClose={() => setEditItem(null)} onSave={handleSave} wide>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 20px' }}>
            <Field label="Classification">
              <Select value={form.classification} onChange={e => setForm(p => ({ ...p, classification: e.target.value }))}>
                {RAID_CLASSIFICATIONS.map(c => <option key={c}>{c}</option>)}
              </Select>
            </Field>
            <Field label="Status">
              <Select value={form.status} onChange={e => setForm(p => ({ ...p, status: e.target.value }))}>
                {RAID_STATUSES.map(s => <option key={s}>{s}</option>)}
              </Select>
            </Field>
            <Field label="Item" required>
              <Input value={form.item} onChange={e => setForm(p => ({ ...p, item: e.target.value }))} placeholder="Short title…" />
            </Field>
            <Field label="Pending With">
              <Select value={form.pendingWith} onChange={e => setForm(p => ({ ...p, pendingWith: e.target.value }))}>
                <option value="">—</option>
                {RAID_PENDING.map(p => <option key={p}>{p}</option>)}
              </Select>
            </Field>
            <Field label="Raised On"><Input type="date" value={form.raisedOn} onChange={e => setForm(p => ({ ...p, raisedOn: e.target.value }))} /></Field>
            <Field label="ETA"><Input type="date" value={form.eta} onChange={e => setForm(p => ({ ...p, eta: e.target.value }))} /></Field>
            <Field label="Details"><Textarea value={form.details} onChange={e => setForm(p => ({ ...p, details: e.target.value }))} /></Field>
            <Field label="Updates"><Textarea value={form.updates} onChange={e => setForm(p => ({ ...p, updates: e.target.value }))} /></Field>
            <Field label="Comments"><Textarea value={form.comments} onChange={e => setForm(p => ({ ...p, comments: e.target.value }))} /></Field>
          </div>
        </Modal>
      )}
    </div>
  );
}
