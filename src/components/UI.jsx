import React, { useState } from 'react';

// ── Toast ────────────────────────────────────────────────────
export function Toast({ toast }) {
  if (!toast) return null;
  const bg = toast.type === 'error' ? '#c62828' : '#2e7d32';
  return (
    <div style={{
      position: 'fixed', bottom: 20, right: 20, background: bg,
      color: '#fff', padding: '10px 16px', borderRadius: 5,
      fontSize: 13, zIndex: 500, boxShadow: '0 2px 8px rgba(0,0,0,.3)',
    }}>
      {toast.msg}
    </div>
  );
}

// ── Modal ────────────────────────────────────────────────────
export function Modal({ title, onClose, onSave, children, wide }) {
  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200,
    }} onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{
        background: '#fff', borderRadius: 8,
        width: wide ? 820 : 560, maxWidth: '95vw', maxHeight: '88vh',
        display: 'flex', flexDirection: 'column', overflow: 'hidden',
        boxShadow: '0 8px 32px rgba(0,0,0,.18)',
      }}>
        <div style={{ padding: '14px 18px', borderBottom: '1px solid #f0f0f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <span style={{ fontSize: 15, fontWeight: 500, color: '#343c86' }}>{title}</span>
          <span style={{ cursor: 'pointer', color: '#999', fontSize: 20 }} onClick={onClose}>✕</span>
        </div>
        <div style={{ padding: 18, overflowY: 'auto', flex: 1 }}>{children}</div>
        {onSave && (
          <div style={{ padding: '12px 18px', borderTop: '1px solid #f0f0f0', display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
            <button className="btn btn-ghost" onClick={onClose}>Cancel</button>
            <button className="btn btn-primary" onClick={onSave}>Save</button>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Confirm Dialog ───────────────────────────────────────────
export function ConfirmDialog({ message, onConfirm, onCancel }) {
  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,.4)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 300,
    }}>
      <div style={{ background: '#fff', borderRadius: 8, padding: 22, width: 380, textAlign: 'center' }}>
        <div style={{ fontSize: 28 }}>⚠</div>
        <div style={{ fontSize: 14, color: '#444', margin: '8px 0 16px' }}>{message}</div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'center' }}>
          <button className="btn btn-ghost" onClick={onCancel}>Cancel</button>
          <button className="btn btn-danger" onClick={onConfirm}>Confirm</button>
        </div>
      </div>
    </div>
  );
}

// ── FormField ────────────────────────────────────────────────
export function Field({ label, required, children }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
      <label style={{ fontSize: 12, color: '#666' }}>
        {label}{required && <span style={{ color: '#d32f2f' }}> *</span>}
      </label>
      {children}
    </div>
  );
}

export function Input({ style, ...props }) {
  return (
    <input style={{
      fontFamily: 'Roboto,sans-serif', fontSize: 13, color: '#212121',
      border: '1px solid #c4c4c4', borderRadius: 4, padding: '7px 9px',
      width: '100%', outline: 'none', background: '#fff', ...style,
    }} {...props} />
  );
}

export function Select({ style, children, ...props }) {
  return (
    <select style={{
      fontFamily: 'Roboto,sans-serif', fontSize: 13, color: '#212121',
      border: '1px solid #c4c4c4', borderRadius: 4, padding: '7px 9px',
      width: '100%', outline: 'none', background: '#fff', ...style,
    }} {...props}>{children}</select>
  );
}

export function Textarea({ style, ...props }) {
  return (
    <textarea style={{
      fontFamily: 'Roboto,sans-serif', fontSize: 13, color: '#212121',
      border: '1px solid #c4c4c4', borderRadius: 4, padding: '7px 9px',
      width: '100%', outline: 'none', background: '#fff', minHeight: 60, resize: 'vertical', ...style,
    }} {...props} />
  );
}

// ── Status Badge ─────────────────────────────────────────────
const BADGE_STYLES = {
  'Done':         { color: '#2e7d32', dot: '#4caf50' },
  'Completed':    { color: '#2e7d32', dot: '#4caf50' },
  'In Progress':  { color: '#1565c0', dot: '#2278cf' },
  'Blocked':      { color: '#b71c1c', dot: '#d32f2f' },
  'Delayed':      { color: '#b71c1c', dot: '#d32f2f' },
  'Not Started':  { color: '#888',    dot: '#bbb' },
  'Not Initiated':{ color: '#888',    dot: '#bbb' },
  'Escalated':    { color: '#6a1b9a', dot: '#9c27b0' },
  'Closed':       { color: '#555',    dot: '#999' },
  'Under Review': { color: '#e67700', dot: '#e8a025' },
};

export function StatusBadge({ status }) {
  const s = BADGE_STYLES[status] || { color: '#888', dot: '#bbb' };
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: 11, fontWeight: 500, color: s.color }}>
      <span style={{ width: 7, height: 7, borderRadius: '50%', background: s.dot, display: 'inline-block' }} />
      {status}
    </span>
  );
}

// ── Progress Bar ─────────────────────────────────────────────
export function ProgressBar({ pct, color }) {
  const c = color || (pct >= 80 ? '#4caf50' : pct >= 50 ? '#404789' : '#e8a025');
  return (
    <div style={{ background: '#eef0ff', borderRadius: 10, height: 8, width: '100%', overflow: 'hidden' }}>
      <div style={{ height: '100%', width: `${Math.min(pct, 100)}%`, background: c, transition: 'width .3s' }} />
    </div>
  );
}

// ── KPI Card ─────────────────────────────────────────────────
export function KpiCard({ label, value, delta, variant }) {
  const colors = { ok: '#4caf50', warn: '#e8a025', fail: '#d32f2f', default: '#404789' };
  return (
    <div style={{
      background: '#fff', borderRadius: 8, padding: 14,
      boxShadow: '0 1px 4px rgba(0,0,0,.08)',
      borderLeft: `4px solid ${colors[variant] || colors.default}`,
    }}>
      <div style={{ fontSize: 11, color: '#666', textTransform: 'uppercase', letterSpacing: '.4px', marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 24, fontWeight: 700, color: '#404041', lineHeight: 1 }}>{value}</div>
      {delta && <div style={{ fontSize: 11, marginTop: 5, color: '#666' }}>{delta}</div>}
    </div>
  );
}

// ── Widget Card ───────────────────────────────────────────────
export function WidgetCard({ title, actions, children, noPad }) {
  return (
    <div style={{ background: '#fff', borderRadius: 8, boxShadow: '0 1px 4px rgba(0,0,0,.08)', overflow: 'hidden', marginBottom: 16 }}>
      {title && (
        <div style={{
          background: '#f4f4f9', color: '#343c86', fontSize: 14, fontWeight: 500,
          padding: '10px 16px', borderBottom: '1px solid #f1f2fc',
          display: 'flex', justifyContent: 'space-between', alignItems: 'center',
        }}>
          <span>{title}</span>
          {actions && <div style={{ display: 'flex', gap: 6 }}>{actions}</div>}
        </div>
      )}
      <div style={noPad ? {} : { padding: 16 }}>{children}</div>
    </div>
  );
}

// ── Page Header ───────────────────────────────────────────────
export function PageHeader({ breadcrumb, title, subtitle, actions }) {
  return (
    <div style={{
      background: '#fff', borderBottom: '1px solid #e0e0e0',
      padding: '14px 20px', display: 'flex', alignItems: 'flex-start',
      justifyContent: 'space-between', flexShrink: 0,
    }}>
      <div>
        {breadcrumb && <div style={{ fontSize: 12, color: '#888', marginBottom: 3 }}>{breadcrumb}</div>}
        <div style={{ fontSize: 17, fontWeight: 500, color: '#404041' }}>{title}</div>
        {subtitle && <div style={{ fontSize: 12, color: '#666', marginTop: 3 }}>{subtitle}</div>}
      </div>
      {actions && <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>{actions}</div>}
    </div>
  );
}

// ── Empty State ───────────────────────────────────────────────
export function EmptyState({ icon, message, onAction, actionLabel }) {
  return (
    <div style={{ textAlign: 'center', padding: '60px 20px', color: '#999' }}>
      <div style={{ fontSize: 48, marginBottom: 14, opacity: 0.5 }}>{icon || '📋'}</div>
      <div style={{ fontSize: 14, marginBottom: 18 }}>{message}</div>
      {onAction && <button className="btn btn-primary" onClick={onAction}>{actionLabel || 'Get Started'}</button>}
    </div>
  );
}

// ── useConfirm hook ───────────────────────────────────────────
export function useConfirm() {
  const [confirm, setConfirm] = useState(null);
  const ask = (message, onConfirm) => setConfirm({ message, onConfirm });
  const dialog = confirm ? (
    <ConfirmDialog
      message={confirm.message}
      onConfirm={() => { confirm.onConfirm(); setConfirm(null); }}
      onCancel={() => setConfirm(null)}
    />
  ) : null;
  return { ask, dialog };
}
