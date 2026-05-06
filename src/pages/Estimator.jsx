import React, { useState, useMemo } from 'react';
import { useApp } from '../hooks/useApp';
import { PageHeader, WidgetCard, EmptyState } from '../components/UI';

const MULTIPLIERS = {
  env: { '1 env': 1, '2 envs': 1.5, '3+ envs': 2 },
  security: { 'Simple (1 type)': 0.6, 'Standard (2-3 types)': 1, 'Complex (4+ types)': 1.5 },
  positionSrc: { '1 source': 0.7, '2-3 sources': 1, '4+ sources': 1.4 },
  pricingSrc: { '1-2 sources': 0.7, '3-4 sources': 1, '5+ sources': 1.5 },
  customCode: { 'None': 0, 'Minor (<200 LOC)': 0.5, 'Moderate (200-1000 LOC)': 1, 'Heavy (1000+ LOC)': 2 },
  pricingRules: (n) => Math.max(0.5, n / 42),
  exceptions: (n) => Math.max(0.5, n / 14),
  dashboards: { 'Standard': 0.8, 'Custom light': 1, 'Custom heavy': 1.5 },
  challenge: { 'None': 0, 'Standard config': 1, 'Custom': 1.8 },
  access: { 'Simple': 0.7, 'Standard': 1, 'SSO + complex': 1.5 },
  downstream: (n) => Math.max(0.5, n / 6),
  reporting: { 'Standard only': 0.7, 'Standard + custom': 1, 'Heavily custom': 1.6 },
  uat: { 'Low complexity': 0.7, 'Medium complexity': 1, 'High complexity': 1.4 },
  golive: { 'Simple': 0.7, 'Standard': 1, 'Complex (parallel run)': 1.5 },
};

const RISK_FACTORS = {
  clientNew: { label: 'New client (no prior PM exp)', pct: 10 },
  clientITBottlenecks: { label: 'Client IT bottlenecks expected', pct: 15 },
  multiSignoff: { label: 'Multi-level sign-off required', pct: 8 },
  livingProd: { label: 'Migration from live production', pct: 12 },
  regDeadline: { label: 'Hard regulatory deadline', pct: 20 },
  greenfield: { label: 'Greenfield (no existing PM base)', pct: 10 },
  archetypeReusable: { label: 'Archetype reusable (discount)', pct: -8 },
};

function getMul(key, value) {
  if (typeof MULTIPLIERS[key] === 'function') return MULTIPLIERS[key](parseFloat(value) || 0);
  if (typeof MULTIPLIERS[key] === 'object') return MULTIPLIERS[key][value] ?? 1;
  return 1;
}

const STEP_DEFAULTS = {
  env: '2 envs', security: 'Standard (2-3 types)', positionSrc: '2-3 sources',
  pricingSrc: '3-4 sources', customCode: 'Moderate (200-1000 LOC)',
  pricingRules: 42, exceptions: 14, dashboards: 'Standard', challenge: 'Standard config',
  access: 'Standard', downstream: 6, reporting: 'Standard + custom',
  uat: 'Medium complexity', golive: 'Standard',
};

export default function Estimator() {
  const { activeProject, updateEstimator, showToast } = useApp();
  const [inputs, setInputs] = useState(() => ({ ...STEP_DEFAULTS, ...(activeProject?.estimator?.inputs || {}) }));
  const [risks, setRisks] = useState(() => ({ ...(activeProject?.estimator?.risks || {}) }));

  if (!activeProject) return <EmptyState message="No project selected." />;

  const steps = activeProject.estimator?.stepsBase || [];

  const stepHours = useMemo(() => steps.map(s => {
    const mul = getMul(s.mul, inputs[s.mul]);
    return Math.round(s.base * mul);
  }), [steps, inputs]);

  const baseTotal = stepHours.reduce((a, b) => a + b, 0);
  const riskPct = Object.entries(risks).filter(([, v]) => v).reduce((s, [k]) => s + (RISK_FACTORS[k]?.pct || 0), 0);
  const finalTotal = Math.round(baseTotal * (1 + riskPct / 100));

  const baHours = steps.reduce((s, st, i) => s + (st.role.includes('BA') ? stepHours[i] : 0), 0);
  const engHours = steps.reduce((s, st, i) => s + (st.role.includes('Engineer') ? stepHours[i] : 0), 0);

  const handleSave = async () => {
    await updateEstimator({ inputs, risks, stepsBase: steps });
    showToast('Estimator saved');
  };

  const SelectInput = ({ mulKey, options }) => (
    <select value={inputs[mulKey] || ''} onChange={e => setInputs(p => ({ ...p, [mulKey]: e.target.value }))}
      style={{ fontFamily: 'Roboto,sans-serif', fontSize: 12, border: '1px solid #c4c4c4', borderRadius: 3, padding: '3px 5px', background: '#fff', width: '100%' }}>
      {Object.keys(options).map(o => <option key={o}>{o}</option>)}
    </select>
  );

  const NumberInput = ({ mulKey }) => (
    <input type="number" value={inputs[mulKey] || 0} onChange={e => setInputs(p => ({ ...p, [mulKey]: e.target.value }))}
      style={{ fontFamily: 'Roboto,sans-serif', fontSize: 12, border: '1px solid #c4c4c4', borderRadius: 3, padding: '3px 5px', width: 70 }} />
  );

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
      <PageHeader
        breadcrumb={`Implementation Hub › ${activeProject.name}`}
        title="Implementation Estimator"
        subtitle="Adjust inputs to calculate project effort"
        actions={<button className="btn btn-primary" onClick={handleSave}>💾 Save</button>}
      />
      <div style={{ flex: 1, overflowY: 'auto', padding: '18px 20px' }}>
        {/* Summary */}
        <div style={{ background: 'linear-gradient(135deg,#404789 0%,#5259b3 100%)', color: '#fff', padding: 20, borderRadius: 8, marginBottom: 14 }}>
          <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: '.5px', opacity: .85, marginBottom: 5 }}>Total Estimated Effort</div>
          <div style={{ fontSize: 30, fontWeight: 700, lineHeight: 1 }}>{finalTotal}h</div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 14, marginTop: 14, paddingTop: 14, borderTop: '1px solid rgba(255,255,255,.2)' }}>
            <div><div style={{ fontSize: 11, opacity: .8 }}>Base (before risk)</div><div style={{ fontSize: 16, fontWeight: 500, marginTop: 3 }}>{baseTotal}h</div></div>
            <div><div style={{ fontSize: 11, opacity: .8 }}>BA Effort</div><div style={{ fontSize: 16, fontWeight: 500, marginTop: 3 }}>{baHours}h</div></div>
            <div><div style={{ fontSize: 11, opacity: .8 }}>Engineer Effort</div><div style={{ fontSize: 16, fontWeight: 500, marginTop: 3 }}>{engHours}h</div></div>
          </div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 14 }}>
          {/* Step Table */}
          <WidgetCard title="Step-by-Step Breakdown">
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ background: '#f5f5f5', borderBottom: '1px solid #d0d0d0' }}>
                  <th style={{ padding: '8px', textAlign: 'left', fontWeight: 500, color: '#444' }}>Step</th>
                  <th style={{ padding: '8px', textAlign: 'left', fontWeight: 500, color: '#444' }}>Input</th>
                  <th style={{ padding: '8px', textAlign: 'left', fontWeight: 500, color: '#444' }}>Role</th>
                  <th style={{ padding: '8px', textAlign: 'right', fontWeight: 500, color: '#444' }}>Hours</th>
                </tr>
              </thead>
              <tbody>
                {steps.map((s, i) => (
                  <tr key={i} style={{ borderBottom: '1px solid #f0f0f0', background: i % 2 === 0 ? '#fafafe' : '#fff' }}>
                    <td style={{ padding: '6px 8px', color: '#404041' }}>{s.step}</td>
                    <td style={{ padding: '6px 8px' }}>
                      {typeof MULTIPLIERS[s.mul] === 'function'
                        ? <NumberInput mulKey={s.mul} />
                        : <SelectInput mulKey={s.mul} options={MULTIPLIERS[s.mul]} />}
                    </td>
                    <td style={{ padding: '6px 8px', color: '#666' }}>{s.role}</td>
                    <td style={{ padding: '6px 8px', textAlign: 'right', fontWeight: 500, color: '#404789' }}>{stepHours[i]}h</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </WidgetCard>

          {/* Risk Factors */}
          <WidgetCard title="Risk Adjustments">
            {Object.entries(RISK_FACTORS).map(([key, rf]) => (
              <label key={key} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 0', fontSize: 12, cursor: 'pointer', borderBottom: '1px solid #f5f5f5' }}>
                <input type="checkbox" checked={!!risks[key]} onChange={e => setRisks(p => ({ ...p, [key]: e.target.checked }))} />
                <span style={{ flex: 1 }}>{rf.label}</span>
                <span style={{ fontWeight: 500, color: rf.pct < 0 ? '#2e7d32' : '#e65100' }}>{rf.pct > 0 ? '+' : ''}{rf.pct}%</span>
              </label>
            ))}
            <div style={{ marginTop: 12, padding: '10px', background: '#f4f4f9', borderRadius: 4, fontSize: 12 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span style={{ color: '#666' }}>Total risk adjustment</span>
                <span style={{ fontWeight: 600, color: riskPct < 0 ? '#2e7d32' : '#e65100' }}>{riskPct > 0 ? '+' : ''}{riskPct}%</span>
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 6 }}>
                <span style={{ color: '#404041', fontWeight: 500 }}>Final Estimate</span>
                <span style={{ fontWeight: 700, color: '#404789', fontSize: 14 }}>{finalTotal}h</span>
              </div>
            </div>
          </WidgetCard>
        </div>
      </div>
    </div>
  );
}
