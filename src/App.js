import React from 'react';
import { AppProvider, useApp } from './hooks/useApp';
import AppShell from './components/AppShell';
import { Toast } from './components/UI';
import Overview from './pages/Overview';
import Kickoff from './pages/Kickoff';
import Tasks from './pages/Tasks';
import RaidLog from './pages/RaidLog';
import Estimator from './pages/Estimator';
import { Dashboard, Templates, Docs, DataSources } from './pages/OtherPages';
import GanttChart from './pages/GanttChart';

const PAGE_COMPONENTS = {
  overview: Overview,
  kickoff: Kickoff,
  tasks: Tasks,
  raidlog: RaidLog,
  estimator: Estimator,
  dashboard: Dashboard,
  templates: Templates,
  docs: Docs,
  datasources: DataSources,
  gantt: GanttChart,
};

function AppContent() {
  const { currentPage, loading, toast } = useApp();
  const PageComponent = PAGE_COMPONENTS[currentPage] || Overview;

  return (
    <AppShell>
      {loading
        ? <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', flex: 1, color: '#888', fontSize: 14 }}>Loading…</div>
        : <PageComponent />
      }
      <Toast toast={toast} />
    </AppShell>
  );
}

export default function App() {
  return (
    <AppProvider>
      <AppContent />
    </AppProvider>
  );
}
