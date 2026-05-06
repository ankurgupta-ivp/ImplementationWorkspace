-- ============================================================
-- PriceMaster Implementation Hub — Supabase Schema
-- Run this entire file in: Supabase Dashboard → SQL Editor
-- ============================================================

-- Projects table
CREATE TABLE IF NOT EXISTS projects (
  id          TEXT PRIMARY KEY,
  name        TEXT NOT NULL,
  created_at  DATE DEFAULT CURRENT_DATE,
  metadata    JSONB DEFAULT '{}'::jsonb,
  questionnaire JSONB DEFAULT '{}'::jsonb,
  estimator   JSONB DEFAULT '{}'::jsonb
);

-- Tasks table (one row per task, linked to a project)
CREATE TABLE IF NOT EXISTS tasks (
  id              SERIAL PRIMARY KEY,
  project_id      TEXT NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
  sort_order      INTEGER DEFAULT 0,
  phase           TEXT DEFAULT '',
  item            TEXT DEFAULT '',
  task            TEXT DEFAULT '',
  task_type       TEXT DEFAULT '',
  tags            TEXT DEFAULT '',
  responsible     TEXT DEFAULT '',
  owner           TEXT DEFAULT '',
  reviewer        TEXT DEFAULT '',
  owner_status    TEXT DEFAULT 'Not Started',
  reviewer_status TEXT DEFAULT 'Not Started',
  comments        TEXT DEFAULT '',
  actual_start    TEXT DEFAULT '',
  actual_end      TEXT DEFAULT '',
  expected_start  TEXT DEFAULT '',
  expected_end    TEXT DEFAULT '',
  actual_effort   TEXT DEFAULT '',
  expected_effort TEXT DEFAULT '',
  actual_elapsed  TEXT DEFAULT '',
  expected_elapsed TEXT DEFAULT ''
);

-- RAID Log table
CREATE TABLE IF NOT EXISTS raid_items (
  id              SERIAL PRIMARY KEY,
  project_id      TEXT NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
  classification  TEXT DEFAULT 'Risk',
  item            TEXT DEFAULT '',
  details         TEXT DEFAULT '',
  raised_on       TEXT DEFAULT '',
  pending_with    TEXT DEFAULT '',
  updates         TEXT DEFAULT '',
  status          TEXT DEFAULT 'Not Initiated',
  eta             TEXT DEFAULT '',
  comments        TEXT DEFAULT ''
);

-- Global templates table (one row per template type)
CREATE TABLE IF NOT EXISTS templates (
  id          TEXT PRIMARY KEY,
  name        TEXT NOT NULL,
  version     TEXT DEFAULT '1.0',
  updated     TEXT DEFAULT '',
  data        JSONB NOT NULL DEFAULT '{}'::jsonb
);

-- App state table (stores global settings like active project)
CREATE TABLE IF NOT EXISTS app_state (
  key   TEXT PRIMARY KEY,
  value TEXT
);

-- ── Indexes for performance ─────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_tasks_project_id    ON tasks(project_id);
CREATE INDEX IF NOT EXISTS idx_raid_project_id     ON raid_items(project_id);

-- ── Row Level Security ──────────────────────────────────────
-- All team members share the same data (no per-user isolation).
-- Enable RLS but allow all operations for the anon key.
-- If you want per-user data in future, replace these policies.

ALTER TABLE projects    ENABLE ROW LEVEL SECURITY;
ALTER TABLE tasks       ENABLE ROW LEVEL SECURITY;
ALTER TABLE raid_items  ENABLE ROW LEVEL SECURITY;
ALTER TABLE templates   ENABLE ROW LEVEL SECURITY;
ALTER TABLE app_state   ENABLE ROW LEVEL SECURITY;

-- Public (team-wide) access policies
CREATE POLICY "team_access_projects"   ON projects    FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "team_access_tasks"      ON tasks       FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "team_access_raid"       ON raid_items  FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "team_access_templates"  ON templates   FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "team_access_app_state"  ON app_state   FOR ALL USING (true) WITH CHECK (true);
