-- AnchorCast adaptive local transcript schema
-- Current build runtime uses JSON files under Data/adaptive/ for portability.
-- This SQL schema documents the intended relational structure for future migration.

CREATE TABLE speaker_profiles (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  description TEXT,
  is_default INTEGER DEFAULT 0,
  is_active INTEGER DEFAULT 1,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE transcript_sessions (
  id TEXT PRIMARY KEY,
  title TEXT,
  speaker_profile_id TEXT,
  mode TEXT NOT NULL,
  started_at TEXT NOT NULL,
  ended_at TEXT,
  notes TEXT
);

CREATE TABLE transcript_chunks (
  id TEXT PRIMARY KEY,
  session_id TEXT NOT NULL,
  speaker_profile_id TEXT,
  chunk_index INTEGER NOT NULL,
  raw_text TEXT NOT NULL,
  normalized_text TEXT,
  corrected_text TEXT,
  final_text TEXT,
  confidence REAL,
  created_at TEXT NOT NULL
);

CREATE TABLE correction_rules (
  id TEXT PRIMARY KEY,
  scope TEXT NOT NULL,
  speaker_profile_id TEXT,
  rule_type TEXT NOT NULL,
  source_text TEXT NOT NULL,
  target_text TEXT NOT NULL,
  priority INTEGER DEFAULT 100,
  confidence REAL DEFAULT 1.0,
  hit_count INTEGER DEFAULT 0,
  approved_count INTEGER DEFAULT 0,
  rejected_count INTEGER DEFAULT 0,
  is_active INTEGER DEFAULT 1,
  created_by TEXT DEFAULT 'system',
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE correction_events (
  id TEXT PRIMARY KEY,
  transcript_chunk_id TEXT NOT NULL,
  rule_id TEXT,
  before_text TEXT NOT NULL,
  after_text TEXT NOT NULL,
  was_user_edited INTEGER DEFAULT 0,
  accepted INTEGER,
  created_at TEXT NOT NULL
);

CREATE TABLE custom_vocabulary (
  id TEXT PRIMARY KEY,
  scope TEXT NOT NULL,
  speaker_profile_id TEXT,
  canonical_term TEXT NOT NULL,
  aliases_json TEXT,
  category TEXT,
  weight REAL DEFAULT 1.0,
  usage_count INTEGER DEFAULT 0,
  is_active INTEGER DEFAULT 1,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE verse_detection_events (
  id TEXT PRIMARY KEY,
  transcript_chunk_id TEXT NOT NULL,
  detected_ref TEXT NOT NULL,
  detected_method TEXT NOT NULL,
  confidence REAL,
  accepted INTEGER,
  created_at TEXT NOT NULL
);

CREATE TABLE learning_jobs (
  id TEXT PRIMARY KEY,
  started_at TEXT NOT NULL,
  ended_at TEXT,
  status TEXT NOT NULL,
  summary_json TEXT
);
