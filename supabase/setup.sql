-- Supabase の SQL エディタで 1 回だけ実行してください。
-- Table Editor → lingers_word_cloud に id=me の行ができる想定です。

create table if not exists lingers_word_cloud (
  id text primary key default 'me',
  body jsonb not null default '{}'::jsonb,
  updated_at timestamptz default now()
);

alter table lingers_word_cloud enable row level security;

-- 個人利用向け（anon キーで誰でも読み書き可）。プロジェクトURLは他人に共有しないでください。
create policy "lingers_rw" on lingers_word_cloud
  for all using (true) with check (true);
