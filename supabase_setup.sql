-- Run this in Supabase SQL Editor

-- Access control
CREATE TABLE promolab_access (
  user_id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  solo_ads BOOLEAN DEFAULT false,
  facebook BOOLEAN DEFAULT false,
  email_sequence BOOLEAN DEFAULT false,
  launchjacking BOOLEAN DEFAULT false,
  affiliate_launch_guide BOOLEAN DEFAULT false,
  is_admin BOOLEAN DEFAULT false,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Projects
CREATE TABLE promolab_projects (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  name TEXT NOT NULL,
  url TEXT,
  channel TEXT DEFAULT 'solo_ads',
  angle TEXT,
  analysis JSONB,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Project content (each piece saved separately)
CREATE TABLE promolab_project_content (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  project_id UUID NOT NULL REFERENCES promolab_projects(id) ON DELETE CASCADE,
  user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  content_type TEXT NOT NULL,
  content JSONB,
  updated_at TIMESTAMPTZ DEFAULT NOW(),
  UNIQUE(project_id, content_type)
);

-- RLS
ALTER TABLE promolab_access ENABLE ROW LEVEL SECURITY;
ALTER TABLE promolab_projects ENABLE ROW LEVEL SECURITY;
ALTER TABLE promolab_project_content ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users see own access" ON promolab_access FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "Users manage own projects" ON promolab_projects FOR ALL USING (auth.uid() = user_id);
CREATE POLICY "Users manage own content" ON promolab_project_content FOR ALL USING (auth.uid() = user_id);

-- Set yourself as admin (replace YOUR-USER-ID with your actual Supabase user ID)
-- INSERT INTO promolab_access (user_id, solo_ads, facebook, email_sequence, launchjacking, affiliate_launch_guide, is_admin)
-- VALUES ('YOUR-USER-ID', true, true, true, true, true, true);
