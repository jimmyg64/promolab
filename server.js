require('dotenv').config();
const express = require('express');
const path = require('path');
const cookieParser = require('cookie-parser');
const { createClient } = require('@supabase/supabase-js');
const Anthropic = require('@anthropic-ai/sdk');
const fetch = require('node-fetch');

const app = express();
const PORT = process.env.PORT || 3000;

const adminClient = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SECRET_KEY);
const publicClient = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_PUBLISHABLE_KEY);
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

// ── Knowledge base — baked into every solo ads prompt ────────
const SOLO_ADS_KB = `
You are the Solo Ads Funnel Architect, an expert copywriter and funnel strategist specializing in cold traffic and the Make Money Online (MMO) niche.

GLOBAL WRITING RULES:
- Action-Driven & Angle-Based: Every piece of copy drives a specific action. Do not overcomplicate.
- Cold MMO Traffic Mindset: The audience is cold, skeptical, and looking for ways to make money online. Hooks must be punchy, curiosity-driven, and benefit-heavy.
- Voice: Relatable, conversational, and direct.
- No hallucinations. Formatting must be clean.

THE THREE ANGLES:
1. Consultative/Pathfinder: Best for overwhelmed audiences. Position offer as the most logical, clear path. Guide them to the click as the obvious next step.
2. Pain & Agitation: Best for audiences stuck using outdated/broken methods. Focus on frustration of current situation and present offer as instant relief.
3. Pure Value & Bonus: Best for highly competitive offers. Stack immense value of the main offer alongside custom gap-filling bonuses to make purchase a no-brainer.

BONUS STRATEGY RULES:
- Bonuses are PRIMARY conversion tools, not afterthoughts. They push leads over the edge to purchase.
- Bonuses are STRICTLY for buyers only — never mention on opt-in page.
- Fill the Value Gaps: Create bonuses that specifically fill missing pieces in the main offer.
- Bonus stack must be defined BEFORE writing bridge page or non-buyer email sequence so hooks can be woven in.

OPT-IN PAGE RULES:
- Goal: Capture contact details using curiosity and big promise. That's it.
- NO bonuses mentioned on opt-in page — buyers only.
- Ultra-concise above the fold: Headline + Sub-headline + CTA must fit one screen, no scrolling.
- Always generate TWO headline variations (Version A and Version B) for A/B split testing.
- Collect ONLY First Name and Email Address.

EMAIL SEQUENCE RULES (The Frequency Method):
- Always generate THREE distinct subject line options per email.
- Non-Buyers Sequence (Convert to Sale): Day 1 thank you + soft link, Day 2 story/solution, Day 3 offer + bonuses hard pitch, Day 4 authority + scarcity.
- Buyers Sequence (Onboard & Ascend): Day 1 welcome + bonus delivery, Day 2 advanced tip, Day 3 introduce upsell, Day 4 upsell benefits.
`;

// ── Auth helpers ──────────────────────────────────────────────
function cookieOptions() {
  return { httpOnly: true, secure: true, sameSite: 'lax', maxAge: 60*60*24*7*1000, path: '/' };
}

async function getCurrentUser(req) {
  const token = req.cookies && req.cookies.sb_access_token;
  if (!token) return null;
  try {
    const { data, error } = await adminClient.auth.getUser(token);
    if (error || !data || !data.user) return null;
    return data.user;
  } catch { return null; }
}

async function requireUser(req, res) {
  const user = await getCurrentUser(req);
  if (!user) { res.status(401).json({ success: false, message: 'Not logged in' }); return null; }
  if (!user.email_confirmed_at) { res.status(403).json({ success: false, message: 'Please confirm your email' }); return null; }
  return user;
}

async function getUserAccess(userId) {
  const { data } = await adminClient.from('promolab_access').select('*').eq('user_id', userId).single();
  return data || { solo_ads: false, facebook: false, email_sequence: false, launchjacking: false, affiliate_launch_guide: false };
}

// ── Page routes ───────────────────────────────────────────────
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('/login.html', (req, res) => res.sendFile(path.join(__dirname, 'public', 'login.html')));
app.get('/signup.html', (req, res) => res.sendFile(path.join(__dirname, 'public', 'signup.html')));
app.get('/health', (req, res) => res.json({ status: 'ok' }));

app.get('/app.html', async (req, res) => {
  const user = await getCurrentUser(req);
  if (!user) return res.redirect('/login.html');
  if (!user.email_confirmed_at) return res.redirect('/login.html?unconfirmed=1');
  res.sendFile(path.join(__dirname, 'public', 'app.html'));
});

app.get('/admin.html', async (req, res) => {
  const user = await getCurrentUser(req);
  if (!user) return res.redirect('/login.html');
  const { data: profile } = await adminClient.from('promolab_access').select('is_admin').eq('user_id', user.id).single();
  if (!profile || !profile.is_admin) return res.redirect('/app.html');
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

// ── Auth routes ───────────────────────────────────────────────
app.post('/signup', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, message: 'Email and password required' });
  if (password.length < 8) return res.status(400).json({ success: false, message: 'Password must be at least 8 characters' });
  try {
    const { data, error } = await publicClient.auth.signUp({ email: email.toLowerCase().trim(), password });
    if (error) return res.status(400).json({ success: false, message: error.message });
    res.json({ success: true, message: 'Account created. Check your email to confirm.' });
  } catch (err) { res.status(500).json({ success: false, message: 'Signup failed: ' + err.message }); }
});

app.post('/login', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, message: 'Email and password required' });
  try {
    const { data, error } = await publicClient.auth.signInWithPassword({ email: email.toLowerCase().trim(), password });
    if (error) return res.status(401).json({ success: false, message: error.message });
    if (!data.session) return res.status(401).json({ success: false, message: 'Login failed' });
    res.cookie('sb_access_token', data.session.access_token, cookieOptions());
    res.cookie('sb_refresh_token', data.session.refresh_token, cookieOptions());
    res.json({ success: true, email: data.user.email });
  } catch (err) { res.status(500).json({ success: false, message: 'Login failed: ' + err.message }); }
});

app.post('/logout', (req, res) => {
  res.clearCookie('sb_access_token', { path: '/' });
  res.clearCookie('sb_refresh_token', { path: '/' });
  res.json({ success: true });
});

app.get('/me', async (req, res) => {
  const user = await getCurrentUser(req);
  if (!user) return res.json({ success: false, user: null });
  const access = await getUserAccess(user.id);
  res.json({ success: true, user: { id: user.id, email: user.email, email_confirmed: !!user.email_confirmed_at }, access });
});

// ── ANALYZE URL ───────────────────────────────────────────────
app.post('/api/analyze', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { url } = req.body;
  if (!url) return res.status(400).json({ success: false, message: 'URL required' });

  try {
    let pageText = '';
    try {
      const r = await fetch(url, { timeout: 8000, headers: { 'User-Agent': 'Mozilla/5.0' } });
      const html = await r.text();
      pageText = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 5000);
    } catch { pageText = 'Could not fetch page automatically.'; }

    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 800,
      messages: [{
        role: 'user',
        content: `Analyze this affiliate sales page and return ONLY a JSON object with these exact fields:
{
  "product_name": "name of the product",
  "price": "price shown",
  "commission": "commission percentage or amount",
  "niche": "primary niche (e.g. Make Money Online, Health, etc)",
  "target_audience": "who this is for in one sentence",
  "main_pain_point": "the biggest problem this solves",
  "main_benefit": "the single strongest benefit or promise",
  "summary": "2-3 sentence summary of what this is and why someone would buy it",
  "offer_score": {
    "overall": 8,
    "commission_rating": 9,
    "niche_demand": 8,
    "conversion_potential": 7,
    "tier1_suitability": 8,
    "notes": "brief explanation of scores"
  },
  "value_gaps": ["gap 1 the offer doesn't address", "gap 2", "gap 3"],
  "recommended_angle": "Consultative or Pain & Agitation or Pure Value & Bonus",
  "angle_reason": "one sentence explaining why this angle fits"
}
Page content: ${pageText}`
      }]
    });

    let parsed;
    try {
      const text = message.content[0].text;
      parsed = JSON.parse(text.replace(/```json|```/g, '').trim());
    } catch {
      parsed = {
        product_name: 'Affiliate product', price: 'See page', commission: 'See page',
        niche: 'Make Money Online', target_audience: 'Beginner affiliate marketers',
        main_pain_point: 'Not making money online', main_benefit: 'Start earning commissions',
        summary: 'An affiliate marketing product with training and tools.',
        offer_score: { overall: 7, commission_rating: 7, niche_demand: 7, conversion_potential: 7, tier1_suitability: 7, notes: 'Score estimated' },
        value_gaps: ['No implementation support', 'No traffic strategy', 'No email templates'],
        recommended_angle: 'Pain & Agitation', angle_reason: 'MMO audiences respond well to pain-based hooks'
      };
    }

    res.json({ success: true, analysis: parsed });
  } catch (err) { res.status(500).json({ success: false, message: 'Analysis failed: ' + err.message }); }
});

// ── GENERATE BONUS STACK ──────────────────────────────────────
app.post('/api/generate/bonuses', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, project_id } = req.body;
  if (!analysis || !angle) return res.status(400).json({ success: false, message: 'Analysis and angle required' });

  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 4000,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Create a buyer-only bonus stack for this affiliate offer. The chosen angle is: ${angle}.

Offer details:
- Product: ${analysis.product_name}
- Niche: ${analysis.niche}
- Main benefit: ${analysis.main_benefit}
- Value gaps to fill: ${(analysis.value_gaps || []).join(', ')}
- Target audience: ${analysis.target_audience}

Generate exactly 3 bonuses. Keep full_content concise (150-200 words max each). Return ONLY a valid JSON array with no text before or after it:
[
  {
    "number": 1,
    "title": "Bonus title here",
    "type": "1-Page Checklist",
    "description": "One sentence: what this bonus does and why it fills a gap",
    "full_content": "The full checklist content here in 150-200 words",
    "cover_prompt": "A professional book cover: dark navy background, gold title text, checklist icon, clean modern design, 1080x1080"
  },
  {
    "number": 2,
    "title": "Bonus title here",
    "type": "2-Page Guide",
    "description": "One sentence description",
    "full_content": "The full guide content here in 150-200 words",
    "cover_prompt": "A professional book cover: deep blue background, white title text, open book icon, modern design, 1080x1080"
  },
  {
    "number": 3,
    "title": "Bonus title here",
    "type": "AI Prompt Pack",
    "description": "One sentence description",
    "full_content": "5 detailed AI prompts with brief explanations",
    "cover_prompt": "A professional book cover: dark purple background, gold title text, AI circuit icon, futuristic design, 1080x1080"
  }
]`
      }]
    });

    let bonuses;
    const rawText = message.content[0].text;
    try {
      bonuses = JSON.parse(rawText.replace(/```json|```/g, '').trim());
    } catch {
      try {
        const match = rawText.match(/\[[\s\S]*\]/);
        if (match) { bonuses = JSON.parse(match[0]); }
        else { throw new Error('No JSON array found'); }
      } catch {
        bonuses = [
          { number:1, title:`${analysis.product_name} Quick Start Checklist`, type:'1-Page Checklist',
            description:`Step-by-step checklist to get fast results from ${analysis.product_name}`,
            full_content:`Step 1: Complete your setup\nStep 2: Go through core training in order\nStep 3: Apply each lesson before moving on\nStep 4: Track your results daily\nStep 5: Scale what is working`,
            cover_prompt:'Professional book cover, dark navy background, gold checklist icon, bold white title, clean modern, 1080x1080' },
          { number:2, title:`${analysis.niche} Fast Track Guide`, type:'2-Page Guide',
            description:`Practical guide to getting your first result in ${analysis.niche}`,
            full_content:`This guide walks you through the fastest path to results. Focus on one strategy at a time. Consistent daily action beats shortcuts every time.`,
            cover_prompt:'Professional book cover, deep blue background, white title, rocket icon, modern clean, 1080x1080' },
          { number:3, title:'AI Prompt Pack — Faster Results', type:'AI Prompt Pack',
            description:'5 copy-paste AI prompts to speed up your results',
            full_content:`Prompt 1: Write a social media post promoting [product] to [audience]\nPrompt 2: Write a follow-up email to someone who opted in but did not buy\nPrompt 3: Write 5 subject line options for a promotional email\nPrompt 4: Write a bridge page script using the pain and agitation angle\nPrompt 5: Write a Facebook post using a personal story to promote [product]`,
            cover_prompt:'Professional book cover, dark purple background, gold AI icon, bold white title, futuristic, 1080x1080' }
        ];
      }
    }

    // Generate cover images for each bonus
    for (let i = 0; i < bonuses.length; i++) {
      try {
        const imgResponse = await fetch('https://api.stability.ai/v1/generation/stable-diffusion-v1-6/text-to-image', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + process.env.STABILITY_API_KEY,
            'Accept': 'application/json'
          },
          body: JSON.stringify({
            text_prompts: [{ text: bonuses[i].cover_prompt + ', professional book cover, high quality, sharp text, marketing material', weight: 1 }],
            cfg_scale: 7, height: 1024, width: 1024, samples: 1, steps: 30
          })
        });

        if (imgResponse.ok) {
          const imgData = await imgResponse.json();
          if (imgData.artifacts && imgData.artifacts[0]) {
            bonuses[i].cover_image = 'data:image/png;base64,' + imgData.artifacts[0].base64;
          }
        }
      } catch { bonuses[i].cover_image = null; }
    }

    // Save to project if project_id provided
    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({
        project_id, user_id: user.id, content_type: 'bonus_stack',
        content: JSON.stringify(bonuses), updated_at: new Date().toISOString()
      }, { onConflict: 'project_id,content_type' });
    }

    res.json({ success: true, bonuses });
  } catch (err) { res.status(500).json({ success: false, message: 'Bonus generation failed: ' + err.message }); }
});

// ── GENERATE OPT-IN PAGE ──────────────────────────────────────
app.post('/api/generate/optin', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, project_id } = req.body;
  if (!analysis || !angle) return res.status(400).json({ success: false, message: 'Analysis and angle required' });

  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Write an opt-in squeeze page for cold solo ad traffic using the ${angle} angle.

Offer: ${analysis.product_name}
Niche: ${analysis.niche}
Main pain point: ${analysis.main_pain_point}
Main benefit: ${analysis.main_benefit}
Target audience: ${analysis.target_audience}

RULES: No bonuses mentioned. Ultra-concise above the fold. Two A/B headline variations.

Return ONLY JSON:
{
  "headline_a": "Version A headline — curiosity/benefit driven",
  "headline_b": "Version B headline — different angle/hook",
  "subheadline": "Supporting subheadline that expands the promise",
  "cta_button": "CTA button text",
  "form_fields": ["First Name", "Email Address"],
  "above_fold_note": "Brief note on how to lay this out above the fold",
  "split_test_tip": "Which headline to test first and why based on the angle"
}`
      }]
    });

    let content;
    try { content = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch { return res.status(500).json({ success: false, message: 'Failed to parse opt-in content' }); }

    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({
        project_id, user_id: user.id, content_type: 'optin_page',
        content: JSON.stringify(content), updated_at: new Date().toISOString()
      }, { onConflict: 'project_id,content_type' });
    }

    res.json({ success: true, content });
  } catch (err) { res.status(500).json({ success: false, message: 'Opt-in generation failed: ' + err.message }); }
});

// ── GENERATE BRIDGE PAGE ──────────────────────────────────────
app.post('/api/generate/bridge', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, format, bonuses, project_id } = req.body;
  if (!analysis || !angle || !format) return res.status(400).json({ success: false, message: 'Analysis, angle, and format required' });

  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  const bonusContext = bonuses && bonuses.length > 0
    ? 'Buyer bonuses to weave in: ' + bonuses.map(b => b.title + ' — ' + b.description).join(', ')
    : 'No bonuses defined — write without bonus references.';

  const formatInstructions = {
    text: 'Write text-only copy: Headline, relatable story (2-3 paragraphs), benefit bullets, bonus pitch, CTA.',
    vsl: 'Write a VSL video script (2-3 minutes when read) PLUS supporting text copy. Script: Hook, intro, transition, bonus reveal, CTA.',
    short: 'Write a short punchy intro blurb (3-5 sentences max) with a strong CTA button. Keep it extremely tight.'
  };

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 2000,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Write a bridge page for this affiliate offer using the ${angle} angle.

Offer: ${analysis.product_name}
Main benefit: ${analysis.main_benefit}
Main pain point: ${analysis.main_pain_point}
${bonusContext}

Format: ${formatInstructions[format] || formatInstructions.text}

Critical: Bridge the gap from opt-in to sales page. Highlight strongest angle AND pitch the bonus stack as reward for buying through affiliate link today.

Return ONLY JSON:
{
  "headline": "Bridge page headline",
  "content": "Full bridge page copy as requested in format",
  "cta_text": "CTA button text",
  "above_fold_note": "What must appear above fold"
}`
      }]
    });

    let content;
    try { content = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch { return res.status(500).json({ success: false, message: 'Failed to parse bridge content' }); }

    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({
        project_id, user_id: user.id, content_type: 'bridge_page',
        content: JSON.stringify(content), updated_at: new Date().toISOString()
      }, { onConflict: 'project_id,content_type' });
    }

    res.json({ success: true, content });
  } catch (err) { res.status(500).json({ success: false, message: 'Bridge generation failed: ' + err.message }); }
});

// ── GENERATE EMAIL SEQUENCES ──────────────────────────────────
app.post('/api/generate/emails', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, sequence_type, bonuses, project_id } = req.body;
  if (!analysis || !angle || !sequence_type) return res.status(400).json({ success: false, message: 'Analysis, angle, and sequence type required' });

  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  const bonusContext = bonuses && bonuses.length > 0
    ? 'Buyer bonuses: ' + bonuses.map(b => b.title + ' — ' + b.description).join(', ')
    : 'No bonuses defined.';

  const sequenceInstructions = {
    non_buyers: `Non-Buyers 4-Day Sequence (Goal: Convert to Sale):
Day 1: Thank you for opting in, check for questions, soft link to offer
Day 2: Story/Solution — relate a problem and pitch the offer
Day 3: Offer & Bonuses — hard pitch using buyer bonuses as main incentive
Day 4: Follow-up & Authority — final push with scarcity`,
    buyers: `Buyers 4-Day Sequence (Goal: Onboard & Ascend):
Day 1: Welcome, thank them for buying, tell them how to access their 3 bonuses
Day 2: Advanced Tip — provide standalone value
Day 3: Introduce related solution or upsell
Day 4: Demo and benefits of new solution`
  };

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 3000,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Write a ${sequence_type === 'non_buyers' ? 'Non-Buyers' : 'Buyers'} email sequence using the ${angle} angle.

Offer: ${analysis.product_name}
Niche: ${analysis.niche}
Main pain point: ${analysis.main_pain_point}
Main benefit: ${analysis.main_benefit}
${bonusContext}

${sequenceInstructions[sequence_type]}

CRITICAL: Generate THREE distinct subject line options for EVERY email. Mix: curiosity, direct benefit, relatable question.

Return ONLY a JSON array of 4 emails:
[
  {
    "day": 1,
    "subject_lines": ["Option 1", "Option 2", "Option 3"],
    "body": "Full email body with [FirstName] personalization and [YourAffiliateLink] placeholder"
  }
]`
      }]
    });

    let emails;
    try { emails = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch { return res.status(500).json({ success: false, message: 'Failed to parse email content' }); }

    const contentType = sequence_type === 'non_buyers' ? 'emails_non_buyers' : 'emails_buyers';
    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({
        project_id, user_id: user.id, content_type: contentType,
        content: JSON.stringify(emails), updated_at: new Date().toISOString()
      }, { onConflict: 'project_id,content_type' });
    }

    res.json({ success: true, emails });
  } catch (err) { res.status(500).json({ success: false, message: 'Email generation failed: ' + err.message }); }
});

// ── PROJECT MANAGEMENT ────────────────────────────────────────
app.post('/api/projects', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { name, url, analysis, angle, channel } = req.body;

  const { data, error } = await adminClient.from('promolab_projects').insert({
    user_id: user.id, name: name || analysis.product_name,
    url: url || '', channel: channel || 'solo_ads',
    angle, analysis: JSON.stringify(analysis),
    created_at: new Date().toISOString(), updated_at: new Date().toISOString()
  }).select().single();

  if (error) return res.status(400).json({ success: false, message: error.message });
  res.json({ success: true, project: data });
});

app.get('/api/projects', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { data } = await adminClient.from('promolab_projects').select('*').eq('user_id', user.id).order('updated_at', { ascending: false });
  res.json({ success: true, projects: data || [] });
});

app.get('/api/projects/:id', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { data: project } = await adminClient.from('promolab_projects').select('*').eq('id', req.params.id).eq('user_id', user.id).single();
  if (!project) return res.status(404).json({ success: false, message: 'Project not found' });
  const { data: content } = await adminClient.from('promolab_project_content').select('*').eq('project_id', req.params.id);
  const contentMap = {};
  (content || []).forEach(c => { contentMap[c.content_type] = JSON.parse(c.content); });
  res.json({ success: true, project, content: contentMap });
});

app.delete('/api/projects/:id', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  await adminClient.from('promolab_projects').delete().eq('id', req.params.id).eq('user_id', user.id);
  res.json({ success: true });
});

// ── ADMIN ─────────────────────────────────────────────────────
app.post('/api/admin/access', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { data: profile } = await adminClient.from('promolab_access').select('is_admin').eq('user_id', user.id).single();
  if (!profile || !profile.is_admin) return res.status(403).json({ success: false, message: 'Admin only' });
  const { target_email, solo_ads, facebook, email_sequence, launchjacking, affiliate_launch_guide } = req.body;
  const { data: users } = await adminClient.auth.admin.listUsers();
  const found = users.users.find(u => u.email === target_email);
  if (!found) return res.status(404).json({ success: false, message: 'User not found' });
  const { error } = await adminClient.from('promolab_access').upsert({
    user_id: found.id,
    solo_ads: solo_ads || false, facebook: facebook || false,
    email_sequence: email_sequence || false, launchjacking: launchjacking || false,
    affiliate_launch_guide: affiliate_launch_guide || false,
    updated_at: new Date().toISOString()
  }, { onConflict: 'user_id' });
  if (error) return res.status(400).json({ success: false, message: error.message });
  res.json({ success: true, message: 'Access updated for ' + target_email });
});

app.get('/api/admin/users', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { data: profile } = await adminClient.from('promolab_access').select('is_admin').eq('user_id', user.id).single();
  if (!profile || !profile.is_admin) return res.status(403).json({ success: false, message: 'Admin only' });
  const { data } = await adminClient.from('promolab_access').select('*');
  res.json({ success: true, users: data || [] });
});

app.listen(PORT, () => console.log('PromoLab v2 running on port ' + PORT));
