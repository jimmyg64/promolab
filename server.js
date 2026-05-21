require('dotenv').config();
const express = require('express');
const path = require('path');
const cookieParser = require('cookie-parser');
const { createClient } = require('@supabase/supabase-js');
const Anthropic = require('@anthropic-ai/sdk');
const fetch = require('node-fetch');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType, ShadingType } = require('docx');

const app = express();
const PORT = process.env.PORT || 3000;

const adminClient = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SECRET_KEY);
const publicClient = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_PUBLISHABLE_KEY);
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

// ── Knowledge base ────────────────────────────────────────────
const SOLO_ADS_KB = `
You are the Solo Ads Funnel Architect, an expert copywriter and funnel strategist specializing in cold traffic and the Make Money Online (MMO) niche.

GLOBAL WRITING RULES:
- Action-Driven & Angle-Based: Every piece of copy drives a specific action. Do not overcomplicate.
- Cold MMO Traffic Mindset: The audience is cold, skeptical, and has likely tried and failed before. Hooks must be punchy, curiosity-driven, and benefit-heavy.
- Voice: Relatable, conversational, direct. Write like a helpful friend, not a salesperson.
- No hallucinations. No fluff. No filler sentences. Every sentence must earn its place.
- Use short paragraphs. Use white space. Make content easy to scan and read.

THE THREE ANGLES:
1. Consultative/Pathfinder: Best for overwhelmed audiences. Position offer as the most logical, clear path.
2. Pain & Agitation: Best for audiences stuck using outdated/broken methods. Focus on frustration and present offer as instant relief.
3. Pure Value & Bonus: Best for highly competitive offers. Stack immense value alongside custom gap-filling bonuses.

BONUS & LEAD MAGNET CONTENT RULES:
- Content must be genuinely valuable — not filler, not vague advice
- Use clear headers, numbered steps, bullet points, real examples
- Write like an expert coaching a friend — specific, direct, practical
- Every section must give the reader something they can immediately use
- Format for readability: short paragraphs, white space, scannable structure
- These are standalone deliverables — they must feel premium and complete
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

// ── DALL-E 3 image generation ─────────────────────────────────
async function generateImage(prompt) {
  if (process.env.OPENAI_API_KEY) {
    try {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 35000);
      const response = await fetch('https://api.openai.com/v1/images/generations', {
        method: 'POST', signal: ctrl.signal,
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + process.env.OPENAI_API_KEY },
        body: JSON.stringify({ model: 'dall-e-3', prompt, n: 1, size: '1024x1024', response_format: 'b64_json' })
      });
      clearTimeout(timer);
      if (response.ok) {
        const data = await response.json();
        if (data.data && data.data[0] && data.data[0].b64_json) return 'data:image/png;base64,' + data.data[0].b64_json;
      } else { console.log('DALL-E error:', response.status, await response.text()); }
    } catch(e) { console.log('DALL-E failed:', e.message); }
  }
  // Fallback Stability AI
  if (process.env.STABILITY_API_KEY) {
    try {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 25000);
      const response = await fetch('https://api.stability.ai/v1/generation/stable-diffusion-xl-1024-v1-0/text-to-image', {
        method: 'POST', signal: ctrl.signal,
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + process.env.STABILITY_API_KEY, 'Accept': 'application/json' },
        body: JSON.stringify({ text_prompts: [{ text: prompt, weight: 1 }, { text: 'blurry, low quality, people, faces, watermark', weight: -1 }], cfg_scale: 7, height: 1024, width: 1024, samples: 1, steps: 25 })
      });
      clearTimeout(timer);
      if (response.ok) {
        const data = await response.json();
        if (data.artifacts && data.artifacts[0]) return 'data:image/png;base64,' + data.artifacts[0].base64;
      }
    } catch(e) { console.log('Stability failed:', e.message); }
  }
  return null;
}

// ── DOCX builder helpers ──────────────────────────────────────
function makeParagraph(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: opts.spaceBefore || 100, after: opts.spaceAfter || 100 },
    children: [new TextRun({
      text: text || '',
      bold: opts.bold || false,
      italic: opts.italic || false,
      size: opts.size || 24,
      color: opts.color || '1A1A2E',
      font: 'Calibri'
    })]
  });
}

function makeHeading(text, level = 1) {
  const sizes = { 1: 40, 2: 32, 3: 28 };
  const colors = { 1: '1A1A2E', 2: '4F46B8', 3: '333333' };
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
    spacing: { before: level === 1 ? 400 : 240, after: 160 },
    children: [new TextRun({ text, bold: true, size: sizes[level] || 28, color: colors[level] || '1A1A2E', font: 'Calibri' })]
  });
}

function makeBullet(text) {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 24, color: '333333', font: 'Calibri' })]
  });
}

function makeDivider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { color: '4F46B8', space: 1, style: BorderStyle.SINGLE, size: 6 } },
    children: []
  });
}

function makeCoverPage(title, subtitle, type, productName) {
  return [
    new Paragraph({ spacing: { before: 800, after: 200 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: type.toUpperCase(), size: 20, color: 'AAAACC', font: 'Calibri', bold: true })] }),
    new Paragraph({ spacing: { before: 0, after: 300 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: title, size: 52, bold: true, color: '1A1A2E', font: 'Calibri' })] }),
    new Paragraph({ spacing: { before: 0, after: 200 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: subtitle || '', size: 28, italic: true, color: '555555', font: 'Calibri' })] }),
    makeDivider(),
    new Paragraph({ spacing: { before: 200, after: 100 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'Exclusive bonus for buyers of ' + productName, size: 20, color: '4F46B8', font: 'Calibri' })] }),
    new Paragraph({ spacing: { before: 0, after: 100 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'PromoLab by Jimmy Griffith, JGAffiliate', size: 18, color: 'AAAACC', font: 'Calibri' })] }),
    new Paragraph({ pageBreakBefore: true, children: [] })
  ];
}

function cleanMarkdownLine(line) {
  return line
    .replace(/^#{1,6}\s+/, '')         // strip # headers (handled separately)
    .replace(/\*\*(.*?)\*\*/g, '$1') // **bold** -> bold
    .replace(/\*(.*?)\*/g, '$1')       // *italic* -> italic
    .replace(/`([^`]+)`/g, '$1')         // `code` -> code
    .replace(/^[-*+]\s+/, '')           // strip bullet markers (handled separately)
    .trim();
}

function parseContentToDocx(content, productName, title, subtitle, type) {
  const children = [];
  children.push(...makeCoverPage(title, subtitle, type, productName));

  const lines = content.split('\n');
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) { children.push(new Paragraph({ spacing: { before: 60, after: 60 }, children: [] })); continue; }
    if (trimmed.startsWith('### ')) { children.push(makeHeading(trimmed.slice(4), 3)); }
    else if (trimmed.startsWith('## ')) { children.push(makeHeading(trimmed.slice(3), 2)); }
    else if (trimmed.startsWith('# ')) { children.push(makeHeading(trimmed.slice(2), 1)); }
    else if (trimmed.startsWith('- ') || trimmed.startsWith('• ')) { children.push(makeBullet(trimmed.slice(2))); }
    else if (/^\d+\.\s/.test(trimmed)) { children.push(makeBullet(cleanMarkdownLine(trimmed))); }
    else if (trimmed.startsWith('**') && trimmed.endsWith('**')) {
      children.push(makeParagraph(trimmed.slice(2, -2), { bold: true, size: 26, spaceBefore: 160, spaceAfter: 80 }));
    }
    else {
      // Strip any remaining inline markdown: **bold**, *italic*, `code`
      let clean = trimmed
        .replace(/\*\*(.*?)\*\*/g, '$1')   // remove **bold**
        .replace(/\*(.*?)\*/g, '$1')          // remove *italic*
        .replace(/`([^`]+)`/g, '$1')            // remove `code`
        .replace(/^#+\s+/, '')                 // remove leading # symbols
        .replace(/^[-*+]\s+/, '')              // remove leading - * + bullets
        .trim();
      if (clean) children.push(makeParagraph(clean, { spaceBefore: 80, spaceAfter: 80 }));
    }
  }

  children.push(makeDivider());
  children.push(makeParagraph('© 2026 PromoLab · Jimmy Griffith, JGAffiliate · All rights reserved', { center: true, size: 18, color: 'AAAAAA', spaceBefore: 200 }));
  return children;
}

// ── DOCX route helper ─────────────────────────────────────────
async function buildAndSendDocx(res, children, filename) {
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'Calibri', size: 24, color: '1A1A2E' } }
      }
    },
    sections: [{ properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children }]
  });
  const buffer = await Packer.toBuffer(doc);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(buffer);
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

// ── ANALYZE ───────────────────────────────────────────────────
app.post('/api/analyze', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { url } = req.body;
  if (!url) return res.status(400).json({ success: false, message: 'URL required' });
  try {
    let pageText = '';
    try {
      const r = await fetch(url, { timeout: 10000, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' } });
      const html = await r.text();
      pageText = html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ').replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 8000);
    } catch { pageText = 'Could not fetch page. Analyze based on URL context.'; }

    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514', max_tokens: 2000,
      messages: [{ role: 'user', content: `Analyze this affiliate sales page deeply. Return ONLY valid JSON, no text before or after:
{
  "product_name": "exact name",
  "price": "price shown",
  "commission": "commission rate",
  "niche": "primary niche",
  "target_audience": "detailed who this is for",
  "main_pain_point": "biggest frustration",
  "secondary_pain_points": ["pain 2", "pain 3", "pain 4"],
  "main_benefit": "strongest promise",
  "unique_mechanism": "what makes this different",
  "audience_psychology": "2-3 sentences on emotional state, past failures, what they really want",
  "summary": "3-4 sentence summary",
  "offer_score": { "overall": 8, "commission_rating": 7, "niche_demand": 8, "conversion_potential": 7, "tier1_suitability": 8, "notes": "2-3 sentence explanation" },
  "value_gaps": ["gap 1", "gap 2", "gap 3", "gap 4"],
  "recommended_angle": "Consultative or Pain & Agitation or Pure Value & Bonus",
  "angle_reason": "2-3 sentences why this angle fits"
}
Page: ${pageText}` }]
    });
    let parsed;
    try { parsed = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch { parsed = { product_name:'Affiliate product', price:'See page', commission:'See page', niche:'Make Money Online', target_audience:'Beginner affiliate marketers', main_pain_point:'Not making money online despite trying', secondary_pain_points:['Tech overwhelm','Wasted money','No clear path'], main_benefit:'Start earning commissions', unique_mechanism:'Done-for-you system', audience_psychology:'Skeptical beginners who have tried and failed before.', summary:'An affiliate marketing product for beginners.', offer_score:{overall:7,commission_rating:7,niche_demand:8,conversion_potential:7,tier1_suitability:7,notes:'Solid MMO offer.'}, value_gaps:['No traffic strategy','No email templates','No checklist','No content prompts'], recommended_angle:'Pain & Agitation', angle_reason:'Pain hooks work well for burned-out MMO beginners.' }; }
    res.json({ success: true, analysis: parsed });
  } catch (err) { res.status(500).json({ success: false, message: 'Analysis failed: ' + err.message }); }
});

// ── GENERATE LEAD MAGNET ──────────────────────────────────────
app.post('/api/generate/lead-magnet', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, project_id } = req.body;
  if (!analysis || !angle) return res.status(400).json({ success: false, message: 'Analysis and angle required' });
  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514', max_tokens: 6000, system: SOLO_ADS_KB,
      messages: [{ role: 'user', content: `Create a high-value free lead magnet for cold solo ad traffic in the ${analysis.niche} niche.

Offer: ${analysis.product_name}
Target audience: ${analysis.target_audience}
Main pain point: ${analysis.main_pain_point}
Main benefit: ${analysis.main_benefit}
Audience psychology: ${analysis.audience_psychology}
Angle: ${angle}

LEAD MAGNET REQUIREMENTS:
- 1,500-2,000 words of genuinely useful, actionable content
- Professional document format with clear sections
- Use headers (## for sections, ### for subsections), numbered steps, bullet points
- Give real value — specific steps, real examples, things they can do TODAY
- Do NOT pitch the paid offer — this stands alone as a valuable free resource
- Title must be compelling, benefit-driven, and specific (include numbers where natural)
- Write short paragraphs — 2-3 sentences max per paragraph
- Every section must give them something concrete to implement

Return ONLY valid JSON:
{
  "title": "Compelling benefit-driven title with specifics",
  "subtitle": "Supporting subtitle that expands the promise",
  "cover_prompt": "3D rendered hardcover book product mockup on dark reflective surface. Navy blue cover, gold embossed title, gold decorative icon center, visible thick pages on right side. Studio lighting, soft shadow, bokeh background. Product photography style.",
  "full_content": "Complete 1500-2000 word lead magnet with proper markdown formatting — real headers, numbered steps, bullet points, examples. Write every word. No placeholders."
}` }]
    });

    let lm;
    try { lm = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch { return res.status(500).json({ success: false, message: 'Failed to parse lead magnet' }); }

    lm.cover_image = await generateImage(lm.cover_prompt);

    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({ project_id, user_id: user.id, content_type: 'lead_magnet', content: JSON.stringify(lm), updated_at: new Date().toISOString() }, { onConflict: 'project_id,content_type' });
    }
    res.json({ success: true, lead_magnet: lm });
  } catch (err) { res.status(500).json({ success: false, message: 'Lead magnet failed: ' + err.message }); }
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
      model: 'claude-sonnet-4-20250514', max_tokens: 9000, system: SOLO_ADS_KB,
      messages: [{ role: 'user', content: `Create a 3-part buyer-only bonus stack for this affiliate offer. Angle: ${angle}.

Product: ${analysis.product_name}
Niche: ${analysis.niche}
Main benefit: ${analysis.main_benefit}
Main pain point: ${analysis.main_pain_point}
Value gaps: ${(analysis.value_gaps||[]).join(', ')}
Audience: ${analysis.target_audience}
Psychology: ${analysis.audience_psychology}

BONUS CONTENT RULES — CRITICAL:
- Each bonus 400-600 words of REAL, usable content. Not outlines. Actual deliverable content.
- Each bonus solves ONE specific gap the main offer does not cover
- Bonus 1 (1-Page Checklist): Step-by-step numbered action checklist format. 15-25 specific action items organized by phase or section. Each item is a concrete action, not vague advice.
- Bonus 2 (2-Page Guide): Educational guide format. Teach a specific strategy or skill. Use headers, sub-sections, examples, and a clear how-to structure.
- Bonus 3 (AI Prompt Pack): 8-10 complete AI prompts. Each prompt has a title, the full prompt text (copy-paste ready), and 1-2 sentence instructions on how to use it.
- Cover prompts must describe DIFFERENT visual styles for each bonus so they look like a set but are distinct

Return ONLY valid JSON array — no text before or after:
[
  {
    "number": 1,
    "title": "Specific benefit-driven title",
    "type": "1-Page Checklist",
    "tagline": "One sentence that sells the bonus",
    "description": "Two sentences: what it is and what specific problem it solves",
    "full_content": "400-600 words of real checklist content. Numbered action steps organized by phase. Every step is specific and actionable.",
    "cover_prompt": "3D rendered hardcover book mockup floating on dark surface with soft reflection. Deep navy blue cover, large gold checkmark icon in center, white title text, gold border frame, visible page thickness on right. Dramatic side lighting. Product photography."
  },
  {
    "number": 2,
    "title": "Specific benefit-driven title",
    "type": "2-Page Guide",
    "tagline": "One sentence that sells the bonus",
    "description": "Two sentences description",
    "full_content": "400-600 words of real guide content with headers, explanation, examples, steps.",
    "cover_prompt": "3D rendered softcover guide mockup at slight angle on dark reflective desk. Royal blue cover, white compass or arrow icon center, white title text, thin white border, pages visible at bottom edge. Overhead studio lighting with soft glow. Product photography."
  },
  {
    "number": 3,
    "title": "Specific benefit-driven title",
    "type": "AI Prompt Pack",
    "tagline": "One sentence that sells the bonus",
    "description": "Two sentences description",
    "full_content": "8-10 complete prompts. Each has: PROMPT TITLE in caps, full copy-paste prompt text, brief usage note.",
    "cover_prompt": "3D rendered hardcover book mockup on dark glossy surface with sharp reflection. Dark purple cover with subtle gold circuit pattern texture, glowing gold lightning bolt icon center, white futuristic title text. Purple and gold studio lighting. Product photography."
  }
]` }]
    });

    let bonuses;
    const raw = message.content[0].text;
    try { bonuses = JSON.parse(raw.replace(/```json|```/g, '').trim()); }
    catch {
      const match = raw.match(/\[[\s\S]*\]/);
      if (match) { try { bonuses = JSON.parse(match[0]); } catch { bonuses = null; } }
    }
    if (!bonuses) return res.status(500).json({ success: false, message: 'Failed to parse bonuses' });

    // Generate individual bonus cover images
    for (let i = 0; i < bonuses.length; i++) {
      bonuses[i].cover_image = await generateImage(bonuses[i].cover_prompt);
    }

    // Generate bonus stack summary copy
    const summaryMsg = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514', max_tokens: 800, system: SOLO_ADS_KB,
      messages: [{ role: 'user', content: `Write a compelling 250-300 word bonus stack pitch for a bridge page. This sells all 3 bonuses as a package.

Product: ${analysis.product_name}
Angle: ${angle}
Bonuses:
${bonuses.map(b => `${b.number}. ${b.title} — ${b.description}`).join('\n')}

Structure:
- Open with the exact problem buyers still face after getting the main offer
- One sentence per bonus explaining the specific benefit
- Close with a 3-line summary of what they now have (what to do first / where to get traffic / what to say — or equivalent for this niche)
- Final line creates urgency

Return only the plain copy text. No JSON.` }]
    });
    const stackSummary = summaryMsg.content[0].text.trim();

    // Generate bonus stack image — all 3 together with descriptions
    const stackSummaryText = bonuses.map(b => b.title + ': ' + (b.tagline || b.description || '')).join(' | ');
    const stackPrompt = `Three 3D book product mockups arranged side by side on dark navy surface with reflections. Left book: navy/gold with checkmark. Center book: royal blue/white with arrow icon, slightly taller. Right book: purple/gold with lightning bolt. Gold sparkle particles floating around books. Below each book, short white label text. Gold banner above reading EXCLUSIVE BONUS PACKAGE. Studio lighting, dramatic shadows. Marketing display image.`;
    const stackImage = await generateImage(stackPrompt);

    const result = { bonuses, stack_summary: stackSummary, stack_image: stackImage };
    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({ project_id, user_id: user.id, content_type: 'bonus_stack', content: JSON.stringify(result), updated_at: new Date().toISOString() }, { onConflict: 'project_id,content_type' });
    }
    res.json({ success: true, ...result });
  } catch (err) { res.status(500).json({ success: false, message: 'Bonus generation failed: ' + err.message }); }
});

// ── GENERATE OPT-IN PAGE ──────────────────────────────────────
app.post('/api/generate/optin', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, lead_magnet, project_id } = req.body;
  if (!analysis || !angle) return res.status(400).json({ success: false, message: 'Analysis and angle required' });
  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  const lmContext = lead_magnet ? `Lead magnet: "${lead_magnet.title}" — ${lead_magnet.subtitle || ''}` : 'No lead magnet — use main offer promise as incentive.';

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514', max_tokens: 1500, system: SOLO_ADS_KB,
      messages: [{ role: 'user', content: `Write a high-converting opt-in squeeze page for cold solo ad traffic. Angle: ${angle}.

Offer: ${analysis.product_name}
Niche: ${analysis.niche}
Main pain point: ${analysis.main_pain_point}
Secondary pains: ${(analysis.secondary_pain_points||[]).join(', ')}
Main benefit: ${analysis.main_benefit}
Audience: ${analysis.target_audience}
Psychology: ${analysis.audience_psychology}
${lmContext}

Rules: No bonuses mentioned. Everything above the fold. Two A/B headlines.

Return ONLY valid JSON:
{
  "headline_a": "Version A — curiosity/pain driven, punchy",
  "headline_b": "Version B — different emotional trigger, benefit focused",
  "subheadline": "One sentence expanding the promise — specific and believable",
  "bullets": ["Specific benefit 1", "Specific benefit 2", "Specific benefit 3", "Specific benefit 4", "Specific benefit 5"],
  "cta_button": "Action-oriented CTA — urgency or curiosity",
  "microcopy": "Short objection-removing line under button",
  "form_fields": ["First Name", "Email Address"],
  "above_fold_note": "Layout guidance",
  "split_test_tip": "Which to run first and why"
}` }]
    });

    let content;
    try { content = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch { return res.status(500).json({ success: false, message: 'Failed to parse opt-in' }); }

    // Build HTML preview
    const coverImg = lead_magnet && lead_magnet.cover_image
      ? `<img src="${lead_magnet.cover_image}" style="width:200px;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.3);">`
      : '';

    content.html_preview = `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${content.headline_a}</title>
<style>*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Arial',sans-serif;background:linear-gradient(135deg,#0F0F23 0%,#1A1A3E 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.page{max-width:600px;width:100%;text-align:center}
.eyebrow{display:inline-block;background:rgba(255,215,0,0.15);border:1px solid rgba(255,215,0,0.4);color:#FFD700;font-size:12px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;padding:6px 16px;border-radius:20px;margin-bottom:24px}
h1{font-size:clamp(26px,5vw,44px);font-weight:900;color:#FFFFFF;line-height:1.1;letter-spacing:-0.02em;margin-bottom:16px}
.sub{font-size:18px;color:#BBBBDD;line-height:1.5;margin-bottom:28px}
.cover-wrap{margin-bottom:28px}
.bullets{text-align:left;background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:12px;padding:20px 24px;margin-bottom:28px}
.bullet-item{display:flex;align-items:flex-start;gap:10px;margin-bottom:10px;color:#DDDDEE;font-size:15px;line-height:1.4}
.bullet-item:last-child{margin-bottom:0}
.bullet-check{color:#4ADE80;font-size:16px;flex-shrink:0;margin-top:1px}
.form-wrap{margin-bottom:16px}
input[type=text],input[type=email]{width:100%;padding:14px 18px;border-radius:8px;border:1px solid rgba(255,255,255,0.2);background:rgba(255,255,255,0.08);color:#fff;font-size:16px;margin-bottom:10px;outline:none}
input::placeholder{color:#8888AA}
.cta-btn{width:100%;padding:18px;background:linear-gradient(135deg,#FF6B35,#FF4500);color:#fff;border:none;border-radius:10px;font-size:18px;font-weight:900;cursor:pointer;text-transform:uppercase;letter-spacing:0.04em;box-shadow:0 4px 20px rgba(255,107,53,0.4)}
.micro{font-size:12px;color:#8888AA;margin-top:10px}
.ab-toggle{background:rgba(79,70,184,0.3);border:1px solid rgba(79,70,184,0.5);border-radius:8px;padding:12px 16px;margin-bottom:20px;font-size:12px;color:#AAAACC;display:flex;align-items:center;gap:8px}
</style></head>
<body><div class="page">
<div class="ab-toggle">🧪 <strong>Version A</strong> — click to preview Version B: <a href="#" onclick="toggleH();return false;" style="color:#7B7EEE;text-decoration:none;font-weight:700">Switch headline</a></div>
<div class="eyebrow">Free Instant Access</div>
<h1 id="headline">${escHtml(content.headline_a)}</h1>
<p class="sub">${escHtml(content.subheadline)}</p>
${coverImg ? `<div class="cover-wrap">${coverImg}</div>` : ''}
<div class="bullets">${(content.bullets||[]).map(b=>`<div class="bullet-item"><span class="bullet-check">✓</span><span>${escHtml(b)}</span></div>`).join('')}</div>
<div class="form-wrap">
<input type="text" placeholder="Your First Name">
<input type="email" placeholder="Your Email Address">
<button class="cta-btn">${escHtml(content.cta_button)}</button>
<p class="micro">${escHtml(content.microcopy)}</p>
</div>
</div>
<script>var h=['${content.headline_a.replace(/'/g,"\\'")}','${content.headline_b.replace(/'/g,"\\'")}'];var i=0;function toggleH(){i=i===0?1:0;document.getElementById('headline').textContent=h[i];}</script>
</body></html>`;

    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({ project_id, user_id: user.id, content_type: 'optin_page', content: JSON.stringify(content), updated_at: new Date().toISOString() }, { onConflict: 'project_id,content_type' });
    }
    res.json({ success: true, content });
  } catch (err) { res.status(500).json({ success: false, message: 'Opt-in failed: ' + err.message }); }
});

function escHtml(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// ── GENERATE BRIDGE PAGE ──────────────────────────────────────
app.post('/api/generate/bridge', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, format, bonuses, stack_summary, lead_magnet, project_id } = req.body;
  if (!analysis || !angle || !format) return res.status(400).json({ success: false, message: 'Missing required fields' });
  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  const bonusContext = bonuses && bonuses.length > 0
    ? `Buyer bonuses:\n${bonuses.map(b=>`- ${b.title}: ${b.description}`).join('\n')}\nStack summary: ${stack_summary||''}`
    : 'No bonuses defined.';
  const lmContext = lead_magnet ? `Visitor just opted in for: "${lead_magnet.title}". Reference this.` : '';
  const formatInstructions = {
    text: 'Text-only: headline, relatable story/problem (2-3 paragraphs), 4-5 benefit bullets, bonus stack pitch with each bonus named and described, strong CTA. 600+ words.',
    vsl: 'VSL script (3-4 min spoken) + full supporting text copy. Script: hook, relatable problem, solution reveal, each bonus named, urgency close, CTA. 800+ words total.',
    short: 'Punchy 5-7 sentence intro creating massive curiosity + strong CTA. Extremely tight but emotionally powerful.'
  };

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514', max_tokens: 3500, system: SOLO_ADS_KB,
      messages: [{ role: 'user', content: `Write a high-converting bridge page. Angle: ${angle}. Format: ${format}.

Offer: ${analysis.product_name} (${analysis.price||'low ticket'})
Main benefit: ${analysis.main_benefit}
Pain point: ${analysis.main_pain_point}
Mechanism: ${analysis.unique_mechanism||''}
Psychology: ${analysis.audience_psychology}
${lmContext}
${bonusContext}

Format instructions: ${formatInstructions[format]||formatInstructions.text}

Write in first person. Acknowledge skepticism. Present offer as logical solution. Pitch bonuses as reward for buying through your link specifically.

Return ONLY valid JSON:
{
  "headline": "Strong headline",
  "subheadline": "Supporting subheadline",
  "content": "Full bridge copy — every word written, no placeholders",
  "cta_text": "CTA button text",
  "bonus_claim_note": "How to claim bonuses after purchase",
  "above_fold_note": "What appears above fold"
}` }]
    });

    let content;
    try { content = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch { return res.status(500).json({ success: false, message: 'Failed to parse bridge' }); }

    // Build HTML preview
    const stackImgTag = (bonuses && bonuses.length > 0 && bonuses[0].cover_image)
      ? `<div class="bonus-covers">${bonuses.map(b=>b.cover_image?`<img src="${b.cover_image}" style="width:100px;border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,0.4);">`:'').join('')}</div>`
      : '';

    content.html_preview = `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escHtml(content.headline)}</title>
<style>*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#F0F0F8;color:#1A1A2E;padding:0}
.hero{background:linear-gradient(135deg,#0F0F23,#1A1A3E);padding:60px 20px;text-align:center}
.hero h1{font-size:clamp(22px,4vw,38px);font-weight:900;color:#fff;line-height:1.15;max-width:720px;margin:0 auto 16px;letter-spacing:-0.02em}
.hero .sub{font-size:18px;color:#BBBBDD;max-width:560px;margin:0 auto}
.container{max-width:680px;margin:0 auto;padding:40px 20px}
.body-copy{font-size:16px;line-height:1.8;color:#2A2A3E;white-space:pre-wrap;margin-bottom:32px}
.bonus-section{background:#fff;border-radius:16px;padding:28px;box-shadow:0 4px 20px rgba(0,0,0,0.08);margin-bottom:32px;text-align:center}
.bonus-section h2{font-size:22px;font-weight:800;color:#1A1A2E;margin-bottom:20px}
.bonus-covers{display:flex;justify-content:center;gap:12px;flex-wrap:wrap;margin-bottom:20px}
.bonus-list{text-align:left}
.bonus-item{display:flex;gap:12px;padding:12px 0;border-bottom:1px solid #F0F0F8}
.bonus-item:last-child{border-bottom:none}
.b-num{width:28px;height:28px;border-radius:50%;background:#4F46B8;color:#fff;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;flex-shrink:0}
.b-title{font-weight:700;font-size:14px;color:#1A1A2E}
.b-desc{font-size:13px;color:#666;margin-top:2px}
.cta-wrap{text-align:center;margin:32px 0}
.cta-btn{display:inline-block;background:linear-gradient(135deg,#FF6B35,#FF4500);color:#fff;font-size:18px;font-weight:900;padding:18px 40px;border-radius:10px;text-decoration:none;box-shadow:0 4px 20px rgba(255,107,53,0.4);cursor:pointer;border:none;text-transform:uppercase;letter-spacing:0.04em}
.claim-note{font-size:13px;color:#888;margin-top:12px;text-align:center}
</style></head>
<body>
<div class="hero"><h1>${escHtml(content.headline)}</h1><p class="sub">${escHtml(content.subheadline||'')}</p></div>
<div class="container">
<div class="body-copy">${escHtml(content.content)}</div>
${bonuses && bonuses.length > 0 ? `<div class="bonus-section">
<h2>🎁 Your Exclusive Buyer Bonuses</h2>
${stackImgTag}
<div class="bonus-list">${bonuses.map(b=>`<div class="bonus-item"><div class="b-num">${b.number}</div><div><div class="b-title">${escHtml(b.title)}</div><div class="b-desc">${escHtml(b.description)}</div></div></div>`).join('')}</div>
</div>` : ''}
<div class="cta-wrap"><button class="cta-btn">${escHtml(content.cta_text)}</button><p class="claim-note">${escHtml(content.bonus_claim_note||'')}</p></div>
</div></body></html>`;

    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({ project_id, user_id: user.id, content_type: 'bridge_page', content: JSON.stringify(content), updated_at: new Date().toISOString() }, { onConflict: 'project_id,content_type' });
    }
    res.json({ success: true, content });
  } catch (err) { res.status(500).json({ success: false, message: 'Bridge failed: ' + err.message }); }
});

// ── GENERATE EMAILS ───────────────────────────────────────────
app.post('/api/generate/emails', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, sequence_type, bonuses, lead_magnet, project_id } = req.body;
  if (!analysis || !angle || !sequence_type) return res.status(400).json({ success: false, message: 'Missing required fields' });
  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  const bonusContext = bonuses && bonuses.length > 0 ? `Buyer bonuses:\n${bonuses.map(b=>`- ${b.title}: ${b.description}`).join('\n')}` : 'No bonuses.';
  const lmContext = lead_magnet ? `Lead magnet delivered: "${lead_magnet.title}"` : 'No lead magnet.';

  const sequences = {
    non_buyers: `5-Day Non-Buyers Sequence — Goal: Convert to Sale
Day 1: Thank you + deliver lead magnet + soft introduction to offer. Warm, helpful tone. Light link.
Day 2: Story email — share a relatable struggle this audience faces. Transition to how the offer solves it. Include link.
Day 3: Value email — teach one useful thing related to the niche. Build trust. Soft mention of offer at end.
Day 4: Hard pitch — name every bonus specifically. Make the stack the reason to buy TODAY through your link. Clear CTA.
Day 5: Final push — urgency/scarcity. Simple action beats endless research. Last call for bonuses. Strong close.`,
    buyers: `7-Day Buyers Sequence — Goal: Onboard, Build Trust, Ascend
Day 1: Welcome + congrats + deliver all 3 bonuses with access instructions. Tell them exactly what to do first.
Day 2: Quick win tip — teach one specific action that gets a fast result. Practical, no upsell.
Day 3: Mindset/success patterns — what separates people who succeed with this from those who don't. Pure value.
Day 4: Advanced strategy — one deeper tactic they can implement now. Positions you as expert.
Day 5: Social proof + community — stories of others succeeding, invite them to share their progress.
Day 6: Introduce a related solution — frame it as the logical next step after implementing the main offer.
Day 7: Benefits/demo of related solution — make it feel like the obvious continuation of their journey.`
  };

  try {
    const dayCount = sequence_type === 'non_buyers' ? 5 : 7;
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514', max_tokens: 6000, system: SOLO_ADS_KB,
      messages: [{ role: 'user', content: `Write a complete ${sequence_type === 'non_buyers' ? 'Non-Buyers 5-Day' : 'Buyers 7-Day'} email sequence. Angle: ${angle}.

Offer: ${analysis.product_name}
Niche: ${analysis.niche}
Pain: ${analysis.main_pain_point}
Benefit: ${analysis.main_benefit}
Psychology: ${analysis.audience_psychology}
${lmContext}
${bonusContext}

${sequences[sequence_type]}

CRITICAL EMAIL RULES:
- Write EVERY email IN FULL — no placeholders, no [write story here], actual copy
- Use [FirstName], [YourAffiliateLink], [YourName]
- Short paragraphs — 1-3 sentences max
- Generous line breaks — easy to skim on mobile
- 3 subject lines per email: one curiosity, one direct benefit, one relatable question
- Non-buyer emails: every email ends with PS reinforcing main reason to act
- Write emails people actually want to read — not corporate, not hype

Return ONLY a valid JSON array of ${dayCount} emails:
[{"day":1,"subject_lines":["option1","option2","option3"],"body":"complete email body"}]` }]
    });

    let emails;
    try { emails = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch {
      const match = message.content[0].text.match(/\[[\s\S]*\]/);
      if (match) { try { emails = JSON.parse(match[0]); } catch { emails = null; } }
    }
    if (!emails) return res.status(500).json({ success: false, message: 'Failed to parse emails' });

    const contentType = sequence_type === 'non_buyers' ? 'emails_non_buyers' : 'emails_buyers';
    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({ project_id, user_id: user.id, content_type: contentType, content: JSON.stringify(emails), updated_at: new Date().toISOString() }, { onConflict: 'project_id,content_type' });
    }
    res.json({ success: true, emails });
  } catch (err) { res.status(500).json({ success: false, message: 'Email generation failed: ' + err.message }); }
});

// ── GENERATE BLUEPRINT ────────────────────────────────────────
app.post('/api/generate/blueprint', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, bonuses, lead_magnet, project_id } = req.body;
  if (!analysis || !angle) return res.status(400).json({ success: false, message: 'Missing required fields' });
  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514', max_tokens: 4000, system: SOLO_ADS_KB,
      messages: [{ role: 'user', content: `Create a complete copy-paste funnel blueprint.

Offer: ${analysis.product_name} | Angle: ${angle}
Lead magnet: ${lead_magnet ? lead_magnet.title : 'None'}
Bonuses: ${bonuses ? bonuses.map(b=>b.title).join(', ') : 'None'}

Write a practical implementation guide covering:
1. Funnel flow overview
2. Opt-in page setup (structure, design notes, what goes above fold)
3. Bridge page setup (layout, video placement, bonus section)
4. Lead magnet delivery
5. Email automation setup (triggers, timing, segmentation)
6. Bonus delivery and claim process
7. Asset placement checklist
8. Tracking setup
9. Build order (step by step)
10. Pre-launch test checklist

Be specific. Write like someone who has built 100 of these funnels.

Return ONLY valid JSON:
{"title":"Blueprint title","sections":[{"heading":"Section heading","content":"Full detailed content"}]}` }]
    });

    let blueprint;
    try { blueprint = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch { return res.status(500).json({ success: false, message: 'Failed to parse blueprint' }); }

    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({ project_id, user_id: user.id, content_type: 'blueprint', content: JSON.stringify(blueprint), updated_at: new Date().toISOString() }, { onConflict: 'project_id,content_type' });
    }
    res.json({ success: true, blueprint });
  } catch (err) { res.status(500).json({ success: false, message: 'Blueprint failed: ' + err.message }); }
});

// ── DOWNLOAD: SINGLE BONUS as DOCX ───────────────────────────
app.post('/api/download/bonus', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { bonus, product_name } = req.body;
  if (!bonus) return res.status(400).json({ success: false, message: 'Bonus required' });
  try {
    const children = parseContentToDocx(bonus.full_content, product_name || 'this offer', bonus.title, bonus.tagline || bonus.type, `Exclusive Buyer Bonus #${bonus.number}`);
    const filename = `Bonus-${bonus.number}-${(bonus.title||'bonus').replace(/[^a-z0-9]/gi,'_').slice(0,40)}.docx`;
    await buildAndSendDocx(res, children, filename);
  } catch (err) { res.status(500).json({ success: false, message: 'Download failed: ' + err.message }); }
});

// ── DOWNLOAD: ALL BONUSES — 3 separate DOCX in zip, or individual ──
app.post('/api/download/bonuses-all', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { bonuses, product_name } = req.body;
  if (!bonuses || !bonuses.length) return res.status(400).json({ success: false, message: 'Bonuses required' });
  // Return a combined HTML index with download links for each — simpler than zip
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Bonus Stack — ${escHtml(product_name||'Your Offer')}</title>
<style>body{font-family:Arial,sans-serif;max-width:600px;margin:60px auto;padding:20px;color:#1A1A2E}
h1{font-size:24px;margin-bottom:8px}p{color:#555;margin-bottom:32px}
.item{border:1px solid #EEEDFE;border-radius:10px;padding:20px;margin-bottom:16px;display:flex;align-items:center;gap:16px}
.num{width:36px;height:36px;border-radius:50%;background:#4F46B8;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;flex-shrink:0}
.info{flex:1}.title{font-weight:700;font-size:15px}.desc{font-size:13px;color:#666;margin-top:3px}
.tag{font-size:11px;background:#EEEDFE;color:#3C3489;padding:2px 8px;border-radius:4px;font-weight:700}</style></head>
<body>
<h1>Your Bonus Stack</h1>
<p>Your 3 exclusive buyer bonuses for ${escHtml(product_name||'this offer')}. Each is downloaded as a separate document.</p>
${bonuses.map(b=>`<div class="item"><div class="num">${b.number}</div><div class="info"><div class="title">${escHtml(b.title)}</div><div class="desc">${escHtml(b.description)}</div><div style="margin-top:6px"><span class="tag">${escHtml(b.type)}</span></div></div></div>`).join('')}
<p style="margin-top:24px;font-size:13px;color:#888">To download individual bonuses as Word documents, use the Download button next to each bonus.</p>
</body></html>`;
  res.setHeader('Content-Type', 'text/html');
  res.setHeader('Content-Disposition', `attachment; filename="Bonus-Stack-Index.html"`);
  res.send(html);
});

// ── DOWNLOAD: LEAD MAGNET as DOCX ─────────────────────────────
app.post('/api/download/lead-magnet', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { lead_magnet, product_name } = req.body;
  if (!lead_magnet) return res.status(400).json({ success: false, message: 'Lead magnet required' });
  try {
    const children = parseContentToDocx(lead_magnet.full_content, product_name || 'JGAffiliate', lead_magnet.title, lead_magnet.subtitle || '', 'Free Guide');
    const filename = `Lead-Magnet-${(lead_magnet.title||'guide').replace(/[^a-z0-9]/gi,'_').slice(0,40)}.docx`;
    await buildAndSendDocx(res, children, filename);
  } catch (err) { res.status(500).json({ success: false, message: 'Download failed: ' + err.message }); }
});

// ── DOWNLOAD: OPT-IN PAGE HTML ────────────────────────────────
app.post('/api/download/optin-html', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { html_content, product_name } = req.body;
  if (!html_content) return res.status(400).json({ success: false, message: 'HTML required' });
  res.setHeader('Content-Type', 'text/html');
  res.setHeader('Content-Disposition', `attachment; filename="Opt-In-Page-${(product_name||'offer').replace(/[^a-z0-9]/gi,'_').slice(0,30)}.html"`);
  res.send(html_content);
});

// ── DOWNLOAD: BRIDGE PAGE HTML ────────────────────────────────
app.post('/api/download/bridge-html', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { html_content, product_name } = req.body;
  if (!html_content) return res.status(400).json({ success: false, message: 'HTML required' });
  res.setHeader('Content-Type', 'text/html');
  res.setHeader('Content-Disposition', `attachment; filename="Bridge-Page-${(product_name||'offer').replace(/[^a-z0-9]/gi,'_').slice(0,30)}.html"`);
  res.send(html_content);
});

// ── DOWNLOAD: BLUEPRINT ───────────────────────────────────────
app.post('/api/download/blueprint', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { blueprint, product_name } = req.body;
  if (!blueprint) return res.status(400).json({ success: false, message: 'Blueprint required' });
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${escHtml(blueprint.title||'Blueprint')}</title>
<style>body{font-family:Arial,sans-serif;max-width:720px;margin:0 auto;padding:48px 32px;color:#1A1A2E}
.cover{text-align:center;padding:60px 0 40px;border-bottom:3px solid #4F46B8;margin-bottom:48px}
.eyebrow{font-size:11px;color:#4F46B8;text-transform:uppercase;letter-spacing:0.08em;font-weight:700;margin-bottom:12px}
h1{font-size:30px;font-weight:700;margin:0 0 12px}
.section{margin-bottom:40px;padding-bottom:32px;border-bottom:1px solid #EEEDFE}
h2{font-size:20px;font-weight:700;color:#4F46B8;margin:0 0 16px}
.content{font-size:14px;line-height:1.9;white-space:pre-wrap}
.footer{margin-top:48px;padding-top:20px;border-top:1px solid #EEEDFE;font-size:11px;color:#aaa;text-align:center}</style></head>
<body><div class="cover"><div class="eyebrow">Funnel Implementation Guide</div><h1>${escHtml(blueprint.title||'Blueprint')}</h1><p style="color:#555">${escHtml(product_name||'')}</p></div>
${(blueprint.sections||[]).map(s=>`<div class="section"><h2>${escHtml(s.heading)}</h2><div class="content">${escHtml(s.content)}</div></div>`).join('')}
<div class="footer">PromoLab by Jimmy Griffith, JGAffiliate</div></body></html>`;
  res.setHeader('Content-Type', 'text/html');
  res.setHeader('Content-Disposition', `attachment; filename="Blueprint-${(product_name||'offer').replace(/[^a-z0-9]/gi,'_').slice(0,30)}.html"`);
  res.send(html);
});

// ── PROJECT MANAGEMENT ────────────────────────────────────────
app.post('/api/projects', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { name, url, analysis, angle, channel } = req.body;
  const { data, error } = await adminClient.from('promolab_projects').insert({ user_id: user.id, name: name || analysis.product_name, url: url || '', channel: channel || 'solo_ads', angle, analysis: JSON.stringify(analysis), created_at: new Date().toISOString(), updated_at: new Date().toISOString() }).select().single();
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
  (content || []).forEach(c => { try { contentMap[c.content_type] = JSON.parse(c.content); } catch { contentMap[c.content_type] = c.content; } });
  res.json({ success: true, project, content: contentMap });
});
app.delete('/api/projects/:id', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  await adminClient.from('promolab_project_content').delete().eq('project_id', req.params.id);
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
  const { error } = await adminClient.from('promolab_access').upsert({ user_id: found.id, solo_ads: solo_ads||false, facebook: facebook||false, email_sequence: email_sequence||false, launchjacking: launchjacking||false, affiliate_launch_guide: affiliate_launch_guide||false, updated_at: new Date().toISOString() }, { onConflict: 'user_id' });
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

app.listen(PORT, () => console.log('PromoLab v4 running on port ' + PORT));
