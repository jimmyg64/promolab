require('dotenv').config();

const express = require('express');
const path = require('path');
const cookieParser = require('cookie-parser');
const { createClient } = require('@supabase/supabase-js');
const Anthropic = require('@anthropic-ai/sdk');
const fetch = require('node-fetch');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle
} = require('docx');

const app = express();
const PORT = process.env.PORT || 3000;

const adminClient = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SECRET_KEY);
const publicClient = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_PUBLISHABLE_KEY);
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

const MODEL = process.env.ANTHROPIC_MODEL || 'claude-sonnet-4-20250514';

const SOLO_ADS_KB = `
You are the Solo Ads IQ Funnel Architect inside PromoLab.

Core principles:
- Every output must drive a specific action.
- Write for cold MMO solo ad traffic: skeptical, impatient, burned before, looking for a clear path.
- Use short paragraphs, direct language, concrete steps, and no filler.
- Do not make vague promises. Be useful, specific, and practical.
- Buyer-only bonuses are conversion tools. They are never opt-in incentives.
- Lead magnets are opt-in incentives. They must not pitch the paid offer too aggressively.
- Opt-in copy must be ultra concise, curiosity-driven, and above the fold.
- Bridge copy must transition from opt-in to sales page and pitch the buyer-only bonus stack.
- Emails must be segmented: non-buyers convert to sale, buyers onboard and ascend.

Angles:
1. Consultative/Pathfinder: for overwhelmed audiences. Position the offer as the logical clear path.
2. Pain & Agitation: for audiences stuck using broken methods. Focus on frustration and relief.
3. Pure Value & Bonus: for competitive offers. Stack value and gap-filling bonuses.

Formatting:
- Return only the requested JSON.
- Do not include markdown fences.
- Do not use placeholders like [write story here].
- Write finished copy, not outlines, unless the requested output is a plan.
`;

function cookieOptions() {
  return {
    httpOnly: true,
    secure: true,
    sameSite: 'lax',
    maxAge: 60 * 60 * 24 * 7 * 1000,
    path: '/'
  };
}

async function getCurrentUser(req) {
  const token = req.cookies && req.cookies.sb_access_token;
  if (!token) return null;
  try {
    const { data, error } = await adminClient.auth.getUser(token);
    if (error || !data || !data.user) return null;
    return data.user;
  } catch {
    return null;
  }
}

async function requireUser(req, res) {
  const user = await getCurrentUser(req);
  if (!user) {
    res.status(401).json({ success: false, message: 'Not logged in' });
    return null;
  }
  if (!user.email_confirmed_at) {
    res.status(403).json({ success: false, message: 'Please confirm your email' });
    return null;
  }
  return user;
}

async function getUserAccess(userId) {
  const { data } = await adminClient.from('promolab_access').select('*').eq('user_id', userId).single();
  return data || {
    solo_ads: false,
    facebook: false,
    email_sequence: false,
    launchjacking: false,
    affiliate_launch_guide: false,
    is_admin: false
  };
}

async function requireSoloAds(req, res) {
  const user = await requireUser(req, res);
  if (!user) return null;
  const access = await getUserAccess(user.id);
  if (!access.solo_ads && !access.is_admin) {
    res.status(403).json({ success: false, message: 'Solo Ads module is not unlocked for this account.' });
    return null;
  }
  return user;
}

function jsonFromText(text) {
  const raw = String(text || '').replace(/```json|```/g, '').trim();
  try {
    return JSON.parse(raw);
  } catch {}
  const firstArray = raw.indexOf('[');
  const firstObject = raw.indexOf('{');
  const isArray = firstArray !== -1 && (firstObject === -1 || firstArray < firstObject);
  const start = isArray ? firstArray : firstObject;
  const end = isArray ? raw.lastIndexOf(']') : raw.lastIndexOf('}');
  if (start >= 0 && end > start) {
    try {
      return JSON.parse(raw.slice(start, end + 1));
    } catch {}
  }
  return null;
}

async function askJson(prompt, maxTokens = 5000, repairInstruction = 'Return valid JSON only.') {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: maxTokens,
    system: SOLO_ADS_KB,
    messages: [{ role: 'user', content: prompt }]
  });
  const text = msg.content[0].text;
  const parsed = jsonFromText(text);
  if (parsed) return parsed;

  const repair = await anthropic.messages.create({
    model: MODEL,
    max_tokens: maxTokens,
    system: 'You repair malformed AI output into strict valid JSON. Return JSON only.',
    messages: [{
      role: 'user',
      content: `${repairInstruction}\n\nBroken output:\n${text}`
    }]
  });
  const fixed = jsonFromText(repair.content[0].text);
  if (!fixed) throw new Error('AI returned invalid JSON and repair failed.');
  return fixed;
}

async function fetchSalesPage(url, pastedText) {
  if (pastedText && pastedText.trim()) return pastedText.trim().slice(0, 14000);
  if (!url) return '';
  try {
    const r = await fetch(url, {
      timeout: 12000,
      headers: { 'User-Agent': 'Mozilla/5.0 PromoLab Offer Analyzer' }
    });
    const html = await r.text();
    return html
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ')
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 14000);
  } catch {
    return `Could not fetch the page. Analyze based on this URL and user context: ${url}`;
  }
}

function esc(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function wrapText(value, maxChars, maxLines) {
  const words = String(value || '').replace(/\s+/g, ' ').trim().split(' ');
  const lines = [];
  let line = '';
  for (const word of words) {
    const next = line ? `${line} ${word}` : word;
    if (next.length > maxChars && line) {
      lines.push(line);
      line = word;
    } else {
      line = next;
    }
    if (lines.length === maxLines) break;
  }
  if (line && lines.length < maxLines) lines.push(line);
  return lines.length ? lines : ['PromoLab'];
}

function svgUri(svg) {
  return 'data:image/svg+xml;base64,' + Buffer.from(svg, 'utf8').toString('base64');
}

function coverImage({ title, subtitle, badge, theme = 'blue', number }) {
  const palettes = {
    blue: ['#0f172a', '#1d4ed8', '#bfdbfe', '#ffffff'],
    gold: ['#111827', '#1e3a8a', '#f8c75a', '#fff7d6'],
    purple: ['#1e1235', '#6d28d9', '#f0abfc', '#ffffff'],
    green: ['#10231b', '#047857', '#bbf7d0', '#ffffff']
  };
  const p = palettes[theme] || palettes.blue;
  const titleLines = wrapText(title, 18, 4)
    .map((line, i) => `<text x="512" y="${360 + i * 58}" text-anchor="middle" font-family="Arial" font-size="48" font-weight="900" fill="${p[3]}">${esc(line)}</text>`)
    .join('');
  const subLines = wrapText(subtitle || 'Exclusive resource', 30, 2)
    .map((line, i) => `<text x="512" y="${630 + i * 32}" text-anchor="middle" font-family="Arial" font-size="24" font-weight="700" fill="${p[2]}">${esc(line)}</text>`)
    .join('');
  const num = number ? `<circle cx="306" cy="262" r="40" fill="${p[2]}"/><text x="306" y="276" text-anchor="middle" font-family="Arial" font-size="34" font-weight="900" fill="${p[0]}">${number}</text>` : '';
  return svgUri(`<svg xmlns="http://www.w3.org/2000/svg" width="1080" height="1080" viewBox="0 0 1080 1080">
<defs><linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#050816"/><stop offset="1" stop-color="#111827"/></linearGradient><linearGradient id="c" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="${p[0]}"/><stop offset="1" stop-color="${p[1]}"/></linearGradient><filter id="s" x="-30%" y="-30%" width="160%" height="160%"><feDropShadow dx="0" dy="28" stdDeviation="28" flood-color="#000" flood-opacity=".55"/></filter></defs>
<rect width="1080" height="1080" fill="url(#bg)"/><ellipse cx="540" cy="865" rx="330" ry="60" fill="#000" opacity=".42"/>
<g filter="url(#s)" transform="translate(78 12)"><path d="M275 160 L748 116 Q804 114 825 170 L825 790 Q805 846 748 858 L275 900 Q222 894 214 835 L214 238 Q222 180 275 160Z" fill="url(#c)"/><path d="M748 116 Q804 114 825 170 L825 790 Q805 846 748 858 L748 116Z" fill="#fff" opacity=".17"/><rect x="260" y="218" width="455" height="620" rx="24" fill="none" stroke="${p[2]}" stroke-width="7"/><text x="512" y="270" text-anchor="middle" font-family="Arial" font-size="22" font-weight="900" letter-spacing="4" fill="${p[2]}">${esc(badge || 'EXCLUSIVE BONUS')}</text>${num}<circle cx="512" cy="315" r="34" fill="${p[2]}"/><path d="M492 315 L508 331 L538 295" fill="none" stroke="${p[0]}" stroke-width="10" stroke-linecap="round" stroke-linejoin="round"/>${titleLines}${subLines}<text x="512" y="770" text-anchor="middle" font-family="Arial" font-size="22" font-weight="700" fill="#fff" opacity=".7">PromoLab</text></g></svg>`);
}

function stackImage(bonuses) {
  const items = (bonuses || []).slice(0, 3);
  const cards = items.map((b, i) => {
    const colors = [['#1e3a8a', '#f8c75a'], ['#2563eb', '#dbeafe'], ['#6d28d9', '#f0abfc']][i];
    const x = [165, 420, 675][i];
    const lines = wrapText(b.title, 14, 3).map((line, n) => `<text x="${x + 105}" y="${350 + n * 30}" text-anchor="middle" font-family="Arial" font-size="25" font-weight="900" fill="#fff">${esc(line)}</text>`).join('');
    return `<g transform="rotate(${[-7, 0, 7][i]} ${x + 105} 500)"><rect x="${x}" y="245" width="210" height="345" rx="18" fill="${colors[0]}" filter="url(#s)"/><rect x="${x + 18}" y="272" width="174" height="288" rx="10" fill="none" stroke="${colors[1]}" stroke-width="5"/><text x="${x + 105}" y="312" text-anchor="middle" font-family="Arial" font-size="15" font-weight="900" fill="${colors[1]}">BONUS ${i + 1}</text>${lines}<text x="${x + 105}" y="534" text-anchor="middle" font-family="Arial" font-size="14" font-weight="700" fill="${colors[1]}">${esc(b.type)}</text></g>`;
  }).join('');
  return svgUri(`<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="800" viewBox="0 0 1200 800"><defs><filter id="s" x="-30%" y="-30%" width="160%" height="160%"><feDropShadow dx="0" dy="24" stdDeviation="24" flood-color="#000" flood-opacity=".55"/></filter></defs><rect width="1200" height="800" fill="#07111f"/><ellipse cx="600" cy="635" rx="460" ry="70" fill="#000" opacity=".42"/><text x="600" y="130" text-anchor="middle" font-family="Arial" font-size="50" font-weight="900" fill="#f8c75a">EXCLUSIVE BONUS PACKAGE</text><text x="600" y="178" text-anchor="middle" font-family="Arial" font-size="22" font-weight="700" fill="#dbeafe">Three buyer-only resources included with your campaign</text>${cards}</svg>`);
}

function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: opts.before || 80, after: opts.after || 100 },
    children: [new TextRun({
      text: String(text || ''),
      bold: !!opts.bold,
      italic: !!opts.italic,
      size: opts.size || 24,
      color: opts.color || '1A1A2E',
      font: 'Calibri'
    })]
  });
}

function h(text, level = 1) {
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
    spacing: { before: level === 1 ? 320 : 220, after: 120 },
    children: [new TextRun({
      text: String(text || ''),
      bold: true,
      size: level === 1 ? 36 : level === 2 ? 30 : 26,
      color: level === 1 ? '1A1A2E' : '4F46B8',
      font: 'Calibri'
    })]
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 220, after: 220 },
    border: { bottom: { color: '4F46B8', space: 1, style: BorderStyle.SINGLE, size: 6 } },
    children: []
  });
}

function contentToDoc(title, subtitle, type, content) {
  const children = [
    p(type.toUpperCase(), { center: true, bold: true, size: 20, color: '4F46B8', before: 700, after: 120 }),
    p(title, { center: true, bold: true, size: 44, before: 0, after: 180 }),
    p(subtitle || '', { center: true, italic: true, size: 26, color: '555555', before: 0, after: 160 }),
    divider(),
    new Paragraph({ pageBreakBefore: true, children: [] })
  ];

  String(content || '').split('\n').forEach((line) => {
    const t = line.trim();
    if (!t) {
      children.push(p('', { before: 30, after: 30 }));
    } else if (t.startsWith('### ')) {
      children.push(h(t.slice(4), 3));
    } else if (t.startsWith('## ')) {
      children.push(h(t.slice(3), 2));
    } else if (t.startsWith('# ')) {
      children.push(h(t.slice(2), 1));
    } else if (t.startsWith('- ') || t.startsWith('* ')) {
      children.push(new Paragraph({
        bullet: { level: 0 },
        spacing: { before: 50, after: 50 },
        children: [new TextRun({ text: t.slice(2), size: 24, font: 'Calibri', color: '333333' })]
      }));
    } else {
      children.push(p(t.replace(/\*\*/g, ''), { before: 70, after: 70 }));
    }
  });

  children.push(divider());
  children.push(p('PromoLab by Jimmy Griffith, JGAffiliate', { center: true, size: 18, color: '999999' }));
  return children;
}

async function sendDocx(res, filename, title, subtitle, type, content) {
  const doc = new Document({
    styles: { default: { document: { run: { font: 'Calibri', size: 24, color: '1A1A2E' } } } },
    sections: [{
      properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: contentToDoc(title, subtitle, type, content)
    }]
  });
  const buffer = await Packer.toBuffer(doc);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(buffer);
}

async function saveContent(userId, projectId, type, content) {
  if (!projectId) return;
  await adminClient.from('promolab_project_content').upsert({
    project_id: projectId,
    user_id: userId,
    content_type: type,
    content: JSON.stringify(content),
    updated_at: new Date().toISOString()
  }, { onConflict: 'project_id,content_type' });
}

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('/login.html', (req, res) => res.sendFile(path.join(__dirname, 'public', 'login.html')));
app.get('/signup.html', (req, res) => res.sendFile(path.join(__dirname, 'public', 'signup.html')));
app.get('/health', (req, res) => res.json({ status: 'ok', version: 'promolab-2' }));

app.get('/app.html', async (req, res) => {
  const user = await getCurrentUser(req);
  if (!user || !user.email_confirmed_at) return res.redirect('/login.html');
  res.sendFile(path.join(__dirname, 'public', 'app.html'));
});

app.get('/admin.html', async (req, res) => {
  const user = await getCurrentUser(req);
  if (!user) return res.redirect('/login.html');
  const access = await getUserAccess(user.id);
  if (!access.is_admin) return res.redirect('/app.html');
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

app.post('/signup', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, message: 'Email and password required' });
  if (password.length < 8) return res.status(400).json({ success: false, message: 'Password must be at least 8 characters' });
  const { data, error } = await publicClient.auth.signUp({ email: email.toLowerCase().trim(), password });
  if (error) return res.status(400).json({ success: false, message: error.message });
  res.json({ success: true, message: 'Account created. Check your email to confirm.', user: data.user });
});

app.post('/login', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ success: false, message: 'Email and password required' });
  const { data, error } = await publicClient.auth.signInWithPassword({ email: email.toLowerCase().trim(), password });
  if (error || !data.session) return res.status(401).json({ success: false, message: error ? error.message : 'Login failed' });
  res.cookie('sb_access_token', data.session.access_token, cookieOptions());
  res.cookie('sb_refresh_token', data.session.refresh_token, cookieOptions());
  res.json({ success: true, email: data.user.email });
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

app.post('/api/solo/analyze', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  const { url, pasted_text, affiliate_link, audience_note, tone_note } = req.body;
  try {
    const pageText = await fetchSalesPage(url, pasted_text);
    const analysis = await askJson(`Analyze this affiliate offer for a solo ads funnel.

Sales page URL: ${url || 'not provided'}
Affiliate link: ${affiliate_link || 'not provided'}
Audience note: ${audience_note || 'none'}
Tone note: ${tone_note || 'none'}

Sales page text:
${pageText}

Return JSON:
{
  "product_name": "",
  "niche": "",
  "price": "",
  "commission": "",
  "main_promise": "",
  "main_pain_point": "",
  "secondary_pain_points": ["", "", ""],
  "target_audience": "",
  "audience_psychology": "",
  "unique_mechanism": "",
  "value_gaps": ["", "", ""],
  "offer_summary": "",
  "recommended_angle": "Consultative/Pathfinder Angle or Pain & Agitation Angle or Pure Value & Bonus Angle",
  "angle_reason": "",
  "offer_score": {
    "overall": 0,
    "commission_rating": 0,
    "niche_demand": 0,
    "conversion_potential": 0,
    "tier1_suitability": 0,
    "notes": ""
  }
}`, 3500);

    const { data: project, error } = await adminClient.from('promolab_projects').insert({
      user_id: user.id,
      name: analysis.product_name || 'Solo Ads Project',
      url: url || '',
      channel: 'solo_ads',
      angle: analysis.recommended_angle || '',
      analysis: JSON.stringify(analysis),
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    }).select().single();
    if (error) return res.status(400).json({ success: false, message: error.message });
    res.json({ success: true, project, analysis });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/project/:id/angle', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  const { angle } = req.body;
  await adminClient.from('promolab_projects').update({ angle, updated_at: new Date().toISOString() }).eq('id', req.params.id).eq('user_id', user.id);
  res.json({ success: true });
});

app.post('/api/solo/bonus-plan', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  const { project_id, analysis, angle } = req.body;
  try {
    const plan = await askJson(`Create the buyer-only bonus stack plan.

Offer analysis:
${JSON.stringify(analysis)}

Chosen angle: ${angle}

Rules:
- Bonuses are strictly for buyers who buy through the affiliate link.
- Bonus 1 is a 1-page checklist, about 300-400 words when generated.
- Bonus 2 is a 2-page guide, about 600-800 words when generated.
- Bonus 3 is an AI prompt pack, 8-10 detailed prompts, about 600-800 words when generated.
- Each bonus must fill a specific value gap in the main offer.

Return JSON:
{
  "bonuses": [
    {
      "number": 1,
      "title": "",
      "type": "1-Page Checklist",
      "description": "",
      "value_gap_filled": "",
      "why_it_sells_the_offer": ""
    },
    {
      "number": 2,
      "title": "",
      "type": "2-Page Guide",
      "description": "",
      "value_gap_filled": "",
      "why_it_sells_the_offer": ""
    },
    {
      "number": 3,
      "title": "",
      "type": "AI Prompt Pack",
      "description": "",
      "value_gap_filled": "",
      "why_it_sells_the_offer": ""
    }
  ]
}`, 3000);
    await saveContent(user.id, project_id, 'bonus_plan', plan);
    res.json({ success: true, plan });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/bonus', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  const { project_id, analysis, angle, bonus } = req.body;
  try {
    const full = await askJson(`Generate the complete text for this buyer-only bonus.

Offer analysis:
${JSON.stringify(analysis)}

Chosen angle: ${angle}

Bonus to generate:
${JSON.stringify(bonus)}

Rules:
- Write finished, useful content.
- Use clear headers, steps, bullets, and examples.
- Do not write an outline.
- Do not include placeholders.
- Bonus 1 checklist: 15-25 specific action steps organized by phase.
- Bonus 2 guide: practical teaching, examples, how-to structure.
- Bonus 3 prompt pack: 8-10 complete copy-paste prompts with usage notes.

Return JSON:
{
  "number": ${Number(bonus.number || 1)},
  "title": "",
  "type": "",
  "tagline": "",
  "description": "",
  "full_content": ""
}`, 7000);
    full.cover_image = coverImage({
      title: full.title,
      subtitle: full.tagline || full.type,
      badge: full.type || `BONUS ${full.number}`,
      number: full.number,
      theme: ['gold', 'blue', 'purple'][Number(full.number || 1) - 1] || 'blue'
    });
    await saveContent(user.id, project_id, `bonus_${full.number}`, full);
    res.json({ success: true, bonus: full });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/bonus-stack', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  const { project_id, analysis, angle, bonuses } = req.body;
  try {
    const summary = await askJson(`Write the approved buyer-only bonus stack summary for the bridge page.

Offer: ${analysis.product_name}
Angle: ${angle}
Bonuses:
${JSON.stringify(bonuses)}

Return JSON:
{
  "headline": "",
  "summary": "About 300 words. Sell the stack as the reason to buy through this affiliate link today.",
  "bullets": ["", "", ""]
}`, 1800);
    summary.stack_image = stackImage(bonuses);
    await saveContent(user.id, project_id, 'bonus_stack', summary);
    res.json({ success: true, stack: summary });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/lead-magnet', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  const { project_id, analysis, angle } = req.body;
  try {
    const lm = await askJson(`Create a high-value lead magnet for the opt-in page.

Offer analysis:
${JSON.stringify(analysis)}

Chosen angle: ${angle}

Rules:
- This is the free opt-in incentive.
- Do not heavily pitch the paid offer.
- Write a complete practical resource, about 1200-1800 words.
- Use headers, bullets, numbered steps, examples, and a strong promise.

Return JSON:
{
  "title": "",
  "subtitle": "",
  "description": "",
  "full_content": ""
}`, 8000);
    lm.cover_image = coverImage({ title: lm.title, subtitle: lm.subtitle, badge: 'FREE GUIDE', theme: 'green' });
    await saveContent(user.id, project_id, 'lead_magnet', lm);
    res.json({ success: true, lead_magnet: lm });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/optin', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  const { project_id, analysis, angle, lead_magnet } = req.body;
  try {
    const optin = await askJson(`Generate opt-in squeeze page copy for cold solo ad traffic.

Offer analysis:
${JSON.stringify(analysis)}

Chosen angle: ${angle}
Lead magnet: ${lead_magnet ? JSON.stringify(lead_magnet) : 'No lead magnet selected.'}

Rules:
- Do not mention buyer-only bonuses.
- Must fit above the fold.
- Form collects first name and email address only.
- Provide two A/B headline versions.

Return JSON:
{
  "headline_a": "",
  "headline_b": "",
  "subheadline": "",
  "form_fields": ["First Name", "Email Address"],
  "cta_button": "",
  "below_fold_bullets": ["", "", ""],
  "trust_element": "",
  "privacy_statement": "",
  "design_notes": ["", "", ""]
}`, 2500);
    await saveContent(user.id, project_id, 'optin_page', optin);
    res.json({ success: true, optin });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/bridge', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  const { project_id, analysis, angle, bridge_format, bonuses, bonus_stack, lead_magnet } = req.body;
  try {
    const bridge = await askJson(`Generate the bridge page.

Offer analysis:
${JSON.stringify(analysis)}

Chosen angle: ${angle}
Bridge format: ${bridge_format || 'text'}
Lead magnet: ${lead_magnet ? lead_magnet.title : 'none'}
Bonuses:
${JSON.stringify(bonuses || [])}
Bonus stack:
${JSON.stringify(bonus_stack || {})}

Rules:
- Bridge from opt-in to sales page.
- Mention the strongest angle.
- Prominently pitch the buyer-only bonus stack.
- Include CTA and bonus claim instructions.
- If video format, include a 3-5 minute script plus supporting text.

Return JSON:
{
  "headline": "",
  "subheadline": "",
  "confirmation_section": "",
  "bridge_content": "",
  "video_script": "",
  "cta_text": "",
  "bonus_claim_note": "",
  "exit_intent_popup": {
    "trigger": "Mouse moves toward browser close",
    "offer": "",
    "link_note": ""
  }
}`, 5500);
    await saveContent(user.id, project_id, 'bridge_page', bridge);
    res.json({ success: true, bridge });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/emails', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  const { project_id, analysis, angle, bonuses, lead_magnet } = req.body;
  try {
    const emails = await askJson(`Generate both segmented email sequences.

Offer analysis:
${JSON.stringify(analysis)}

Chosen angle: ${angle}
Lead magnet: ${lead_magnet ? lead_magnet.title : 'none'}
Bonuses:
${JSON.stringify(bonuses || [])}

Rules:
- Non-buyers sequence has 4 days: thank you/delivery, story-solution, offer/bonuses, final push.
- Buyers sequence has 4 days: welcome/bonus delivery, advanced tip, related solution, demo/benefits.
- Every email needs 3 subject lines.
- Write full body copy.
- Use [FirstName], [YourAffiliateLink], and [YourName] where useful.

Return JSON:
{
  "non_buyers": [
    {"day": 1, "purpose": "", "subject_lines": ["", "", ""], "body": ""}
  ],
  "buyers": [
    {"day": 1, "purpose": "", "subject_lines": ["", "", ""], "body": ""}
  ]
}`, 8500);
    await saveContent(user.id, project_id, 'emails', emails);
    res.json({ success: true, emails });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/download/docx', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { filename, title, subtitle, type, content } = req.body;
  await sendDocx(
    res,
    (filename || 'PromoLab-Document.docx').replace(/[^a-z0-9_.-]/gi, '_'),
    title || 'PromoLab Document',
    subtitle || '',
    type || 'PromoLab',
    content || ''
  );
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
  const { data: rows } = await adminClient.from('promolab_project_content').select('*').eq('project_id', req.params.id).eq('user_id', user.id);
  const content = {};
  (rows || []).forEach((row) => {
    try { content[row.content_type] = JSON.parse(row.content); } catch { content[row.content_type] = row.content; }
  });
  res.json({ success: true, project, content });
});

app.delete('/api/projects/:id', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  await adminClient.from('promolab_project_content').delete().eq('project_id', req.params.id).eq('user_id', user.id);
  await adminClient.from('promolab_projects').delete().eq('id', req.params.id).eq('user_id', user.id);
  res.json({ success: true });
});

app.get('/api/admin/users', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const access = await getUserAccess(user.id);
  if (!access.is_admin) return res.status(403).json({ success: false, message: 'Admin only' });
  const { data: users } = await adminClient.auth.admin.listUsers();
  const { data: accessRows } = await adminClient.from('promolab_access').select('*');
  res.json({ success: true, users: users.users || [], access: accessRows || [] });
});

app.post('/api/admin/access', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const access = await getUserAccess(user.id);
  if (!access.is_admin) return res.status(403).json({ success: false, message: 'Admin only' });
  const { target_email, solo_ads, facebook, email_sequence, launchjacking, affiliate_launch_guide } = req.body;
  const { data: users } = await adminClient.auth.admin.listUsers();
  const found = (users.users || []).find((u) => u.email === target_email);
  if (!found) return res.status(404).json({ success: false, message: 'User not found' });
  const { error } = await adminClient.from('promolab_access').upsert({
    user_id: found.id,
    solo_ads: !!solo_ads,
    facebook: !!facebook,
    email_sequence: !!email_sequence,
    launchjacking: !!launchjacking,
    affiliate_launch_guide: !!affiliate_launch_guide,
    updated_at: new Date().toISOString()
  }, { onConflict: 'user_id' });
  if (error) return res.status(400).json({ success: false, message: error.message });
  res.json({ success: true });
});

app.listen(PORT, () => console.log(`PromoLab 2.0 running on port ${PORT}`));
