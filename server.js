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
const CLAUDE_QUEUE_INTERVAL_MS = Number(process.env.CLAUDE_QUEUE_INTERVAL_MS || 22000);
const CLAUDE_RATE_LIMIT_WAIT_MS = Number(process.env.CLAUDE_RATE_LIMIT_WAIT_MS || 95000);
const DEFAULT_MONTHLY_GENERATION_LIMIT = Number(process.env.DEFAULT_MONTHLY_GENERATION_LIMIT || 100);
let claudeQueue = Promise.resolve();
let lastClaudeStart = 0;

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
    is_admin: false,
    plan_name: 'solo_ads_basic',
    monthly_generation_limit: null
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

function monthStartIso() {
  const start = new Date();
  start.setUTCDate(1);
  start.setUTCHours(0, 0, 0, 0);
  return start.toISOString();
}

async function getMonthlyUsageSummary(userId) {
  const { data, error } = await adminClient
    .from('promolab_usage')
    .select('status, estimated_input_tokens, estimated_output_tokens, action')
    .eq('user_id', userId)
    .gte('created_at', monthStartIso());
  if (error) throw error;
  const rows = data || [];
  return rows.reduce((acc, row) => {
    acc.total_events += 1;
    acc.success += row.status === 'success' ? 1 : 0;
    acc.failed += row.status === 'failed' ? 1 : 0;
    acc.estimated_input_tokens += row.estimated_input_tokens || 0;
    acc.estimated_output_tokens += row.estimated_output_tokens || 0;
    acc.by_action[row.action] = (acc.by_action[row.action] || 0) + 1;
    return acc;
  }, { total_events: 0, success: 0, failed: 0, estimated_input_tokens: 0, estimated_output_tokens: 0, by_action: {} });
}

async function requireUsageAvailable(user, res) {
  const access = await getUserAccess(user.id);
  const userLimit = access.monthly_generation_limit === null || access.monthly_generation_limit === undefined
    ? DEFAULT_MONTHLY_GENERATION_LIMIT
    : Number(access.monthly_generation_limit);
  const limit = access.is_admin ? Infinity : userLimit;
  if (!Number.isFinite(limit)) return true;
  const summary = await getMonthlyUsageSummary(user.id);
  if (summary.success >= limit) {
    res.status(429).json({
      success: false,
      message: `Monthly generation limit reached (${summary.success}/${limit}). Please upgrade or wait until next month.`,
      usage_limit_reached: true,
      limit,
      used: summary.success
    });
    return false;
  }
  return true;
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

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isRateLimitError(err) {
  const text = String((err && (err.message || err.error && err.error.message)) || '');
  return err && (err.status === 429 || text.includes('rate_limit') || text.toLowerCase().includes('rate limit'));
}

async function runClaudeQueued(payload) {
  const run = claudeQueue.then(async () => {
    const elapsed = Date.now() - lastClaudeStart;
    if (elapsed < CLAUDE_QUEUE_INTERVAL_MS) await sleep(CLAUDE_QUEUE_INTERVAL_MS - elapsed);
    lastClaudeStart = Date.now();

    for (let attempt = 1; attempt <= 4; attempt++) {
      try {
        return await anthropic.messages.create(payload);
      } catch (err) {
        if (!isRateLimitError(err) || attempt === 4) throw err;
        console.log(`Claude rate limit hit. Waiting before retry ${attempt + 1}/4.`);
        await sleep(CLAUDE_RATE_LIMIT_WAIT_MS);
      }
    }
  });

  claudeQueue = run.catch(() => {});
  return run;
}

async function askJson(prompt, maxTokens = 5000, repairInstruction = 'Return valid JSON only.') {
  const msg = await runClaudeQueued({
    model: MODEL,
    max_tokens: maxTokens,
    system: SOLO_ADS_KB,
    messages: [{ role: 'user', content: prompt }]
  });
  const text = msg.content[0].text;
  const parsed = jsonFromText(text);
  if (parsed) return parsed;

  const repair = await runClaudeQueued({
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

function estimateTokensFromChars(chars) {
  return Math.ceil(Number(chars || 0) / 4);
}

async function logUsage({ userId, projectId, module = 'solo_ads', action, provider = 'anthropic', model = MODEL, status = 'success', inputText = '', outputText = '', error }) {
  try {
    const inputChars = String(inputText || '').length;
    const outputChars = String(outputText || '').length;
    await adminClient.from('promolab_usage').insert({
      user_id: userId,
      project_id: projectId || null,
      module,
      action,
      provider,
      model,
      status,
      input_chars: inputChars,
      output_chars: outputChars,
      estimated_input_tokens: estimateTokensFromChars(inputChars),
      estimated_output_tokens: estimateTokensFromChars(outputChars),
      error_message: error ? String(error).slice(0, 1000) : null,
      created_at: new Date().toISOString()
    });
  } catch (err) {
    console.log('Usage logging skipped:', err.message);
  }
}

async function trackedAskJson({ userId, projectId, action, prompt, maxTokens, repairInstruction }) {
  try {
    const result = await askJson(prompt, maxTokens, repairInstruction);
    await logUsage({
      userId,
      projectId,
      action,
      inputText: prompt,
      outputText: JSON.stringify(result),
      status: 'success'
    });
    return result;
  } catch (err) {
    await logUsage({
      userId,
      projectId,
      action,
      inputText: prompt,
      status: 'failed',
      error: err.message || err
    });
    throw err;
  }
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
  const titleLines = wrapText(title, 20, 4)
    .map((line, i) => `<text x="540" y="${390 + i * 54}" text-anchor="middle" font-family="Arial" font-size="44" font-weight="900" fill="${p[3]}">${esc(line)}</text>`)
    .join('');
  const subLines = wrapText(subtitle || 'Exclusive resource', 34, 3)
    .map((line, i) => `<text x="540" y="${650 + i * 31}" text-anchor="middle" font-family="Arial" font-size="23" font-weight="700" fill="${p[2]}">${esc(line)}</text>`)
    .join('');
  const num = number ? `<circle cx="340" cy="255" r="38" fill="${p[2]}"/><text x="340" y="268" text-anchor="middle" font-family="Arial" font-size="32" font-weight="900" fill="${p[0]}">${number}</text>` : '';
  return svgUri(`<svg xmlns="http://www.w3.org/2000/svg" width="1080" height="1080" viewBox="0 0 1080 1080">
<defs><radialGradient id="glow" cx="50%" cy="38%" r="58%"><stop offset="0" stop-color="${p[1]}" stop-opacity=".55"/><stop offset="1" stop-color="#050816" stop-opacity="1"/></radialGradient><linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#050816"/><stop offset="1" stop-color="#111827"/></linearGradient><linearGradient id="c" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="${p[0]}"/><stop offset="1" stop-color="${p[1]}"/></linearGradient><filter id="s" x="-30%" y="-30%" width="160%" height="160%"><feDropShadow dx="0" dy="30" stdDeviation="28" flood-color="#000" flood-opacity=".58"/></filter></defs>
<rect width="1080" height="1080" fill="url(#bg)"/><rect width="1080" height="1080" fill="url(#glow)" opacity=".75"/><circle cx="180" cy="170" r="3" fill="${p[2]}" opacity=".8"/><circle cx="890" cy="220" r="4" fill="${p[2]}" opacity=".7"/><circle cx="850" cy="760" r="3" fill="#fff" opacity=".45"/><ellipse cx="540" cy="875" rx="345" ry="62" fill="#000" opacity=".44"/>
<g filter="url(#s)"><path d="M326 150 L752 118 Q810 116 830 172 L830 792 Q808 850 750 862 L326 898 Q275 894 266 836 L266 228 Q275 170 326 150Z" fill="url(#c)"/><path d="M752 118 Q810 116 830 172 L830 792 Q808 850 750 862 L750 118Z" fill="#fff" opacity=".18"/><path d="M763 176 L763 812" stroke="#fff" stroke-width="5" opacity=".18"/><rect x="318" y="212" width="445" height="620" rx="26" fill="none" stroke="${p[2]}" stroke-width="7"/><rect x="345" y="238" width="390" height="86" rx="18" fill="#000" opacity=".2"/><text x="540" y="292" text-anchor="middle" font-family="Arial" font-size="21" font-weight="900" letter-spacing="4" fill="${p[2]}">${esc(badge || 'EXCLUSIVE BONUS')}</text>${num}<circle cx="540" cy="340" r="31" fill="${p[2]}"/><path d="M522 340 L536 354 L562 324" fill="none" stroke="${p[0]}" stroke-width="9" stroke-linecap="round" stroke-linejoin="round"/>${titleLines}<rect x="365" y="605" width="350" height="108" rx="18" fill="#000" opacity=".18"/>${subLines}<text x="540" y="775" text-anchor="middle" font-family="Arial" font-size="19" font-weight="700" fill="#fff" opacity=".62">Exclusive Buyer Resource</text></g></svg>`);
}

function stackImage(bonuses) {
  const items = (bonuses || []).slice(0, 3);
  const cards = items.map((b, i) => {
    const colors = [['#1e3a8a', '#f8c75a'], ['#2563eb', '#dbeafe'], ['#6d28d9', '#f0abfc']][i];
    const x = [130, 390, 650][i];
    const lines = wrapText(b.title, 14, 3).map((line, n) => `<text x="${x + 105}" y="${322 + n * 28}" text-anchor="middle" font-family="Arial" font-size="23" font-weight="900" fill="#fff">${esc(line)}</text>`).join('');
    const desc = wrapText(b.description || b.tagline || '', 32, 3).map((line, n) => `<text x="${x + 105}" y="${642 + n * 22}" text-anchor="middle" font-family="Arial" font-size="17" font-weight="700" fill="#dbeafe">${esc(line)}</text>`).join('');
    return `<g><g transform="rotate(${[-6, 0, 6][i]} ${x + 105} 430)"><rect x="${x}" y="215" width="210" height="310" rx="18" fill="${colors[0]}" filter="url(#s)"/><rect x="${x + 18}" y="240" width="174" height="258" rx="10" fill="none" stroke="${colors[1]}" stroke-width="5"/><text x="${x + 105}" y="282" text-anchor="middle" font-family="Arial" font-size="15" font-weight="900" fill="${colors[1]}">BONUS ${i + 1}</text>${lines}<text x="${x + 105}" y="480" text-anchor="middle" font-family="Arial" font-size="14" font-weight="700" fill="${colors[1]}">${esc(b.type || '')}</text></g><rect x="${x - 20}" y="585" width="250" height="110" rx="16" fill="#0f172a" stroke="${colors[1]}" stroke-opacity=".55"/><text x="${x + 105}" y="622" text-anchor="middle" font-family="Arial" font-size="18" font-weight="900" fill="${colors[1]}">${esc(b.type || `Bonus ${i + 1}`)}</text>${desc}</g>`;
  }).join('');
  return svgUri(`<svg xmlns="http://www.w3.org/2000/svg" width="1400" height="900" viewBox="0 0 1400 900"><defs><radialGradient id="g" cx="50%" cy="32%" r="70%"><stop offset="0" stop-color="#1e3a8a" stop-opacity=".62"/><stop offset="1" stop-color="#07111f"/></radialGradient><filter id="s" x="-30%" y="-30%" width="160%" height="160%"><feDropShadow dx="0" dy="24" stdDeviation="24" flood-color="#000" flood-opacity=".55"/></filter></defs><rect width="1400" height="900" fill="url(#g)"/><circle cx="230" cy="170" r="4" fill="#f8c75a" opacity=".8"/><circle cx="1120" cy="220" r="5" fill="#dbeafe" opacity=".6"/><circle cx="1030" cy="740" r="3" fill="#f8c75a" opacity=".75"/><ellipse cx="650" cy="560" rx="520" ry="70" fill="#000" opacity=".35"/><text x="650" y="105" text-anchor="middle" font-family="Arial" font-size="50" font-weight="900" fill="#f8c75a">EXCLUSIVE BONUS PACKAGE</text><text x="650" y="150" text-anchor="middle" font-family="Arial" font-size="22" font-weight="700" fill="#dbeafe">Three buyer-only resources that make the main offer easier to use</text><g transform="translate(120 0)">${cards}</g></svg>`);
}

function compactBonusForPrompt(bonus, index) {
  return {
    number: bonus.number || index + 1,
    title: bonus.title || '',
    type: bonus.type || '',
    tagline: bonus.tagline || '',
    description: bonus.description || ''
  };
}

function compactAnalysisForPrompt(analysis) {
  return {
    product_name: analysis.product_name || '',
    niche: analysis.niche || '',
    main_promise: analysis.main_promise || '',
    main_pain_point: analysis.main_pain_point || '',
    target_audience: analysis.target_audience || '',
    audience_psychology: analysis.audience_psychology || '',
    value_gaps: analysis.value_gaps || []
  };
}

function compactLeadMagnetForPrompt(leadMagnet) {
  if (!leadMagnet) return null;
  return {
    title: leadMagnet.title || '',
    subtitle: leadMagnet.subtitle || '',
    description: leadMagnet.description || '',
    bonus_summary_for_prompt: leadMagnet.bonus_summary_for_prompt || leadMagnet.description || leadMagnet.subtitle || leadMagnet.title || ''
  };
}

function compactStackForPrompt(stack) {
  if (!stack) return null;
  return {
    headline: stack.headline || '',
    bullets: stack.bullets || [],
    stack_summary_for_prompt: stack.stack_summary_for_prompt || stack.summary || ''
  };
}

function summarizeForPrompt(text, maxChars = 700) {
  return String(text || '').replace(/\s+/g, ' ').trim().slice(0, maxChars);
}

async function generateAiImage(prompt) {
  if (!process.env.OPENAI_API_KEY) return null;
  try {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 65000);
    const response = await fetch('https://api.openai.com/v1/images/generations', {
      method: 'POST',
      signal: ctrl.signal,
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + process.env.OPENAI_API_KEY
      },
      body: JSON.stringify({
        model: process.env.OPENAI_IMAGE_MODEL || 'gpt-image-1',
        prompt,
        size: '1024x1024',
        quality: 'high',
        n: 1
      })
    });
    clearTimeout(timer);
    if (!response.ok) {
      console.log('OpenAI image failed:', response.status, await response.text());
      return null;
    }
    const data = await response.json();
    const b64 = data && data.data && data.data[0] && data.data[0].b64_json;
    return b64 ? 'data:image/png;base64,' + b64 : null;
  } catch (err) {
    console.log('OpenAI image error:', err.message);
    return null;
  }
}

async function createCoverImage(asset) {
  const fallback = coverImage(asset);
  const prompt = `Create a premium 3D digital product mockup for an affiliate marketing bonus.
Use a polished book/guide cover on a dark studio background with dramatic lighting, realistic shadows, glossy reflections, and a high-value SaaS/course bonus feel.
Make the cover clean and readable. Use this exact main title as the central cover text: "${asset.title}".
Use this small label near the top: "${asset.badge || 'EXCLUSIVE BONUS'}".
Use this supporting line if it fits cleanly: "${asset.subtitle || ''}".
Avoid clutter. Avoid fake author names. Avoid tiny unreadable paragraphs. No people, no faces, no watermarks.`;
  return await generateAiImage(prompt) || fallback;
}

async function createStackImage(bonuses) {
  const fallback = stackImage(bonuses);
  const bonusText = (bonuses || []).map((b, i) => `Bonus ${i + 1}: ${b.title} - ${b.description || b.tagline || b.type}`).join('\n');
  const prompt = `Create a premium promotional bundle image for an affiliate bridge page.
Scene: three 3D book/guide mockups arranged as a high-value bonus stack on a dark navy studio background with gold highlights, realistic shadows, reflections, and polished marketing design.
Top headline text: "EXCLUSIVE BONUS PACKAGE".
Each book should look distinct but cohesive.
Under each book, include a short readable description panel.
Use these bonus titles and descriptions:
${bonusText}
Avoid people, faces, watermarks, distorted gibberish, and clutter. Make it look like a premium digital product bundle used to sell affiliate bonuses.`;
  return await generateAiImage(prompt) || fallback;
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
  if (!(await requireUsageAvailable(user, res))) return;
  const { url, pasted_text, affiliate_link, audience_note, tone_note } = req.body;
  try {
    const pageText = await fetchSalesPage(url, pasted_text);
    const prompt = `Analyze this affiliate offer for a solo ads funnel.

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
}`;
    const analysis = await trackedAskJson({ userId: user.id, projectId: null, action: 'analyze_offer', prompt, maxTokens: 3500 });

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
  if (!(await requireUsageAvailable(user, res))) return;
  const { project_id, analysis, angle } = req.body;
  try {
    const compactAnalysis = compactAnalysisForPrompt(analysis || {});
    const prompt = `Create the buyer-only bonus stack plan.

Offer analysis:
${JSON.stringify(compactAnalysis)}

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
}`;
    const plan = await trackedAskJson({ userId: user.id, projectId: project_id, action: 'bonus_plan', prompt, maxTokens: 3000 });
    await saveContent(user.id, project_id, 'bonus_plan', plan);
    res.json({ success: true, plan });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/bonus', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  if (!(await requireUsageAvailable(user, res))) return;
  const { project_id, analysis, angle, bonus } = req.body;
  try {
    const compactAnalysis = compactAnalysisForPrompt(analysis || {});
    const prompt = `Generate the complete text for this buyer-only bonus.

Offer analysis:
${JSON.stringify(compactAnalysis)}

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
}`;
    const full = await trackedAskJson({ userId: user.id, projectId: project_id, action: `bonus_${Number(bonus.number || 1)}`, prompt, maxTokens: 7000 });
    full.cover_image = await createCoverImage({
      title: full.title,
      subtitle: full.tagline || full.type,
      badge: full.type || `BONUS ${full.number}`,
      number: full.number,
      theme: ['gold', 'blue', 'purple'][Number(full.number || 1) - 1] || 'blue'
    });
    full.bonus_summary_for_prompt = summarizeForPrompt(`${full.title}. ${full.type}. ${full.tagline}. ${full.description}`);
    await saveContent(user.id, project_id, `bonus_${full.number}`, full);
    res.json({ success: true, bonus: full });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/bonus-stack', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  if (!(await requireUsageAvailable(user, res))) return;
  const { project_id, analysis, angle, bonuses } = req.body;
  try {
    const compactBonuses = (bonuses || []).map(compactBonusForPrompt);
    const compactAnalysis = compactAnalysisForPrompt(analysis || {});
    const prompt = `Write the approved buyer-only bonus stack summary for the bridge page.

Offer analysis:
${JSON.stringify(compactAnalysis)}

Angle: ${angle}

Bonuses:
${JSON.stringify(compactBonuses)}

Rules:
- Do not summarize the full bonus documents.
- Sell the 3 bonuses as a buyer-only package.
- Make the stack feel like the missing implementation pieces that make the main offer easier to use.
- Keep the copy direct, benefit-driven, and suitable for bridge page placement.

Return JSON:
{
  "headline": "",
  "summary": "About 300 words. Sell the stack as the reason to buy through this affiliate link today.",
  "bullets": ["", "", ""]
}`;
    const summary = await trackedAskJson({ userId: user.id, projectId: project_id, action: 'bonus_stack', prompt, maxTokens: 1800 });
    summary.stack_summary_for_prompt = summarizeForPrompt(`${summary.headline}. ${summary.summary} ${(summary.bullets || []).join(' ')}`);
    summary.stack_image = await createStackImage(compactBonuses);
    await saveContent(user.id, project_id, 'bonus_stack', summary);
    res.json({ success: true, stack: summary });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/lead-magnet', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  if (!(await requireUsageAvailable(user, res))) return;
  const { project_id, analysis, angle } = req.body;
  try {
    const compactAnalysis = compactAnalysisForPrompt(analysis || {});
    const prompt = `Create a high-value lead magnet for the opt-in page.

Offer analysis:
${JSON.stringify(compactAnalysis)}

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
}`;
    const lm = await trackedAskJson({ userId: user.id, projectId: project_id, action: 'lead_magnet', prompt, maxTokens: 7000 });
    lm.cover_image = await createCoverImage({ title: lm.title, subtitle: lm.subtitle, badge: 'FREE GUIDE', theme: 'green' });
    lm.lead_magnet_summary_for_prompt = summarizeForPrompt(`${lm.title}. ${lm.subtitle}. ${lm.description}`);
    await saveContent(user.id, project_id, 'lead_magnet', lm);
    res.json({ success: true, lead_magnet: lm });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/optin', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  if (!(await requireUsageAvailable(user, res))) return;
  const { project_id, analysis, angle, lead_magnet } = req.body;
  try {
    const compactAnalysis = compactAnalysisForPrompt(analysis || {});
    const compactLeadMagnet = compactLeadMagnetForPrompt(lead_magnet);
    const prompt = `Generate opt-in squeeze page copy for cold solo ad traffic.

Offer analysis:
${JSON.stringify(compactAnalysis)}

Chosen angle: ${angle}
Lead magnet:
${compactLeadMagnet ? JSON.stringify(compactLeadMagnet) : 'No lead magnet selected.'}

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
}`;
    const optin = await trackedAskJson({ userId: user.id, projectId: project_id, action: 'optin_page', prompt, maxTokens: 1600 });
    await saveContent(user.id, project_id, 'optin_page', optin);
    res.json({ success: true, optin });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/bridge', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  if (!(await requireUsageAvailable(user, res))) return;
  const { project_id, analysis, angle, bridge_format, bonuses, bonus_stack, lead_magnet } = req.body;
  try {
    const compactAnalysis = compactAnalysisForPrompt(analysis || {});
    const compactBonuses = (bonuses || []).map(compactBonusForPrompt);
    const compactLeadMagnet = compactLeadMagnetForPrompt(lead_magnet);
    const compactStack = compactStackForPrompt(bonus_stack);
    const prompt = `Generate the bridge page.

Offer analysis:
${JSON.stringify(compactAnalysis)}

Chosen angle: ${angle}
Bridge format: ${bridge_format || 'text'}
Lead magnet:
${compactLeadMagnet ? JSON.stringify(compactLeadMagnet) : 'none'}
Bonuses:
${JSON.stringify(compactBonuses)}
Bonus stack:
${compactStack ? JSON.stringify(compactStack) : 'none'}

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
}`;
    const bridge = await trackedAskJson({ userId: user.id, projectId: project_id, action: 'bridge_page', prompt, maxTokens: 4200 });
    await saveContent(user.id, project_id, 'bridge_page', bridge);
    res.json({ success: true, bridge });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/solo/emails', async (req, res) => {
  const user = await requireSoloAds(req, res);
  if (!user) return;
  if (!(await requireUsageAvailable(user, res))) return;
  const { project_id, analysis, angle, bonuses, lead_magnet } = req.body;
  try {
    const compactAnalysis = compactAnalysisForPrompt(analysis || {});
    const compactBonuses = (bonuses || []).map(compactBonusForPrompt);
    const compactLeadMagnet = compactLeadMagnetForPrompt(lead_magnet);
    const prompt = `Generate both segmented email sequences.

Offer analysis:
${JSON.stringify(compactAnalysis)}

Chosen angle: ${angle}
Lead magnet:
${compactLeadMagnet ? JSON.stringify(compactLeadMagnet) : 'none'}
Bonuses:
${JSON.stringify(compactBonuses)}

Rules:
- Non-buyers sequence has 5 days: thank you/delivery, story-solution, value lesson, offer/bonuses, final push.
- Buyers sequence has 7 days: welcome/bonus delivery, quick win, advanced tip, success mindset, related solution intro, related solution benefits, final next-step invitation.
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
}`;
    const emails = await trackedAskJson({ userId: user.id, projectId: project_id, action: 'emails', prompt, maxTokens: 7600 });
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

app.get('/api/usage/me', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const access = await getUserAccess(user.id);
  const start = new Date();
  start.setUTCDate(1);
  start.setUTCHours(0, 0, 0, 0);
  const { data, error } = await adminClient
    .from('promolab_usage')
    .select('*')
    .eq('user_id', user.id)
    .gte('created_at', start.toISOString())
    .order('created_at', { ascending: false });
  if (error) return res.status(400).json({ success: false, message: error.message });
  const rows = data || [];
  const summary = rows.reduce((acc, row) => {
    acc.total_events += 1;
    acc.success += row.status === 'success' ? 1 : 0;
    acc.failed += row.status === 'failed' ? 1 : 0;
    acc.estimated_input_tokens += row.estimated_input_tokens || 0;
    acc.estimated_output_tokens += row.estimated_output_tokens || 0;
    acc.by_action[row.action] = (acc.by_action[row.action] || 0) + 1;
    return acc;
  }, { total_events: 0, success: 0, failed: 0, estimated_input_tokens: 0, estimated_output_tokens: 0, by_action: {} });
  const accessLimit = access.monthly_generation_limit === null || access.monthly_generation_limit === undefined
    ? DEFAULT_MONTHLY_GENERATION_LIMIT
    : Number(access.monthly_generation_limit);
  const limit = access.is_admin ? null : accessLimit;
  summary.monthly_limit = limit;
  summary.remaining = limit === null ? null : Math.max(0, limit - summary.success);
  summary.plan_name = access.plan_name || (access.is_admin ? 'admin' : 'solo_ads_basic');
  res.json({ success: true, summary, recent: rows.slice(0, 25) });
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

app.get('/api/admin/usage', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const access = await getUserAccess(user.id);
  if (!access.is_admin) return res.status(403).json({ success: false, message: 'Admin only' });
  const start = new Date();
  start.setUTCDate(1);
  start.setUTCHours(0, 0, 0, 0);
  const { data, error } = await adminClient
    .from('promolab_usage')
    .select('*')
    .gte('created_at', start.toISOString())
    .order('created_at', { ascending: false })
    .limit(500);
  if (error) return res.status(400).json({ success: false, message: error.message });
  res.json({ success: true, usage: data || [] });
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
