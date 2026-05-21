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
1. Consultative/Pathfinder: Best for overwhelmed audiences. Position offer as the most logical, clear path. Guide them to the click as the obvious next step.
2. Pain & Agitation: Best for audiences stuck using outdated/broken methods. Focus on frustration of current situation and present offer as instant relief.
3. Pure Value & Bonus: Best for highly competitive offers. Stack immense value of the main offer alongside custom gap-filling bonuses to make purchase a no-brainer.

BONUS STRATEGY RULES:
- Bonuses are PRIMARY conversion tools, not afterthoughts. They push leads over the edge to purchase.
- Bonuses are STRICTLY for buyers only — never mention on opt-in page.
- Fill the Value Gaps: Create bonuses that specifically fill missing pieces in the main offer.
- Each bonus must solve one specific problem the main offer does not fully address.
- Bonus content must be genuinely useful, not filler. Real steps. Real examples. Real value.

OPT-IN PAGE RULES:
- Goal: Capture contact details using curiosity and big promise. That is it.
- NO bonuses mentioned — buyers only.
- Ultra-concise above the fold: Headline + Sub-headline + CTA must fit one screen.
- Always generate TWO headline variations (Version A and Version B) for A/B split testing.
- Collect ONLY First Name and Email Address.
- If a lead magnet exists, it is the primary incentive on the opt-in page.

EMAIL SEQUENCE RULES (The Frequency Method):
- Always generate THREE distinct subject line options per email.
- Mix curiosity, direct benefit, and relatable question across the three options.
- Non-Buyers Sequence: Day 1 thank you + soft link, Day 2 story/solution, Day 3 offer + bonuses hard pitch, Day 4 authority + scarcity final push.
- Buyers Sequence: Day 1 welcome + bonus delivery, Day 2 advanced tip, Day 3 introduce upsell, Day 4 upsell benefits.
- Write emails in full. Every word matters. No placeholder sentences.
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

// ── Image generation via DALL-E 3 ─────────────────────────────
async function generateCoverImage(prompt) {
  if (!process.env.OPENAI_API_KEY) return null;
  try {
    const imgTimeout = new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 30000));
    const imgFetch = fetch('https://api.openai.com/v1/images/generations', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + process.env.OPENAI_API_KEY
      },
      body: JSON.stringify({
        model: 'dall-e-3',
        prompt: prompt,
        n: 1,
        size: '1024x1024',
        response_format: 'b64_json'
      })
    });
    const response = await Promise.race([imgFetch, imgTimeout]);
    if (!response.ok) {
      const err = await response.text();
      console.log('DALL-E error:', response.status, err);
      return null;
    }
    const data = await response.json();
    if (data.data && data.data[0] && data.data[0].b64_json) {
      return 'data:image/png;base64,' + data.data[0].b64_json;
    }
    return null;
  } catch(e) {
    console.log('Image generation failed:', e.message);
    return null;
  }
}

// ── Fallback: Stability AI ────────────────────────────────────
async function generateCoverImageStability(prompt) {
  if (!process.env.STABILITY_API_KEY) return null;
  try {
    const imgTimeout = new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 25000));
    const imgFetch = fetch('https://api.stability.ai/v1/generation/stable-diffusion-xl-1024-v1-0/text-to-image', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + process.env.STABILITY_API_KEY,
        'Accept': 'application/json'
      },
      body: JSON.stringify({
        text_prompts: [
          { text: prompt + ', professional digital product cover, clean typography, high contrast, no people, no faces', weight: 1 },
          { text: 'blurry, low quality, text errors, watermark, people, faces', weight: -1 }
        ],
        cfg_scale: 7, height: 1024, width: 1024, samples: 1, steps: 25
      })
    });
    const response = await Promise.race([imgFetch, imgTimeout]);
    if (!response.ok) { console.log('Stability error:', response.status); return null; }
    const data = await response.json();
    if (data.artifacts && data.artifacts[0]) {
      return 'data:image/png;base64,' + data.artifacts[0].base64;
    }
    return null;
  } catch(e) {
    console.log('Stability failed:', e.message);
    return null;
  }
}

async function generateImage(prompt) {
  // Try DALL-E 3 first, fall back to Stability AI
  const dalleResult = await generateCoverImage(prompt);
  if (dalleResult) return dalleResult;
  return await generateCoverImageStability(prompt);
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
      const r = await fetch(url, { timeout: 10000, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' } });
      const html = await r.text();
      pageText = html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ')
                     .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
                     .replace(/<[^>]+>/g, ' ')
                     .replace(/\s+/g, ' ')
                     .slice(0, 8000);
    } catch { pageText = 'Could not fetch page automatically. Analyze based on URL context.'; }

    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 2000,
      messages: [{
        role: 'user',
        content: `You are an expert affiliate marketing strategist. Analyze this sales page deeply and return ONLY a valid JSON object.

Analyze the offer and provide a thorough breakdown including:
- What the product actually is and does
- The core mechanism or unique selling proposition
- Exact target audience and their psychology
- The biggest frustrations this audience has experienced before finding this offer
- The strongest emotional hook
- Value gaps (what the offer does NOT cover that buyers will still need)

Return ONLY this JSON with no text before or after:
{
  "product_name": "exact product name",
  "price": "front-end price",
  "commission": "commission rate or amount",
  "niche": "primary niche",
  "target_audience": "detailed description of who this is for",
  "main_pain_point": "the single biggest frustration this audience has",
  "secondary_pain_points": ["pain 2", "pain 3", "pain 4"],
  "main_benefit": "the single strongest promise or result",
  "unique_mechanism": "what makes this offer different from others",
  "audience_psychology": "2-3 sentences on the emotional state of this audience — what they have tried, why they are skeptical, what they really want",
  "summary": "3-4 sentence summary covering what this is, who it helps, why they would buy it, and what makes it worth promoting",
  "offer_score": {
    "overall": 8,
    "commission_rating": 9,
    "niche_demand": 8,
    "conversion_potential": 7,
    "tier1_suitability": 8,
    "notes": "2-3 sentence explanation of scores and what would make this offer stronger"
  },
  "value_gaps": ["specific gap 1 the offer does not address", "specific gap 2", "specific gap 3", "specific gap 4"],
  "recommended_angle": "Consultative or Pain & Agitation or Pure Value & Bonus",
  "angle_reason": "2-3 sentences explaining exactly why this angle fits this specific audience and offer"
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
        main_pain_point: 'Not making money online despite trying multiple systems',
        secondary_pain_points: ['Tech overwhelm', 'Wasted money on courses', 'No clear starting point'],
        main_benefit: 'Start earning commissions with a simple proven system',
        unique_mechanism: 'Done-for-you system with built-in traffic',
        audience_psychology: 'This audience has tried multiple MMO systems and failed. They are skeptical of hype but still hopeful. They want something simple that actually works.',
        summary: 'An affiliate marketing product designed to help beginners earn commissions online.',
        offer_score: { overall: 7, commission_rating: 7, niche_demand: 8, conversion_potential: 7, tier1_suitability: 7, notes: 'Solid MMO offer with good demand. Commission details would improve assessment.' },
        value_gaps: ['No traffic strategy provided', 'No email follow-up templates', 'No implementation checklist', 'No content creation guidance'],
        recommended_angle: 'Pain & Agitation',
        angle_reason: 'MMO audiences respond strongly to pain-based hooks that acknowledge their past failures before presenting a simpler solution.'
      };
    }

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
      model: 'claude-sonnet-4-20250514',
      max_tokens: 6000,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Create a high-value free lead magnet for cold solo ad traffic. This is the free incentive given to people who opt in BEFORE they see the sales page. It must deliver real standalone value while naturally leading them toward the paid offer.

Offer details:
- Product: ${analysis.product_name}
- Niche: ${analysis.niche}
- Target audience: ${analysis.target_audience}
- Main pain point: ${analysis.main_pain_point}
- Main benefit: ${analysis.main_benefit}
- Audience psychology: ${analysis.audience_psychology || ''}
- Angle: ${angle}

LEAD MAGNET RULES:
- Must be 1,500-2,000 words of genuinely useful content
- Give real actionable value — steps, examples, specific guidance
- Should make the reader think "this is great, I want more of this"
- Title must be benefit-driven and curiosity-inducing
- Use headers, numbered steps, bullet points for easy reading
- Do NOT pitch the paid offer inside the lead magnet — let it stand alone
- Write short paragraphs. Use white space. Make it feel clean and easy to read.

Return ONLY valid JSON with no text before or after:
{
  "title": "Lead magnet title",
  "subtitle": "Supporting subtitle",
  "cover_prompt": "DALL-E 3 prompt: A professional 3D ebook cover mockup with dark navy background, bold gold title text, clean modern design, digital product mockup style, no people, high quality render",
  "full_content": "The complete lead magnet content here — 1500-2000 words with proper headers, steps, examples. Write the full thing. No placeholders."
}`
      }]
    });

    let leadMagnet;
    try {
      const text = message.content[0].text;
      leadMagnet = JSON.parse(text.replace(/```json|```/g, '').trim());
    } catch {
      return res.status(500).json({ success: false, message: 'Failed to parse lead magnet content' });
    }

    // Generate cover image
    leadMagnet.cover_image = await generateImage(leadMagnet.cover_prompt);

    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({
        project_id, user_id: user.id, content_type: 'lead_magnet',
        content: JSON.stringify(leadMagnet), updated_at: new Date().toISOString()
      }, { onConflict: 'project_id,content_type' });
    }

    res.json({ success: true, lead_magnet: leadMagnet });
  } catch (err) { res.status(500).json({ success: false, message: 'Lead magnet generation failed: ' + err.message }); }
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
      max_tokens: 8000,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Create a buyer-only bonus stack for this affiliate offer. The chosen angle is: ${angle}.

Offer details:
- Product: ${analysis.product_name}
- Niche: ${analysis.niche}
- Main benefit: ${analysis.main_benefit}
- Main pain point: ${analysis.main_pain_point}
- Value gaps to fill: ${(analysis.value_gaps || []).join(', ')}
- Target audience: ${analysis.target_audience}
- Audience psychology: ${analysis.audience_psychology || ''}

BONUS CONTENT RULES — THIS IS CRITICAL:
- Each bonus must be 400-600 words of REAL content. Not summaries. Not outlines. Actual usable content.
- Every bonus solves ONE specific gap the main offer does not cover
- Use headers, numbered steps, bullet points, real examples
- Write like a helpful expert, not a marketer
- Short paragraphs. White space. Scannable.
- The content must be genuinely valuable on its own — not dependent on the main offer

Generate exactly 3 bonuses:
- Bonus 1: A practical 1-page checklist (step-by-step action format)
- Bonus 2: A 2-page guide (educational, teaches a specific skill or strategy)
- Bonus 3: An AI prompt pack (8-10 copy-paste prompts with instructions)

Return ONLY a valid JSON array with no text before or after:
[
  {
    "number": 1,
    "title": "Bonus title",
    "type": "1-Page Checklist",
    "description": "One clear sentence: what this bonus does and exactly which gap it fills",
    "full_content": "400-600 words of complete, real, usable checklist content with headers and steps. Write the whole thing.",
    "cover_prompt": "DALL-E 3 prompt: A professional 3D ebook cover mockup, dark navy and gold color scheme, bold white title text, clean checklist icon, modern digital product style, high quality render, no people"
  },
  {
    "number": 2,
    "title": "Bonus title",
    "type": "2-Page Guide",
    "description": "One clear sentence description",
    "full_content": "400-600 words of complete guide content with headers, explanation, examples, and actionable steps. Write the whole thing.",
    "cover_prompt": "DALL-E 3 prompt: A professional 3D ebook cover mockup, deep blue and white color scheme, bold title text, open book or guide icon, clean modern design, high quality render, no people"
  },
  {
    "number": 3,
    "title": "Bonus title",
    "type": "AI Prompt Pack",
    "description": "One clear sentence description",
    "full_content": "8-10 complete AI prompts with titles, full prompt text, and brief instructions for each. Write all prompts in full. 400-600 words total.",
    "cover_prompt": "DALL-E 3 prompt: A professional 3D ebook cover mockup, dark purple and gold color scheme, bold title text, AI or lightning bolt icon, futuristic clean design, high quality render, no people"
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
        if (match) bonuses = JSON.parse(match[0]);
        else throw new Error('No JSON array found');
      } catch {
        bonuses = [
          { number:1, title:`${analysis.product_name} Quick-Start Checklist`, type:'1-Page Checklist',
            description:`Step-by-step action plan for getting fast results from ${analysis.product_name}`,
            full_content:`Step 1: Complete your account setup\nStep 2: Go through the core training in order\nStep 3: Apply each lesson before moving on\nStep 4: Pick one traffic method and commit to it for 7 days\nStep 5: Track your clicks and results daily\nStep 6: Scale what is working`,
            cover_prompt:'Professional 3D ebook cover mockup, dark navy background, gold checklist icon, bold white title text, clean modern digital product style, no people' },
          { number:2, title:`${analysis.niche} Fast Track Guide`, type:'2-Page Guide',
            description:`Practical guide to getting your first result in ${analysis.niche}`,
            full_content:`This guide covers the fastest path to your first result.\n\nFocus on one strategy at a time. Consistent daily action beats shortcuts.\n\nSection 1: Mindset\nMost beginners fail not because of lack of tools but lack of consistency. Commit to 30 days.\n\nSection 2: Your First Steps\nDo not try to master everything. Pick one method, get good at it, then expand.\n\nSection 3: Daily Routine\nSpend 30-60 minutes per day on your primary activity. Track results weekly.`,
            cover_prompt:'Professional 3D ebook cover mockup, deep blue and white, bold title text, guide or map icon, clean modern design, no people' },
          { number:3, title:'AI Prompt Pack — Faster Results', type:'AI Prompt Pack',
            description:'10 copy-paste AI prompts to create content and promotional material faster',
            full_content:`Prompt 1: Write 5 social media posts promoting [product] to [audience]\nPrompt 2: Write a follow-up email to someone who opted in but did not buy\nPrompt 3: Write 5 subject line options for a promotional email about [product]\nPrompt 4: Write a bridge page script using the [angle] approach\nPrompt 5: Write a Facebook post using a personal story to promote [product]\nPrompt 6: Write 10 curiosity-based hooks for [niche] content\nPrompt 7: Write objection-handling replies for common [niche] doubts\nPrompt 8: Write a 7-day content plan for promoting [product]\nPrompt 9: Write a short video script for TikTok or Reels about [topic]\nPrompt 10: Write a profile bio that attracts [target audience]`,
            cover_prompt:'Professional 3D ebook cover mockup, dark purple and gold, bold title text, AI lightning bolt icon, futuristic clean design, no people' }
        ];
      }
    }

    // Generate cover images for each bonus
    for (let i = 0; i < bonuses.length; i++) {
      bonuses[i].cover_image = await generateImage(bonuses[i].cover_prompt);
    }

    // Generate bonus stack summary
    const summaryMessage = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Write a compelling bonus stack summary for use on a bridge page. This is a 250-350 word pitch that presents all 3 bonuses as a complete package.

Product: ${analysis.product_name}
Angle: ${angle}
Bonuses:
${bonuses.map(b => `- ${b.title}: ${b.description}`).join('\n')}

The summary should:
- Open with the core problem the buyer still faces after getting the main offer
- Position the bonus stack as the implementation shortcut that fills those gaps
- Describe each bonus in 2-3 sentences with specific benefit language
- Close with a powerful "what this means for you" statement
- End with the three things the bonuses give them (what to do first, where to get traffic, what to say — or equivalents for this niche)

Return ONLY the plain text of the summary. No JSON. No headers. Just the copy.`
      }]
    });

    const stackSummary = summaryMessage.content[0].text.trim();

    // Generate stack image
    const stackImagePrompt = `A professional marketing bonus stack display, showing 3 digital ebook covers arranged together, dark navy background, gold and white text, "EXCLUSIVE BUYER BONUSES" label, clean premium marketing design, high quality render, no people`;
    const stackImage = await generateImage(stackImagePrompt);

    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({
        project_id, user_id: user.id, content_type: 'bonus_stack',
        content: JSON.stringify({ bonuses, stack_summary: stackSummary, stack_image: stackImage }),
        updated_at: new Date().toISOString()
      }, { onConflict: 'project_id,content_type' });
    }

    res.json({ success: true, bonuses, stack_summary: stackSummary, stack_image: stackImage });
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

  const leadMagnetContext = lead_magnet
    ? `Lead magnet: "${lead_magnet.title}" — ${lead_magnet.subtitle || 'Free guide for new subscribers'}`
    : 'No lead magnet — use the main offer promise as the opt-in incentive.';

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1500,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Write a high-converting opt-in squeeze page for cold solo ad traffic using the ${angle} angle.

Offer: ${analysis.product_name}
Niche: ${analysis.niche}
Main pain point: ${analysis.main_pain_point}
Secondary pain points: ${(analysis.secondary_pain_points || []).join(', ')}
Main benefit: ${analysis.main_benefit}
Target audience: ${analysis.target_audience}
Audience psychology: ${analysis.audience_psychology || ''}
${leadMagnetContext}

CRITICAL RULES:
- No bonuses mentioned anywhere — buyers only
- Ultra-concise above the fold — everything must fit without scrolling
- Two A/B headline variations testing different hooks
- Bullet points must be specific benefits, not vague claims
- CTA button text must create urgency or curiosity — not just "Submit"
- Microcopy under button removes objections

Return ONLY valid JSON:
{
  "headline_a": "Version A headline — curiosity/benefit driven, speaks directly to the pain",
  "headline_b": "Version B headline — different hook, tests a different emotional trigger",
  "subheadline": "One sentence that expands the promise and makes it feel real and achievable",
  "bullets": ["Specific benefit 1", "Specific benefit 2", "Specific benefit 3", "Specific benefit 4", "Specific benefit 5"],
  "cta_button": "Action-oriented CTA button text",
  "microcopy": "Short reassurance line under the button — removes objections",
  "form_fields": ["First Name", "Email Address"],
  "above_fold_note": "Brief note on above-fold layout and which elements are most critical",
  "split_test_tip": "Which headline to run first and exactly why based on this audience and angle"
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
  const { analysis, angle, format, bonuses, stack_summary, lead_magnet, project_id } = req.body;
  if (!analysis || !angle || !format) return res.status(400).json({ success: false, message: 'Analysis, angle, and format required' });

  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  const bonusContext = bonuses && bonuses.length > 0
    ? `Buyer-only bonuses:\n${bonuses.map(b => `- ${b.title}: ${b.description}`).join('\n')}\n\nBonus stack summary: ${stack_summary || 'Not provided'}`
    : 'No bonuses defined — write without bonus references.';

  const leadMagnetContext = lead_magnet
    ? `The visitor just opted in to receive: "${lead_magnet.title}". Reference this in the bridge to maintain continuity.`
    : 'No lead magnet — visitor came straight from ad.';

  const formatInstructions = {
    text: 'Write text-only copy with: headline, relatable story/problem (2-3 paragraphs), benefit bullets, bonus stack pitch with each bonus described, strong CTA. Minimum 600 words.',
    vsl: 'Write a complete VSL video script (reads as 3-4 minutes when spoken) PLUS full supporting text copy below the video. Script must have: hook, relatable problem, solution reveal, bonus stack pitch with each bonus named, urgency close, CTA. Total minimum 800 words.',
    short: 'Write a punchy 4-6 sentence intro blurb that creates massive curiosity plus a strong CTA. Keep it extremely tight but impactful.'
  };

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 3000,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Write a high-converting bridge page for this affiliate offer using the ${angle} angle.

Offer: ${analysis.product_name}
Price: ${analysis.price || 'low-ticket'}
Main benefit: ${analysis.main_benefit}
Main pain point: ${analysis.main_pain_point}
Unique mechanism: ${analysis.unique_mechanism || ''}
Audience psychology: ${analysis.audience_psychology || ''}
${leadMagnetContext}
${bonusContext}

Format: ${formatInstructions[format] || formatInstructions.text}

CRITICAL RULES:
- Bridge the gap from opt-in to sales page — this is your personal recommendation
- Acknowledge the reader's skepticism and past failures — make them feel understood
- Present the offer as the logical solution they have been looking for
- Pitch the bonus stack as the reason to buy specifically through your link
- Critical info and first CTA must appear above the fold
- Write in first-person conversational tone — like a recommendation from a trusted friend

Return ONLY valid JSON:
{
  "headline": "Bridge page headline — strong, benefit-driven, creates urgency",
  "subheadline": "Supporting subheadline",
  "content": "Full bridge page copy as specified in format — write every word, no placeholders",
  "cta_text": "CTA button text",
  "bonus_claim_note": "Short note under CTA explaining how to claim bonuses after purchase",
  "above_fold_note": "What must appear above fold and why"
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
  const { analysis, angle, sequence_type, bonuses, lead_magnet, project_id } = req.body;
  if (!analysis || !angle || !sequence_type) return res.status(400).json({ success: false, message: 'Analysis, angle, and sequence type required' });

  const access = await getUserAccess(user.id);
  if (!access.solo_ads) return res.status(403).json({ success: false, message: 'Solo Ads channel not unlocked' });

  const bonusContext = bonuses && bonuses.length > 0
    ? `Buyer-only bonuses:\n${bonuses.map(b => `- ${b.title}: ${b.description}`).join('\n')}`
    : 'No bonuses defined.';

  const leadMagnetContext = lead_magnet
    ? `Lead magnet delivered on opt-in: "${lead_magnet.title}"`
    : 'No lead magnet.';

  const sequenceInstructions = {
    non_buyers: `Non-Buyers 4-Day Sequence (Goal: Convert to Sale):
Day 1: Thank you for opting in. Deliver lead magnet if applicable. Check if they have questions. Softly introduce the offer with a link. Keep it warm and low-pressure.
Day 2: Story/Solution email. Tell a relatable story about the struggle this audience faces. Transition to how the offer solves it. Include affiliate link.
Day 3: Hard pitch. Name and describe each buyer-only bonus specifically. Make the bonus stack the main reason to buy today through your link. Clear CTA.
Day 4: Final push. Create gentle scarcity or urgency. Position simple action over endless research. Final reminder of bonuses. Last chance energy.`,
    buyers: `Buyers 4-Day Sequence (Goal: Onboard & Ascend):
Day 1: Welcome and congrats. Tell them how to access all 3 bonuses. Tell them exactly what to do first. Keep it warm and celebratory.
Day 2: Advanced tip email. Teach them one specific insight that makes the main offer work better. Real value, not a sales pitch.
Day 3: Introduce a related solution or natural upsell. Frame it as the logical next step after they have implemented the main offer.
Day 4: Demo/benefits of the related solution. Make it feel like the obvious continuation of their journey.`
  };

  try {
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 5000,
      system: SOLO_ADS_KB,
      messages: [{
        role: 'user',
        content: `Write a complete ${sequence_type === 'non_buyers' ? 'Non-Buyers' : 'Buyers'} email sequence using the ${angle} angle.

Offer: ${analysis.product_name}
Niche: ${analysis.niche}
Main pain point: ${analysis.main_pain_point}
Main benefit: ${analysis.main_benefit}
Audience psychology: ${analysis.audience_psychology || ''}
${leadMagnetContext}
${bonusContext}

${sequenceInstructions[sequence_type]}

EMAIL WRITING RULES — CRITICAL:
- Write every email IN FULL. No placeholders. No "[write story here]". Actual copy.
- Use [FirstName] for personalization
- Use [YourAffiliateLink] for the offer link
- Use [YourName] for sign-off
- Subject lines: one curiosity-based, one direct benefit, one relatable question
- Keep paragraphs short — 1-3 sentences max
- Use line breaks generously — emails must be easy to skim
- End every non-buyers email with a PS that reinforces the main reason to act

Return ONLY a valid JSON array of 4 emails with no text before or after:
[
  {
    "day": 1,
    "subject_lines": ["Curiosity subject option", "Direct benefit subject option", "Relatable question subject option"],
    "body": "Complete full email body — every word written out, no placeholders for content"
  }
]`
      }]
    });

    let emails;
    try { emails = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch {
      try {
        const match = message.content[0].text.match(/\[[\s\S]*\]/);
        if (match) emails = JSON.parse(match[0]);
        else throw new Error('No array found');
      } catch { return res.status(500).json({ success: false, message: 'Failed to parse email content' }); }
    }

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

// ── GENERATE BLUEPRINT ────────────────────────────────────────
app.post('/api/generate/blueprint', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { analysis, angle, bonuses, lead_magnet, project_id } = req.body;
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
        content: `Create a complete copy-paste implementation blueprint for this affiliate funnel.

Offer: ${analysis.product_name}
Angle: ${angle}
Lead magnet: ${lead_magnet ? lead_magnet.title : 'None'}
Bonuses: ${bonuses ? bonuses.map(b => b.title).join(', ') : 'None'}

Write a complete funnel assembly guide that covers:
1. Funnel flow overview (the exact sequence from ad click to buyer)
2. Opt-in page setup (what goes where, design notes)
3. Bridge page setup (what goes where, video placement if applicable)
4. Lead magnet delivery setup
5. Email automation setup (triggers, timing, segments)
6. Bonus claim process
7. Asset placement checklist
8. Tracking setup (what metrics to track and where)
9. Final build order (step by step)
10. Pre-launch test checklist

Be specific and practical. This should read like instructions from someone who has built 100 of these funnels.

Return ONLY valid JSON:
{
  "title": "Blueprint title",
  "sections": [
    {
      "heading": "Section heading",
      "content": "Full section content — specific, practical, actionable"
    }
  ]
}`
      }]
    });

    let blueprint;
    try { blueprint = JSON.parse(message.content[0].text.replace(/```json|```/g, '').trim()); }
    catch { return res.status(500).json({ success: false, message: 'Failed to parse blueprint' }); }

    if (project_id) {
      await adminClient.from('promolab_project_content').upsert({
        project_id, user_id: user.id, content_type: 'blueprint',
        content: JSON.stringify(blueprint), updated_at: new Date().toISOString()
      }, { onConflict: 'project_id,content_type' });
    }

    res.json({ success: true, blueprint });
  } catch (err) { res.status(500).json({ success: false, message: 'Blueprint generation failed: ' + err.message }); }
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
  (content || []).forEach(c => {
    try { contentMap[c.content_type] = JSON.parse(c.content); } catch { contentMap[c.content_type] = c.content; }
  });
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

// ── DOWNLOAD BONUS (single) ───────────────────────────────────
app.post('/api/download/bonus', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { bonus, product_name } = req.body;
  if (!bonus) return res.status(400).json({ success: false, message: 'Bonus data required' });

  const coverImg = bonus.cover_image
    ? `<img src="${bonus.cover_image}" style="width:200px;height:200px;border-radius:8px;margin:0 auto 24px;display:block;">`
    : `<div style="width:200px;height:200px;background:linear-gradient(135deg,#1A1A2E,#4F46B8);border-radius:8px;margin:0 auto 24px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:13px;font-weight:700;text-align:center;padding:20px;">${bonus.title || 'Bonus'}</div>`;

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${bonus.title}</title>
<style>body{font-family:Arial,sans-serif;max-width:680px;margin:0 auto;padding:48px 32px;color:#1A1A2E}
.header{text-align:center;margin-bottom:40px;border-bottom:3px solid #4F46B8;padding-bottom:32px}
.eyebrow{font-size:11px;color:#4F46B8;text-transform:uppercase;letter-spacing:0.08em;font-weight:700;margin-bottom:8px}
.bonus-label{display:inline-block;background:#4F46B8;color:#fff;font-size:11px;font-weight:700;padding:4px 12px;border-radius:20px;margin-bottom:16px}
h1{font-size:26px;font-weight:700;color:#1A1A2E;margin:0 0 12px;line-height:1.3}
.description{font-size:14px;color:#555;line-height:1.6;margin-bottom:0}
.content{font-size:14px;line-height:1.9;color:#333;white-space:pre-wrap;margin-top:32px}
.footer{margin-top:48px;padding-top:20px;border-top:1px solid #EEEDFE;font-size:11px;color:#aaa;text-align:center}
@media print{body{padding:24px}}</style></head>
<body><div class="header">${coverImg}
<div class="eyebrow">Exclusive Buyer Bonus</div>
<span class="bonus-label">Bonus #${bonus.number} &mdash; ${bonus.type}</span>
<h1>${bonus.title}</h1>
<p class="description">${bonus.description}</p></div>
<div class="content">${(bonus.full_content || '').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</div>
<div class="footer">Exclusive bonus for buyers of ${(product_name || 'this offer').replace(/</g,'&lt;')} &bull; PromoLab by Jimmy Griffith, JGAffiliate</div>
</body></html>`;

  res.setHeader('Content-Type', 'text/html');
  res.setHeader('Content-Disposition', `attachment; filename="Bonus-${bonus.number}-${(bonus.title||'bonus').replace(/[^a-z0-9]/gi,'_').slice(0,40)}.html"`);
  res.send(html);
});

// ── DOWNLOAD ALL BONUSES ──────────────────────────────────────
app.post('/api/download/bonuses-all', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { bonuses, product_name } = req.body;
  if (!bonuses || !bonuses.length) return res.status(400).json({ success: false, message: 'Bonuses required' });

  let allBonusesHtml = '';
  bonuses.forEach(function(bonus) {
    const coverImg = bonus.cover_image
      ? `<img src="${bonus.cover_image}" style="width:200px;height:200px;border-radius:8px;margin:0 auto 24px;display:block;">`
      : `<div style="width:200px;height:200px;background:linear-gradient(135deg,#1A1A2E,#4F46B8);border-radius:8px;margin:0 auto 24px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:13px;font-weight:700;text-align:center;padding:20px;">${bonus.title || 'Bonus'}</div>`;
    allBonusesHtml += `<div class="bonus-section">
      <div class="bonus-header">${coverImg}
      <div class="eyebrow">Exclusive Buyer Bonus #${bonus.number}</div>
      <span class="bonus-label">${bonus.type}</span>
      <h2>${bonus.title}</h2>
      <p class="description">${bonus.description}</p></div>
      <div class="content">${(bonus.full_content || '').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</div>
      </div><div class="page-break"></div>`;
  });

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Bonus Stack — ${(product_name||'Your Offer').replace(/</g,'&lt;')}</title>
<style>body{font-family:Arial,sans-serif;max-width:680px;margin:0 auto;padding:48px 32px;color:#1A1A2E}
.cover-page{text-align:center;padding:80px 40px;border-bottom:3px solid #4F46B8;margin-bottom:60px}
.cover-page h1{font-size:32px;font-weight:700;margin:0 0 12px}
.cover-page p{font-size:15px;color:#555}
.bonus-section{margin-bottom:60px}
.bonus-header{text-align:center;margin-bottom:32px;padding-bottom:28px;border-bottom:2px solid #EEEDFE}
.eyebrow{font-size:11px;color:#4F46B8;text-transform:uppercase;letter-spacing:0.08em;font-weight:700;margin-bottom:8px}
.bonus-label{display:inline-block;background:#4F46B8;color:#fff;font-size:11px;font-weight:700;padding:4px 12px;border-radius:20px;margin-bottom:16px}
h2{font-size:24px;font-weight:700;margin:0 0 12px;line-height:1.3}
.description{font-size:14px;color:#555;line-height:1.6;margin:0}
.content{font-size:14px;line-height:1.9;color:#333;white-space:pre-wrap}
.page-break{page-break-after:always;border-bottom:1px solid #EEEDFE;margin:40px 0}
.footer{margin-top:48px;padding-top:20px;border-top:1px solid #EEEDFE;font-size:11px;color:#aaa;text-align:center}
@media print{.page-break{page-break-after:always}.cover-page{page-break-after:always}}</style></head>
<body>
<div class="cover-page">
<div class="eyebrow">Exclusive Buyer Bonus Package</div>
<h1>${(product_name||'Your Offer').replace(/</g,'&lt;')}</h1>
<p>Thank you for your purchase. Here are your ${bonuses.length} exclusive bonuses.</p></div>
${allBonusesHtml}
<div class="footer">Exclusive bonus package &bull; PromoLab by Jimmy Griffith, JGAffiliate</div>
</body></html>`;

  res.setHeader('Content-Type', 'text/html');
  res.setHeader('Content-Disposition', `attachment; filename="Complete-Bonus-Stack-${(product_name||'offer').replace(/[^a-z0-9]/gi,'_').slice(0,30)}.html"`);
  res.send(html);
});

// ── DOWNLOAD LEAD MAGNET ──────────────────────────────────────
app.post('/api/download/lead-magnet', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { lead_magnet, product_name } = req.body;
  if (!lead_magnet) return res.status(400).json({ success: false, message: 'Lead magnet data required' });

  const coverImg = lead_magnet.cover_image
    ? `<img src="${lead_magnet.cover_image}" style="width:220px;height:220px;border-radius:8px;margin:0 auto 24px;display:block;">`
    : `<div style="width:220px;height:220px;background:linear-gradient(135deg,#1A1A2E,#4F46B8);border-radius:8px;margin:0 auto 24px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:14px;font-weight:700;text-align:center;padding:24px;">${lead_magnet.title || 'Free Guide'}</div>`;

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${lead_magnet.title}</title>
<style>body{font-family:Arial,sans-serif;max-width:680px;margin:0 auto;padding:48px 32px;color:#1A1A2E}
.header{text-align:center;margin-bottom:40px;border-bottom:3px solid #4F46B8;padding-bottom:32px}
.free-tag{display:inline-block;background:#EAF3DE;color:#1D6A3A;font-size:11px;font-weight:700;padding:4px 12px;border-radius:20px;margin-bottom:16px;text-transform:uppercase;letter-spacing:0.06em}
h1{font-size:28px;font-weight:700;color:#1A1A2E;margin:0 0 10px;line-height:1.3}
.subtitle{font-size:16px;color:#555;line-height:1.5;margin:0}
.content{font-size:14px;line-height:1.9;color:#333;white-space:pre-wrap;margin-top:32px}
.footer{margin-top:48px;padding-top:20px;border-top:1px solid #EEEDFE;font-size:11px;color:#aaa;text-align:center}
@media print{body{padding:24px}}</style></head>
<body><div class="header">${coverImg}
<span class="free-tag">Free Guide</span>
<h1>${lead_magnet.title}</h1>
<p class="subtitle">${lead_magnet.subtitle || ''}</p></div>
<div class="content">${(lead_magnet.full_content || '').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</div>
<div class="footer">Free guide courtesy of ${(product_name || 'JGAffiliate').replace(/</g,'&lt;')} &bull; PromoLab by Jimmy Griffith, JGAffiliate</div>
</body></html>`;

  res.setHeader('Content-Type', 'text/html');
  res.setHeader('Content-Disposition', `attachment; filename="Lead-Magnet-${(lead_magnet.title||'guide').replace(/[^a-z0-9]/gi,'_').slice(0,40)}.html"`);
  res.send(html);
});

// ── DOWNLOAD BLUEPRINT ────────────────────────────────────────
app.post('/api/download/blueprint', async (req, res) => {
  const user = await requireUser(req, res);
  if (!user) return;
  const { blueprint, product_name } = req.body;
  if (!blueprint) return res.status(400).json({ success: false, message: 'Blueprint data required' });

  let sectionsHtml = (blueprint.sections || []).map(s =>
    `<div class="section"><h2>${s.heading}</h2><div class="content">${(s.content||'').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</div></div>`
  ).join('');

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${blueprint.title || 'Funnel Blueprint'}</title>
<style>body{font-family:Arial,sans-serif;max-width:720px;margin:0 auto;padding:48px 32px;color:#1A1A2E}
.cover{text-align:center;padding:60px 0 40px;border-bottom:3px solid #4F46B8;margin-bottom:48px}
.eyebrow{font-size:11px;color:#4F46B8;text-transform:uppercase;letter-spacing:0.08em;font-weight:700;margin-bottom:12px}
h1{font-size:30px;font-weight:700;margin:0 0 12px;line-height:1.3}
.cover p{font-size:14px;color:#555}
.section{margin-bottom:40px;padding-bottom:32px;border-bottom:1px solid #EEEDFE}
h2{font-size:20px;font-weight:700;color:#4F46B8;margin:0 0 16px}
.content{font-size:14px;line-height:1.9;color:#333;white-space:pre-wrap}
.footer{margin-top:48px;padding-top:20px;border-top:1px solid #EEEDFE;font-size:11px;color:#aaa;text-align:center}
@media print{body{padding:24px}}</style></head>
<body>
<div class="cover">
<div class="eyebrow">Funnel Implementation Guide</div>
<h1>${blueprint.title || 'Copy-Paste Blueprint'}</h1>
<p>Complete assembly instructions for your ${(product_name||'affiliate').replace(/</g,'&lt;')} funnel</p></div>
${sectionsHtml}
<div class="footer">PromoLab by Jimmy Griffith, JGAffiliate</div>
</body></html>`;

  res.setHeader('Content-Type', 'text/html');
  res.setHeader('Content-Disposition', `attachment; filename="Funnel-Blueprint-${(product_name||'offer').replace(/[^a-z0-9]/gi,'_').slice(0,30)}.html"`);
  res.send(html);
});

app.listen(PORT, () => console.log('PromoLab v3 running on port ' + PORT));
