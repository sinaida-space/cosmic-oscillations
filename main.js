import PptxGenJS from 'pptxgenjs';

// ── STATE ─────────────────────────────────────────────────────────
const S = {
    ans: JSON.parse(sessionStorage.getItem('sn_a') || '{}'),
    projectName: '', presenterName: '', deck: null,
    save() { sessionStorage.setItem('sn_a', JSON.stringify(this.ans)); }
};

// ── QUESTIONS ─────────────────────────────────────────────────────
const Q = [
  { id:'problem', num:'01', slide:'Problem Statement',
    opener:"Every great project picks a fight with something broken.",
    title:"What opportunities for improvement do you see, and what makes you say that?",
    guidance:"Name the specific pain point and who experiences it daily. Make it feel urgent and provable — not a vague observation but a daily frustration.",
    placeholder:"e.g. Product managers spend 6h/week manually consolidating feedback from 4+ tools into one report. It's error-prone and soul-crushing, but nobody questions it anymore.",
    chips:["Who is the main victim here?","What evidence proves it's real?","What happens if no one steps in?","Why is it urgent right now?","How often does this occur?","What does inaction cost in time or money?","Is this your personal pain too?","Why is everyone accepting this as normal?"],
    refine:"Rewrite this as a crisp 2-sentence problem statement naming: the specific person affected, the specific friction, and one piece of evidence. Avoid 'inefficiency' or 'challenges'. Be concrete. Source:"},
  { id:'solution', num:'02', slide:'Solution',
    opener:"Now for the vanish trick.",
    title:"When you win, what annoying thing completely disappears from the world?",
    guidance:"Describe the BEFORE (current friction, workarounds, tears) and the AFTER (what's measurably better). One hard number makes it unforgettable.",
    placeholder:"e.g. Before: 4 tabs open, copy-paste into Excel, 2h of re-formatting. After: one dashboard, all sources merged, report auto-generates in 3 seconds.",
    chips:["Current painful process step by step?","The single biggest friction point?","Before vs. after in a hard number","What gets completely removed?","What becomes possible that wasn't?","What's the 'aha' moment for users?","What would make it irreplaceable?","What are you deliberately NOT doing?"],
    refine:"Rewrite as a tight BEFORE/AFTER contrast. BEFORE: current state including friction and workarounds (3-4 bullets). AFTER: improved state with one measurable improvement (3-4 bullets). Parallel structure, no buzzwords. Source:"},
  { id:'clients', num:'03', slide:'Target Clients',
    opener:"The hero's welcome.",
    title:"Who is the specific person that will literally hug you when you launch this?",
    guidance:"Describe 2–3 segments with precision. Not demographics, but behaviors and daily frustrations. For each: what's their specific pain, and what concrete change does your solution bring?",
    placeholder:"e.g. (1) Solo PMs at SaaS startups — drown in Notion + Jira + Intercom. (2) Heads of Product — need exec-ready weekly reports, currently built manually in Google Slides at 8pm on Sundays.",
    chips:["Their specific daily frustration?","What do they currently pay to fix this?","Who is the desperate early adopter?","Where do they complain about this the most?","How does their week change after?","What's their absolute win condition?","Who refers them to tools today?","What would make them switch instantly?"],
    refine:"For each client segment write: [Role/context] — [2-sentence pain specific to that role] — [1 outcome sentence starting 'With this, they can...']. Keep each under 60 words. Source:"},
  { id:'goals', num:'04', slide:'Goal & Success Metrics',
    opener:"The 12-month truth serum. Fast forward a year.",
    title:"What single number proves you weren't crazy to start this?",
    guidance:"Goal: what, by when, measured how. Then 3 Critical-to-Quality metrics — the non-negotiables — each with a target number and a measurement method. Vague metrics get ignored.",
    placeholder:"e.g. Goal: 500 active paying users by Dec 2026. Metric 1: user activates within 24h (target >70%). Metric 2: 30-day retention (>60%). Metric 3: NPS >45.",
    chips:["Main goal in one sentence?","3 metrics you'd be aggressively proud to report?","How do you actually measure each?","What early signal would trigger a pivot?","The make-or-break 6-month milestone?","What does failure look like early?","How do you catch failure before it's too late?","What is the absolute moonshot version?"],
    refine:"Write one goal sentence: 'By [date], achieve [outcome] measured by [metric].' Then 3 CTQ rows: name (3-5 words), description, numeric target, measurement method. No 'improve satisfaction'. Source:"},
  { id:'scope', num:'05', slide:'Scope Definition',
    opener:"The power of 'no'. Clarity about what you're NOT building is as strategic as what you are.",
    title:"What is the one highly requested feature you are explicitly refusing to build in v1?",
    guidance:"List 3–5 things you ARE definitely doing in the initial launch phase, and 3–5 things you're aggressively not doing yet, and tell me why.",
    placeholder:"e.g. IN: invoice generation from Figma exports, PDF download, basic dashboard. OUT: payment collection (distraction/phase 2), team accounts (too complex for now).",
    chips:["V1 vital features in 3 bullets?","Reason for each out-of-scope item?","Time to the first usable raw product?","What is the biggest risk of scope creep?","Systems or teams completely excluded?","Geographic limits?","What's the natural phase 2?","What are you saying no to permanently?"],
    refine:"Rewrite scope items so each IN SCOPE item is verifiable (not 'improve UX' but 'redesign checkout for mobile') and each OUT OF SCOPE item names the reason for exclusion. Max 5 per column. Source:"},
  { id:'competitors', num:'06', slide:'Competitive Landscape',
    opener:"The 'good enough' enemy.",
    title:"Let's be honest — what's the painful workaround people are putting up with right now?",
    guidance:"Name 2–3 real alternatives. Pick 4 criteria that matter most to your users. Be honest about where competitors do well. Your advantage must be specific — not just 'better UX'.",
    placeholder:"e.g. Alternatives: ProductBoard, Notion, manual Excel. Criteria: setup time, integrations, report automation, price. Gap: none auto-generate PM reports from live data.",
    chips:["Direct competitors by name?","What hacky spreadsheet do people use today?","4 criteria that matter most to buyers?","Where does the giant incumbent actually win?","Where do they fail your user?","Your completely unfair advantage (be ruthless)?","What can't they literally ever copy?","What trend is working in your favor right now?"],
    refine:"Create a comparison: 4 criteria rows × 3 columns (Competitor A, Competitor B, Our Approach). Each cell: 5-10 words, honest. If no direct competitors, frame as 'Current Alternatives'. Source:"},
  { id:'progress', num:'07', slide:'Progress & Findings',
    opener:"Leave the building. Evidence beats smart plans taking place on whiteboards.",
    title:"What's the most surprising thing you learned when you actually talked to the people who need this?",
    guidance:"Replace claims with evidence. Three headline numbers + three key findings demonstrate rigor. The most compelling finding is one that surprised you or changed your thinking.",
    placeholder:"e.g. 18 user interviews. 2 prototype rounds (230 Figma plugin installs). Finding: 14/18 said report consolidation was their #1 weekly time-sink. Surprise: they don't want AI to write — just gather and organize.",
    chips:["Research, interviews, or prototypes done so far?","Total number of stakeholders spoken to?","Prototype or MVP current status?","3 most critical findings?","What absolutely shocked you?","What completely changed your original thinking?","What miserably failed?","Hard evidence it's technically feasible?"],
    refine:"Extract 3 headline statistics. Then write 3 key findings: (1) what research confirmed with a quote or data point, (2) a surprising insight or pivot, (3) the signal that justifies moving forward. Be specific. Source:"},
  { id:'team', num:'08', slide:'Team',
    opener:"The secret weapon.",
    title:"What hard-learned lesson from your past makes you the exact right person for this?",
    guidance:"Not a generic LinkedIn bio. What makes each person specifically essential for this exact problem? The story of how you came to it matters as much as formal credentials.",
    placeholder:"e.g. [Name]: PM at Atlassian 6y, built first Jira-Slack integration. This is literally her own daily problem. [Name]: ex-Stripe engineer, shipped 3 complex B2B SaaS architecture tools.",
    chips:["Core team in 2 potent lines each?","Why this problem — the deep personal connection?","Each person's one indispensable superpower?","Full-time or nights and weekends?","History of the founders working together?","Advisors or sponsors adding massive credibility?","What's the gaping hole in the team today?","Biggest team-related risk?"],
    refine:"Rewrite each person as: [Name] — [Role] — [One sentence: why THIS person is essential to THIS project's success, not just 'experienced']. Make it feel personal. Source:"},
  { id:'resources', num:'09', slide:'Resources',
    opener:"The back pocket.",
    title:"What unfair advantages do you already have, and what is the glaring missing piece?",
    guidance:"HAVE: specific items with amounts or names. NEED: specific items with estimated cost or effort. Not 'some budget' but '€50k from Q3 fund'. Not 'more devs' but '2 backend engineers for 6 months'.",
    placeholder:"e.g. HAVE: €40k savings, 1 part-time killer designer, Figma pro license. NEED: €150k seed (€80k eng hire, €40k pilot program, €30k ops). 6-month runway to first revenue.",
    chips:["Current actual budget in hard numbers?","Already committed partnerships or sponsors?","Specific roles desperately still needed?","Tools, audience, or IP you legitimately already own?","Top 3 fatal resource gaps?","Regulatory or legal requirements?","Strategic relationships you have in your network?","First use of new funding? What's the fire?"],
    refine:"Two parallel lists. HAVE: specific items with amounts or names. NEED: specific items with estimated cost or effort. Every entry should be actionable, not vague. Source:"},
  { id:'risks', num:'10', slide:'Risk Assessment',
    opener:"The post-mortem. A team that names risks proactively is a thousand times more credible.",
    title:"Pretend it's a year from now and the project completely failed. What actually killed it?",
    guidance:"Cover 4 categories: Delivery (execution), Operational (post-launch), Market (external), Dependency (vendors/approvals). Each risk needs a specific mitigation — not 'we'll monitor it'.",
    placeholder:"e.g. Market: Figma API change (Medium) → mitigation: export-agnostic core architecture. Delivery: scope creep (High) → mitigation: hard scope doc, brutal weekly review. Dependency: GDPR (Medium) → legal review Q2.",
    chips:["Biggest single existential risk?","Technical or delivery meltdown risks?","Market timing risks?","Vendor or API sudden dependency risks?","Regulatory or legal constraints?","Early warning signal for each major risk?","Concrete mitigation for your top 3 risks?","What would make you completely stop and pivot?"],
    refine:"Organize as a table: Category (Delivery/Operational/Market/Dependency), Risk Level (H/M/L), Specific Risk Scenario, Mitigation (owned action, not 'monitor'). 3-4 rows. Mitigations must be concrete. Source:"},
  { id:'market', num:'11', slide:'Market Potential',
    opener:"The bar napkin math. Size the prize.",
    title:"If this works perfectly, how huge does this actually get?",
    guidance:"TAM → SAM → SOM with the logic behind each number. Start with: how many target users exist, what they spend on this problem, your realistic share in 3 years. Source your numbers.",
    placeholder:"e.g. TAM: PM tooling market $4.8B (Gartner 2024). SAM: mid-market SaaS PMs in EU+NA, ~400k users. SOM: 5k users at €29/mo by end of Y2 = €1.74M ARR.",
    chips:["Total market size + a real source?","Your highly specific addressable segment?","Realistic absolutely conservative user count in 3 years?","Revenue per user per year realistically?","Is the market growing or shrinking rapidly?","Macro trend forcibly driving this now?","Adjacent markets to expand into later?","Why now is profoundly better than 2 years ago?"],
    refine:"Structure as TAM/SAM/SOM. For each level, one sentence explaining the logic behind the number, not just the number. Flag numbers that need source citation. Source:"},
  { id:'model', num:'12', slide:'Business Model',
    opener:"The money flow. The last question — first one investors literally ask.",
    title:"How do the economics actually make sense? Where does the first dollar come from?",
    guidance:"Pricing model + price point + how customers discover and buy + rough unit economics + Year 1–3 projections with ONE key assumption stated explicitly. Simpler always wins.",
    placeholder:"e.g. SaaS: €29/mo individual / €99/mo team. Discovery: PLG via Figma plugin store. CAC ~€80, LTV ~€580 (20mo avg). Y1: €60k ARR. Y2: €240k. Key assumption: 3% free-to-paid conversion.",
    chips:["Exact pricing model + price point?","How does the profoundly busy customer discover you?","How do they actually swipe their card?","Customer acquisition cost very rough estimate?","Lifetime value rough estimate?","When does the literal first revenue arrive?","Y1/Y2/Y3 realistic revenue targets?","The one massive key assumption behind the projections?"],
    refine:"Four sections: (1) Pricing Model — model name + price point justified in one sentence. (2) Sales Model — how discovered and purchased. (3) Unit Economics — CAC, LTV, payback (estimates ok). (4) Revenue Forecast — Y1/2/3 with ONE stated key assumption. Source:"}
];

const PHRASES = ['Reading between the lines...','Connecting the dots...','Building your narrative...','Sharpening the language...','Almost there — polishing...','Worth the wait.'];

// ── IMPROVE PANEL ─────────────────────────────────────────────────
let improveEl = null;
function createImprovePanel() {
    improveEl = document.createElement('div');
    improveEl.style.cssText = 'position:fixed;inset:0;z-index:8000;background:rgba(8,8,8,.97);display:flex;align-items:center;justify-content:center;padding:40px;opacity:0;pointer-events:none;transition:opacity .3s;';
    improveEl.innerHTML = `
        <div style="max-width:680px;width:100%;border:1px solid rgba(255,255,255,.1);padding:48px;position:relative;background:#0d0d0d;">
            <button id="imp-close" style="position:absolute;top:16px;right:20px;font-size:22px;background:none;border:none;color:rgba(255,255,255,.4);cursor:default;">×</button>
            <p style="font-size:11px;font-weight:700;letter-spacing:.22em;text-transform:uppercase;color:#E30613;margin-bottom:10px;">REFINE WITH AI</p>
            <h3 style="font-size:22px;font-weight:700;margin-bottom:8px;">Copy this prompt into Claude or ChatGPT</h3>
            <p style="font-size:14px;color:rgba(255,255,255,.5);margin-bottom:24px;">Paste the prompt below into any AI to get a sharper, deck-ready version of your answer.</p>
            <textarea id="imp-prompt" style="width:100%;height:220px;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.1);color:#fff;padding:18px;font-size:13px;font-family:Inter,sans-serif;resize:none;outline:none;line-height:1.6;" readonly></textarea>
            <button id="imp-copy" style="margin-top:16px;width:100%;height:50px;background:#fff;color:#000;border:none;font-size:13px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;">Copy to clipboard</button>
        </div>`;
    document.body.appendChild(improveEl);
    improveEl.querySelector('#imp-close').onclick = closeImprove;
    improveEl.onclick = e => { if (e.target === improveEl) closeImprove(); };
    improveEl.querySelector('#imp-copy').onclick = () => {
        navigator.clipboard.writeText(improveEl.querySelector('#imp-prompt').value).then(() => {
            improveEl.querySelector('#imp-copy').textContent = '✓ Copied!';
            setTimeout(() => { improveEl.querySelector('#imp-copy').textContent = 'Copy to clipboard'; }, 2000);
        });
    };
}
function openImprove(prompt) {
    if (!improveEl) createImprovePanel();
    improveEl.querySelector('#imp-prompt').value = prompt;
    improveEl.style.opacity = '1'; improveEl.style.pointerEvents = 'all';
}
function closeImprove() { improveEl.style.opacity = '0'; improveEl.style.pointerEvents = 'none'; }

// ── CURSOR ────────────────────────────────────────────────────────
const cursor = document.getElementById('cursor');
if (cursor) {
    const HERO_LABELS = ['Think →', 'Build →', 'Begin →', 'Start →'];
    let heroLabelIdx = 0;
    setInterval(() => {
        heroLabelIdx = (heroLabelIdx + 1) % HERO_LABELS.length;
        if (cursor.classList.contains('label')) cursor.textContent = HERO_LABELS[heroLabelIdx];
    }, 2200);

    // Initial state
    cursor.style.opacity = '0';

    document.addEventListener('mousemove', e => {
        cursor.style.opacity = '1';
        cursor.style.left = e.clientX + 'px';
        cursor.style.top = e.clientY + 'px';
        
        const bgl = document.getElementById('bg-layer');
        if (bgl) {
            bgl.style.setProperty('--mx', (e.clientX / window.innerWidth * 100) + '%');
            bgl.style.setProperty('--my', (e.clientY / window.innerHeight * 100) + '%');
        }

        const hero = document.getElementById('hero');
        if (hero) {
            const heroRect = hero.getBoundingClientRect();
            const overInteractive = e.target.closest('a,button,textarea,input,.q-chip');
            const inHero = e.clientY >= heroRect.top && e.clientY <= heroRect.bottom && !overInteractive;
            
            if (inHero) {
                cursor.classList.add('label');
                cursor.classList.remove('ring');
                cursor.textContent = HERO_LABELS[heroLabelIdx];
            } else {
                cursor.classList.remove('label');
                cursor.textContent = '';
                // Only show ring for links and buttons, hide cursor entirely for inputs to show system I-beam
                const isBtn = e.target.closest('a,button,.q-chip');
                const isInput = e.target.closest('textarea,input');
                
                cursor.classList.toggle('ring', !!isBtn);
                cursor.style.opacity = isInput ? '0' : '1';
            }
        }
    });

    document.addEventListener('mouseleave', () => cursor.style.opacity = '0');
    document.addEventListener('mouseenter', () => cursor.style.opacity = '1');
}

// Hero button click target
const heroCta = document.getElementById('hero-cta');
if(heroCta) heroCta.addEventListener('click', e => {
    document.getElementById('generator').scrollIntoView({ behavior: 'smooth' });
});

// ── HEADER ────────────────────────────────────────────────────────
window.addEventListener('scroll', () => {
    document.getElementById('hdr').style.borderBottomColor = window.scrollY > 10 ? 'rgba(255,255,255,.12)' : 'rgba(255,255,255,.06)';
});

// ── COOKIE ────────────────────────────────────────────────────────
function initCookie() { if (!localStorage.getItem('sn_c')) document.getElementById('cookie-banner').classList.remove('hidden'); }
document.getElementById('ck-accept').onclick = () => { localStorage.setItem('sn_c','accepted'); document.getElementById('cookie-banner').classList.add('hidden'); };
document.getElementById('ck-reject').onclick = () => { localStorage.setItem('sn_c','rejected'); document.getElementById('cookie-banner').classList.add('hidden'); };

// ── MODAL ─────────────────────────────────────────────────────────
const modal = document.getElementById('privacy-modal');
function openModal(e) { e.preventDefault(); modal.classList.add('open'); }
document.getElementById('fp-legal').onclick = openModal;
document.querySelectorAll('.fp-legal-link').forEach(a => a.onclick = openModal);
document.getElementById('modal-close').onclick = () => modal.classList.remove('open');
modal.onclick = e => { if (e.target === modal) modal.classList.remove('open'); };

// ── VIEWS ─────────────────────────────────────────────────────────
function show(id) {
    ['wizard-view','confirm-view','loading-view','result-view'].forEach(v => document.getElementById(v).classList.add('hidden'));
    document.getElementById(id).classList.remove('hidden');
}

// ── WIZARD ────────────────────────────────────────────────────────
function renderStep(idx) {
    const q = Q[idx];
    document.getElementById('prog-wrap').classList.remove('hidden');
    document.getElementById('prog-count').textContent = `${q.num} / 12`;
    document.getElementById('prog-fill').style.width = `${((idx + 1) / 12) * 100}%`;
    const wv = document.getElementById('wizard-view');
    wv.classList.remove('hidden');
    wv.innerHTML = `
        <div class="m-grid">
            <div class="m-col-sidebar">
                <div class="m-side-meta">
                    <span class="q-label">${q.num} / 12</span>
                    <span class="q-section-name">${q.slide}</span>
                </div>
            </div>
            <div class="m-col-main">
                <p class="q-opener">${q.opener}</p>
                <h2 class="q-title" style="font-family:var(--font-serif); font-size:clamp(42px,5vw,72px); font-weight:400; font-style:italic; line-height:1.0; margin-bottom:48px;">${q.title}</h2>
                <p class="q-guidance" style="font-size:18px; line-height:1.7; color:var(--dim); margin-bottom:48px; max-width:640px;">${q.guidance}</p>
                <div class="chips">${q.chips.map(c => `<span class="q-chip" data-chip="${c.replace(/"/g,'&quot;')}">${c}</span>`).join('')}</div>
                <textarea id="q-ta" placeholder="${q.placeholder}" style="width:100%; height:200px; background:rgba(255,255,255,0.02); border:1px solid var(--border); border-radius:12px; padding:24px; color:var(--white); font-family:var(--font-sans); font-size:16px; line-height:1.6; resize:none; outline:none; transition:border-color .3s;"></textarea>
                <div class="char-meta" style="margin-top:16px; display:flex; justify-content:space-between; align-items:center; font-size:12px; color:var(--muted); font-weight:700; letter-spacing:0.05em;">
                    <button class="btn-improve" id="btn-improve" style="background:none; border:none; color:var(--red); text-transform:uppercase; letter-spacing:0.1em; font-weight:700;">✦ Refine this answer</button>
                    <span><span id="char-n">${(S.ans[q.id]||'').length}</span> characters</span>
                </div>
                <div class="wiz-nav" style="margin-top:64px; display:flex; gap:20px;">
                    <button class="btn-back" id="btn-back" style="visibility:${idx===0?'hidden':'visible'}; background:none; border:1px solid var(--border); color:var(--dim); padding:18px 32px; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:0.1em;">Back</button>
                    <button class="btn-next" id="btn-next" style="background:var(--white); color:var(--black); border:none; padding:18px 48px; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:0.1em; flex:1;">${idx===11?'Review narrative →':'Continue →'}</button>
                </div>
            </div>
        </div>`;
    const ta = document.getElementById('q-ta');
    wv.querySelectorAll('.q-chip').forEach(chip => {
        chip.onclick = () => {
            ta.value += (ta.value.trim() ? '\n\n' : '') + chip.dataset.chip + ': ';
            ta.focus(); ta.selectionStart = ta.selectionEnd = ta.value.length;
            S.ans[q.id] = ta.value; S.save();
            document.getElementById('char-n').textContent = ta.value.length;
        };
    });
    ta.addEventListener('input', () => { S.ans[q.id] = ta.value; S.save(); document.getElementById('char-n').textContent = ta.value.length; });
    document.getElementById('btn-improve').onclick = () => {
        const raw = ta.value.trim();
        const prompt = `${q.refine}\n\n"${raw || '[No answer yet — add your notes first]'}"`;
        openImprove(prompt);
    };
    document.getElementById('btn-next').onclick = () => idx < 11 ? renderStep(idx + 1) : showConfirm();
    document.getElementById('btn-back').onclick = () => renderStep(idx - 1);
    ta.focus();
}

function showConfirm() { show('confirm-view'); document.getElementById('prog-wrap').classList.add('hidden'); }

document.getElementById('ans-toggle').onclick = () => {
    const list = document.getElementById('ans-list');
    const open = list.classList.toggle('open');
    document.getElementById('ans-arrow').textContent = open ? '↑' : '↓';
    if (open) list.innerHTML = Q.map(q => `
        <div class="ans-row">
            <div class="ans-q">${q.num}. ${q.slide}</div>
            <div class="ans-a">${S.ans[q.id] || '<em style="color:rgba(255,255,255,.25)">Not answered</em>'}</div>
        </div>`).join('');
};

// ── GENERATE ──────────────────────────────────────────────────────
document.getElementById('btn-gen').onclick = async () => {
    const name = document.getElementById('proj-name').value.trim();
    if (!name) { document.getElementById('proj-name').style.borderColor = '#E30613'; document.getElementById('proj-name').focus(); return; }
    S.projectName = name; S.presenterName = document.getElementById('pres-name').value.trim();
    show('loading-view');
    let pi = 0; const pe = document.getElementById('load-phrase');
    const iv = setInterval(() => { pi = (pi + 1) % PHRASES.length; pe.textContent = PHRASES[pi]; }, 2600);
    try { S.deck = await apiGen(S); } finally { clearInterval(iv); }
    showResults(S.deck);
};

// ── API ───────────────────────────────────────────────────────────
async function apiGen(st) {
    const answersText = Q.map(q => `[${q.slide}]:\n${st.ans[q.id] || 'not answered'}`).join('\n\n');
    const prompt = `[INST] You are a senior product strategist at a top startup accelerator. Transform these raw project notes into a board-ready presentation deck.

Project: "${st.projectName}"${st.presenterName ? `, by ${st.presenterName}` : ''}

RAW ANSWERS:
${answersText}

CRITICAL RULES:
- Rewrite everything in confident, evidence-first executive language. Never parrot the input.
- Each slide has ONE core insight — a bold, specific headline that stands alone without the body.
- If data is missing, write [placeholder: what's needed here] — never leave a blank.
- No buzzwords: no 'synergy', 'leverage', 'disruptive', 'innovative', 'holistic'.
- Titles should be active and specific: not 'Our Solution' but 'From 4-Hour Reviews to 3-Minute Automated Reports'.
- Every claim backed by evidence or labeled as an estimate.

Return ONLY valid JSON (no markdown, no preamble):
{
  "deck_title": "string",
  "tagline": "One sharp line",
  "slides": [
    { "n": 1, "title": "string", "headline": "Bold specific headline 8-12 words", "insight": "One sentence insight", "points": ["point","point","point","point"], "stat": "Optional bold datapoint or callout" }
  ]
}

Generate exactly 15 slides. Arc: Cover → Problem Statement → Solution → Target Clients → Goal & Success Metrics → Scope Definition → Competitive Landscape → Progress & Findings → Team → Resources → Risk Assessment → Market Potential → Business Model → Executive Summary → Thank You & Contacts. [/INST]`;

    const HF = 'https://api-inference.huggingface.co/models/mistralai/Mistral-7B-Instruct-v0.2';
    const deadline = Date.now() + 150000;
    while (Date.now() < deadline) {
        try {
            document.getElementById('load-sub').textContent = 'Contacting AI...';
            const res = await fetch(HF, { method:'POST', headers:{'Content-Type':'application/json'},
                body: JSON.stringify({ inputs: prompt, parameters: { max_new_tokens: 3500, temperature: 0.62, return_full_text: false } }) });
            if (res.status === 503) { document.getElementById('load-sub').textContent = 'AI warming up — hold tight...'; await sleep(5000); continue; }
            if (!res.ok) throw new Error('HTTP ' + res.status);
            const data = await res.json();
            const text = data[0]?.generated_text || '';
            const j0 = text.indexOf('{'), j1 = text.lastIndexOf('}');
            if (j0 === -1) throw new Error('No JSON in response');
            return JSON.parse(text.substring(j0, j1 + 1));
        } catch(e) {
            console.warn('AI attempt failed:', e.message);
            document.getElementById('load-sub').textContent = 'Having trouble connecting — trying again...';
            if (Date.now() > deadline - 120000) break;
            await sleep(3500);
        }
    }
    return fallback(st);
}

function fallback(st) {
    const slides = Q.map((q, i) => ({
        n: i + 1, title: q.slide,
        headline: `[AI couldn't connect — use the prompt below to complete this slide]`,
        insight: st.ans[q.id] ? st.ans[q.id].substring(0, 180).trim() + (st.ans[q.id].length > 180 ? '...' : '') : '[No answer — return to questionnaire to add your notes]',
        points: [`Paste into Claude / ChatGPT to complete this slide:\n\n${q.refine}\n\n"${st.ans[q.id] || '[Add your answer here]'}"`],
        stat: '← Copy this prompt into any AI'
    }));
    slides.push({
        n: 13, title: 'Validate Your Assumptions',
        headline: '5 prompts to pressure-test your deck before presenting',
        insight: 'Copy any prompt below into Claude or ChatGPT to verify data, competition, and assumptions.',
        points: [
            `"What is the total addressable market for ${st.projectName} in [target geography]? Cite sources and note the year."`,
            `"List the top 5 competitors or alternatives in the [your market] space. Compare pricing, key features, and their main weaknesses."`,
            `"What are the most common failure modes for projects like ${st.projectName}? How should a team mitigate them?"`,
            `"What macro trends in 2025-2026 are driving demand for [your solution]? Include named datasets or reports."`,
            `"If a team had 6 months and €50k, what's the fastest way to validate the core assumption behind ${st.projectName}?"`
        ],
        stat: 'AI-verified data = credible pitch'
    });
    slides.push({
        n: 14, title: 'Executive Summary',
        headline: 'A one-page synthesis of the entire pitch',
        insight: 'Copy the prompt below to generate a tight executive summary of your answers.',
        points: [
            `"Based on all my previous answers, write a single-slide Executive Summary for ${st.projectName}."`,
            `"It should include: the core problem, our specific solution, the target market size, and our timeline or ask."`,
            `"Keep it under 100 words total, using 4 bullet points."`
        ],
        stat: 'Cut to the chase'
    });
    slides.push({
        n: 15, title: 'Thank You & Contacts',
        headline: 'Ready to build together?',
        insight: `Prepared by ${st.presenterName ? st.presenterName : 'the founding team'}. Reach out to discuss next steps.`,
        points: [
            '[Your Email Here]',
            '[Your Phone / LinkedIn Here]',
            '[Website or Demo Link Here]'
        ],
        stat: 'Let\'s talk'
    });
    return { deck_title: st.projectName, tagline: '[Your one-line pitch — what this project does and for whom]', slides };
}

// ── RESULTS ───────────────────────────────────────────────────────
function showResults(deck) {
    show('result-view');
    document.getElementById('slide-preview').innerHTML = deck.slides.map((s, i) => `
        <div class="slide-row" style="background:rgba(255,255,255,0.01); border:1px solid var(--border); border-radius:8px; margin-bottom:12px; padding:32px; display:flex; gap:32px; align-items:center;">
            <div class="slide-num" style="font-family:var(--font-serif); font-size:32px; color:var(--red); font-style:italic; border-right:1px solid var(--border); padding-right:32px; min-width:80px; text-align:center;">${String(s.n || i+1).padStart(2,'0')}</div>
            <div style="flex:1">
                <div class="slide-title" style="font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:0.2em; color:var(--muted); margin-bottom:8px;">${s.title}</div>
                <div class="slide-sub" style="font-family:var(--font-serif); font-size:24px; font-style:italic; line-height:1.2; color:var(--white);">${s.headline || s.insight || ''}</div>
            </div>
        </div>`).join('');
}

document.getElementById('btn-dl').onclick = () => S.deck ? buildPPTX(S.deck, S.projectName, S.presenterName) : null;
document.getElementById('btn-restart').onclick = () => { sessionStorage.removeItem('sn_a'); location.reload(); };

// ── PPTX — Swiss modular grid ─────────────────────────────────────
function buildPPTX(deck, projectName, presenterName) {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33 × 7.5 in
    const C = { black:'0A0A0A', white:'FFFFFF', red:'E30613', gray:'666666', lgray:'E8E8E8', mgray:'BBBBBB', offblack:'1A1A1A' };

    deck.slides.forEach((s, idx) => {
        const slide = pptx.addSlide();
        const num = s.n || idx + 1;
        const total = deck.slides.length;

        if (idx === 0) {
            // ── COVER ──────────────────────────────────────────────
            slide.background = { fill: C.black };
            slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.2, h:'100%', fill:C.red });
            slide.addText((deck.deck_title || projectName).toUpperCase(), {
                x:0.5, y:1.7, w:12, fontSize:56, bold:true, color:C.white, fontFace:'Helvetica Neue', charSpacing:-1 });
            slide.addShape(pptx.ShapeType.rect, { x:0.5, y:3.5, w:3.2, h:0.06, fill:C.red });
            slide.addText(deck.tagline || '[One-line pitch]', {
                x:0.5, y:3.7, w:9.5, fontSize:20, color:'999999', fontFace:'Helvetica Neue', italic:true });
            if (presenterName) slide.addText(presenterName.toUpperCase(), { x:0.5, y:5.7, w:6, fontSize:10, color:C.gray, fontFace:'Helvetica Neue', bold:true, charSpacing:2.5 });
            slide.addText(new Date().toLocaleDateString('en-GB',{month:'long',year:'numeric'}).toUpperCase(), { x:0.5, y:6.1, w:6, fontSize:9, color:'555555', fontFace:'Helvetica Neue' });
            slide.addText(`01 / ${total}`, { x:11.3, y:7.0, w:1.8, align:'right', fontSize:8, color:'444444', fontFace:'Helvetica Neue' });

        } else if (s.title === 'Validate Your Assumptions') {
            // ── VALIDATE SLIDE ─────────────────────────────────────
            slide.background = { fill: C.black };
            slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:0.2, fill:C.red });
            slide.addText('VALIDATE YOUR ASSUMPTIONS', { x:0.5, y:0.38, w:12, fontSize:10, bold:true, color:C.red, fontFace:'Helvetica Neue', charSpacing:3 });
            slide.addText(s.headline || s.title, { x:0.5, y:0.72, w:12, fontSize:32, bold:true, color:C.white, fontFace:'Helvetica Neue', charSpacing:-0.5 });
            slide.addText(s.insight || '', { x:0.5, y:1.55, w:12, fontSize:14, color:'888888', fontFace:'Helvetica Neue', italic:true });
            slide.addShape(pptx.ShapeType.rect, { x:0.5, y:2.0, w:12.3, h:0.008, fill:'2A2A2A' });
            (s.points||[]).slice(0,5).forEach((pt, pi) => {
                const y = 2.2 + pi * 0.95;
                slide.addShape(pptx.ShapeType.rect, { x:0.5, y, w:0.05, h:0.55, fill:C.red });
                slide.addText(`PROMPT ${pi+1}`, { x:0.7, y, w:1.5, fontSize:8, bold:true, color:C.red, fontFace:'Helvetica Neue', charSpacing:2 });
                slide.addText(pt, { x:0.7, y:y+0.2, w:12.1, fontSize:11.5, color:'CCCCCC', fontFace:'Helvetica Neue', lineSpacingMultiple:1.35 });
            });
            slide.addText(`${String(num).padStart(2,'0')} / ${total}`, { x:11.3, y:7.0, w:1.8, align:'right', fontSize:8, color:'444444' });

        } else {
            // ── CONTENT SLIDES — Swiss grid ────────────────────────
            slide.background = { fill: C.white };
            // Red rule top
            slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:'100%', h:0.15, fill:C.red });
            // Slide number
            slide.addText(String(num).padStart(2,'0'), { x:0.44, y:0.28, w:0.7, fontSize:10, bold:true, color:C.red, fontFace:'Helvetica Neue' });
            // Slide name label
            slide.addText((s.title||'').toUpperCase(), { x:1.1, y:0.28, w:10, fontSize:10, bold:true, color:'AAAAAA', fontFace:'Helvetica Neue', charSpacing:1.5 });
            // Main headline — large and specific
            slide.addText(s.headline || s.title, {
                x:0.44, y:0.62, w:12.3, fontSize:30, bold:true, color:C.black,
                fontFace:'Helvetica Neue', charSpacing:-0.4, lineSpacingMultiple:1.1 });
            // Horizontal rule
            slide.addShape(pptx.ShapeType.rect, { x:0.44, y:1.6, w:12.3, h:0.01, fill:C.lgray });
            // LEFT COLUMN — Core Insight
            slide.addText('CORE INSIGHT', { x:0.44, y:1.74, w:5.6, fontSize:8, bold:true, color:C.red, fontFace:'Helvetica Neue', charSpacing:2.5 });
            slide.addText(s.insight || '[Core insight for this slide]', {
                x:0.44, y:2.06, w:5.4, fontSize:17, color:C.black,
                fontFace:'Helvetica Neue', lineSpacingMultiple:1.55, italic:true });
            // Stat callout (bottom left), if present
            if (s.stat) {
                slide.addShape(pptx.ShapeType.rect, { x:0.44, y:5.0, w:5.4, h:1.28, fill:C.black });
                slide.addText(s.stat, { x:0.56, y:5.12, w:5.2, fontSize:20, bold:true, color:C.white, fontFace:'Helvetica Neue', lineSpacingMultiple:1.25 });
            }
            // Vertical divider
            slide.addShape(pptx.ShapeType.rect, { x:6.2, y:1.6, w:0.01, h:5.6, fill:C.lgray });
            // RIGHT COLUMN — Key Points
            slide.addText('KEY POINTS', { x:6.4, y:1.74, w:6.5, fontSize:8, bold:true, color:C.gray, fontFace:'Helvetica Neue', charSpacing:2.5 });
            (s.points||[]).slice(0,5).forEach((pt, pi) => {
                const y = 2.06 + pi * 1.06;
                slide.addShape(pptx.ShapeType.rect, { x:6.4, y:y+0.04, w:0.18, h:0.18, fill:C.red });
                slide.addText(pt || '[Key point]', {
                    x:6.74, y:y, w:6.3, fontSize:13.5, color:C.black,
                    fontFace:'Helvetica Neue', lineSpacingMultiple:1.45 });
            });
            // Footer rule
            slide.addShape(pptx.ShapeType.rect, { x:0, y:7.22, w:'100%', h:0.008, fill:C.lgray });
            slide.addText((projectName || deck.deck_title || '').toUpperCase(), { x:0.44, y:7.3, w:7, fontSize:7.5, color:C.mgray, fontFace:'Helvetica Neue', bold:true, charSpacing:2 });
            slide.addText(`${String(num).padStart(2,'0')} / ${total}`, { x:11.3, y:7.3, w:1.8, align:'right', fontSize:7.5, color:C.mgray });
        }
    });

    const safe = (deck.deck_title || 'deck').replace(/[^a-zA-Z0-9_-]/g, '_');
    pptx.writeFile({ fileName: `Sinaida_${safe}.pptx` }).catch(err => alert('Download error: ' + err.message));
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ── INIT ──────────────────────────────────────────────────────────
initCookie();
renderStep(0);
document.getElementById('prog-wrap').classList.remove('hidden');
