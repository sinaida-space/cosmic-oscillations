import PptxGenJS from 'pptxgenjs';

// ─── STATE ───────────────────────────────────────────────────────
const state = {
    answers: JSON.parse(sessionStorage.getItem('sn_answers') || '{}'),
    projectName: '', presenterName: '', deck: null,
    save() { sessionStorage.setItem('sn_answers', JSON.stringify(this.answers)); }
};

// ─── QUESTIONS (YC-advisor tone, 8+ chips each) ───────────────────
const QUESTIONS = [
    {
        id: 'problem', num: '01',
        opener: 'Every great project starts with something broken. Let's find yours.',
        title: 'What problem are you obsessed with?',
        guidance: 'Don\'t describe a feature. Describe the frustration. The best problems are ones people work around every day and have accepted as normal.',
        placeholder: 'e.g. Freelance designers spend 6h/week on invoicing. They hate it, but there\'s no tool that fits their workflow.',
        chips: ['Who wakes up frustrated by this?', 'What\'s the workaround people use today?', 'How expensive is this problem to ignore?', 'When did you first notice this gap?', 'What\'s getting worse, not better?', 'Is this a nice-to-have or a must-fix?', 'Why hasn\'t it been solved yet?', 'How do you know this is real — not just your problem?']
    },
    {
        id: 'solution', num: '02',
        opener: 'Now for your moment of clarity. What do you actually do?',
        title: 'What\'s your answer to the mess?',
        guidance: 'Resist the urge to list features. Describe the shift — the before and after. What becomes possible that wasn\'t before?',
        placeholder: 'e.g. We automate invoice creation from design files. One click. The designer never opens a spreadsheet again.',
        chips: ['The core mechanic in one sentence', 'Before vs. after for your user', 'What\'s the \'oh wow\' moment?', 'What would make it irreplaceable?', 'What does day 1 look like?', 'Is it a vitamin or a painkiller?', 'What makes it sticky over time?', 'What are you choosing NOT to do?']
    },
    {
        id: 'clients', num: '03',
        opener: 'Be specific. "Everyone" is not a customer.',
        title: 'Who is this actually for?',
        guidance: 'Describe the early adopter specifically — not a demographic, but a behavior. Who would pay for this the moment it exists?',
        placeholder: 'e.g. Independent designers who invoice >5 clients/month, use Figma, and currently use Excel or FreshBooks reluctantly.',
        chips: ['Early adopter in one sentence', 'What do they do Monday morning?', 'Where do they hang out online?', 'What do they complain about publicly?', 'What do they already pay for?', 'What are they NOT willing to change?', 'Who refers them to solutions today?', 'What would make them switch instantly?']
    },
    {
        id: 'goals', num: '04',
        opener: 'What does winning look like for you? Not in theory — in numbers.',
        title: 'What does success look like in 12 months?',
        guidance: 'Specific targets create accountability. Vague goals stall projects. What\'s the number you\'d be proud to share at a dinner table?',
        placeholder: 'e.g. 500 paying users, €5k MRR, 80% weekly active rate. Secondary: featured at one major design conference.',
        chips: ['Revenue target?', 'User / customer count?', 'Geography?', 'Retention goal?', 'What\'s the signal this is working?', 'What would make you pivot?', 'What\'s the 3-year version?', 'What does failure look like, and how will you catch it early?']
    },
    {
        id: 'scope', num: '05',
        opener: 'Boundaries are not limitations — they\'re your strategy made visible.',
        title: 'What are you building first, and what are you explicitly leaving out?',
        guidance: 'The fastest projects ship a narrow, opinionated version first. What\'s in v1? What goes on the "not yet" list?',
        placeholder: 'e.g. V1: invoice generation only, PDF export. NOT in scope: payment collection, CRM, team accounts.',
        chips: ['V1 in 3 bullets', 'What\'s deliberately excluded?', 'Time to first usable product?', 'What could derail scope?', 'What are you saying no to (and why)?', 'What\'s the next natural expansion?', 'Geographic or language constraints?', 'Regulatory or compliance limits?']
    },
    {
        id: 'competitors', num: '06',
        opener: 'If you think you have no competitors, you haven\'t looked hard enough.',
        title: 'Who else is playing in this space — and why isn\'t that good enough?',
        guidance: 'Name real alternatives honestly. Then name your unfair advantage. If it\'s just "better UX", go deeper.',
        placeholder: 'e.g. FreshBooks, HoneyBook, plain Excel. Their weakness: built for accountants, not designers. Our edge: integrates directly with Figma workflow.',
        chips: ['Direct competitors?', 'What do people use today instead?', 'What does the incumbent do well?', 'Where does the incumbent fail your user?', 'Your unfair advantage (be ruthlessly honest)', 'What would a competitor copy first?', 'What can\'t they copy?', 'Market trend working in your favor?']
    },
    {
        id: 'progress', num: '07',
        opener: 'Show me you\'ve left the building. Momentum matters more than plans.',
        title: 'What have you built or learned already?',
        guidance: 'Early signal — even a conversation, a prototype, a rejection — proves you\'re learning. What do you know now that you didn\'t 3 months ago?',
        placeholder: 'e.g. 12 user interviews. Built a Figma plugin prototype (230 installs). 3 designers paying €29/mo for early access.',
        chips: ['Users spoken to?', 'Prototype or MVP status?', 'First paying customer?', 'Key learning that changed your thinking?', 'What surprised you?', 'What didn\'t work?', 'Partners or advisors involved?', 'Press, community, or organic interest?']
    },
    {
        id: 'team', num: '08',
        opener: 'Ideas are common. Execution is rare. Who\'s doing the work?',
        title: 'Who is building this, and why are you the right team?',
        guidance: 'The honest version. Not a LinkedIn bio — what makes this team uniquely equipped for this specific problem?',
        placeholder: 'e.g. Two co-founders: one ex-Figma designer (10y exp), one Rails engineer (shipped 4 products). We met building freelance tools for ourselves.',
        chips: ['Core team in 2 lines', 'What is each person\'s superpower?', 'Why this problem — what\'s personal about it?', 'What\'s missing from the team today?', 'Advisors or mentors?', 'Full-time or part-time?', 'History working together?', 'Biggest team risk?']
    },
    {
        id: 'resources', num: '09',
        opener: 'Money is a resource. So is time, relationships, and IP. What do you have?',
        title: 'What\'s your current position, and what do you need most?',
        guidance: 'Be honest about your runway and your top 3 constraints. Investors and partners respect clarity over optimism.',
        placeholder: 'e.g. €40k personal savings. 6 months runway. Need: €150k seed to hire one more engineer and run paid pilots.',
        chips: ['Current funding / runway?', 'Revenue (if any)?', 'Top 3 resource needs?', 'What\'s the most expensive assumption?', 'Strategic relationships you have?', 'Intellectual property / patents?', 'What would change with 10× the resources?', 'First priority if funded tomorrow?']
    },
    {
        id: 'risks', num: '10',
        opener: 'Every project has a version where it fails. Let\'s find yours first.',
        title: 'What could kill this — and how are you thinking about it?',
        guidance: 'Naming risks shows maturity. For each one, have a mitigation. Investors fund teams who see around corners.',
        placeholder: 'e.g. Risk: Figma changes their API. Mitigation: building export-agnostic core. Risk: designer market too small. Mitigation: expand to all creative freelancers.',
        chips: ['Biggest single risk?', 'Technical risks?', 'Market timing risk?', 'Regulatory risk?', 'Dependency risk (platforms, APIs)?', 'Team / execution risk?', 'What\'s your early warning signal?', 'What would make you stop and pivot?']
    },
    {
        id: 'market', num: '11',
        opener: 'Size matters — but so does the shape of the opportunity.',
        title: 'How big is this, and why is now the right moment?',
        guidance: 'Bottom-up estimates are more credible than TAM slides. Start with: how many of your target user exist, and what would they pay? Then multiply.',
        placeholder: 'e.g. 4M independent designers globally. 20% invoice regularly. At €29/mo, SAM = $278M. Our initial target: EU, 50k designers.',
        chips: ['How many target users exist?', 'What do they currently spend on this problem?', 'What\'s your realistic market share in 3 years?', 'Is the market growing or shrinking?', 'What macro trend is pushing this?', 'Adjacent markets you could expand into?', 'Who else is investing in this space?', 'Why is now better than 2 years ago?']
    },
    {
        id: 'model', num: '12',
        opener: 'The last question is the first one investors ask. Let\'s make it count.',
        title: 'How does this make money — and when does it start?',
        guidance: 'Simple is better. The more complex your revenue model sounds, the less you\'ve tested it. What does the unit economics look like?',
        placeholder: 'e.g. SaaS. €29/mo individual, €99/mo team. First revenue: month 1 from beta users. 12-month target: €5k MRR.',
        chips: ['Pricing model?', 'Revenue streams?', 'When does first revenue arrive?', 'Customer acquisition cost (estimated)?', 'Lifetime value (estimated)?', 'Payback period?', 'Freemium or paid-first?', 'What pricing have you tested?']
    }
];

const PHRASES = [
    'Reading between the lines...', 'Finding the thread...', 'Shaping your narrative...',
    'Trimming the noise...', 'Building the arc...', 'Almost there — polishing...'
];

// ─── CURSOR ──────────────────────────────────────────────────────
const cursor = document.getElementById('cursor');
document.addEventListener('mousemove', e => {
    cursor.style.left = e.clientX + 'px';
    cursor.style.top = e.clientY + 'px';
    document.getElementById('bg-layer').style.setProperty('--mx', (e.clientX / window.innerWidth * 100) + '%');
    document.getElementById('bg-layer').style.setProperty('--my', (e.clientY / window.innerHeight * 100) + '%');
    const t = e.target;
    cursor.classList.toggle('ring', !!(t.closest('a,button') && !t.closest('textarea,input')));
});
document.addEventListener('mouseleave', () => cursor.style.opacity = '0');
document.addEventListener('mouseenter', () => cursor.style.opacity = '1');

// ─── HEADER ──────────────────────────────────────────────────────
window.addEventListener('scroll', () => {
    document.getElementById('hdr').style.borderBottomColor = window.scrollY > 10 ? 'rgba(255,255,255,0.12)' : 'rgba(255,255,255,0.06)';
});
document.getElementById('hdr-cta').onclick = () => document.getElementById('generator').scrollIntoView({ behavior: 'smooth' });
document.getElementById('hero-cta').onclick = e => { e.preventDefault(); document.getElementById('generator').scrollIntoView({ behavior: 'smooth' }); };

// ─── COOKIE ──────────────────────────────────────────────────────
function initCookie() {
    if (!localStorage.getItem('sn_consent')) document.getElementById('cookie-banner').classList.remove('hidden');
}
document.getElementById('ck-accept').onclick = () => { localStorage.setItem('sn_consent', 'accepted'); document.getElementById('cookie-banner').classList.add('hidden'); };
document.getElementById('ck-reject').onclick = () => { localStorage.setItem('sn_consent', 'rejected'); document.getElementById('cookie-banner').classList.add('hidden'); };

// ─── MODALS ──────────────────────────────────────────────────────
const privModal = document.getElementById('privacy-modal');
function openPrivacy(e) { e.preventDefault(); privModal.classList.add('open'); }
document.getElementById('fp-privacy').onclick = openPrivacy;
document.getElementById('fp-terms').onclick = openPrivacy;
document.querySelectorAll('.fp-privacy-link,.fp-terms-link').forEach(a => a.onclick = openPrivacy);
document.getElementById('modal-close').onclick = () => privModal.classList.remove('open');
privModal.onclick = e => { if (e.target === privModal) privModal.classList.remove('open'); };

// ─── VIEWS ───────────────────────────────────────────────────────
function show(id) {
    ['wizard-view', 'confirm-view', 'loading-view', 'result-view'].forEach(v => document.getElementById(v).classList.add('hidden'));
    document.getElementById(id).classList.remove('hidden');
}

// ─── WIZARD ──────────────────────────────────────────────────────
function renderStep(idx) {
    const q = QUESTIONS[idx];
    document.getElementById('prog-wrap').classList.remove('hidden');
    document.getElementById('prog-count').textContent = `${q.num} / 12`;
    document.getElementById('prog-fill').style.width = `${((idx + 1) / 12) * 100}%`;

    const wv = document.getElementById('wizard-view');
    wv.classList.remove('hidden');
    wv.innerHTML = `
        <span class="q-label">${q.num} / 12</span>
        <p class="q-opener">${q.opener}</p>
        <h2 class="q-title">${q.title}</h2>
        <p class="q-guidance">${q.guidance}</p>
        <div class="chips">${q.chips.map(c => `<span class="q-chip" data-chip="${c.replace(/"/g, '&quot;')}">${c}</span>`).join('')}</div>
        <textarea id="q-ta" placeholder="${q.placeholder}">${state.answers[q.id] || ''}</textarea>
        <div class="char-meta"><span id="char-n">${(state.answers[q.id] || '').length}</span> chars</div>
        <div class="wiz-nav">
            <button class="btn-back" id="btn-back" style="visibility:${idx === 0 ? 'hidden' : 'visible'}">← Back</button>
            <button class="btn-next" id="btn-next">${idx === 11 ? 'Review →' : 'Continue →'}</button>
        </div>`;

    const ta = document.getElementById('q-ta');
    wv.querySelectorAll('.q-chip').forEach(chip => {
        chip.onclick = () => {
            ta.value += (ta.value.trim() ? '\n\n' : '') + chip.dataset.chip + ': ';
            ta.focus(); ta.selectionStart = ta.selectionEnd = ta.value.length;
            state.answers[q.id] = ta.value; state.save();
            document.getElementById('char-n').textContent = ta.value.length;
        };
    });
    ta.addEventListener('input', () => {
        state.answers[q.id] = ta.value; state.save();
        document.getElementById('char-n').textContent = ta.value.length;
    });
    document.getElementById('btn-next').onclick = () => idx < 11 ? renderStep(idx + 1) : showConfirm();
    document.getElementById('btn-back').onclick = () => renderStep(idx - 1);
    ta.focus();
}

function showConfirm() {
    show('confirm-view');
    document.getElementById('prog-wrap').classList.add('hidden');
}

document.getElementById('ans-toggle').onclick = () => {
    const list = document.getElementById('ans-list');
    const open = list.classList.toggle('open');
    document.getElementById('ans-arrow').textContent = open ? '↑' : '↓';
    if (open) list.innerHTML = QUESTIONS.map(q => `
        <div class="ans-row">
            <div class="ans-q">${q.num}. ${q.title}</div>
            <div class="ans-a">${state.answers[q.id] || '<em style="color:rgba(255,255,255,.3)">Not answered</em>'}</div>
        </div>`).join('');
};

// ─── GENERATE ────────────────────────────────────────────────────
document.getElementById('btn-gen').onclick = async () => {
    const name = document.getElementById('proj-name').value.trim();
    if (!name) { document.getElementById('proj-name').focus(); document.getElementById('proj-name').style.borderColor = '#E30613'; return; }
    state.projectName = name;
    state.presenterName = document.getElementById('pres-name').value.trim();
    show('loading-view');

    let pi = 0;
    const pe = document.getElementById('load-phrase');
    const interval = setInterval(() => { pi = (pi + 1) % PHRASES.length; pe.textContent = PHRASES[pi]; }, 2600);

    try { state.deck = await generate(state); } finally { clearInterval(interval); }
    showResults(state.deck);
};

// ─── AI ──────────────────────────────────────────────────────────
async function generate(st) {
    const answersText = QUESTIONS.map(q => `[${q.title}]:\n${st.answers[q.id] || 'not provided'}`).join('\n\n');
    const prompt = `[INST] You are a senior partner at a top startup accelerator. You've reviewed thousands of pitches.

Transform these raw project notes into a board-ready presentation. Project: "${st.projectName}"${st.presenterName ? `, by ${st.presenterName}` : ''}.

RAW INPUT:
${answersText}

CRITICAL RULES:
- Rewrite everything in confident, executive language. Never parrot the input.
- Each slide has ONE core insight — not a list of facts.
- If data is missing, output [placeholder text in brackets].
- Tone: clear, human, intelligent. No buzzwords, no filler.
- The last slide (13) should be titled "Validate Your Assumptions" and contain 5 specific prompts the user can paste into ChatGPT or Claude to verify market data, competitive analysis, and assumptions.

Return ONLY valid JSON:
{
  "deck_title": "string",
  "tagline": "One sharp line",
  "slides": [
    { "n": 1, "title": "string", "insight": "One sentence insight", "points": ["point","point","point"], "stat": "Optional bold number or fact" }
  ]
}

Generate exactly 13 slides. Arc: Cover → Problem → Solution → Who It's For → Goals → Scope → Competitive Landscape → Traction → Team → Resources & Needs → Risks & Mitigations → Market Opportunity → Validate Your Assumptions. [/INST]`;

    const HF = 'https://api-inference.huggingface.co/models/mistralai/Mistral-7B-Instruct-v0.2';
    const deadline = Date.now() + 150000;
    while (Date.now() < deadline) {
        try {
            document.getElementById('load-sub').textContent = 'Contacting AI...';
            const res = await fetch(HF, { method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ inputs: prompt, parameters: { max_new_tokens: 3500, temperature: 0.62, return_full_text: false } }) });
            if (res.status === 503) { document.getElementById('load-sub').textContent = 'AI warming up, hold tight...'; await sleep(5000); continue; }
            if (!res.ok) throw new Error('HTTP ' + res.status);
            const data = await res.json();
            const text = data[0]?.generated_text || '';
            const j0 = text.indexOf('{'), j1 = text.lastIndexOf('}');
            if (j0 === -1) throw new Error('No JSON');
            return JSON.parse(text.substring(j0, j1 + 1));
        } catch (e) {
            console.warn('AI attempt failed:', e.message);
            document.getElementById('load-sub').textContent = 'Retrying...';
            if (Date.now() > deadline - 120000) break;
            await sleep(3500);
        }
    }
    return fallback(st);
}

function fallback(st) {
    const slides = QUESTIONS.map((q, i) => ({
        n: i + 1, title: q.title,
        insight: '[AI will refine this — add your key insight here]',
        points: (st.answers[q.id] || '').split(/[.\n]/).map(s => s.trim()).filter(s => s.length > 4).slice(0, 4),
        stat: ''
    }));
    slides.push({
        n: 13, title: 'Validate Your Assumptions',
        insight: 'Use these prompts in ChatGPT or Claude to pressure-test your deck before your next meeting.',
        points: [
            `"What is the total addressable market for [${st.projectName}] in [target geography]? Provide sources."`,
            `"List the top 5 competitors in the [your market] space and compare their pricing, features, and weaknesses."`,
            `"What are the most common reasons startups in [your sector] fail? How should ${st.projectName} mitigate these?"`,
            `"What macro trends are driving demand for [your solution] in [current year]? Cite data."`,
            `"If I had 6 months and €50k, what's the fastest way to validate the core assumption of this project: [your core assumption]?"`
        ],
        stat: 'AI-verified data = credible pitch'
    });
    return { deck_title: st.projectName, tagline: '[Your one-line pitch]', slides };
}

// ─── RESULTS ─────────────────────────────────────────────────────
function showResults(deck) {
    show('result-view');
    const labels = ['Cover','Problem','Solution','Audience','Goals','Scope','Competition','Traction','Team','Resources','Risks','Market','Validate'];
    document.getElementById('slide-preview').innerHTML = deck.slides.map((s, i) => `
        <div class="slide-row">
            <div class="slide-num">${String(s.n || i+1).padStart(2,'0')}</div>
            <div>
                <div class="slide-title">${s.title || labels[i] || 'Slide ' + (i+1)}</div>
                <div class="slide-sub">${s.insight || ''}</div>
            </div>
        </div>`).join('');
}

document.getElementById('btn-dl').onclick = () => state.deck ? buildPPTX(state.deck, state.projectName, state.presenterName) : null;
document.getElementById('btn-restart').onclick = () => { sessionStorage.removeItem('sn_answers'); location.reload(); };

// ─── PPTX (Swiss modular grid) ───────────────────────────────────
function buildPPTX(deck, projectName, presenterName) {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33 × 7.5 inches

    const C = { black: '0A0A0A', white: 'FFFFFF', red: 'E30613', gray: '666666', lgray: 'EBEBEB', mgray: 'CCCCCC', offwhite: 'F7F7F7' };
    const F = { bold: 'Helvetica Neue', reg: 'Helvetica Neue' };

    deck.slides.forEach((s, idx) => {
        const slide = pptx.addSlide();

        if (idx === 0) {
            // ── COVER ─────────── full black, large type
            slide.background = { fill: C.black };
            // Left red column
            slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: '100%', fill: C.red });
            // Grid lines (subtle)
            slide.addShape(pptx.ShapeType.rect, { x: 0.18, y: 3.5, w: '100%', h: 0.008, fill: '222222' });
            // Project name massive
            slide.addText((deck.deck_title || projectName).toUpperCase(), {
                x: 0.48, y: 1.6, w: 11.5, h: 1.8,
                fontSize: 64, bold: true, color: C.white, fontFace: F.bold, charSpacing: -1.5
            });
            // Tagline
            slide.addText(deck.tagline || '[Your one-line pitch]', {
                x: 0.48, y: 3.65, w: 9, fontSize: 20, color: '999999', fontFace: F.reg, italic: true
            });
            // Red accent line
            slide.addShape(pptx.ShapeType.rect, { x: 0.48, y: 3.55, w: 2.4, h: 0.06, fill: C.red });
            // Meta info
            if (presenterName) slide.addText(presenterName.toUpperCase(), { x: 0.48, y: 5.8, w: 5, fontSize: 11, color: C.gray, fontFace: F.bold, charSpacing: 2 });
            const dateStr = new Date().toLocaleDateString('en-GB', { month: 'long', year: 'numeric' }).toUpperCase();
            slide.addText(dateStr, { x: 0.48, y: 6.15, w: 5, fontSize: 10, color: '555555', fontFace: F.reg });
            // Slide number
            slide.addText('01 / ' + deck.slides.length, { x: 11.5, y: 6.9, w: 1.6, align: 'right', fontSize: 9, color: '444444', fontFace: F.reg });

        } else if (idx === deck.slides.length - 1) {
            // ── VALIDATE SLIDE ── different treatment
            slide.background = { fill: C.black };
            slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.22, fill: C.red });
            slide.addText('VALIDATE YOUR ASSUMPTIONS', { x: 0.5, y: 0.48, w: 12, fontSize: 11, bold: true, color: C.red, fontFace: F.bold, charSpacing: 3 });
            slide.addText(s.title || 'Validate', { x: 0.5, y: 0.82, w: 12, fontSize: 34, bold: true, color: C.white, fontFace: F.bold, charSpacing: -0.5 });
            slide.addText(s.insight || 'Use these prompts to verify your key assumptions.', { x: 0.5, y: 1.6, w: 12, fontSize: 15, color: '999999', fontFace: F.reg, italic: true });
            slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.05, w: 12, h: 0.008, fill: '2A2A2A' });
            const pts = (s.points || []).slice(0, 5);
            pts.forEach((pt, pi) => {
                const y = 2.25 + pi * 0.96;
                slide.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 0.04, h: 0.58, fill: C.red });
                // Prompt label
                slide.addText(`PROMPT ${pi + 1}`, { x: 0.7, y, w: 1.4, fontSize: 8, bold: true, color: C.red, fontFace: F.bold, charSpacing: 2 });
                slide.addText(pt, { x: 0.7, y: y + 0.18, w: 12, fontSize: 12, color: 'CCCCCC', fontFace: F.reg, lineSpacingMultiple: 1.3 });
            });
            slide.addText(`${String(idx + 1).padStart(2,'0')} / ${deck.slides.length}`, { x: 11.5, y: 6.9, w: 1.6, align: 'right', fontSize: 9, color: '444444' });

        } else {
            // ── CONTENT SLIDES ── Swiss modular grid
            slide.background = { fill: C.white };

            // Top red rule
            slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.16, fill: C.red });

            // Slide number (top left, small)
            slide.addText(String(idx + 1).padStart(2, '0'), { x: 0.45, y: 0.3, w: 0.8, fontSize: 11, bold: true, color: C.red, fontFace: F.bold });

            // Slide title (large, dark)
            const title = (s.title || '').toUpperCase();
            slide.addText(title, {
                x: 0.45, y: 0.58, w: 8.5, fontSize: 32, bold: true, color: C.black,
                fontFace: F.bold, charSpacing: -0.5, lineSpacingMultiple: 1.1
            });

            // Horizontal divider
            slide.addShape(pptx.ShapeType.rect, { x: 0.45, y: 1.58, w: 12.3, h: 0.012, fill: C.lgray });

            // INSIGHT (left column — prominent)
            slide.addText('CORE INSIGHT', { x: 0.45, y: 1.76, w: 5.6, fontSize: 8, bold: true, color: C.red, fontFace: F.bold, charSpacing: 2.5 });
            slide.addText(s.insight || '[Key insight for this slide]', {
                x: 0.45, y: 2.05, w: 5.4, fontSize: 18, color: C.black,
                fontFace: F.reg, lineSpacingMultiple: 1.5, italic: true
            });

            // Stat / callout (left, bottom)
            if (s.stat) {
                slide.addShape(pptx.ShapeType.rect, { x: 0.45, y: 4.8, w: 5.4, h: 1.3, fill: C.black });
                slide.addText(s.stat, { x: 0.55, y: 4.95, w: 5.2, fontSize: 22, bold: true, color: C.white, fontFace: F.bold, lineSpacingMultiple: 1.2 });
            }

            // Vertical column divider
            slide.addShape(pptx.ShapeType.rect, { x: 6.15, y: 1.58, w: 0.012, h: 5.5, fill: C.lgray });

            // KEY POINTS (right column)
            slide.addText('KEY POINTS', { x: 6.35, y: 1.76, w: 6.5, fontSize: 8, bold: true, color: C.gray, fontFace: F.bold, charSpacing: 2.5 });
            const pts = (s.points || []).slice(0, 5);
            pts.forEach((pt, pi) => {
                const y = 2.05 + pi * 1.08;
                slide.addShape(pptx.ShapeType.rect, { x: 6.35, y, w: 0.2, h: 0.2, fill: C.red });
                slide.addText(pt || '[Key point]', {
                    x: 6.72, y: y - 0.05, w: 6.3, fontSize: 14, color: C.black,
                    fontFace: F.reg, lineSpacingMultiple: 1.4
                });
            });

            // Footer rule + project name
            slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.22, w: '100%', h: 0.008, fill: C.lgray });
            slide.addText((projectName || deck.deck_title).toUpperCase(), { x: 0.45, y: 7.3, w: 6, fontSize: 8, color: C.mgray, fontFace: F.bold, charSpacing: 2 });
            slide.addText(`${String(idx + 1).padStart(2,'0')} / ${deck.slides.length}`, { x: 11.5, y: 7.3, w: 1.6, align: 'right', fontSize: 8, color: C.mgray });
        }
    });

    const safe = (deck.deck_title || 'deck').replace(/[^a-zA-Z0-9_-]/g, '_');
    pptx.writeFile({ fileName: `Sinaida_${safe}.pptx` }).catch(err => alert('Download error: ' + err.message));
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ─── INIT ─────────────────────────────────────────────────────────
initCookie();
renderStep(0);
document.getElementById('prog-wrap').classList.remove('hidden');
