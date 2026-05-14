"""
Minimal IEEE paper humanizer.
Strategy: REMOVE AI patterns, don't ADD humanizer patterns.
Only contractions + filler removal + vocabulary simplification + unicode cleanup.
"""
import sys, io, random, re as _re, unicodedata as _ud
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

random.seed(42)

SRC = 'paper_ieee.docx'
DST = 'paper_ieee_rewritten.docx'
doc = Document(SRC)
print(f"Loaded {SRC}: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")

SKIP_INDICES = {0, 1, 2, 3, 5}
REFERENCES_START = 438

def is_section_header(txt):
    return bool(_re.match(r'^(I{1,3}V?|VI{0,3}|IX|X|XI{0,3})\.\s', txt)) or bool(_re.match(r'^\d+\.\d+', txt))

def is_algorithm_line(txt):
    return bool(_re.match(r'^\d{1,2}[\.\)]\s', txt)) or txt.startswith("Input:") or txt.startswith("Output:")

def is_formula(txt):
    if len(txt) > 200:
        return False
    special = ['\u2211','\u220f','\u222b','\u2208','\u2200','\u2203','\u2264','\u2265','\u2260','\u221e','\u2202','\u2207','\u00d7','\u2297','\u2299','\u221d']
    if sum(1 for c in txt if c in special) >= 2:
        return True
    words = txt.split()
    if len(words) < 8 and _re.search(r'[=\u2264\u2265\u2208\u2200\u2203\u2211\u220f\u222b\u2202\u2207\u2297\u2299\u2248\u221d]', txt):
        return True
    if txt.count('=') >= 2 and len(words) < 12:
        return True
    return False

def should_skip(idx, txt):
    if idx in SKIP_INDICES:
        return True
    if idx >= REFERENCES_START:
        return True
    if not txt or len(txt.strip()) < 5:
        return True
    if is_section_header(txt):
        return True
    if is_algorithm_line(txt):
        return True
    if is_formula(txt):
        return True
    return False

# =========================================================
# STEP 1: CONTRACTIONS — most natural human signal
# =========================================================
CONTRACTIONS = [
    (r"\bdoes not\b", "doesn't"), (r"\bDoes not\b", "Doesn't"),
    (r"\bdo not\b", "don't"), (r"\bDo not\b", "Don't"),
    (r"\bdid not\b", "didn't"), (r"\bDid not\b", "Didn't"),
    (r"\bis not\b", "isn't"), (r"\bIs not\b", "Isn't"),
    (r"\bare not\b", "aren't"), (r"\bAre not\b", "Aren't"),
    (r"\bwas not\b", "wasn't"), (r"\bWas not\b", "Wasn't"),
    (r"\bwere not\b", "weren't"), (r"\bWere not\b", "Weren't"),
    (r"\bcannot\b", "can't"), (r"\bCannot\b", "Can't"),
    (r"\bcould not\b", "couldn't"), (r"\bCould not\b", "Couldn't"),
    (r"\bwould not\b", "wouldn't"), (r"\bWould not\b", "Wouldn't"),
    (r"\bwill not\b", "won't"), (r"\bWill not\b", "Won't"),
    (r"\bshould not\b", "shouldn't"), (r"\bShould not\b", "Shouldn't"),
    (r"\bit is\b", "it's"), (r"\bIt is\b", "It's"),
    (r"\bthat is\b", "that's"), (r"\bThat is\b", "That's"),
    (r"\bthere is\b", "there's"), (r"\bThere is\b", "There's"),
    (r"\bwhat is\b", "what's"), (r"\bWhat is\b", "What's"),
    (r"\bthey are\b", "they're"), (r"\bThey are\b", "They're"),
    (r"\bwe are\b", "we're"), (r"\bWe are\b", "We're"),
    (r"\byou are\b", "you're"), (r"\bYou are\b", "You're"),
    (r"\bI am\b", "I'm"), (r"\bI have\b", "I've"),
    (r"\bI would\b", "I'd"), (r"\bI will\b", "I'll"),
    (r"\bwe have\b", "we've"), (r"\bWe have\b", "We've"),
    (r"\bthey have\b", "they've"), (r"\bThey have\b", "They've"),
    (r"\bwho is\b", "who's"), (r"\bWho is\b", "Who's"),
    (r"\bwhere is\b", "where's"), (r"\bWhere is\b", "Where's"),
    (r"\bhe is\b", "he's"), (r"\bHe is\b", "He's"),
    (r"\bshe is\b", "she's"), (r"\bShe is\b", "She's"),
    (r"\bhave not\b", "haven't"), (r"\bHave not\b", "Haven't"),
    (r"\bhas not\b", "hasn't"), (r"\bHas not\b", "Hasn't"),
    (r"\blet us\b", "let's"), (r"\bLet us\b", "Let's"),
]

# =========================================================
# STEP 2: AI FILLER PHRASE REMOVAL — removes AI fingerprints
# =========================================================
FILLER_RULES = [
    (r"(?i)\bIt is important to note that\s*", ""),
    (r"(?i)\bIt should be noted that\s*", ""),
    (r"(?i)\bIt is worth mentioning that\s*", ""),
    (r"(?i)\bIt is worth noting that\s*", ""),
    (r"(?i)\bIt goes without saying that?\s*", ""),
    (r"(?i)\bNeedless to say,?\s*", ""),
    (r"(?i)\bin order to\b", "to"),
    (r"(?i)\bdue to the fact that\b", "because"),
    (r"(?i)\ba wide range of\b", "many"),
    (r"(?i)\ba plethora of\b", "many"),
    (r"(?i)\bin today'?s rapidly evolving\b", "in today's"),
    (r"(?i)\bin the ever-changing landscape of\b", "in"),
    (r"(?i)\bat the present time\b", "now"),
    (r"(?i)\bat this point in time\b", "now"),
    (r"(?i)\bin the realm of\b", "in"),
    (r"(?i)\bfor the purpose of\b", "for"),
    (r"(?i)\bin the event that\b", "if"),
    (r"(?i)\bprior to\b", "before"),
    (r"(?i)\bsubsequent to\b", "after"),
    (r"(?i)\bin conjunction with\b", "with"),
    (r"(?i)\bon the basis of\b", "based on"),
    (r"(?i)\bin the absence of\b", "without"),
    (r"(?i)\bhas the potential to\b", "can"),
    (r"(?i)\bis capable of\b", "can"),
    (r"(?i)\ba considerable amount of\b", "a lot of"),
    (r"(?i)\bthe vast majority of\b", "most"),
    (r"(?i)\bin close proximity to\b", "near"),
    (r"(?i)\bat this juncture\b", "now"),
    (r"(?i)\bgiven the fact that\b", "since"),
    (r"(?i)\bin light of\b", "given"),
]

# =========================================================
# STEP 3: VOCABULARY SIMPLIFICATION — 1:1 formal->simple
# These are deterministic, not random, so they don't create
# detectable patterns of inconsistency.
# =========================================================
VOCAB_RULES = [
    (r"(?i)\butilizes\b", "uses"), (r"(?i)\butilize\b", "use"),
    (r"(?i)\butilized\b", "used"), (r"(?i)\butilizing\b", "using"),
    (r"(?i)\bfacilitates\b", "helps"), (r"(?i)\bfacilitate\b", "help"),
    (r"(?i)\bfacilitated\b", "helped"), (r"(?i)\bfacilitating\b", "helping"),
    (r"(?i)\bdemonstrates\b", "shows"), (r"(?i)\bdemonstrate\b", "show"),
    (r"(?i)\bdemonstrated\b", "showed"), (r"(?i)\bdemonstrating\b", "showing"),
    (r"(?i)\benhances\b", "improves"), (r"(?i)\benhance\b", "improve"),
    (r"(?i)\benhanced\b", "improved"), (r"(?i)\benhancing\b", "improving"),
    (r"(?i)\bleverages\b", "uses"), (r"(?i)\bleverage\b", "use"),
    (r"(?i)\bleveraging\b", "using"), (r"(?i)\bleveraged\b", "used"),
    (r"(?i)\bensures\b", "makes sure"), (r"(?i)\bensure\b", "make sure"),
    (r"(?i)\bensuring\b", "making sure"),
    (r"(?i)\bstreamlines\b", "simplifies"), (r"(?i)\bstreamline\b", "simplify"),
    (r"(?i)\bstreamlining\b", "simplifying"),
    (r"(?i)\bunderscores\b", "highlights"), (r"(?i)\bunderscore\b", "highlight"),
    (r"(?i)\bunderscoring\b", "highlighting"),
    (r"(?i)\baudgments\b", "boosts"), (r"(?i)\baugment\b", "boost"),
    (r"(?i)\baugmenting\b", "boosting"),
    (r"(?i)\bmitigates\b", "reduces"), (r"(?i)\bmitigate\b", "reduce"),
    (r"(?i)\bmitigating\b", "reducing"), (r"(?i)\bmitigation\b", "reduction"),
    (r"(?i)\bcommences\b", "starts"), (r"(?i)\bcommence\b", "start"),
    (r"(?i)\bcommenced\b", "started"), (r"(?i)\bcommencing\b", "starting"),
    (r"(?i)\bterminates\b", "ends"), (r"(?i)\bterminate\b", "end"),
    (r"(?i)\bterminated\b", "ended"),
    (r"(?i)\bexpedites\b", "speeds up"), (r"(?i)\bexpedite\b", "speed up"),
    (r"(?i)\bcorroborates\b", "backs up"), (r"(?i)\bcorroborate\b", "back up"),
    (r"(?i)\belucidated\b", "explained"), (r"(?i)\belucidate\b", "explain"),
    (r"(?i)\bdelineates\b", "outlines"), (r"(?i)\bdelineate\b", "outline"),
    (r"(?i)\bexemplifies\b", "shows"), (r"(?i)\bexemplify\b", "show"),
    (r"(?i)\bnotwithstanding\b", "despite"),
    (r"(?i)\bwherein\b", "where"),
    (r"(?i)\bthereby\b", "by doing this"),
    (r"(?i)\bwhereas\b", "while"),
    (r"(?i)\bpertaining to\b", "about"),
    (r"(?i)\bwith respect to\b", "about"),
    (r"(?i)\bwith regard to\b", "about"),
    (r"(?i)\bin the context of\b", "in"),
    (r"(?i)\bshowcases\b", "shows"), (r"(?i)\bshowcase\b", "show"),
    (r"(?i)\bshowcased\b", "showed"),
    (r"(?i)\bencompasses\b", "covers"), (r"(?i)\bencompass\b", "cover"),
    (r"(?i)\bdelve(?:s)?\b", "dig"), (r"(?i)\bdelving\b", "digging"),
    (r"(?i)\bharnessing\b", "using"), (r"(?i)\bharness(?:es)?\b", "use"),
    (r"(?i)\bfostering\b", "building"), (r"(?i)\bfosters\b", "builds"),
    (r"(?i)\bthe proposed\b", "our"),
    (r"(?i)\bthis paper presents\b", "we present"),
    (r"(?i)\bthis paper proposes\b", "we propose"),
    (r"(?i)\bthis work introduces\b", "we introduce"),
    (r"(?i)\bthis work presents\b", "we present"),
]

# =========================================================
# STEP 4: TRANSITION WORD REPLACEMENT — remove AI-typical ones
# Deterministic 1:1 replacements only
# =========================================================
TRANSITION_RULES = [
    (r"\bFurthermore,?\s", "Also, "),
    (r"\bMoreover,?\s", "Also, "),
    (r"\bAdditionally,?\s", "Also, "),
    (r"\bConsequently,?\s", "So, "),
    (r"\bNevertheless,?\s", "Still, "),
    (r"\bNonetheless,?\s", "Still, "),
    (r"\bSubsequently,?\s", "Then, "),
    (r"\bConversely,?\s", "But "),
    (r"\bUltimately,?\s", "In the end, "),
    (r"(?i)\bIn\s+addition,?\s", "Also, "),
    (r"(?i)\bAs\s+a\s+result,?\s", "So, "),
    (r"\bAccordingly,?\s", "So, "),
    (r"\bHence,?\s", "So, "),
    (r"\bThus,?\s", "So, "),
]

# =========================================================
# STEP 5: PASSIVE -> ACTIVE where safe
# =========================================================
PASSIVE_RULES = [
    (r"(?i)\bwas conducted\b", "took place"),
    (r"(?i)\bwere conducted\b", "took place"),
    (r"(?i)\bwas observed\b", "showed up"),
    (r"(?i)\bwas identified\b", "stood out"),
    (r"(?i)\bwere identified\b", "stood out"),
    (r"(?i)\bis characterized by\b", "stands out for"),
    (r"(?i)\bwas implemented\b", "was put in place"),
    (r"(?i)\bwas utilized\b", "was used"),
]

# =========================================================
# STEP 6: UNICODE CLEANUP — strips AI text artifacts
# =========================================================
def ai_text_clean(text, preserve_boundary_spaces=False):
    leading = ""
    trailing = ""
    if preserve_boundary_spaces:
        lm = _re.match(r'^(\s+)', text)
        tm = _re.search(r'(\s+)$', text)
        if lm: leading = lm.group(1)
        if tm: trailing = tm.group(1)

    text = _ud.normalize("NFKC", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _re.sub(r"[\u200B\u200C\u200D\u2060\uFEFF\u180E]", "", text)
    text = _re.sub(r"[\u00A0\u2007\u202F\u2009\u200A]", " ", text)
    text = _re.sub(r"[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]", "", text)
    text = _re.sub(r"[\u2013\u2014\u2212]", "-", text)
    text = _re.sub(r'[\u201C\u201D\u201E\u00AB\u00BB]', '"', text)
    text = _re.sub(r"[\u2018\u2019\u201A\u201B\u2039\u203A]", "'", text)
    text = text.replace("\u2026", "...")
    text = _re.sub(r"[\u2022\u25E6\u00B7\u2219]", "-", text)
    text = _re.sub(r"[\u00A9\u00AE\u2000-\u3300]|[\U0001F000-\U0001F9FF]", "", text)
    text = _re.sub(r"\s+([.,!?;:])", r"\1", text)
    text = _re.sub(r"[ \t]*\n[ \t]*", "\n", text)
    text = _re.sub(r"[ \t]{2,}", " ", text)
    text = _re.sub(r"\n{3,}", "\n\n", text)
    text = "\n".join(line.rstrip() for line in text.split("\n"))
    text = text.strip()

    if preserve_boundary_spaces:
        if leading and not text.startswith(" "): text = " " + text
        if trailing and not text.endswith(" "): text = text + " "
    return text

# =========================================================
# APPLY — process each run individually to preserve formatting
# =========================================================
def process_text(text):
    """Apply all deterministic rules to text."""
    for pat, repl in CONTRACTIONS:
        text = _re.sub(pat, repl, text)
    for pat, repl in FILLER_RULES:
        text = _re.sub(pat, repl, text)
    for pat, repl in VOCAB_RULES:
        text = _re.sub(pat, repl, text)
    for pat, repl in TRANSITION_RULES:
        text = _re.sub(pat, repl, text)
    for pat, repl in PASSIVE_RULES:
        text = _re.sub(pat, repl, text)
    return text

changed = 0
cleaned = 0
skipped = 0

for idx, para in enumerate(doc.paragraphs):
    raw = para.text.strip()
    if should_skip(idx, raw):
        skipped += 1
        if idx < REFERENCES_START:
            for run in para.runs:
                if run.text:
                    prev = run.text
                    run.text = ai_text_clean(run.text, preserve_boundary_spaces=True)
                    if run.text != prev: cleaned += 1
        continue

    for run in para.runs:
        if run.text and run.text.strip():
            original = run.text
            run.text = process_text(run.text)
            if run.text != original: changed += 1
            prev = run.text
            run.text = ai_text_clean(run.text, preserve_boundary_spaces=True)
            if run.text != prev: cleaned += 1

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    if run.text and run.text.strip():
                        original = run.text
                        run.text = process_text(run.text)
                        if run.text != original: changed += 1
                        prev = run.text
                        run.text = ai_text_clean(run.text, preserve_boundary_spaces=True)
                        if run.text != prev: cleaned += 1

print(f"\nSkipped: {skipped}")
print(f"Runs changed: {changed}")
print(f"Runs unicode-cleaned: {cleaned}")

doc.save(DST)
print(f"\nDone - saved {DST}")
