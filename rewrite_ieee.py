"""
Enhanced port of index.html sanitizeText + humanizeText,
tuned for IEEE academic papers.
Keeps: sanitizer, contractions, vocab simplification, filler removal,
       transition replacement, adverbs, passive voice, n-gram breaking
Removes: casual markers, asides, emphatics (wrong for academic tone)
Adds: heavy word-level synonym perturbation for academic text
"""
import sys, io, re as _re, unicodedata as _ud
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

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
    special = ['\u2211','\u220f','\u222b','\u2208','\u2200','\u2203','\u2264','\u2265',
               '\u2260','\u221e','\u2202','\u2207','\u00d7','\u2297','\u2299','\u221d']
    if sum(1 for c in txt if c in special) >= 2:
        return True
    words = txt.split()
    if len(words) < 8 and _re.search(r'[=\u2264\u2265\u2208\u2200\u2203\u2211\u220f\u222b\u2202\u2207\u2297\u2299\u2248\u221d]', txt):
        return True
    if txt.count('=') >= 2 and len(words) < 12:
        return True
    return False

def should_skip(idx, txt):
    if idx in SKIP_INDICES: return True
    if idx >= REFERENCES_START: return True
    if not txt or len(txt.strip()) < 5: return True
    if is_section_header(txt): return True
    if is_algorithm_line(txt): return True
    if is_formula(txt): return True
    return False


# =========================================================
# Sanitizer — exact port from index.html
# =========================================================
def sanitize_text(text):
    text = _ud.normalize("NFKC", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _re.sub(r"[\u200B\u200C\u200D\u2060\uFEFF\u180E]", "", text)
    text = _re.sub(r"[\u00A0\u2007\u202F\u2009\u200A]", " ", text)
    text = _re.sub(r"[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]", "", text)
    text = _re.sub(r"&[#\w]+;", " ", text)
    text = _re.sub(r"[\u2013\u2014\u2212]", "-", text)
    text = _re.sub(r'[\u201C\u201D\u201E\u00AB\u00BB]', '"', text)
    text = _re.sub(r"[\u2018\u2019\u201A\u201B\u2039\u203A]", "'", text)
    text = text.replace("\u2026", "...")
    text = _re.sub(r"[\u2022\u25E6\u00B7\u2219]", "-", text)
    text = _re.sub(r"\s+([.,!?;:])", r"\1", text)
    text = _re.sub(r"[ \t]{2,}", " ", text)
    return text


# =========================================================
# PRNG — same LCG as index.html
# =========================================================
class PRNG:
    def __init__(self, seed=42):
        self._s = seed
    def rand(self):
        self._s = (self._s * 1103515245 + 12345) & 0x7FFFFFFF
        return self._s / 0x7FFFFFFF
    def pick(self, arr):
        return arr[int(self.rand() * len(arr))]


# =========================================================
# Enhanced humanizer — academic-tuned
# =========================================================
def humanize_text(text, seed=42):
    rng = PRNG(seed)

    # --- 1) Transition words (index.html step 1) — academic alternatives ---
    tr_rules = [
        (r"\bFurthermore,?\s", lambda: rng.pick(["Also, ", "In addition, ", "Equally, "])),
        (r"\bMoreover,?\s", lambda: rng.pick(["Also, ", "Beyond that, ", "On top of that, "])),
        (r"\bAdditionally,?\s", lambda: rng.pick(["Also, ", "On top of this, ", "Equally, "])),
        (r"\bConsequently,?\s", lambda: rng.pick(["As such, ", "For this reason, ", "So "])),
        (r"\bNevertheless,?\s", lambda: rng.pick(["Even so, ", "That said, ", "Still, "])),
        (r"\bNonetheless,?\s", lambda: rng.pick(["Even so, ", "Still, ", "Yet "])),
        (r"\bSubsequently,?\s", lambda: rng.pick(["Then, ", "After that, ", "Next, "])),
        (r"\bConversely,?\s", lambda: rng.pick(["On the other hand, ", "In contrast, ", "Alternatively, "])),
        (r"\bUltimately,?\s", lambda: rng.pick(["In the end, ", "Finally, "])),
        (r"(?i)\bIn\s+addition,?\s", lambda: rng.pick(["Also, ", "Equally, ", "Along with this, "])),
        (r"(?i)\bAs\s+a\s+result,?\s", lambda: rng.pick(["Because of this, ", "For this reason, "])),
        (r"(?i)\bIn\s+particular,?\s", lambda: rng.pick(["Especially, ", "Notably, "])),
        (r"(?i)\bSpecifically,?\s", lambda: rng.pick(["Namely, ", "That is, ", "More precisely, "])),
        (r"(?i)\bAccordingly,?\s", lambda: rng.pick(["For that reason, ", "As such, "])),
        (r"(?i)\bHence,?\s", lambda: rng.pick(["For this reason, ", "As such, "])),
        (r"(?i)\bThus,?\s", lambda: rng.pick(["In this way, ", "So, "])),
    ]
    for pattern, replacer in tr_rules:
        text = _re.sub(pattern, lambda m, fn=replacer: fn(), text)

    # --- 2) Formal vocabulary (index.html step 2) + expanded academic list ---
    voc_rules = [
        (r"(?i)\butilizes\b", "uses"), (r"(?i)\butilize\b", "use"),
        (r"(?i)\butilized\b", "used"), (r"(?i)\butilizing\b", "using"),
        (r"(?i)\bfacilitates\b", "helps"), (r"(?i)\bfacilitate\b", "help"),
        (r"(?i)\bfacilitated\b", "helped"), (r"(?i)\bfacilitating\b", "helping"),
        (r"(?i)\bdemonstrates\b", "shows"), (r"(?i)\bdemonstrate\b", "show"),
        (r"(?i)\bdemonstrated\b", "showed"), (r"(?i)\bdemonstrating\b", "showing"),
        (r"(?i)\bcomprehensive\b", lambda: rng.pick(["thorough", "complete", "full"])),
        (r"(?i)\bsignificantly\b", lambda: rng.pick(["greatly", "noticeably", "markedly"])),
        (r"(?i)\bsignificant\b", lambda: rng.pick(["major", "notable", "marked"])),
        (r"(?i)\bsubstantially\b", lambda: rng.pick(["greatly", "largely", "considerably"])),
        (r"(?i)\bsubstantial\b", lambda: rng.pick(["large", "considerable", "sizable"])),
        (r"(?i)\bapproximately\b", lambda: rng.pick(["about", "around", "roughly"])),
        (r"(?i)\bnumerous\b", lambda: rng.pick(["many", "several", "a number of"])),
        (r"(?i)\boptimal\b", lambda: rng.pick(["best", "ideal"])),
        (r"(?i)\benhances\b", "improves"), (r"(?i)\benhance\b", "improve"),
        (r"(?i)\benhanced\b", "improved"), (r"(?i)\benhancing\b", "improving"),
        (r"(?i)\bleverages\b", "uses"), (r"(?i)\bleverage\b", "use"),
        (r"(?i)\bleveraging\b", "using"), (r"(?i)\bleveraged\b", "used"),
        (r"(?i)\bensures\b", "makes sure"), (r"(?i)\bensure\b", "make sure"),
        (r"(?i)\bensuring\b", "making sure"),
        (r"(?i)\bpivotal\b", lambda: rng.pick(["key", "central", "critical"])),
        (r"(?i)\bmultifaceted\b", lambda: rng.pick(["complex", "varied"])),
        (r"(?i)\bstreamlines\b", "simplifies"), (r"(?i)\bstreamline\b", "simplify"),
        (r"(?i)\bstreamlining\b", "simplifying"),
        (r"(?i)\bparamount\b", lambda: rng.pick(["critical", "most important"])),
        (r"(?i)\brobust\b", lambda: rng.pick(["strong", "solid", "reliable"])),
        (r"(?i)\bseamlessly\b", "smoothly"), (r"(?i)\bseamless\b", "smooth"),
        (r"(?i)\binnovative\b", lambda: rng.pick(["new", "novel", "creative"])),
        (r"(?i)\bholistic\b", lambda: rng.pick(["overall", "complete"])),
        (r"(?i)\bparadigms\b", "models"), (r"(?i)\bparadigm\b", "model"),
        (r"(?i)\bmethodologies\b", "methods"), (r"(?i)\bmethodology\b", "method"),
        (r"(?i)\bunderscores\b", "highlights"), (r"(?i)\bunderscore\b", "highlight"),
        (r"(?i)\bunderscoring\b", "highlighting"),
        (r"(?i)\bendeavors\b", "efforts"), (r"(?i)\bendeavor\b", "effort"),
        (r"(?i)\baugments\b", "boosts"), (r"(?i)\baugment\b", "boost"),
        (r"(?i)\baugmenting\b", "boosting"),
        (r"(?i)\bmitigates\b", "reduces"), (r"(?i)\bmitigate\b", "reduce"),
        (r"(?i)\bmitigating\b", "reducing"), (r"(?i)\bmitigation\b", "reduction"),
        (r"(?i)\bascertains\b", "finds out"), (r"(?i)\bascertain\b", "find out"),
        (r"(?i)\bascertained\b", "found out"),
        (r"(?i)\bcommences\b", "starts"), (r"(?i)\bcommence\b", "start"),
        (r"(?i)\bcommenced\b", "started"), (r"(?i)\bcommencing\b", "starting"),
        (r"(?i)\bterminates\b", "ends"), (r"(?i)\bterminate\b", "end"),
        (r"(?i)\bterminated\b", "ended"),
        (r"(?i)\bexpedites\b", "speeds up"), (r"(?i)\bexpedite\b", "speed up"),
        (r"(?i)\bexpediting\b", "speeding up"),
        (r"(?i)\brequisite\b", "needed"),
        (r"(?i)\bpredominantly\b", lambda: rng.pick(["mostly", "mainly"])),
        (r"(?i)\bpredominant\b", lambda: rng.pick(["main", "primary"])),
        (r"(?i)\bcorroborates\b", "backs up"), (r"(?i)\bcorroborate\b", "back up"),
        (r"(?i)\bcorroborating\b", "backing up"),
        (r"(?i)\belucidated\b", "explained"), (r"(?i)\belucidate\b", "explain"),
        (r"(?i)\belucidating\b", "explaining"),
        (r"(?i)\bameliorate\b", "improve"), (r"(?i)\bameliorates\b", "improves"),
        (r"(?i)\bamelioration\b", "improvement"),
        (r"(?i)\bdelineates\b", "outlines"), (r"(?i)\bdelineate\b", "outline"),
        (r"(?i)\bdelineating\b", "outlining"),
        (r"(?i)\bpropensity\b", "tendency"),
        (r"(?i)\bdiscernible\b", "noticeable"),
        (r"(?i)\bpertaining to\b", "about"),
        (r"(?i)\bwith respect to\b", "about"),
        (r"(?i)\bwith regard to\b", "about"),
        (r"(?i)\bin the context of\b", "in"),
        (r"(?i)\bthe proposed\b", "our"),
        (r"(?i)\bthis paper presents\b", "we present"),
        (r"(?i)\bthis paper proposes\b", "we propose"),
        (r"(?i)\bthis work introduces\b", "we introduce"),
        (r"(?i)\bthis work presents\b", "we present"),
        (r"(?i)\bnotwithstanding\b", "despite"),
        (r"(?i)\bwherein\b", "where"),
        (r"(?i)\bthereby\b", "and so"),
        (r"(?i)\bwhereas\b", "while"),
        (r"(?i)\bshowcases\b", "shows"), (r"(?i)\bshowcase\b", "show"),
        (r"(?i)\bshowcased\b", "showed"),
        (r"(?i)\bencompasses\b", "covers"), (r"(?i)\bencompass\b", "cover"),
        (r"(?i)\bdelves\b", "digs"), (r"(?i)\bdelve\b", "dig"),
        (r"(?i)\bdelving\b", "digging"),
        (r"(?i)\bharnessing\b", "using"), (r"(?i)\bharnesses\b", "uses"),
        (r"(?i)\bfostering\b", "building"), (r"(?i)\bfosters\b", "builds"),
        (r"(?i)\bexemplifies\b", "illustrates"), (r"(?i)\bexemplify\b", "illustrate"),
        (r"(?i)\bexemplified\b", "illustrated"),
        (r"(?i)\bpertinent\b", "relevant"),
        (r"(?i)\bconcomitant\b", "accompanying"),
        (r"(?i)\bpreclude\b", "prevent"), (r"(?i)\bprecludes\b", "prevents"),
        (r"(?i)\bprecluding\b", "preventing"),
        (r"(?i)\baforementioned\b", "mentioned earlier"),
        (r"(?i)\baforesaid\b", "mentioned above"),
        (r"(?i)\bherein\b", "here"),
        (r"(?i)\binsofar\b", "to the extent"),
        (r"(?i)\binasmuch\b", "since"),
        (r"(?i)\bexhibits\b", "shows"), (r"(?i)\bexhibit\b", "show"),
        (r"(?i)\bexhibited\b", "showed"), (r"(?i)\bexhibiting\b", "showing"),
        (r"(?i)\bmanifests\b", "shows"), (r"(?i)\bmanifest\b", "show"),
        (r"(?i)\bmanifested\b", "showed"),
        (r"(?i)\belicits\b", "draws out"), (r"(?i)\belicit\b", "draw out"),
        (r"(?i)\belicited\b", "drawn out"),
        (r"(?i)\bameliorating\b", "improving"),
        (r"(?i)\bconducive\b", "helpful"),
        (r"(?i)\bpossesses\b", "has"), (r"(?i)\bpossess\b", "have"),
        (r"(?i)\bcommencing\b", "starting"),
        (r"(?i)\brenders\b", "makes"), (r"(?i)\brendered\b", "made"),
        (r"(?i)\bgarners\b", "gets"), (r"(?i)\bgarner\b", "get"),
        (r"(?i)\bgarnered\b", "got"),
        (r"(?i)\badjacent to\b", "next to"),
        (r"(?i)\bin lieu of\b", "instead of"),
        (r"(?i)\bvis-a-vis\b", "compared to"),
        (r"(?i)\bplethora\b", "abundance"),
        (r"(?i)\bmyriad\b", lambda: rng.pick(["wide range of", "many"])),
        (r"(?i)\bample\b", lambda: rng.pick(["enough", "plenty of"])),
        (r"(?i)\bcategorically\b", "firmly"),
        (r"(?i)\bexponentially\b", "rapidly"),
        (r"(?i)\bintrinsically\b", "by nature"),
        (r"(?i)\binherently\b", "by nature"),
        (r"(?i)\boverarching\b", "broad"),
        (r"(?i)\bunderpinning\b", "supporting"),
        (r"(?i)\bunderpins\b", "supports"),
        (r"(?i)\bencapsulates\b", "captures"), (r"(?i)\bencapsulate\b", "capture"),
    ]
    for pattern, repl in voc_rules:
        if callable(repl):
            text = _re.sub(pattern, lambda m, fn=repl: fn(), text)
        else:
            text = _re.sub(pattern, repl, text)

    # --- 3) AI filler phrases (index.html step 3) ---
    ph_rules = [
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
        (r"(?i)\bplays a (?:crucial|vital|key|important|significant) role\b", "matters"),
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
        (r"(?i)\ba considerable amount of\b", "much"),
        (r"(?i)\bthe vast majority of\b", "most"),
        (r"(?i)\bin close proximity to\b", "near"),
        (r"(?i)\bat this juncture\b", "now"),
        (r"(?i)\bgiven the fact that\b", lambda: rng.pick(["since", "because"])),
        (r"(?i)\bin light of\b", lambda: rng.pick(["given", "considering"])),
        (r"(?i)\bfor the purpose of\b", "for"),
        (r"(?i)\bwith a view to\b", "to"),
        (r"(?i)\bin an effort to\b", "to"),
        (r"(?i)\bas a means of\b", "to"),
        (r"(?i)\bin such a manner that\b", "so that"),
        (r"(?i)\bin a manner that\b", "so that"),
        (r"(?i)\bto a large extent\b", "largely"),
        (r"(?i)\bto a great extent\b", "greatly"),
        (r"(?i)\bto a certain extent\b", "partly"),
        (r"(?i)\bby means of\b", "through"),
        (r"(?i)\bwith the aim of\b", "to"),
        (r"(?i)\bin the case of\b", "for"),
        (r"(?i)\bin the process of\b", "while"),
        (r"(?i)\bas a consequence of\b", "because of"),
        (r"(?i)\bin terms of\b", "regarding"),
        (r"(?i)\bthrough the use of\b", "using"),
    ]
    for pattern, repl in ph_rules:
        if callable(repl):
            text = _re.sub(pattern, lambda m, fn=repl: fn(), text)
        else:
            text = _re.sub(pattern, repl, text)

    # --- 4) Contractions (index.html step 4) ---
    c_rules = [
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
    for pattern, repl in c_rules:
        text = _re.sub(pattern, repl, text)

    # --- 5) Adverb variation (index.html step 9) ---
    adv_rules = [
        (r"(?i)\bconsiderably\b", lambda: rng.pick(["noticeably", "markedly"])),
        (r"(?i)\bremarkably\b", lambda: rng.pick(["surprisingly", "strikingly"])),
        (r"(?i)\bparticularly\b", lambda: rng.pick(["especially", "notably"])),
        (r"(?i)\bprimarily\b", lambda: rng.pick(["mainly", "mostly", "chiefly"])),
        (r"(?i)\bfundamentally\b", lambda: rng.pick(["at its core", "basically"])),
        (r"(?i)\bincreasingly\b", lambda: rng.pick(["more and more", "progressively"])),
        (r"(?i)\bsystematically\b", lambda: rng.pick(["step by step", "methodically"])),
        (r"(?i)\beffectively\b", lambda: rng.pick(["in practice", "successfully"])),
        (r"(?i)\bcritically\b", lambda: rng.pick(["importantly", "crucially"])),
        (r"(?i)\bextensively\b", lambda: rng.pick(["broadly", "widely"])),
        (r"(?i)\bempirically\b", lambda: rng.pick(["through experiments", "experimentally"])),
        (r"(?i)\brespectively\b", lambda: rng.pick(["in that order", "correspondingly"])),
        (r"(?i)\bnotably\b", lambda: rng.pick(["especially", "in particular"])),
        (r"(?i)\bconversely\b", lambda: rng.pick(["in contrast", "on the other hand"])),
        (r"(?i)\binversely\b", lambda: rng.pick(["in reverse", "the other way around"])),
        (r"(?i)\baccordingly\b", lambda: rng.pick(["for that reason", "as expected"])),
        (r"(?i)\bconcurrently\b", lambda: rng.pick(["at the same time", "in parallel"])),
        (r"(?i)\bexplicitly\b", lambda: rng.pick(["directly", "clearly"])),
        (r"(?i)\bimplicitly\b", lambda: rng.pick(["indirectly", "without stating"])),
    ]
    for pattern, replacer in adv_rules:
        text = _re.sub(pattern, lambda m, fn=replacer: fn(), text)

    # --- 6) Passive voice hints (index.html step 10) ---
    pv_rules = [
        (r"(?i)\bwas conducted\b", lambda: rng.pick(["took place", "was carried out"])),
        (r"(?i)\bwere conducted\b", lambda: rng.pick(["took place", "were carried out"])),
        (r"(?i)\bwas observed\b", lambda: rng.pick(["showed up", "was noticed"])),
        (r"(?i)\bwas identified\b", lambda: rng.pick(["stood out", "was spotted"])),
        (r"(?i)\bwere identified\b", lambda: rng.pick(["stood out", "were found"])),
        (r"(?i)\bis characterized by\b", lambda: rng.pick(["stands out for", "is known for"])),
        (r"(?i)\bwas implemented\b", lambda: rng.pick(["was built", "was put in place"])),
        (r"(?i)\bwas utilized\b", "was used"),
        (r"(?i)\bwas employed\b", "was used"),
        (r"(?i)\bwere employed\b", "were used"),
        (r"(?i)\bwas performed\b", lambda: rng.pick(["was done", "was carried out"])),
        (r"(?i)\bwere performed\b", lambda: rng.pick(["were done", "were carried out"])),
        (r"(?i)\bwas evaluated\b", lambda: rng.pick(["was tested", "was assessed"])),
        (r"(?i)\bwere evaluated\b", lambda: rng.pick(["were tested", "were assessed"])),
        (r"(?i)\bwas obtained\b", lambda: rng.pick(["was found", "was gathered"])),
        (r"(?i)\bwere obtained\b", lambda: rng.pick(["were found", "were gathered"])),
        (r"(?i)\bwas achieved\b", lambda: rng.pick(["was reached", "was attained"])),
        (r"(?i)\bwere achieved\b", lambda: rng.pick(["were reached", "were attained"])),
    ]
    for pattern, replacer in pv_rules:
        if callable(replacer):
            text = _re.sub(pattern, lambda m, fn=replacer: fn(), text)
        else:
            text = _re.sub(pattern, replacer, text)

    # --- 7) AI n-gram breaking (index.html step 13) ---
    ng_rules = [
        (r"(?i)\bIt is worth noting\b", lambda: rng.pick(["Note that", "Notably"])),
        (r"(?i)\bThis is particularly true\b", lambda: rng.pick(["This holds especially", "This especially applies"])),
        (r"(?i)\bA growing body of\b", lambda: rng.pick(["More and more", "Increasing"])),
        (r"(?i)\bIn light of\b", lambda: rng.pick(["Given", "Considering"])),
        (r"(?i)\bOn the other hand\b", lambda: rng.pick(["In contrast", "Alternatively"])),
        (r"(?i)\bIn this regard\b", lambda: rng.pick(["On this point", "Here"])),
        (r"(?i)\bThis suggests that\b", lambda: rng.pick(["This points to", "This hints that"])),
        (r"(?i)\bThis indicates that\b", lambda: rng.pick(["This shows that", "This reveals that"])),
        (r"(?i)\bThis implies that\b", lambda: rng.pick(["This means", "This points to"])),
        (r"(?i)\bIt is essential to\b", lambda: rng.pick(["It's critical to", "One must"])),
        (r"(?i)\bserves as a\b", lambda: rng.pick(["works as a", "acts as a"])),
        (r"(?i)\ba key factor\b", lambda: rng.pick(["a major factor", "a main driver"])),
        (r"(?i)\bin today's (?:digital|modern|contemporary) (?:age|era|world)\b",
         lambda: rng.pick(["nowadays", "today", "currently"])),
        (r"(?i)\bpaves the way for\b", lambda: rng.pick(["opens the door to", "enables"])),
        (r"(?i)\bsheds? light on\b", lambda: rng.pick(["clarifies", "helps explain"])),
        (r"(?i)\bthe findings suggest\b", lambda: rng.pick(["the results point to", "the data shows"])),
        (r"(?i)\bthe results indicate\b", lambda: rng.pick(["the data shows", "the numbers show"])),
        (r"(?i)\bas a whole\b", lambda: rng.pick(["overall", "taken together"])),
        (r"(?i)\bgiven the fact that\b", lambda: rng.pick(["since", "because"])),
        (r"(?i)\bit is evident that\b", lambda: rng.pick(["clearly,", "it's clear that"])),
        (r"(?i)\bit is clear that\b", lambda: rng.pick(["clearly,", "plainly,"])),
        (r"(?i)\bplay(?:s)? a (?:crucial|vital|critical|key|significant|important) role\b",
         lambda: rng.pick(["matter greatly", "is central", "is key"])),
    ]
    for pattern, replacer in ng_rules:
        text = _re.sub(pattern, lambda m, fn=replacer: fn(), text)

    # --- 8) Word-level synonym perturbation for academic text ---
    # 30% chance to swap common academic words with alternatives
    ACADEMIC_SYNONYMS = {
        "achieved": ["reached", "attained", "obtained"],
        "achieves": ["reaches", "attains", "gets"],
        "achieve": ["reach", "attain", "get"],
        "achieving": ["reaching", "attaining", "getting"],
        "accuracy": ["precision", "correctness"],
        "adequate": ["sufficient", "enough"],
        "analyze": ["examine", "study", "inspect"],
        "analyzed": ["examined", "studied", "inspected"],
        "analyzes": ["examines", "studies"],
        "analyzing": ["examining", "studying"],
        "analysis": ["examination", "study", "review"],
        "approach": ["technique", "strategy", "method"],
        "approaches": ["techniques", "strategies", "methods"],
        "applicable": ["relevant", "suitable", "fitting"],
        "appropriate": ["suitable", "fitting", "proper"],
        "architecture": ["design", "structure", "layout"],
        "capability": ["ability", "capacity", "potential"],
        "capabilities": ["abilities", "capacities"],
        "challenge": ["difficulty", "hurdle", "obstacle"],
        "challenges": ["difficulties", "hurdles", "obstacles"],
        "component": ["part", "element", "piece"],
        "components": ["parts", "elements", "pieces"],
        "computation": ["calculation", "processing"],
        "computational": ["computing", "processing"],
        "configuration": ["setup", "arrangement", "layout"],
        "constraint": ["limitation", "restriction", "bound"],
        "constraints": ["limitations", "restrictions", "bounds"],
        "critical": ["crucial", "vital", "essential"],
        "dataset": ["data collection", "corpus"],
        "datasets": ["data collections", "corpora"],
        "degradation": ["decline", "drop", "weakening"],
        "deployment": ["rollout", "release", "launch"],
        "detrimental": ["harmful", "damaging", "negative"],
        "domain": ["field", "area", "sphere"],
        "domains": ["fields", "areas"],
        "dynamic": ["changing", "shifting", "evolving"],
        "efficacy": ["effectiveness", "potency"],
        "efficient": ["effective", "capable", "fast"],
        "efficiency": ["effectiveness", "performance", "speed"],
        "evident": ["clear", "apparent", "obvious"],
        "framework": ["system", "structure", "setup"],
        "frameworks": ["systems", "structures"],
        "furthermore": ["also", "beyond this"],
        "granularity": ["detail level", "resolution"],
        "heterogeneous": ["diverse", "mixed", "varied"],
        "homogeneous": ["uniform", "consistent", "alike"],
        "hypothesis": ["theory", "assumption", "idea"],
        "illustrate": ["show", "depict", "highlight"],
        "illustrates": ["shows", "depicts", "highlights"],
        "illustrated": ["shown", "depicted"],
        "impact": ["effect", "influence", "consequence"],
        "impacts": ["effects", "influences"],
        "implement": ["build", "create", "set up"],
        "implemented": ["built", "created", "set up"],
        "implements": ["builds", "creates"],
        "implementing": ["building", "creating", "setting up"],
        "implementation": ["setup", "realization", "execution"],
        "indicate": ["show", "suggest", "point to"],
        "indicates": ["shows", "suggests", "points to"],
        "indicated": ["showed", "suggested", "pointed to"],
        "indicating": ["showing", "suggesting", "pointing to"],
        "inherent": ["built-in", "intrinsic", "natural"],
        "investigate": ["explore", "study", "look into"],
        "investigated": ["explored", "studied", "looked into"],
        "investigation": ["exploration", "study", "review"],
        "limitation": ["drawback", "shortcoming", "weakness"],
        "limitations": ["drawbacks", "shortcomings", "weaknesses"],
        "magnitude": ["size", "scale", "extent"],
        "mechanism": ["process", "procedure", "method"],
        "mechanisms": ["processes", "procedures", "methods"],
        "modality": ["mode", "type", "channel"],
        "modalities": ["modes", "types", "channels"],
        "module": ["unit", "block", "segment"],
        "modules": ["units", "blocks", "segments"],
        "obtain": ["get", "gather", "collect"],
        "obtained": ["got", "gathered", "collected"],
        "obtains": ["gets", "gathers"],
        "obtaining": ["getting", "gathering", "collecting"],
        "parameter": ["setting", "variable", "factor"],
        "parameters": ["settings", "variables", "factors"],
        "perform": ["carry out", "do", "execute"],
        "performed": ["carried out", "did", "executed"],
        "performs": ["carries out", "does", "executes"],
        "performing": ["carrying out", "doing", "executing"],
        "phenomenon": ["event", "occurrence", "pattern"],
        "pipeline": ["workflow", "process chain", "sequence"],
        "prevalent": ["common", "widespread", "frequent"],
        "prior": ["earlier", "previous", "past"],
        "procedure": ["process", "method", "step"],
        "procedures": ["processes", "methods", "steps"],
        "propagation": ["spread", "transmission", "flow"],
        "rationale": ["reasoning", "basis", "logic"],
        "scenario": ["situation", "case", "setting"],
        "scenarios": ["situations", "cases", "settings"],
        "scheme": ["plan", "strategy", "setup"],
        "straightforward": ["simple", "direct", "easy"],
        "subsequent": ["later", "following", "next"],
        "sufficient": ["enough", "adequate"],
        "superior": ["better", "higher", "stronger"],
        "threshold": ["cutoff", "limit", "boundary"],
        "thresholds": ["cutoffs", "limits", "boundaries"],
        "trajectory": ["path", "course", "direction"],
        "utilize": ["use", "apply", "employ"],
        "validates": ["confirms", "verifies", "checks"],
        "validate": ["confirm", "verify", "check"],
        "validated": ["confirmed", "verified", "checked"],
        "validating": ["confirming", "verifying", "checking"],
        "various": ["different", "diverse", "multiple"],
        "yield": ["produce", "give", "generate"],
        "yields": ["produces", "gives", "generates"],
    }

    def _perturb_word(m):
        word = m.group(0)
        lower = word.lower()
        if lower in ACADEMIC_SYNONYMS and rng.rand() < 0.30:
            replacement = rng.pick(ACADEMIC_SYNONYMS[lower])
            if word[0].isupper():
                replacement = replacement[0].upper() + replacement[1:]
            return replacement
        return word

    pattern = r"\b(" + "|".join(_re.escape(w) for w in ACADEMIC_SYNONYMS.keys()) + r")\b"
    text = _re.sub(pattern, _perturb_word, text, flags=_re.IGNORECASE)

    return text


# =========================================================
# PROCESS DOCUMENT — whole paragraph, distribute back to runs
# =========================================================
changed = 0
skipped = 0
seed_counter = 42

def distribute_text_to_runs(runs, new_text):
    orig_lengths = [len(r.text) for r in runs]
    total_orig = sum(orig_lengths)
    if total_orig == 0:
        runs[0].text = new_text
        return
    new_len = len(new_text)
    pos = 0
    for i, run in enumerate(runs):
        if i == len(runs) - 1:
            run.text = new_text[pos:]
        else:
            share = max(1, int(round(new_len * orig_lengths[i] / total_orig)))
            end = min(pos + share, new_len)
            if end < new_len:
                space = new_text.rfind(' ', pos, end + 15)
                if space > pos:
                    end = space + 1
            run.text = new_text[pos:end]
            pos = end

def process_para(para):
    global changed, seed_counter
    runs = [r for r in para.runs if r.text]
    if not runs:
        return
    full_text = "".join(r.text for r in runs)
    if not full_text.strip():
        return
    seed_counter += 1
    processed = sanitize_text(full_text)
    processed = humanize_text(processed, seed=seed_counter)
    if processed != full_text:
        changed += 1
        if len(runs) == 1:
            runs[0].text = processed
        else:
            distribute_text_to_runs(runs, processed)

for idx, para in enumerate(doc.paragraphs):
    raw = para.text.strip()
    if should_skip(idx, raw):
        skipped += 1
        continue
    process_para(para)

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                if p.text.strip():
                    process_para(p)

print(f"\nSkipped: {skipped}")
print(f"Paragraphs changed: {changed}")

doc.save(DST)
print(f"\nDone - saved {DST}")
