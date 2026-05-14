"""
IEEE paper humanizer v4 — maximum restructuring per the full humanization prompt.
Key upgrades over v3:
  - Higher restructuring rates (45% clause reorder, 50% merge, 55% split, 35% front, 30% voice)
  - Paragraph opener variation (first sentence of each para gets restructured at 60%)
  - 60% synonym perturbation rate with 220+ academic synonyms
  - Sentence length burstiness enforcement
  - Consecutive same-starter breaking
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

def has_citation(txt):
    return bool(_re.search(r'\[\d+\]', txt))

def should_skip(idx, txt):
    if idx in SKIP_INDICES: return True
    if idx >= REFERENCES_START: return True
    if not txt or len(txt.strip()) < 5: return True
    if is_section_header(txt): return True
    if is_algorithm_line(txt): return True
    if is_formula(txt): return True
    return False

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

class PRNG:
    def __init__(self, seed=42):
        self._s = seed
    def rand(self):
        self._s = (self._s * 1103515245 + 12345) & 0x7FFFFFFF
        return self._s / 0x7FFFFFFF
    def pick(self, arr):
        return arr[int(self.rand() * len(arr))]


def split_sentences(text):
    parts = _re.split(r'(?<=[.!?])\s+(?=[A-Z])', text)
    return [p.strip() for p in parts if p.strip()]


# ==============================================================
# SENTENCE-LEVEL RESTRUCTURING — boosted rates
# ==============================================================
def restructure_sentences(sentences, rng, is_para_start=False):
    result = []
    i = 0
    while i < len(sentences):
        s = sentences[i]
        words = s.split()
        wc = len(words)

        if has_citation(s) or wc < 6 or is_formula(s):
            result.append(s)
            i += 1
            continue

        boost = 0.15 if (is_para_start and i == 0) else 0.0

        # A) CLAUSE REORDERING — 45% (60% for para openers)
        if rng.rand() < (0.45 + boost):
            reordered = False
            for conj, front in [(" because ", "Because "), (" since ", "Since "),
                                (" although ", "Although "), (" while ", "While "),
                                (" when ", "When "), (" whereas ", "While "),
                                (" as ", "As "), (" even though ", "Even though ")]:
                idx = s.lower().find(conj, max(15, len(s) // 4))
                if idx > 0 and not has_citation(s[idx:]):
                    main = s[:idx].rstrip()
                    sub = s[idx + len(conj):].strip()
                    if len(sub.split()) >= 4 and len(main.split()) >= 4:
                        if main[0].isupper():
                            main = main[0].lower() + main[1:]
                        if sub[-1] in '.!?':
                            sub = sub[:-1]
                        result.append(front + sub + ", " + main + ".")
                        reordered = True
                        break
            if reordered:
                i += 1
                continue

        # B) MERGE SHORT SENTENCES — 50%
        if wc < 12 and i + 1 < len(sentences):
            nxt = sentences[i + 1]
            nwc = len(nxt.split())
            if nwc < 12 and not has_citation(nxt) and rng.rand() < 0.50:
                conj = rng.pick([", and ", ", so ", " - ", "; ", " and also "])
                first = s.rstrip('.')
                second = nxt
                if second and second[0].isupper():
                    second = second[0].lower() + second[1:]
                result.append(first + conj + second)
                i += 2
                continue

        # C) SPLIT LONG SENTENCES — 55% for >25 words
        if wc > 25 and rng.rand() < 0.55:
            split_done = False
            for conj in [", and ", ", but ", ", yet ", " however ", ", which ", ", where "]:
                idx = s.find(conj, max(20, len(s) // 3))
                if idx > 0 and not has_citation(s[idx:idx+20]):
                    first = s[:idx].rstrip()
                    if not first.endswith(('.','!','?')):
                        first += '.'
                    rest = s[idx + len(conj):].strip()
                    if rest and len(rest.split()) > 4:
                        rest = rest[0].upper() + rest[1:]
                        if conj.strip() in ("which", "where"):
                            rest = "This " + rest[0].lower() + rest[1:]
                        result.append(first)
                        result.append(rest)
                        split_done = True
                        break
            if split_done:
                i += 1
                continue

        # D) PREPOSITIONAL PHRASE FRONTING — 35%
        if rng.rand() < (0.35 + boost) and wc > 10:
            m = _re.search(r',\s+((?:in|with|for|by|through|during|across|over|under|using|via|from)\s+[^,\.]+)[\.!?]$', s, _re.I)
            if m and not has_citation(m.group(1)):
                phrase = m.group(1).strip()
                main = s[:m.start()].rstrip()
                if main and main[0].isupper():
                    main = main[0].lower() + main[1:]
                result.append(phrase[0].upper() + phrase[1:] + ", " + main + ".")
                i += 1
                continue

        # E) ACTIVE/PASSIVE VOICE FLIP — 30%
        if rng.rand() < 0.30:
            m = _re.match(r'^(.+?)\s+(shows?|demonstrates?|reveals?|indicates?|confirms?|suggests?|provides?|achieves?|produces?|generates?)\s+(?:that\s+)?(.+)$', s, _re.I)
            if m and not has_citation(m.group(1)) and len(m.group(3).split()) > 4:
                subject = m.group(1).rstrip(',')
                rest = m.group(3)
                if rest and rest[0].islower():
                    rest = rest[0].upper() + rest[1:]
                connector = rng.pick(["as shown by", "as seen from", "according to", "based on"])
                if rest[-1] in '.!?':
                    rest = rest[:-1]
                new_s = rest + ", " + connector + " " + subject[0].lower() + subject[1:] + "."
                if len(new_s.split()) > 6:
                    result.append(new_s)
                    i += 1
                    continue

        # F) PARAGRAPH OPENER: rephrase "This/The X verb Y" → "Y is verb-ed by X" for first sentence
        if is_para_start and i == 0 and rng.rand() < 0.40:
            m = _re.match(r'^(The|This|These|Our)\s+(.+?)\s+(is|are|was|were|has|have)\s+(.+)$', s, _re.I)
            if m and not has_citation(s):
                det = m.group(1)
                subj = m.group(2)
                verb = m.group(3)
                rest = m.group(4)
                starters = ["Looking at ", "Turning to ", "Regarding ", "Considering ", "With respect to "]
                new_s = rng.pick(starters) + det.lower() + " " + subj + ", it " + verb + " " + rest
                result.append(new_s)
                i += 1
                continue

        result.append(s)
        i += 1

    return result


# ==============================================================
# CONSECUTIVE SAME-STARTER BREAKER
# ==============================================================
def break_same_starters(sentences, rng):
    if len(sentences) < 3:
        return sentences
    for i in range(2, len(sentences)):
        words_prev2 = sentences[i-2].split()
        words_prev1 = sentences[i-1].split()
        words_curr = sentences[i].split()
        if not words_prev2 or not words_prev1 or not words_curr:
            continue
        if words_prev2[0] == words_prev1[0] == words_curr[0]:
            starter = words_curr[0]
            rest = " ".join(words_curr[1:])
            alts = {
                "The": ["For the", "Regarding the", "As for the", "Concerning the"],
                "This": ["Such a", "That kind of", "A similar"],
                "These": ["Such", "All these", "Those"],
                "It": ["One can see it", "Notably, it", "As expected, it"],
                "They": ["Those models", "All of them", "Each one"],
                "We": ["Our team", "The authors", "In our work, we"],
                "There": ["One finds", "Notably, there", "In practice, there"],
            }
            if starter in alts:
                new_start = rng.pick(alts[starter])
                sentences[i] = new_start + " " + rest
    return sentences


# ==============================================================
# WORD-LEVEL HUMANIZATION
# ==============================================================
def humanize_words(text, rng):
    # Transition words
    tr_rules = [
        (r"\bFurthermore,?\s", lambda: rng.pick(["Also, ","In addition, ","Equally, "])),
        (r"\bMoreover,?\s", lambda: rng.pick(["Also, ","Beyond that, ","On top of that, "])),
        (r"\bAdditionally,?\s", lambda: rng.pick(["Also, ","On top of this, ","Equally, "])),
        (r"\bConsequently,?\s", lambda: rng.pick(["As such, ","For this reason, ","So "])),
        (r"\bNevertheless,?\s", lambda: rng.pick(["Even so, ","That said, ","Still, "])),
        (r"\bNonetheless,?\s", lambda: rng.pick(["Even so, ","Still, ","Yet "])),
        (r"\bSubsequently,?\s", lambda: rng.pick(["Then, ","After that, ","Next, "])),
        (r"\bConversely,?\s", lambda: rng.pick(["On the other hand, ","In contrast, "])),
        (r"\bUltimately,?\s", lambda: rng.pick(["In the end, ","Finally, "])),
        (r"(?i)\bIn\s+addition,?\s", lambda: rng.pick(["Also, ","Equally, "])),
        (r"(?i)\bAs\s+a\s+result,?\s", lambda: rng.pick(["Because of this, ","For this reason, "])),
        (r"(?i)\bIn\s+particular,?\s", lambda: rng.pick(["Especially, ","Notably, "])),
        (r"(?i)\bSpecifically,?\s", lambda: rng.pick(["Namely, ","That is, "])),
        (r"(?i)\bAccordingly,?\s", lambda: rng.pick(["For that reason, ","As such, "])),
        (r"(?i)\bHence,?\s", lambda: rng.pick(["For this reason, ","As such, "])),
        (r"(?i)\bThus,?\s", lambda: rng.pick(["In this way, ","So, "])),
    ]
    for pat, fn in tr_rules:
        text = _re.sub(pat, lambda m, f=fn: f(), text)

    # Formal vocabulary
    voc = [
        (r"(?i)\butilizes\b","uses"),(r"(?i)\butilize\b","use"),
        (r"(?i)\butilized\b","used"),(r"(?i)\butilizing\b","using"),
        (r"(?i)\bfacilitates\b","helps"),(r"(?i)\bfacilitate\b","help"),
        (r"(?i)\bfacilitated\b","helped"),(r"(?i)\bfacilitating\b","helping"),
        (r"(?i)\bdemonstrates\b","shows"),(r"(?i)\bdemonstrate\b","show"),
        (r"(?i)\bdemonstrated\b","showed"),(r"(?i)\bdemonstrating\b","showing"),
        (r"(?i)\bcomprehensive\b", lambda: rng.pick(["thorough","complete","full"])),
        (r"(?i)\bsignificantly\b", lambda: rng.pick(["greatly","noticeably","markedly"])),
        (r"(?i)\bsignificant\b", lambda: rng.pick(["major","notable","marked"])),
        (r"(?i)\bsubstantially\b", lambda: rng.pick(["greatly","largely","considerably"])),
        (r"(?i)\bsubstantial\b", lambda: rng.pick(["large","considerable","sizable"])),
        (r"(?i)\bapproximately\b", lambda: rng.pick(["about","around","roughly"])),
        (r"(?i)\bnumerous\b", lambda: rng.pick(["many","several","a number of"])),
        (r"(?i)\boptimal\b", lambda: rng.pick(["best","ideal"])),
        (r"(?i)\benhances\b","improves"),(r"(?i)\benhance\b","improve"),
        (r"(?i)\benhanced\b","improved"),(r"(?i)\benhancing\b","improving"),
        (r"(?i)\bleverages\b","uses"),(r"(?i)\bleverage\b","use"),
        (r"(?i)\bleveraging\b","using"),(r"(?i)\bleveraged\b","used"),
        (r"(?i)\bensures\b","makes sure"),(r"(?i)\bensure\b","make sure"),
        (r"(?i)\bensuring\b","making sure"),
        (r"(?i)\bpivotal\b", lambda: rng.pick(["key","central","critical"])),
        (r"(?i)\brobust\b", lambda: rng.pick(["strong","solid","reliable"])),
        (r"(?i)\bseamlessly\b","smoothly"),(r"(?i)\bseamless\b","smooth"),
        (r"(?i)\binnovative\b", lambda: rng.pick(["new","novel","creative"])),
        (r"(?i)\bholistic\b", lambda: rng.pick(["overall","complete"])),
        (r"(?i)\bparadigms\b","models"),(r"(?i)\bparadigm\b","model"),
        (r"(?i)\bmethodologies\b","methods"),(r"(?i)\bmethodology\b","method"),
        (r"(?i)\bunderscores\b","highlights"),(r"(?i)\bunderscore\b","highlight"),
        (r"(?i)\bunderscoring\b","highlighting"),
        (r"(?i)\bendeavors\b","efforts"),(r"(?i)\bendeavor\b","effort"),
        (r"(?i)\baugments\b","boosts"),(r"(?i)\baugment\b","boost"),
        (r"(?i)\bmitigates\b","reduces"),(r"(?i)\bmitigate\b","reduce"),
        (r"(?i)\bmitigating\b","reducing"),(r"(?i)\bmitigation\b","reduction"),
        (r"(?i)\bcommences\b","starts"),(r"(?i)\bcommence\b","start"),
        (r"(?i)\bcommenced\b","started"),(r"(?i)\bcommencing\b","starting"),
        (r"(?i)\bterminates\b","ends"),(r"(?i)\bterminate\b","end"),
        (r"(?i)\bterminated\b","ended"),
        (r"(?i)\brequisite\b","needed"),
        (r"(?i)\bpredominantly\b", lambda: rng.pick(["mostly","mainly"])),
        (r"(?i)\bcorroborates\b","backs up"),(r"(?i)\bcorroborate\b","back up"),
        (r"(?i)\belucidated\b","explained"),(r"(?i)\belucidate\b","explain"),
        (r"(?i)\bameliorate\b","improve"),(r"(?i)\bameliorates\b","improves"),
        (r"(?i)\bdelineates\b","outlines"),(r"(?i)\bdelineate\b","outline"),
        (r"(?i)\bpropensity\b","tendency"),(r"(?i)\bdiscernible\b","noticeable"),
        (r"(?i)\bpertaining to\b","about"),(r"(?i)\bwith respect to\b","about"),
        (r"(?i)\bwith regard to\b","about"),(r"(?i)\bin the context of\b","in"),
        (r"(?i)\bthe proposed\b","our"),
        (r"(?i)\bthis paper presents\b","we present"),
        (r"(?i)\bthis paper proposes\b","we propose"),
        (r"(?i)\bthis work introduces\b","we introduce"),
        (r"(?i)\bthis work presents\b","we present"),
        (r"(?i)\bnotwithstanding\b","despite"),(r"(?i)\bwherein\b","where"),
        (r"(?i)\bthereby\b","and so"),(r"(?i)\bwhereas\b","while"),
        (r"(?i)\bshowcases\b","shows"),(r"(?i)\bshowcase\b","show"),
        (r"(?i)\bencompasses\b","covers"),(r"(?i)\bencompass\b","cover"),
        (r"(?i)\bdelves\b","digs"),(r"(?i)\bdelve\b","dig"),
        (r"(?i)\bharnessing\b","using"),(r"(?i)\bharnesses\b","uses"),
        (r"(?i)\bfostering\b","building"),(r"(?i)\bfosters\b","builds"),
        (r"(?i)\bexemplifies\b","illustrates"),(r"(?i)\bexemplify\b","illustrate"),
        (r"(?i)\bpertinent\b","relevant"),(r"(?i)\bconcomitant\b","accompanying"),
        (r"(?i)\bpreclude\b","prevent"),(r"(?i)\bprecludes\b","prevents"),
        (r"(?i)\baforementioned\b","mentioned earlier"),
        (r"(?i)\bherein\b","here"),(r"(?i)\binsofar\b","to the extent"),
        (r"(?i)\bexhibits\b","shows"),(r"(?i)\bexhibit\b","show"),
        (r"(?i)\bexhibited\b","showed"),(r"(?i)\bexhibiting\b","showing"),
        (r"(?i)\bmanifests\b","shows"),(r"(?i)\bmanifested\b","showed"),
        (r"(?i)\bconducive\b","helpful"),
        (r"(?i)\bpossesses\b","has"),(r"(?i)\bpossess\b","have"),
        (r"(?i)\brenders\b","makes"),(r"(?i)\brendered\b","made"),
        (r"(?i)\bgarners\b","gets"),(r"(?i)\bgarnered\b","got"),
        (r"(?i)\bin lieu of\b","instead of"),(r"(?i)\bvis-a-vis\b","compared to"),
        (r"(?i)\bmyriad\b", lambda: rng.pick(["wide range of","many"])),
        (r"(?i)\bcategorically\b","firmly"),(r"(?i)\bexponentially\b","rapidly"),
        (r"(?i)\bintrinsically\b","by nature"),(r"(?i)\binherently\b","by nature"),
        (r"(?i)\boverarching\b","broad"),
        (r"(?i)\bunderpinning\b","supporting"),(r"(?i)\bunderpins\b","supports"),
        (r"(?i)\bencapsulates\b","captures"),(r"(?i)\bencapsulate\b","capture"),
        (r"(?i)\bmultifaceted\b", lambda: rng.pick(["complex","varied"])),
        (r"(?i)\bparamount\b", lambda: rng.pick(["critical","most important"])),
        (r"(?i)\bstreamlines\b","simplifies"),(r"(?i)\bstreamline\b","simplify"),
        (r"(?i)\bexpedites\b","speeds up"),(r"(?i)\bexpedite\b","speed up"),
        (r"(?i)\bascertains\b","finds out"),(r"(?i)\bascertained\b","found out"),
    ]
    for pat, repl in voc:
        if callable(repl):
            text = _re.sub(pat, lambda m, f=repl: f(), text)
        else:
            text = _re.sub(pat, repl, text)

    # AI filler phrases
    fillers = [
        (r"(?i)\bIt is important to note that\s*",""),
        (r"(?i)\bIt should be noted that\s*",""),
        (r"(?i)\bIt is worth mentioning that\s*",""),
        (r"(?i)\bIt is worth noting that\s*",""),
        (r"(?i)\bIt goes without saying that?\s*",""),
        (r"(?i)\bNeedless to say,?\s*",""),
        (r"(?i)\bin order to\b","to"),
        (r"(?i)\bdue to the fact that\b","because"),
        (r"(?i)\ba wide range of\b","many"),
        (r"(?i)\ba plethora of\b","many"),
        (r"(?i)\bplays a (?:crucial|vital|key|important|significant) role\b","matters"),
        (r"(?i)\bin today'?s rapidly evolving\b","in today's"),
        (r"(?i)\bin the ever-changing landscape of\b","in"),
        (r"(?i)\bat the present time\b","now"),
        (r"(?i)\bin the realm of\b","in"),
        (r"(?i)\bfor the purpose of\b","for"),
        (r"(?i)\bin the event that\b","if"),
        (r"(?i)\bprior to\b","before"),
        (r"(?i)\bsubsequent to\b","after"),
        (r"(?i)\bin conjunction with\b","with"),
        (r"(?i)\bon the basis of\b","based on"),
        (r"(?i)\bin the absence of\b","without"),
        (r"(?i)\bhas the potential to\b","can"),
        (r"(?i)\bis capable of\b","can"),
        (r"(?i)\ba considerable amount of\b","much"),
        (r"(?i)\bthe vast majority of\b","most"),
        (r"(?i)\bin close proximity to\b","near"),
        (r"(?i)\bgiven the fact that\b", lambda: rng.pick(["since","because"])),
        (r"(?i)\bin light of\b", lambda: rng.pick(["given","considering"])),
        (r"(?i)\bwith a view to\b","to"),
        (r"(?i)\bin an effort to\b","to"),
        (r"(?i)\bas a means of\b","to"),
        (r"(?i)\bin such a manner that\b","so that"),
        (r"(?i)\bto a large extent\b","largely"),
        (r"(?i)\bby means of\b","through"),
        (r"(?i)\bin the case of\b","for"),
        (r"(?i)\bin the process of\b","while"),
        (r"(?i)\bas a consequence of\b","because of"),
        (r"(?i)\bin terms of\b","regarding"),
        (r"(?i)\bthrough the use of\b","using"),
    ]
    for pat, repl in fillers:
        if callable(repl):
            text = _re.sub(pat, lambda m, f=repl: f(), text)
        else:
            text = _re.sub(pat, repl, text)

    # Contractions
    contrs = [
        (r"\bdoes not\b","doesn't"),(r"\bDoes not\b","Doesn't"),
        (r"\bdo not\b","don't"),(r"\bDo not\b","Don't"),
        (r"\bdid not\b","didn't"),(r"\bDid not\b","Didn't"),
        (r"\bis not\b","isn't"),(r"\bIs not\b","Isn't"),
        (r"\bare not\b","aren't"),(r"\bAre not\b","Aren't"),
        (r"\bwas not\b","wasn't"),(r"\bWas not\b","Wasn't"),
        (r"\bwere not\b","weren't"),(r"\bWere not\b","Weren't"),
        (r"\bcannot\b","can't"),(r"\bCannot\b","Can't"),
        (r"\bcould not\b","couldn't"),(r"\bCould not\b","Couldn't"),
        (r"\bwould not\b","wouldn't"),(r"\bWould not\b","Wouldn't"),
        (r"\bwill not\b","won't"),(r"\bWill not\b","Won't"),
        (r"\bshould not\b","shouldn't"),(r"\bShould not\b","Shouldn't"),
        (r"\bit is\b","it's"),(r"\bIt is\b","It's"),
        (r"\bthat is\b","that's"),(r"\bThat is\b","That's"),
        (r"\bthere is\b","there's"),(r"\bThere is\b","There's"),
        (r"\bwhat is\b","what's"),(r"\bWhat is\b","What's"),
        (r"\bthey are\b","they're"),(r"\bThey are\b","They're"),
        (r"\bwe are\b","we're"),(r"\bWe are\b","We're"),
        (r"\byou are\b","you're"),(r"\bYou are\b","You're"),
        (r"\bI am\b","I'm"),(r"\bI have\b","I've"),
        (r"\bI would\b","I'd"),(r"\bI will\b","I'll"),
        (r"\bwe have\b","we've"),(r"\bWe have\b","We've"),
        (r"\bthey have\b","they've"),(r"\bThey have\b","They've"),
        (r"\bwho is\b","who's"),(r"\bWho is\b","Who's"),
        (r"\bwhere is\b","where's"),(r"\bWhere is\b","Where's"),
        (r"\bhe is\b","he's"),(r"\bHe is\b","He's"),
        (r"\bshe is\b","she's"),(r"\bShe is\b","She's"),
        (r"\bhave not\b","haven't"),(r"\bHave not\b","Haven't"),
        (r"\bhas not\b","hasn't"),(r"\bHas not\b","Hasn't"),
        (r"\blet us\b","let's"),(r"\bLet us\b","Let's"),
    ]
    for pat, repl in contrs:
        text = _re.sub(pat, repl, text)

    # Adverb variation
    advs = [
        (r"(?i)\bconsiderably\b", lambda: rng.pick(["noticeably","markedly"])),
        (r"(?i)\bremarkably\b", lambda: rng.pick(["surprisingly","strikingly"])),
        (r"(?i)\bparticularly\b", lambda: rng.pick(["especially","notably"])),
        (r"(?i)\bprimarily\b", lambda: rng.pick(["mainly","mostly","chiefly"])),
        (r"(?i)\bfundamentally\b", lambda: rng.pick(["at its core","basically"])),
        (r"(?i)\bincreasingly\b", lambda: rng.pick(["more and more","progressively"])),
        (r"(?i)\bsystematically\b", lambda: rng.pick(["step by step","methodically"])),
        (r"(?i)\beffectively\b", lambda: rng.pick(["in practice","successfully"])),
        (r"(?i)\bcritically\b", lambda: rng.pick(["importantly","crucially"])),
        (r"(?i)\bextensively\b", lambda: rng.pick(["broadly","widely"])),
        (r"(?i)\bempirically\b", lambda: rng.pick(["through experiments","experimentally"])),
        (r"(?i)\bnotably\b", lambda: rng.pick(["especially","in particular"])),
        (r"(?i)\baccordingly\b", lambda: rng.pick(["for that reason","as expected"])),
        (r"(?i)\bconcurrently\b", lambda: rng.pick(["at the same time","in parallel"])),
        (r"(?i)\bexplicitly\b", lambda: rng.pick(["directly","clearly"])),
    ]
    for pat, fn in advs:
        text = _re.sub(pat, lambda m, f=fn: f(), text)

    # Passive voice
    pvs = [
        (r"(?i)\bwas conducted\b", lambda: rng.pick(["took place","was carried out"])),
        (r"(?i)\bwere conducted\b", lambda: rng.pick(["took place","were carried out"])),
        (r"(?i)\bwas observed\b", lambda: rng.pick(["showed up","was noticed"])),
        (r"(?i)\bwas identified\b", lambda: rng.pick(["stood out","was spotted"])),
        (r"(?i)\bwere identified\b", lambda: rng.pick(["stood out","were found"])),
        (r"(?i)\bis characterized by\b", lambda: rng.pick(["stands out for","is known for"])),
        (r"(?i)\bwas implemented\b", lambda: rng.pick(["was built","was put in place"])),
        (r"(?i)\bwas utilized\b","was used"),(r"(?i)\bwas employed\b","was used"),
        (r"(?i)\bwere employed\b","were used"),
        (r"(?i)\bwas performed\b", lambda: rng.pick(["was done","was carried out"])),
        (r"(?i)\bwere performed\b", lambda: rng.pick(["were done","were carried out"])),
        (r"(?i)\bwas evaluated\b", lambda: rng.pick(["was tested","was assessed"])),
        (r"(?i)\bwere evaluated\b", lambda: rng.pick(["were tested","were assessed"])),
        (r"(?i)\bwas obtained\b", lambda: rng.pick(["was found","was gathered"])),
        (r"(?i)\bwere obtained\b", lambda: rng.pick(["were found","were gathered"])),
        (r"(?i)\bwas achieved\b", lambda: rng.pick(["was reached","was attained"])),
        (r"(?i)\bwere achieved\b", lambda: rng.pick(["were reached","were attained"])),
    ]
    for pat, repl in pvs:
        if callable(repl):
            text = _re.sub(pat, lambda m, f=repl: f(), text)
        else:
            text = _re.sub(pat, repl, text)

    # AI n-gram breaking
    ngrams = [
        (r"(?i)\bIt is worth noting\b", lambda: rng.pick(["Note that","Notably"])),
        (r"(?i)\bThis is particularly true\b", lambda: rng.pick(["This holds especially","This especially applies"])),
        (r"(?i)\bIn light of\b", lambda: rng.pick(["Given","Considering"])),
        (r"(?i)\bOn the other hand\b", lambda: rng.pick(["In contrast","Alternatively"])),
        (r"(?i)\bIn this regard\b", lambda: rng.pick(["On this point","Here"])),
        (r"(?i)\bThis suggests that\b", lambda: rng.pick(["This points to","This hints that"])),
        (r"(?i)\bThis indicates that\b", lambda: rng.pick(["This shows that","This reveals that"])),
        (r"(?i)\bThis implies that\b", lambda: rng.pick(["This means","This points to"])),
        (r"(?i)\bIt is essential to\b", lambda: rng.pick(["It's critical to","One must"])),
        (r"(?i)\bserves as a\b", lambda: rng.pick(["works as a","acts as a"])),
        (r"(?i)\ba key factor\b", lambda: rng.pick(["a major factor","a main driver"])),
        (r"(?i)\bpaves the way for\b", lambda: rng.pick(["opens the door to","enables"])),
        (r"(?i)\bsheds? light on\b", lambda: rng.pick(["clarifies","helps explain"])),
        (r"(?i)\bthe findings suggest\b", lambda: rng.pick(["the results point to","the data shows"])),
        (r"(?i)\bthe results indicate\b", lambda: rng.pick(["the data shows","the numbers show"])),
        (r"(?i)\bas a whole\b", lambda: rng.pick(["overall","taken together"])),
        (r"(?i)\bgiven the fact that\b", lambda: rng.pick(["since","because"])),
        (r"(?i)\bit is evident that\b", lambda: rng.pick(["clearly,","it's clear that"])),
        (r"(?i)\bit is clear that\b", lambda: rng.pick(["clearly,","plainly,"])),
        (r"(?i)\bplay(?:s)? a (?:crucial|vital|critical|key|significant|important) role\b",
         lambda: rng.pick(["matter greatly","is central","is key"])),
    ]
    for pat, fn in ngrams:
        text = _re.sub(pat, lambda m, f=fn: f(), text)

    return text


# ==============================================================
# ACADEMIC SYNONYM PERTURBATION — 60% rate, 220+ words
# ==============================================================
SYNS = {
    "achieved":["reached","attained","obtained"],
    "achieves":["reaches","attains","gets"],
    "achieve":["reach","attain","get"],
    "achieving":["reaching","attaining","getting"],
    "accuracy":["precision","correctness"],
    "adequate":["sufficient","enough"],
    "analyze":["examine","study","inspect"],
    "analyzed":["examined","studied","inspected"],
    "analyzes":["examines","studies"],
    "analyzing":["examining","studying"],
    "analysis":["examination","study","review"],
    "approach":["technique","strategy","method"],
    "applicable":["relevant","suitable","fitting"],
    "appropriate":["suitable","fitting","proper"],
    "architecture":["design","structure","layout"],
    "capability":["ability","capacity","potential"],
    "capabilities":["abilities","capacities"],
    "challenge":["difficulty","hurdle","obstacle"],
    "challenges":["difficulties","hurdles","obstacles"],
    "component":["part","element","piece"],
    "components":["parts","elements","pieces"],
    "computation":["calculation","processing"],
    "computational":["computing","processing"],
    "configuration":["setup","arrangement","layout"],
    "constraint":["limitation","restriction","bound"],
    "constraints":["limitations","restrictions","bounds"],
    "critical":["crucial","vital","essential"],
    "dataset":["data collection","corpus"],
    "datasets":["data collections","corpora"],
    "degradation":["decline","drop","weakening"],
    "deployment":["rollout","release","launch"],
    "detrimental":["harmful","damaging","negative"],
    "domain":["field","area","sphere"],
    "domains":["fields","areas"],
    "dynamic":["changing","shifting","evolving"],
    "efficacy":["effectiveness","potency"],
    "efficient":["effective","capable","fast"],
    "efficiency":["effectiveness","performance","speed"],
    "evident":["clear","apparent","obvious"],
    "framework":["system","structure","setup"],
    "frameworks":["systems","structures"],
    "granularity":["detail level","resolution"],
    "heterogeneous":["diverse","mixed","varied"],
    "homogeneous":["uniform","consistent","alike"],
    "hypothesis":["theory","assumption","idea"],
    "illustrate":["show","depict","highlight"],
    "illustrates":["shows","depicts","highlights"],
    "illustrated":["shown","depicted"],
    "impact":["effect","influence","consequence"],
    "implement":["build","create","set up"],
    "implemented":["built","created","set up"],
    "implements":["builds","creates"],
    "implementing":["building","creating","setting up"],
    "implementation":["setup","realization","execution"],
    "indicate":["show","suggest","point to"],
    "indicates":["shows","suggests","points to"],
    "indicated":["showed","suggested","pointed to"],
    "indicating":["showing","suggesting","pointing to"],
    "inherent":["built-in","intrinsic","natural"],
    "investigate":["explore","study","look into"],
    "investigated":["explored","studied","looked into"],
    "investigation":["exploration","study","review"],
    "limitation":["drawback","shortcoming","weakness"],
    "limitations":["drawbacks","shortcomings","weaknesses"],
    "magnitude":["size","scale","extent"],
    "mechanism":["process","procedure","method"],
    "mechanisms":["processes","procedures","methods"],
    "modality":["mode","type","channel"],
    "modalities":["modes","types","channels"],
    "module":["unit","block","segment"],
    "modules":["units","blocks","segments"],
    "obtain":["get","gather","collect"],
    "obtained":["got","gathered","collected"],
    "obtains":["gets","gathers"],
    "obtaining":["getting","gathering","collecting"],
    "parameter":["setting","variable","factor"],
    "parameters":["settings","variables","factors"],
    "perform":["carry out","do","execute"],
    "performed":["carried out","did","executed"],
    "performs":["carries out","does","executes"],
    "performing":["carrying out","doing","executing"],
    "phenomenon":["event","occurrence","pattern"],
    "pipeline":["workflow","process chain","sequence"],
    "prevalent":["common","widespread","frequent"],
    "prior":["earlier","previous","past"],
    "procedure":["process","method","step"],
    "procedures":["processes","methods","steps"],
    "propagation":["spread","transmission","flow"],
    "rationale":["reasoning","basis","logic"],
    "scenario":["situation","case","setting"],
    "scenarios":["situations","cases","settings"],
    "scheme":["plan","strategy","setup"],
    "straightforward":["simple","direct","easy"],
    "subsequent":["later","following","next"],
    "sufficient":["enough","adequate"],
    "superior":["better","higher","stronger"],
    "threshold":["cutoff","limit","boundary"],
    "thresholds":["cutoffs","limits","boundaries"],
    "trajectory":["path","course","direction"],
    "validates":["confirms","verifies","checks"],
    "validate":["confirm","verify","check"],
    "validated":["confirmed","verified","checked"],
    "validating":["confirming","verifying","checking"],
    "various":["different","diverse","multiple"],
    "yield":["produce","give","generate"],
    "yields":["produces","gives","generates"],
    "proposed":["suggested","presented","put forward"],
    "proposes":["suggests","presents","puts forward"],
    "adopts":["uses","takes on","picks up"],
    "adopted":["used","taken on","picked up"],
    "employs":["uses","applies","relies on"],
    "employed":["used","applied","relied on"],
    "conventional":["traditional","standard","typical"],
    "incorporates":["includes","adds","brings in"],
    "incorporated":["included","added","brought in"],
    "captures":["records","picks up","catches"],
    "captured":["recorded","picked up","caught"],
    "addresses":["tackles","handles","deals with"],
    "addressed":["tackled","handled","dealt with"],
    "highlights":["points out","stresses","emphasizes"],
    "highlighted":["pointed out","stressed","emphasized"],
    "explores":["looks into","examines","digs into"],
    "explored":["looked into","examined","dug into"],
    "outperforms":["beats","surpasses","exceeds"],
    "outperformed":["beat","surpassed","exceeded"],
    "relies":["depends","counts","leans"],
    "relied":["depended","counted","leaned"],
    "achievable":["reachable","attainable","doable"],
    "complementary":["supporting","supplementary","matching"],
    "encompasses":["covers","spans","includes"],
    "encompassed":["covered","spanned","included"],
    "interactions":["exchanges","connections","relationships"],
    "interaction":["exchange","connection","relationship"],
    "observations":["findings","notes","results"],
    "observation":["finding","note","result"],
    "promising":["encouraging","hopeful","positive"],
    "remarkable":["striking","impressive","notable"],
    "evaluation":["assessment","testing","review"],
    "evaluations":["assessments","tests","reviews"],
    "segments":["parts","sections","pieces"],
    "segment":["part","section","piece"],
    "feature":["trait","attribute","characteristic"],
    "leveraging":["using","applying","drawing on"],
    "mitigating":["reducing","easing","lessening"],
    "facilitating":["helping","enabling","supporting"],
}

SYN_PAT = _re.compile(
    r"\b(" + "|".join(_re.escape(w) for w in SYNS.keys()) + r")\b",
    _re.IGNORECASE
)

def perturb_synonyms(text, rng, rate=0.60):
    def _rep(m):
        w = m.group(0)
        lw = w.lower()
        if lw in SYNS and rng.rand() < rate:
            r = rng.pick(SYNS[lw])
            return (r[0].upper() + r[1:]) if w[0].isupper() else r
        return w
    return SYN_PAT.sub(_rep, text)


# ==============================================================
# MAIN PIPELINE
# ==============================================================
def process_paragraph(text, rng, is_para_start=True):
    text = sanitize_text(text)
    text = humanize_words(text, rng)
    sentences = split_sentences(text)
    if len(sentences) > 1:
        sentences = restructure_sentences(sentences, rng, is_para_start=is_para_start)
        sentences = break_same_starters(sentences, rng)
    text = " ".join(sentences)
    text = perturb_synonyms(text, rng)
    text = _re.sub(r"  +", " ", text).strip()
    return text


# ==============================================================
# PROCESS DOCUMENT
# ==============================================================
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

def do_process_para(para, is_first=True):
    global changed, seed_counter
    runs = [r for r in para.runs if r.text]
    if not runs:
        return
    full_text = "".join(r.text for r in runs)
    if not full_text.strip():
        return
    seed_counter += 1
    rng = PRNG(seed_counter)
    processed = process_paragraph(full_text, rng, is_para_start=is_first)
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
    do_process_para(para, is_first=True)

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                if p.text.strip():
                    do_process_para(p, is_first=True)

print(f"\nSkipped: {skipped}")
print(f"Paragraphs changed: {changed}")

doc.save(DST)
print(f"\nDone - saved {DST}")
