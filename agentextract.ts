// PLAIN TEXT VERSION 

//
const ATTRIBUTION_WRAP_WINDOW = 2

// An "On ... wrote:" attribution can get glued onto the sender's reply text.
// The cutter scans line-by-line, so it can't see it. 
// This splits it onto its own line. Guards ensure we only split genuine headers (date + email + verb).
// Supports all languages in the data set.
const GLUED_ATTRIBUTION_RE = /([^\r\n>])[ \t]+((?:On|Le|Il|W dniu|Op|Am|På|Den|Em|El|Vào|في)\s+(?=[^\r\n]{0,240}(?:wrote|sent|a écrit|escribió|ha scritto|escreveu|schrieb|schreef|geschreven|verzond|skrev|napisał|написал|đã viết|كتب)\s*:)(?=[^\r\n]{0,240}(?:<[^>]+@[^>]+>|\b[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+))(?=[^\r\n]{0,240}(?:\b20\d\d\b|\b\d{1,2}:\d{2}\b|\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\b|[\u0660-\u0669]{1,2}[:：][\u0660-\u0669]{2}))[^\r\n]{0,240}(?:wrote|sent|a écrit|escribió|ha scritto|escreveu|schrieb|schreef|geschreven|verzond|skrev|napisał|написал|đã viết|كتب)\s*:)/gi
const splitGluedAttributions = (text: string): string => text.replace(GLUED_ATTRIBUTION_RE, '$1\n$2')

// Returns whether this line looks like the START of quoted history.
// The 'cutter' walks the email and cuts at the first line that returns true.
// Takes all lines + index i because some signals need to peek ahead at i+1, i+2.
const isQuoteSignal = (lines: string[], i: number): boolean => {
    const line = lines[i].trim()
    // #1. Leading ">".
    if (line.startsWith('>')) {
        const nextLine = (lines[i + 1] ?? '').trim()
        // Guard: Two ">" lines in a row = a real quoted block, not an innocent ">" in the email. 
        if (nextLine.startsWith('>')) return true
        // Nested quoting (">>", ">>>"): strip the markers and re-test recursively. 
        const dequoted = line.replace(/^(?:>\s?)+/, '')
        // Guard: only recurse if stripping > changed the line — avoids wasted re-checks and infinite self-calls.
        if (dequoted !== line && isQuoteSignal([dequoted], 0)) return true
    }
    // #2. -----Original Message-----
    // i: Matches any casing (Original / original / ORIGINAL).
    // Supports all languages in the data set.
    if (/^-+\s*(original message|ursprüngliche nachricht|oprindelig meddelelse|reply message|antwort nachricht)/i.test(line)) return true
    // #3. _____________________ (ambiguous - could be a signature divider).
    // Therefore, we only treat it as quoted history IFF a real quote signal follows it.
    // Scan the next 3 non-empty lines (blanks skipped, not counted) and recurse on each.
    if (/^_{5,}\s*$/.test(line)) {
        for (let j = i + 1, seen = 0; j < lines.length && seen < 3; j++) {
            if (lines[j].trim() === '') continue
            seen++
            if (isQuoteSignal(lines, j)) return true
        }
    }
    // #4. On ... wrote: 
    if (/^On\b/i.test(line)) {
        // This line can wrap across 2-3 lines, splitting "On" from "wrote:". 
        // Stitch line i + up to ATTRIBUTION_WRAP_WINDOW more into one string.
        const parts: string[] = []
        for (let j = i; j < lines.length && j <= i + ATTRIBUTION_WRAP_WINDOW; j++) {
            const part = lines[j].trim()
            // Stop at the first blank line — real attribution has no blanks inside it.
            if (part === '') break
            parts.push(part)
        }
        const joined = parts.join(' ')
        // "On ... wrote:" trusted on its own. 
        // Adding an email-address guard here would be net-negative: more false negatives.
        if (/^On\b.*wrote:/i.test(joined)) return true
        // "Sent:" variant: weaker signal, so require an email address as corroboration. 
        // Here that same guard is net-positive: blocks "sent:" miscuts in prose.
        if (/^On\b.*sent:/i.test(joined) && /\S+@\S+/.test(joined)) return true
    }
    // -------- On ... wrote: -------- 
    if (/^-{2,}.*\bwrote\b:?\s*-{2,}\s*$/i.test(line)) {
        return true
    }
    //  Samsung phone signature (Sent from Samsung Galaxy ... [email] ... wrote)
    if (/sent from samsung\b.*\S+@\S+.*\bwrote\b/i.test(line)) return true

    // #5. "From:" header block pasted above a quoted message.
    // "From" is also a normal word, so the label alone isn't enough — we corroborate.
    // Supports all languages in the data set.
    const fromLabel = /^\s*(from|von|de|van|da|fra|från)\s*:/i.exec(line)
    if (fromLabel) {
        // Risky short labels (de/van/da = common words) require more proof.
        const shortLabel = /^(de|van|da|fra|från)$/i.test(fromLabel[1])
        const emailRe = /<[^>]+@[^>]+>|\b[^\s@]+@[^\s@]+\.[^\s@]+/
        const hasEmail = emailRe.test(line) // email on the From: line itself
        const next = lines.slice(i + 1, i + 1 + 3).map((l) => l.trim())  // look at the 3 lines below
        const hasSentOrDate = next.some((l) => /^\s*(sent|date|gesendet|envoyé|enviado|enviada|datum|verzonden|inviato|skickat|sendt)\s*:/i.test(l)) // a Sent:/Date: line in those 3
        const emailInBlock = next.some((l) => emailRe.test(l)) // an email in those 3 lines
        // Essentially: the weaker the trigger word, the more evidence we demand before we mark it as quoted history.
        if (hasEmail || (!shortLabel && hasSentOrDate) || (shortLabel && hasSentOrDate && emailInBlock)) return true
    }
    // #6. Foreign attribution verbs — same job as #4, for non-English emails.
    // These don't start with "On", so we anchor on the verb instead of the opener.
    const FOREIGN_VERB = /(a écrit|escribió|ha scritto|escreveu|schrieb|schreef|geschreven|verzond|skrev|napisał|написал|đã viết|yazdı|a scris)/i
    // a. Foreign "wrote" verb + line ends in a colon (the colon filters out mid-sentence prose).
    if (FOREIGN_VERB.test(line) && /:\s*$/.test(line)) return true
    // b. Same wrap problem as #4: the verb+colon can go onto the next line.
    // Only stitch when the line has a year + time - stitching is risky, so we pay that
    // cost only when the line truly looks like an attribution start. Date+time is the tell-tale sign.
    if (/\b20\d\d\b/.test(line) && /\d{1,2}:\d{2}/.test(line)) {
        const parts: string[] = []
        for (let j = i; j < lines.length && j <= i + ATTRIBUTION_WRAP_WINDOW; j++) {
            const part = lines[j].trim()
            if (part === '') break
            parts.push(part)
            const joined = parts.join(' ')
            // Check after each line glued on — return the instant verb+colon appears.
            if (FOREIGN_VERB.test(joined) && /:\s*$/.test(joined)) return true
        }
    }
    // #7. Chinese and Arabic 
    // Note [:：]: Chinese/Arabic text often uses a full-width colon (：), so every test accepts both : and ：.
    // a. Chinese 
    if (/写道\s*[:：]\s*$/.test(line)) return true // 写道 = "wrote", line ends in a colon 
    if (/^发件人[\s:：]/.test(line)) return true // 发件人 = "sender" header; followed by space OR colon 
    // b. Arabic (stored in logical order, so verb-before-colon holds despite right-to-left display)  
    if (/كتب\s*[:：]\s*$/.test(line)) return true  // كتب = "wrote", line ends in a colon 
    // Wrapped attribution whose visible tail is just "<email>:" — date + time + a bracketed
    // address ending in a colon is enough to flag it as quoted history.
    if (/\b20\d\d\b/.test(line) && /\d{1,2}:\d{2}/.test(line) && /<[^>]*@[^>]*>\s*[:：]\s*$/.test(line)) {
        return true 
    }
    return false
}

// Forwards 
// Detects ---------- Forwarded message ---------- or begin forwarded message:
const RE_FWD_LINE = /^(?:-+\s*forwarded message\s*-+|begin forwarded message:?)\s*$/i

// DSNs
// Machine-generated technical headers (most reliable). 
const DSN_FIELD = /^(reporting-mta:|final-recipient:|original-recipient:|diagnostic-code:|action:\s*failed|status:\s*[45]\.\d)/i
// Human-readable opening sentence (sometimes no DSN_FIELD's). 
const DSN_HEADLINE = /^[\s*]*(delivery (has )?failed|address not found|undelivered mail|mail delivery (failed|subsystem)|delivery status notification|message not delivered|your message\b[^\n]{0,80}?(?:wasn'?t|couldn'?t be|could not be) delivered|returned mail|message blocked|this is the mail system|this message was created automatically by mail|i'?m sorry to have to inform you|could not be delivered)/i
// Mid-line phrase the two others can't reach. 
const DSN_HEADLINE_MIDLINE = /\brejected your message\b/i
// Bounces quote the original email back, so the 'cutter' would wrongly chop them.
// This gate catches a bounce so the caller can keep it whole instead of cutting.
// Only the first real line is checked (skipping [bracketed] tags and bare URLs) —
// bounces announce at the top, which avoids false-positives from emails that just
// mention "delivery failed" in the body. 
export const isTrueDsn = (text: string): boolean => {
    const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean)
    const first = lines.find((l) => !/^\[.*\]$/.test(l) && !/^https?:\/\/\S+$/i.test(l)) ?? ''
    return DSN_HEADLINE.test(first) || DSN_FIELD.test(first) || DSN_HEADLINE_MIDLINE.test(first)
}

// Keep inline replies whole. 
// Pattern we look for: quote -> prose -> quote. 
const hasInlineReply = (lines: string[]): boolean => {
    let quotedAbove = false // Passed through a real quote block. 
    let proseSince = false // Saw genuine reply-prose since that quote. 
    for (let i = 0; i < lines.length; i++) {
        const t = lines[i].trim()
        if (t === '') continue
        if (t.startsWith('>')) {
            // Confirm a real quoted block (neighbour line also ">"), not a lone ">".
            const run = (lines[i + 1] ?? '').trim().startsWith('>') || (lines[i - 1] ?? '').trim().startsWith('>')
            if (run) {
                if (quotedAbove && proseSince) return true // quote -> prose -> quote: it's interleaved
                quotedAbove = true // record this quote
                proseSince = false // reset: prose must reappear AFTER this quote to count - avoiding false-positives
            }
            continue
        }
        // Skip attribution lines (On...wrote: etc.) — that's quote machinery, not reply prose.
        if (isQuoteSignal(lines, i)) continue
        // Proxy for "substantial NEW content worth keeping the whole quote for" — not just "is it a reply?".
        // Validated the ≥20/≥4 bar against the dataset: short inter-quote lines are 0.25% of mail and essentially all debris, 
        // so loosening would add ~86 false keep-wholes and catch ~0 real cases.  
        if (quotedAbove && (lines[i - 1] ?? '').trim() === '' && t.length >= 20 && t.split(/\s+/).length >= 4) {
            proseSince = true
        }
    }
    return false
}

// Plain-text 'cutter' function.
export const extractNewContent = (text: string): string => {
    // Delivery-failure bounces (keep whole).
    if (isTrueDsn(text)) return text.trim()
    
    // Un-glue any runaway "On...wrote:" so the cutter can see it (see splitGluedAttributions).
    const preparedText = splitGluedAttributions(text)
    // Split into lines (handles \n and Outlook's \r\n).
    const lines = preparedText.split(/\r?\n/) // Note: when we decide to "keep-whole" we return `text`, not these lines — by design. 

    // Inline-reply guard (keep-whole).
    if (hasInlineReply(lines)) return text.trim()

    // FW:/Fwd: pasted forward (keep-whole).
    const firstContent = lines.find((l) => l.trim() !== '')?.trim() ?? ''
    if (/^(?:fw|fwd):\s*\S/i.test(firstContent)) return text.trim()
    
    // 'Cutter' loop.
    for (let i = 0; i < lines.length; i++) {
        // Forwarded-message divider (keep whole) — inside the loop because it can appear anywhere, not just the top.
        if (RE_FWD_LINE.test(lines[i].trim())) {
            return text.trim()
        }
        // Cut point — the first quote signal marks where old history begins.
        if (isQuoteSignal(lines, i)) {
            
            // Rescue a stranded signature: sometimes the sender's real sign-off sits
            // BELOW the quoted block (reply text → quote → "--" → signature). A naive
            // cut at the quote would delete that signature too. So look for it.
            //
            // Candidate "--" markers below the cut whose preceding block (from the cut) is ≥50% quoted
            // history — the ≥50% guard confirms the marker sits just after the quote, not some unrelated
            // "--". Among those, choosing WHICH one to reattach from handles two mirror-image layouts:
            //   • quote → "--" → sig → "--" → disclaimer/list-footer: reattaching from the last "--"
            //     would drop the sender's real sig with the quote. So we anchor on the EARLIEST marker
            //     whose signature body is substantial (multi-line) — the real sig — keeping it.
            //   • quote → "--" → inner (quoted) sig → "--" → real sig: the intervening line is short
            //     quoted junk (a quoted sig that lost its ">"), so no early marker is "substantial" and
            //     we fall back to the latest qualifying marker — the inner sig stays dropped.
            const cut = i
            const markers: number[] = []
            for (let j = cut + 1; j < lines.length; j++) {
                if (!/^--\s*$/.test(lines[j])) continue
                const body = lines.slice(cut + 1, j).filter((l) => l.trim())
                const quoted = body.filter((l) => l.startsWith('>')).length
                if (body.length && quoted >= body.length * 0.5) markers.push(j)
            }
            let sigIdx = -1
            // Earliest marker whose sig body (up to the next marker) is a substantial ≥2-line block.
            for (const j of markers) {
                const nextMarker = markers.find((k) => k > j) ?? lines.length
                const sigBody = lines.slice(j + 1, nextMarker).filter((l) => l.trim())
                if (sigBody.length >= 2) {
                    sigIdx = j
                    break
                }
            }
            // Fallback: latest qualifying marker (the original behavior) — single-line sigs, inner-sig case.
            if (sigIdx < 0 && markers.length) sigIdx = markers[markers.length - 1]
            if (sigIdx >= 0) {
                // Keep reply (above cut) + signature (from "--" down), drop the quoted middle.
                return [...lines.slice(0, cut), ...lines.slice(sigIdx)].join('\n').trim()
            }
            // Normal cut: keep everything above the quote, drop everything below.
            return lines.slice(0, cut).join('\n').trim()
        }
    }

    return text.trim() // no signal found -> keep the whole email
}

/////////////////////////////////////////////////////////////
// HTML VERSION 
// HTML counterpart to isQuoteSignal. HTML has no lines, so we search for a LIST of
// markers and cut at whichever appears earliest.
const HTML_QUOTE_MARKERS: RegExp[] = [
    // Client wrappers — each client marks quotes its own way.
    /<div[^>]*gmail_quote/i,                     // Gmail
    /<blockquote/i,                             // standard quote tag (Apple Mail etc.)
    /<div[^>]*id=["']?appendonsend/i,          // Outlook: new message ends here        
    /<hr[^>]*width:\s*98%/i,                   // Outlook: separator rule        
    /<div[^>]*id=["']?(x_)?divRplyFwdMsg/i,   // Outlook: reply/forward marker
    
    // Mirrors #4 — "On ... wrote:". Tags can sit between them, so allow 200 chars;
    // the inner guard stops "On" pairing with a "wrote:" in a different block.
    /\bOn\b(?:(?!<\/(?:div|p|td|tr|table|blockquote|body|ul|ol|li))[\s\S]){0,200}?wrote:/i,
    // Mirrors #5 — "From:" header. Lookbehind keeps "From" a standalone word, then needs an email.
    /(?<![A-Za-z])(?:From|Von|De|Van|Da|Fra|Från):(?:(?!<\/(?:div|p|td|tr|table|blockquote|body))[\s\S]){0,250}?[^\s@]+@[^\s@]+\.[^\s@]+/i,
    // Mirrors #6 — foreign "wrote" verbs ending in a colon.
    /(a [ée]crit|escribi[óo]|ha scritto|escreveu|schrieb|schreef|geschreven|verzond|skrev|napisał|написал|đã viết|yazd[ıi]|a scris)(?:(?!<\/(?:div|p|td|tr|table|blockquote|body))[\s\S]){0,80}?[:：]/i,
    // Mirrors #7 — Chinese + Arabic ([:：] accepts the full-width colon).
    /写道\s*[:：]/,
    /发件人\s*[:：]/,
    /كتب\s*[:：]/,
    // Mirrors #2 — "-----Original Message-----" divider.
    /-+\s*(original message|ursprüngliche nachricht|oprindelig meddelelse|reply message|antwort nachricht)\s*-*/i,
]

// Strips HTML/CSS comments out of the sliced fragment so the result doesn't carry invisible junk.
const cleanHTML = (html: string): string => {
    return html 
        .replace(/<!--[\s\S]*?-->/g, '') // removes HTML comments
        .replace(/\/\*[\s\S]*?\*\//g, '') // removes CSS comments
        .trim()
}

// Flatten an HTML fragment to plain visible text (block tags → newlines, others stripped).
const visibleText = (html: string): string =>
    html.replace(/<\/(?:div|p|blockquote|tr|td|li|table)>/gi, '\n').replace(/<(?:br|hr)\s*\/?>/gi, '\n')
        .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&amp;/gi, '&')
        .replace(/[ \t]+/g, ' ').trim()


// Scans blocks BOTTOM-up and drops the trailing run of noise/empty blocks,
// stopping at the first real content. Trailing-only (not top-down like stripNoise) so a sig sitting
// BELOW a leading noise line ("Sent from my iPhone" then the sig) is kept, not amputated.
const trimTrailingNoise = (frag: string): string => {
    const parts = frag.split(/(?<=<\/(?:div|p|tr|td|li|table|blockquote)>)/i)
    let cut = parts.length
    for (let i = parts.length - 1; i >= 0; i--) {
        const v = parts[i].replace(/<[^>]+>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/&[a-z]+;/gi, ' ').replace(/\s+/g, ' ').trim()
        if (!v) continue // empty trailing block → keep scanning up
        if (NOISE_MARKERS.some((re) => re.test(v))) { cut = i; continue } // trailing noise → drop, keep going
        break // hit real content → everything above stays
    }
    return parts.slice(0, cut).join('')
}

// Sometimes the sender's real sign-off sits BELOW the quoted block (reply → quote → signature).
// A plain cut at the quote would delete it too. This reattaches it, but ONLY when the trailing
// chunk is unmistakably a signature, never a forward or more quoting (see guards below).
const reattachTrailingSignature = (html: string, cutAt: number): string | undefined => {
    // Quoted history is wrapped in <blockquote>. The trailing sig sits after the LAST one —
    // lastIndexOf jumps past the whole nested reply chain in one step.
    const close = html.toLowerCase().lastIndexOf('</blockquote>')
    if (close < cutAt) return undefined // no blockquote after the cut → nothing to reattach
    const quoteEnd = close + '</blockquote>'.length // position just past the last </blockquote>
    const rawTail = html.slice(quoteEnd) // // grab the raw HTML after the quote (tags kept — this is what we reattach)
    const tail = visibleText(rawTail) // flattened to plain text, for the guard checks below

    // Guard 1: must LOOK like a signature — Gmail's signature class (highest-precision), a "--"
    // delimiter, or a contact token / sign-off word for the non-Gmail clients.
    const looksLikeSig = /class=["'][^"']*gmail_signature/i.test(rawTail) || /^\s*--\s*$/m.test(tail) ||
        /[^\s@]+@[^\s@]+\.[^\s@]|\+?\d[\d\-\s().]{7,}|https?:\/\/|linkedin\.com|\b(?:regards|thanks|sincerely|cheers|best)\b/i.test(tail)
    // Guard 2: reject anything that is really MORE quoting or a forward — not a signature.
    const hasMoreQuote = /\bOn\b[\s\S]{0,200}?wrote:|-+\s*forwarded message|begin forwarded message:|^\s*From:\s.{0,80}@/im.test(tail)
    if (looksLikeSig && !hasMoreQuote) {
        // Skip the empty wrapper junk between the quote close and the real trailing content
        // (stray </div>, empty <div><br></div> gaps) so the spliced HTML stays clean.
        const lead = /^(?:\s*<\/(?:div|p|span|blockquote|td|tr|table|li)>|\s*<(?:div|p)[^>]*>\s*(?:<br[^>]*>\s*)*<\/(?:div|p)>|\s*<br[^>]*>)*/i.exec(rawTail)?.[0].length ?? 0
        const sig = trimTrailingNoise(rawTail.slice(lead)) // drop any disclaimer/footer below the sig
        const sigVis = visibleText(sig)
        // Guard 3: what's LEFT after trimming noise must be a real, short signature — bounded by VISIBLE
        // CHARACTER count (not line count: HTML splits sigs into many short lines/table-cells, so a line
        // cap misfires on whitespace-padded sigs). Empty (tail was pure noise) → bail to a plain cut.
        if (sigVis && sigVis.length <= 400) {
            return html.slice(0, cutAt) + sig // reply + signature, quoted middle + footer dropped
        }
    }
    return undefined // not confidently a signature → caller does a plain cut
}

// HTML 'cutter' functon. 
export const extractFromHtml = (html: string): string => {
    // Bounce guard (keep whole). isTrueDsn only reads text, so flatten the HTML first.
    const asText = html
        .replace(/<(?:br|\/p|\/div|\/tr|\/td|\/li|hr)[^>]*>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&amp;/gi, '&').replace(/&nbsp;/gi, ' ')
    if (isTrueDsn(asText)) return cleanHTML(html).trim()

    // Find the EARLIEST quote marker. Markers are a list (not in document order),
    // so we check all of them and keep the minimum position.
    let cutAt = html.length
    for (const marker of HTML_QUOTE_MARKERS) {
        const pos = html.search(marker)
        if (pos !== -1 && pos < cutAt) cutAt = pos
    }

    // Forwarded-message divider before the cut (keep whole).
    const fwd = html.search(/-+\s*forwarded message\s*-+|begin forwarded message:/i)
    if (fwd !== -1 && fwd < cutAt) {
        return cleanHTML(html).trim()
    }

    // Did we find a quote marker? If so, back up to the tag start and try the signature reattach.
    if (cutAt < html.length) {
        const tagStart = html.lastIndexOf('<', cutAt)
        if (tagStart !== -1) cutAt = tagStart // back up so we don't slice mid-tag
        const rescued = reattachTrailingSignature(html, cutAt)
        if (rescued !== undefined) return cleanHTML(rescued).trim()
    }

    // No reattach (or no marker found): plain cut — keep everything before cutAt.
    return cleanHTML(html.slice(0, cutAt)).trim()
}

/////////////////////////////////////////////////////////////
// NOISE STRIPPING (final cleanup) - "sent from iphone", subscribe buttons, etc. 
const NOISE_MARKERS: RegExp[] = [
    // 
    /^sent from (?:my |outlook|mail|yahoo|samsung)/i,
    /^get outlook for (?:ios|android)\b/i, 
    // 
    /\bview (?:this email|it|this)?\s*in (?:your )?(?:web )?browser\b/i,
    /\byou(?:'re| are) receiving this (?:e-?mail|message|because)\b/i,
    /\byou received this (?:e-?mail|message)\b/i,
    /\bmanage (?:your )?(?:email )?(?:preferences|subscription)\b/i,
    /(?:to unsubscribe|unsubscribe (?:here|from this|at any|below|now)\b|click [^\n]{0,25}\bunsubscribe\b|\bunsubscribe\b[^\n]{0,40}(?:https?:\/\/|www\.))/i,
    // 
    /\bthis (?:e-?mail|email|message|communication)\b[^.\n]{0,60}\b(?:confidential|privileged)\b/i,
    /\bconfidentiality (?:notice|statement|warning)\b/i,
    /\bif you are not the intended recipient\b/i,
    // Disclaimer phrasings seen stranded under quotes (Mass General, corporate legal footers).
    /\b(?:intended|designated) (?:solely|only) for\b/i,
    /\bthis (?:e-?mail|email|message) is not secure\b/i,
    /\b(?:please )?(?:do not|don'?t) print this (?:e-?mail|email|message)\b/i,
    /^\s*your personal data\b/i,
]

export const stripNoise = (text: string): string => {
    const lines = text.split(/\r?\n/)
    let cut = lines.length
    for (let i = 0; i < lines.length; i++) {
        if (NOISE_MARKERS.some((re) => re.test(lines[i]))) { cut = i; break }
    }
    return lines.slice(0, cut).join('\n').trim()
}

// HTML counterpart to stripNoise — same NOISE_MARKERS, but block-aware. Walk the blocks top-down
// and cut at the FIRST one whose visible text is boilerplate noise (mobile sig, footer, disclaimer),
// dropping it and everything after. Mirrors stripNoise's "cut at the first noise line" on HTML.
export const stripNoiseHtml = (html: string): string => {
    const parts = html.split(/(?<=<\/(?:div|p|tr|td|li|table|blockquote)>)/i)
    let out = ''
    for (const part of parts) {
        const v = part.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/&[a-z]+;/gi, ' ').replace(/\s+/g, ' ').trim()
        if (v && NOISE_MARKERS.some((re) => re.test(v))) break // first noise block → cut here, drop the rest
        out += part
    }
    return out.trim()
}

/////////////////////////////////////////////////////////////
// Email entry point — cut the quote and strip trailing noise for the text and/or html
// body. Named extractEmailBody (rather than a generic "extract") to leave room for
// future extractors over non-email content, e.g. text-like attachments.
export const extractEmailBody = ({ text, html }: { text?: string; html?: string }) => {
    const result: { extractedText?: string; extractedHtml?: string } = {}

    if (text) {
        try {
            // Cut the quote, then strip trailing boilerplate noise (mobile sigs, footers, disclaimers).
            result.extractedText = stripNoise(extractNewContent(text))
        } catch (error) {
            console.warn('extractEmailBody failed on text:', error)
        }
    }

    if (html) {
        try {
            result.extractedHtml = stripNoiseHtml(extractFromHtml(html))
        } catch (error) {
            console.warn('extractEmailBody failed on html:', error)
        }
    }

    return result
} 