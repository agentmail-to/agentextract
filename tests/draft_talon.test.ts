import { describe, it, expect } from 'vitest'

import { extractContent, stripNoise } from '../draft_talon'

// ---------------------------------------------------------------------------
// Synthetic Tests for my code (draft_talon.ts)
// ---------------------------------------------------------------------------

// Build a multi-line body without fighting template-literal indentation.
const t = (...lines: string[]) => lines.join('\n')
// Run the text path and return just the extracted text.
const ex = (text: string) => extractContent({ text }).extracted_text
// Run the HTML path and return just the extracted html.
const exHtml = (html: string) => extractContent({ html }).extracted_html

// Plain Text Tests

describe('extractContent — plain text — core signals', () => {
    // The one big CHANGE (from our data). Agent-generated mail uses ISO timestamps
    // with NO comma: "On 2025-05-15T17:36:04.391Z ... wrote:". Talon's attribution
    // regex requires a comma (it was built for human-written mail), so it never
    // fires here and leaves the "On ... wrote:" line dangling at the bottom of its
    // output. We cut on "On ... wrote:" comma-or-not, so we remove the isolated line.
    // Source: our own data.
    it('cuts at an "On ... wrote:" line WITHOUT a comma (the beachhead)', () => {
        const text = t(
            'Hi,',
            '',
            'Thanks for following up — this works for me.',
            '',
            'Best,',
            'Neil',
            '',
            'On 2025-05-15T17:36:04.391Z Neil Banerjee <neil@mercorrecruiting.com> wrote:',
            '',
            '> Hey Corban,',
            '> some older content here',
        )
        expect(ex(text)).toBe(t('Hi,', '', 'Thanks for following up — this works for me.', '', 'Best,', 'Neil'))
    })

    // Human-typed Gmail attribution "On <date>, ... wrote:" — has the comma Talon
    // expects, so Talon handles it and we must too. We cut at the line and keep
    // everything above. Canonical Talon case (from Talon README / test suite). 
    it('cuts at a Gmail "On <date>, ... wrote:" line (with commas)', () => {
        const text = t(
            'Hi Hunter,',
            '',
            'Thank you for rescheduling. I’ll connect with Holly today.',
            '',
            'Best regards,',
            'Deepthee',
            '',
            'On Fri, Nov 7, 2025 at 8:58 AM Hunter <hunter@hunterscouts.co> wrote:',
            '',
            '> Hi Deepthee,',
            '> the screening chat has been rescheduled',
        )
        expect(ex(text)).toBe(t('Hi Hunter,', '', 'Thank you for rescheduling. I’ll connect with Holly today.', '', 'Best regards,', 'Deepthee'))
    })

    // Wrapped attribution.
    // When the sender's name/email is long, the client wraps the line, so "wrote:"
    // lands on the NEXT line and the "On ..." line alone has no "wrote:". We peek at a
    // tiny window (this line + up to 2 more) so the wrapped attribution still cuts.
    // Source: our own data (eval_missing.ts). 
    it('cuts at a WRAPPED "On ... wrote:" attribution (wrote: on the next line)', () => {
        const text = t(
            'Hi Muhammad,',
            '',
            'Yes, I am interested in the position.',
            '',
            'On Fri, Apr 11, 2025 at 6:31 PM Unsightly Jacket <',
            'unsightlyjacket@gmail.com> wrote:',
            '',
            '> Hi Muhammad,',
            '> old quoted content',
        )
        expect(ex(text)).toBe(t('Hi Muhammad,', '', 'Yes, I am interested in the position.'))
    })

    // Classic Outlook "-----Original Message-----" separator. The most canonical
    // quote-strip case there is ("Reply\n-----Original Message-----\nQuote" ->
    // "Reply"). We cut at the separator and keep everything above.
    // Source: Talon README.
    it('cuts at an Outlook "-----Original Message-----" line', () => {
        const text = t(
            'Sounds good, see you then.',
            '',
            '-----Original Message-----',
            'From: Bob <bob@example.com>',
            'Sent: Monday, June 1, 2026 9:00 AM',
        )
        expect(ex(text)).toBe('Sounds good, see you then.')
    })

    // The "-----Original Message-----" rule is too strict — real mail varies the dash
    // count and casing ("-------- Original message --------"). We GENERALIZE to: leading
    // dashes + "original message" (any case). Without this the separator is missed and a
    // later signal (the From: line) cuts one line late, orphaning the separator.
    // Source: our data (eval_missing.ts "other" bucket).
    it('cuts at an "Original message" separator with a different dash count and casing', () => {
        const text = t(
            'Thanks, talk soon.',
            '',
            '-------- Original message --------',
            'From: Bob <bob@example.com>',
            'Sent: Monday, June 1, 2026 9:00 AM',
        )
        expect(ex(text)).toBe('Thanks, talk soon.')
    })

    // Localized "Original Message" separators (Signal 2): German "Ursprüngliche Nachricht",
    // Danish "Oprindelig meddelelse". Same dashed divider, just translated. Source: our data + Talon.
    it('cuts at a German "-----Ursprüngliche Nachricht-----" separator', () => {
        const text = t('Danke, bis bald.', '', '-----Ursprüngliche Nachricht-----', 'Von: Bob <bob@example.com>', 'Gesendet: Montag')
        expect(ex(text)).toBe('Danke, bis bald.')
    })

    it('cuts at a Danish "-------- Oprindelig meddelelse --------" separator', () => {
        const text = t('Tak, vi tales ved.', '', '-------- Oprindelig meddelelse --------', 'Fra: Bob <bob@example.com>', 'Sendt: mandag')
        expect(ex(text)).toBe('Tak, vi tales ved.')
    })

    // Mobile variant: the separator and the "From:" header sit on the SAME line, glued
    // together ("-------- Original message --------From: Hunter <...> Date: ..."). The
    // rule must match the separator at the START of the line. Source: our data.
    it('cuts at an "Original message" separator glued to a From: header on one line', () => {
        const text = t(
            'Sure, calling you now.',
            '',
            '-------- Original message --------From: Hunter <hunter@hunterscouts.co> Date: 2025-04-19 6:00 a.m.',
            '> earlier quoted content',
        )
        expect(ex(text)).toBe('Sure, calling you now.')
    })

    // Outlook underscore separator. 
    // Outlook drops a long run of underscores on its own line right above the quoted "From:" header block. 
    // We cut at that line and keep everything above. 
    // The run must be LONG enough that a stray "___" in real content won't trip it.
    it('cuts at an Outlook underscore-separator line', () => {
        const text = t(
            'Thanks, that works for me.',
            '',
            '________________________________',
            'From: Hunter <hunter@hunterscouts.co>',
            'Sent: Monday, June 1, 2026 9:00 AM',
            'Subject: Re: interview',
            '',
            '> earlier quoted content',
        )
        expect(ex(text)).toBe('Thanks, that works for me.')
    })

    // Signal 4 TIGHTENED (over-cut audit fix). A "____" line is a cut signal ONLY when it
    // introduces a real quoted header/attribution — i.e. one of the next few non-blank lines is
    // itself a quote signal. An underscore above a sender DISCLAIMER (boilerplate, no header
    // after) must stay whole: disclaimers are kept in extraction, removed only downstream.
    it('does NOT cut at an underscore line followed by a disclaimer (no header after)', () => {
        const text = t(
            'Thanks!',
            '',
            '________________________________',
            '',
            'This email and its attachments are confidential and intended solely for the addressee.',
            'If you are not the intended recipient, please delete it.',
        )
        // The quote-cutter does NOT cut at the lone underscore; stripNoise then drops the disclaimer.
        expect(ex(text)).toBe(t('Thanks!', '', '________________________________'))
    })

    // Regression: underscore followed by an "On … wrote:" attribution still cuts at the underscore.
    it('still cuts at an underscore line followed by an "On … wrote:" attribution', () => {
        const text = t(
            'Sounds good.',
            '',
            '________________________________',
            'On Mon, Jun 1, 2026 at 9:00 AM Bob <bob@example.com> wrote:',
            '> earlier quoted content',
        )
        expect(ex(text)).toBe('Sounds good.')
    })

    // Edge: an underscore line at the very end with nothing after it is not a quote separator.
    it('does NOT cut at a trailing underscore line with nothing after it', () => {
        const text = t('Got it, talk soon.', '', '________________________________')
        expect(ex(text)).toBe(text)
    })

    // Outlook "From:" header block — the biggest remaining gap (58% of "cut less").
    // Outlook quotes by pasting the original headers (From:/Sent:/To:/Subject:) with no
    // ">". We cut at the "From:" line, but ONLY when it looks like a real header — here,
    // it carries an email address. The guard is what stops us cutting innocent content
    // that merely starts with "From:" (see the keep/guard test below).
    // Source: our own data (eval_missing.ts) + Talon's FromColonRegexp.
    it('cuts at an Outlook "From:" header block (guarded by an email address)', () => {
        const text = t(
            'Thanks, will review and revert shortly.',
            '',
            'From: Hunter <hunter@hunterscouts.co>',
            'Sent: Monday, June 1, 2026 9:00 AM',
            'To: Me',
            'Subject: Re: interview',
            '',
            '> earlier quoted content',
        )
        expect(ex(text)).toBe('Thanks, will review and revert shortly.')
    })

    // Same Outlook header block, but the "From:" line has NO email — it's only
    // recognizable as a header because a "Sent:" line follows. The guard must accept
    // this too (a following Sent:/Date: line), matching Talon. Without the broader guard
    // these ~6.5% of From: blocks would be missed. Source: our data (scan_from.ts).
    it('cuts at a "From:" header block identified by a following "Sent:" line (no email)', () => {
        const text = t(
            'Got it, thanks.',
            '',
            'From: Hunter',
            'Sent: Monday, June 1, 2026 9:00 AM',
            'Subject: Re: interview',
            '',
            '> earlier quoted content',
        )
        expect(ex(text)).toBe('Got it, thanks.')
    })

    // Oldest, most universal convention: a ">"-prefixed line marks quoted text.
    // Our baseline cuts on the FIRST ">". Source: shared email convention (also
    // handled by Talon).
    it('cuts at the first bare ">" quote line', () => {
        const text = t(
            'Yes, that works.',
            '',
            '> previous message line',
            '> another quoted line',
        )
        expect(ex(text)).toBe('Yes, that works.')
    })

    // CHANGE 1 — robustness: a LONE ">" (not followed by another ">" line) must NOT
    // cut. A lone ">" appears in real content ("experience required: > 5 years") and is
    // too noisy as a standalone signal. Real single-line quotes almost always sit under
    // an "On…wrote:" attribution that Signal 3 already catches. FAILING before fix.
    it('does not cut on a single isolated ">" line (no preceding attribution)', () => {
        const text = t(
            'Here are the requirements:',
            '',
            '> 5 years of experience required',
            '',
            'Let me know if you qualify.',
        )
        expect(ex(text)).toBe(text)
    })

    // CHANGE 1 regression — 2+ consecutive ">" lines still cut at the first.
    // The run-of-2 rule must not break the normal multi-line quoting case.
    it('still cuts at the first ">" in a run of 2+ consecutive ">" lines (regression)', () => {
        const text = t('Got it.', '', '> first quoted line', '> second quoted line')
        expect(ex(text)).toBe('Got it.')
    })

    // CHANGE 1 regression — single ">" under an attribution still cuts (at Signal 3).
    // Signal 3 fires at the "On…wrote:" line (which comes first), so the cut lands there
    // before Signal 1 even sees the lone ">". The lone-">" guard must not break this.
    it('cuts at the attribution when a single ">" follows it (Signal 3 takes priority)', () => {
        const text = t(
            'Sounds good!',
            '',
            'On Mon, Jun 1, 2026 at 9:00 AM Bob <bob@example.com> wrote:',
            '> single quoted line',
        )
        expect(ex(text)).toBe('Sounds good!')
    })
})

// Foreign-language attribution tests (the Latin batch). The trick: key on the distinctive
// VERB (a écrit, escribió, schrieb, ...) which never appears in English content, plus a
// trailing colon. That one shape handles BOTH word orders — verb-last (French/Spanish/
// Italian/Portuguese) AND verb-then-name (German/Dutch/Scandinavian) — because we only
// require "foreign verb somewhere, ':' at the end". We deliberately do NOT key on begin-
// words (Le/Il/Am) — too common, too false-positive-prone. Spanish + Italian are bonus:
// Python-Talon misses those too, so we beat both libs. Source: our data (scan_corpus_langs.js).
describe('extractContent — plain text — foreign-language attributions', () => {
    const cases: Array<[string, string]> = [
        ['French', 'Le dim. 13 avr. 2025 à 12:24, Hunter <hunter@hunterscouts.co> a écrit :'],
        ['Spanish', 'El 13 abr 2025 a las 12:24, Hunter <hunter@hunterscouts.co> escribió:'],
        ['German (verb-then-name)', 'Am Mo., 14. Apr. 2025 um 10:18 Uhr schrieb Hunter <hunter@hunterscouts.co>:'],
        ['Italian', 'Il giorno 9 apr 2025 alle ore 15:08 Hunter <hunter@hunterscouts.co> ha scritto:'],
        ['Portuguese', 'Em qui., 5 de fev. de 2026, Hunter <hunter@hunterscouts.co> escreveu:'],
        ['Dutch', 'Op wo 14 mei 2025 om 14:05 schreef Hunter <hunter@hunterscouts.co>:'],
        ['Dutch (geschreven)', 'Op 28 feb 2026 om 14:22 heeft Hunter <hunter@hunterscouts.co> geschreven:'],
        ['Scandinavian', 'Den 13.08.2025 22:27, skrev AgentMail:'],
        ['Polish', 'niedz., 8 lut 2026, 10:57 użytkownik Hemiro <hemiro@agentmail.to> napisał:'],
        ['Vietnamese', 'Vào Th 6, 28 thg 2 2026 lúc 14:22 Hunter <hunter@hunterscouts.co> đã viết:'],
    ]
    for (const [lang, attribution] of cases) {
        it(`cuts at a ${lang} attribution line`, () => {
            const text = t('Reply body here.', '', attribution, '> quoted old content')
            expect(ex(text)).toBe('Reply body here.')
        })
    }
})

// WRAPPED foreign attribution — a long name/email pushed the verb+colon onto the NEXT line, so the
// date-led lead-in line has no verb. Signal 6a (verb + ':' on one line) fired on the verb line, cutting
// one line too LOW and KEEPING the lead-in ("El mar, ... Hunter (email)") as if it were sender content
// (over-keep). Signal 6b mirrors Signal 3's wrap window, anchored on the date fingerprint (year + clock),
// to cut at the lead-in instead. Source: foreign_verb over-keep residuals (31 cases).
describe('extractContent — plain text — WRAPPED foreign attributions', () => {
    it('cuts at the LEAD-IN line of a wrapped Spanish "escribió:" attribution', () => {
        const text = t(
            'Yeah! Do you think that you can give me a call?',
            '',
            'El mar, 22 abr 2025 a la(s) 3:46 p.m., Hunter (hunter@hunterscouts.co)',
            'escribió:',
            '',
            '> Dear Gabriel,',
            '> I noticed there were connectivity issues during your interview.',
        )
        expect(ex(text)).toBe('Yeah! Do you think that you can give me a call?')
    })

    it('cuts at the lead-in line of a wrapped Portuguese "escreveu:" attribution', () => {
        const text = t(
            'Obrigado, combinado!',
            '',
            'Em qui., 5 de fev. de 2026 às 14:22, Hunter <hunter@hunterscouts.co>',
            'escreveu:',
            '> conteúdo antigo citado',
        )
        expect(ex(text)).toBe('Obrigado, combinado!')
    })

    it('cuts at the lead-in of a wrapped attribution even with the quote flush under the verb (no blank)', () => {
        const text = t(
            'Sorry but it still is not working.',
            'El mar, 22 abr 2025 a la(s) 4:04 p.m., Hunter (hunter@hunterscouts.co)',
            'escribió:',
            '> Hi Gabriel, I rescheduled your interview.',
        )
        expect(ex(text)).toBe('Sorry but it still is not working.')
    })

    // SAFETY: the wrap window only fires when a foreign VERB completes the attribution, so a plain
    // sender line that merely carries a year + clock time (directly above a quote) must NOT be cut.
    it('does NOT over-cut a sender line that has a date/time but no foreign verb', () => {
        const text = t(
            'Let us meet on 12 May 2025 at 3:45 to finalize the plan.',
            'On Mon, 1 Jun 2026 at 9:00 AM Bob <bob@example.com> wrote:',
            '> old quoted content',
        )
        expect(ex(text)).toBe('Let us meet on 12 May 2025 at 3:45 to finalize the plan.')
    })
})

// Chinese attribution / header — a separate signal because Chinese uses no Latin verb
// and FULL-WIDTH punctuation (：not :). Two shapes, mirroring our English On…wrote: + From::
//   - 写道 ("wrote"), verb-LAST, at end of line:  "...<email>于2025年12月13日...写道："
//   - 发件人 ("From") header at start of line:     "发件人：Marco F. <...>"
// Both markers are distinctive (they don't occur in normal content), so no guard needed
// beyond anchoring 写道 to the end. Source: our data (scan_chinese.js).
describe('extractContent — plain text — Chinese', () => {
    // 写道 (wrote) reply attribution — verb at the end, full-width colon.
    it('cuts at a Chinese "写道：" (wrote) attribution line', () => {
        const text = t(
            '好的，谢谢。',
            '',
            "Daming's Hunter <daming.jiang@viahiringagents.com>于2025年12月13日 周六下午8:55写道：",
            '> 引用的旧内容',
        )
        expect(ex(text)).toBe('好的，谢谢。')
    })

    // 发件人 (From) header block — start of line, full-width colon.
    it('cuts at a Chinese "发件人：" (From) header line', () => {
        const text = t(
            '收到，谢谢！',
            '',
            '发件人：Marco F. <marcoferraro@agentmail.to>',
            '发送时间：2026-01-01 23:01',
            '收件人：someone@example.com',
            '',
            '> 引用内容',
        )
        expect(ex(text)).toBe('收到，谢谢！')
    })

    // Space-separated 发件人 variant (one client, cnbn.cn): all header fields on ONE line,
    // separated by spaces instead of colons. The rule must accept whitespace OR a colon
    // after 发件人 — but still require *some* delimiter so content like "发件人是张三" is safe.
    // Source: our data (scan_chinese.js).
    it('cuts at a Chinese "发件人" header with space-separated fields (no colon)', () => {
        const text = t(
            '好的。',
            '',
            '发件人  吕姝婷<lvshuting@cnbn.cn> 发送日期  2026年04月24日 15:54 收件人  guoyoumeng<guoyoumeng@cnbn.cn>',
            '> 引用内容',
        )
        expect(ex(text)).toBe('好的。')
    })

    // Guard: 写道 mid-sentence (not at line end) must NOT cut — anchoring it to the end
    // is what separates a real attribution from content that merely mentions "wrote".
    it('does not cut a line that mentions 写道 mid-sentence', () => {
        const text = t('大家好，', '', '他写道的内容非常好：请看下面。', '谢谢！')
        expect(ex(text)).toBe(text)
    })
})

// Arabic (RTL). كتب ("wrote") is verb-last like Chinese 写道, with the colon at the LOGICAL end of
// the stored string. We key on verb + end-colon only, so RTL order and Arabic-Indic date digits
// (٢٠٢٥) never matter. Source: our data (discover_patterns.ts — Arabic attributions survived).
describe('extractContent — plain text — Arabic', () => {
    // كتب: attribution (verb-last, ASCII colon at the logical end).
    it('cuts at an Arabic "… كتب:" (wrote) attribution', () => {
        const text = t(
            'شكرا لك.',
            '',
            'في الجمعة، ١٨ أبريل ٢٠٢٥، ٤:٥٤ م Hunter <hunter@hunterscouts.co> كتب:',
            '> المحتوى المقتبس القديم',
        )
        expect(ex(text)).toBe('شكرا لك.')
    })

    // Guard: كتب mid-sentence (not verb-last at the end-colon) must NOT cut — the end-colon anchor
    // is what separates a real attribution from content mentioning the root كتب ("wrote/book").
    it('does not cut a line that mentions كتب mid-sentence', () => {
        const text = t('مرحبا،', '', 'لقد كتب التقرير بشكل جيد: انظر أدناه.', 'شكرا!')
        expect(ex(text)).toBe(text)
    })
})

// Localized header blocks — the foreign-language siblings of Signal 5's English "From:".
// Same header shape, just a localized From-label (German Von, FR/ES/PT De, Dutch Van, IT Da).
// GUARD asymmetry: from/von cut on email OR a following localized Sent/Date; the short common-
// word labels de/van/da require an EMAIL on the line (no Sent/Date-only path), since "de"/"van"/
// "da" are everyday words. Source: our data (combined_audit.ts) + the de/van/da over-cut audit.
describe('extractContent — plain text — localized headers', () => {
    // CUT: German Von: header, email on the line.
    it('cuts at a German "Von:" header block (email on the line)', () => {
        const text = t(
            'Danke, das passt für mich.',
            '',
            'Von: Hunter <hunter@hunterscouts.co>',
            'Gesendet: Montag, 1. Juni 2026 09:00',
            'An: mich',
            'Betreff: Re: Interview',
            '',
            '> alter zitierter Inhalt',
        )
        expect(ex(text)).toBe('Danke, das passt für mich.')
    })

    // CUT: "von" keeps the FULL guard — no email here, but a following "Gesendet:" line makes
    // it cut. Contrast with the short de/van/da labels, which would NOT cut without an email.
    it('cuts at a German "Von:" header with no email but a following "Gesendet:" line', () => {
        const text = t(
            'Alles klar, danke.',
            '',
            'Von: Hunter',
            'Gesendet: Montag, 1. Juni 2026 09:00',
            'Betreff: Re: Interview',
            '',
            '> alter zitierter Inhalt',
        )
        expect(ex(text)).toBe('Alles klar, danke.')
    })

    // CUT: Spanish De: header — the short labels cut only with an email on the line.
    it('cuts at a Spanish "De:" header block (email on the line)', () => {
        const text = t(
            'Gracias, perfecto.',
            '',
            'De: Hunter <hunter@hunterscouts.co>',
            'Enviado: lunes, 1 de junio de 2026 9:00',
            'Asunto: Re: entrevista',
            '',
            '> contenido citado',
        )
        expect(ex(text)).toBe('Gracias, perfecto.')
    })

    // CUT: real Dutch Van: header, email on the line.
    it('cuts at a Dutch "Van:" header block (email on the line)', () => {
        const text = t(
            'Bedankt, dat werkt.',
            '',
            'Van: Hunter <hunter@hunterscouts.co>',
            'Verzonden: maandag 1 juni 2026 09:00',
            'Aan: mij',
            'Onderwerp: Re: gesprek',
            '',
            '> oude geciteerde inhoud',
        )
        expect(ex(text)).toBe('Bedankt, dat werkt.')
    })

    // CUT: real Italian Da: header, email on the line.
    it('cuts at an Italian "Da:" header block (email on the line)', () => {
        const text = t(
            'Grazie, va bene.',
            '',
            'Da: Hunter <hunter@hunterscouts.co>',
            'Inviato: lunedì 1 giugno 2026 09:00',
            'A: me',
            'Oggetto: Re: colloquio',
            '',
            '> vecchio contenuto citato',
        )
        expect(ex(text)).toBe('Grazie, va bene.')
    })

    // CUT: Danish/Norwegian "Fra:" header, email on the line (short label — needs email).
    it('cuts at a Danish/Norwegian "Fra:" header block (email on the line)', () => {
        const text = t(
            'Tak, det virker.',
            '',
            'Fra: Hunter <hunter@hunterscouts.co>',
            'Sendt: mandag 1. juni 2026 09:00',
            'Til: mig',
            'Emne: Re: samtale',
            '',
            '> gammelt citeret indhold',
        )
        expect(ex(text)).toBe('Tak, det virker.')
    })

    // CUT: Swedish "Från:" header, email on the line.
    it('cuts at a Swedish "Från:" header block (email on the line)', () => {
        const text = t(
            'Tack, det fungerar.',
            '',
            'Från: Hunter <hunter@hunterscouts.co>',
            'Skickat: måndag 1 juni 2026 09:00',
            'Till: mig',
            'Ämne: Re: samtal',
            '',
            '> gammalt citerat innehåll',
        )
        expect(ex(text)).toBe('Tack, det fungerar.')
    })

    // GUARD: "Fra:" as innocent content — "fra" is an everyday Danish/Norwegian word. Trap: over-cut.
    it('keeps a "Fra:" content line that is not a real header', () => {
        const text = t('Hej,', '', 'Fra: og med i morgen er kontoret lukket', '', 'Tak.')
        expect(ex(text)).toBe(text)
    })

    // GUARD: Spanish content line starting with "De:" — no email, no Sent/Date — stays whole.
    // Trap: keying on "De:" alone over-cuts ("de" is a common word).
    it('keeps a "De:" content line that is not a real header', () => {
        const text = t('Hola,', '', 'Lista de tareas:', 'De: revisar el informe y enviarlo', '', 'Gracias.')
        expect(ex(text)).toBe(text)
    })

    // GUARD: a bare "Von:" line, no email, no Sent/Date — stays whole. Trap: a From-word alone cutting.
    it('keeps a bare "Von:" content line that is not a real header', () => {
        const text = t('Hallo,', '', 'Von: hier kommt die Liste der Aufgaben', '', 'Danke.')
        expect(ex(text)).toBe(text)
    })

    // GUARD: "Van:" as innocent content — "van" is an everyday Dutch word / name particle. Trap: over-cut.
    it('keeps a "Van:" content line that is not a real header', () => {
        const text = t('Hoi,', '', 'Van: alles wat we vandaag besproken hebben', '', 'Groeten.')
        expect(ex(text)).toBe(text)
    })

    // GUARD: "Da:" as innocent content — "da" is an everyday Italian word. Trap: over-cut.
    it('keeps a "Da:" content line that is not a real header', () => {
        const text = t('Ciao,', '', 'Da: qui parte la lista delle cose da fare', '', 'Grazie.')
        expect(ex(text)).toBe(text)
    })

    // GUARD (the tightening): a Dutch "Van:" with NO email stays whole even when a "Verzonden:"
    // line follows — the short labels de/van/da require an email, dropping the Sent/Date-only
    // path that from/von keep. Trap: the weak Sent/Date-only signal firing on a common word.
    it('keeps a "Van:" header with no email even when a "Verzonden:" line follows (short label needs email)', () => {
        const text = t(
            'Bedankt.',
            '',
            'Van: Jan',
            'Verzonden: maandag 1 juni 2026 09:00',
            'Onderwerp: Re: gesprek',
            '',
            '> oude inhoud',
        )
        expect(ex(text)).toBe(text)
    })

    // CUT (positive counterpart of the guard above): a short-label "Van:" with NO inline email still cuts
    // when a Sent/Date line confirms a real header AND an email appears on a following block line (Aan:/CC:).
    // Fixes gpt-4o calibration case [47] (kooijmanautar.nl), where OURS used to under-cut and keep the thread.
    it('cuts at a short-label "Van:" header whose email is on the next block line (not inline)', () => {
        const text = t(
            'Beste,',
            'Hierbij de stukken voor de levering.',
            'Met vriendelijke groet,',
            'Rob de Gilder',
            'Van: Relinde Meijer | Kooijman Autar Notarissen Namens Rob de Gilder',
            'Verzonden: donderdag 4 juni 2026 11:44',
            "Aan: 'riccardotrinh@gmail.com' <riccardotrinh@gmail.com>",
            'Onderwerp: Schiekade 121A, 3033 BK Rotterdam',
            '',
            'Dear Mr Trinh and Miss Do,',
            'With regard to the purchase of the registrable property...',
        )
        expect(ex(text)).toBe(t('Beste,', 'Hierbij de stukken voor de levering.', 'Met vriendelijke groet,', 'Rob de Gilder'))
    })
})

// Date-led attributions with NO verb (Signal 8). Russian/Japanese Gmail render the attribution
// as "<date+time> Name <email>:" — no "wrote", so the On/verb signals miss it. We require all
// three fingerprints (year + clock time + trailing "<email>:") to keep the over-cut risk low.
// Source: our data (discover_patterns.ts — Cyrillic 0.19% surviving) + Talon's date+email splitters.
describe('extractContent — plain text — date-led attributions (no verb)', () => {
    // Russian Gmail attribution — ends in "<email>:", no verb.
    it('cuts at a Russian "<date>, Name <email>:" attribution (no verb)', () => {
        const text = t(
            'Спасибо, до встречи.',
            '',
            'вс, 18 мая 2025 г. в 12:46, Hunter <hunter@hunterscouts.co>:',
            '> старое цитируемое содержимое',
        )
        expect(ex(text)).toBe('Спасибо, до встречи.')
    })

    // Japanese Gmail attribution — ends in "<email>:", no verb.
    it('cuts at a Japanese "<date> Name <email>:" attribution (no verb)', () => {
        const text = t(
            '了解しました、ありがとうございます。',
            '',
            '2026年6月1日(月) 0:14 MAKOTO <makoto@agentmail.to>:',
            '> 引用された古い内容',
        )
        expect(ex(text)).toBe('了解しました、ありがとうございます。')
    })

    // GUARD: a content line ending in "<email>:" but with NO date/time must stay whole — the
    // trailing email alone can't cut. Trap: over-cutting any sentence that ends with an address.
    it('keeps a content line ending in "<email>:" with no date or time', () => {
        const text = t('Hi,', '', 'For anything urgent reach me at <me@example.com>:', 'thanks!')
        expect(ex(text)).toBe(text)
    })

    // GUARD: a line with a year but NOT ending in "<email>:" stays whole (needs all three signals).
    it('keeps a content line with a year and time but no trailing email', () => {
        const text = t('Team,', '', 'On 2025-06-01 at 14:30 we ship the release.', 'Thanks!')
        expect(ex(text)).toBe(text)
    })
})

// Dashed Android attribution (Signal 9). The Android client wraps the attribution in dashes and
// drops the colon: "---- On <date> … wrote ----". Signal 3 misses it (no leading "On", no
// trailing "wrote:"). Safety hedge: a colon-less "wrote" is accepted ONLY when wrapped in dashes.
describe('extractContent — plain text — dashed Android attribution', () => {
    // CUT: the real Android format — "---- On <date> Name <email> wrote ----" (no colon).
    it('cuts at a dashed "---- On … wrote ----" attribution (no colon)', () => {
        const text = t(
            'Sounds good, talk soon.',
            '',
            '---- On Fri, 30 May 2025 07:44:31 -0400 Hunter <hunter@hunterscouts.co> wrote ----',
            '> earlier quoted content',
        )
        expect(ex(text)).toBe('Sounds good, talk soon.')
    })

    // CUT: the bare dashed form with no "On" — "---- Name wrote ----" (Talon's test_android_wrote).
    it('cuts at a dashed "---- Name wrote ----" attribution', () => {
        const text = t('Got it.', '', '---- John Smith wrote ----', '> old quoted content')
        expect(ex(text)).toBe('Got it.')
    })

    // GUARD: a bare "wrote" in prose (NOT dash-wrapped) must NOT cut — the dash-bracket is the
    // whole safety hedge. Trap: keying on "wrote" alone over-cuts normal sentences.
    it('keeps a sentence containing "wrote" that is not dash-wrapped', () => {
        const text = t('Hi,', '', 'Yesterday John wrote the onboarding doc and shared it.', 'Thanks!')
        expect(ex(text)).toBe(text)
    })

    // GUARD: a dashed decoration with NO "wrote" stays whole (needs both dashes AND wrote).
    it('keeps a dashed separator line that does not contain "wrote"', () => {
        const text = t('Team,', '', '-------- Agenda for Monday --------', 'See you there.')
        expect(ex(text)).toBe(text)
    })
})

// Defensive client formats — ABSENT in our corpus but real clients in the wild, so adding them can't regress
// (0 cases here) yet gives coverage for clients this dataset didn't include. Source: Talon SPLITTER_PATTERNS.
describe('extractContent — plain text — defensive client formats', () => {
    // Exim / some clients render "On <date> <somebody> sent:" instead of "wrote:". Signal 3 accepts "sent:"
    // too — but only with an attribution marker (an email) in the window, since "sent:" is common in prose.
    it('cuts at an "On … <email> sent:" attribution (sent-verb variant)', () => {
        const text = t(
            'Sounds good, thanks.',
            '',
            'On Mon, Jun 1, 2026 at 9:00 AM Bob <bob@example.com> sent:',
            '> earlier quoted content',
        )
        expect(ex(text)).toBe('Sounds good, thanks.')
    })

    // GUARD: a sender line that starts with "On" and ends in "sent:" but has NO email is NOT an attribution
    // (e.g. "On Monday the invoices were sent:") — the email requirement keeps it whole.
    it('does NOT over-cut an "On … sent:" sender line with no email in the window', () => {
        const text = t('Hi team,', '', 'On Monday the invoices were sent:', '- one', '- two', 'Thanks!')
        expect(ex(text)).toBe(text)
    })

    // Samsung mobile reply attribution — "Sent from Samsung … <email> wrote:". The email + "wrote" mark it.
    it('cuts at a Samsung "Sent from Samsung … <email> wrote:" attribution', () => {
        const text = t(
            'Got it, will do.',
            '',
            'Sent from Samsung Galaxy smartphone. Bob <bob@example.com> wrote:',
            '> earlier quoted content',
        )
        expect(ex(text)).toBe('Got it, will do.')
    })

    // The plain "Sent from my Samsung" auto-signature is NOT a quote signal (no email, no "wrote"), so the
    // cutter keeps it — then the inline stripNoise removes it as mobile boilerplate.
    it('strips a plain "Sent from my Samsung" mobile signature (not a quote signal)', () => {
        const text = t('Got it, thanks.', '', 'Sent from my Samsung Galaxy smartphone.')
        expect(ex(text)).toBe('Got it, thanks.')
    })
})

// Talon-parity regression locks. We already pass these implicitly; the tests exist so a future
// change can't silently break behaviors Talon explicitly covers. Source: talon-python tests.
describe('extractContent — plain text — Talon-parity regression locks', () => {
    // 3-line wrapped attribution: a long sender pushes "wrote:" onto a 3rd line. Our window is
    // ATTRIBUTION_WRAP_WINDOW=2 (this line + 2 = 3 lines), so the contiguous join still matches.
    // Mirrors Talon's test_quotation_separator_takes_3_lines.
    it('cuts a 3-line wrapped "On … wrote:" attribution', () => {
        const text = t(
            'Yes, interested.',
            '',
            'On Fri, Apr 11, 2025 at 6:31 PM Unsightly Jacket <',
            'unsightlyjacket@gmail.com>',
            'wrote:',
            '> old quoted content',
        )
        expect(ex(text)).toBe('Yes, interested.')
    })

    // Leading whitespace before "On" — isQuoteSignal trims each line, so the attribution still
    // fires. Mirrors Talon's test_pattern_on_date_somebody_wrote_allows_space_in_front.
    it('cuts an "On … wrote:" attribution with leading whitespace', () => {
        const text = t('Sounds good.', '', '   On Mon, Jun 1, 2026 at 9:00 AM Bob <bob@example.com> wrote:', '> old quoted content')
        expect(ex(text)).toBe('Sounds good.')
    })

    // Date written with slashes — Signal 3 keys on "On … wrote:", not the date format, so it
    // fires. Mirrors Talon's test_pattern_on_date_somebody_wrote_date_with_slashes.
    it('cuts an "On <date-with-slashes> … wrote:" attribution', () => {
        const text = t('Got it.', '', 'On 04/11/2025 Bob <bob@example.com> wrote:', '> old quoted content')
        expect(ex(text)).toBe('Got it.')
    })

    // Nested quotations: a reply above a multi-level (>, >>) quote block. First-signal-wins cuts
    // at the attribution, so ALL nesting is removed. Mirrors Talon's test_reply_wraps_nested_quotations.
    it('removes a nested (>, >>) quote block, keeping only the reply', () => {
        const text = t('My reply.', '', 'On Mon, Jun 1 Bob wrote:', '> first level', '>> second level', '> back to first')
        expect(ex(text)).toBe('My reply.')
    })
})

// Newly folded-in additions (from the fresh eval residual scan): lone/nested ">"-quoted attributions,
// Turkish/Romanian foreign verbs, Office-365 / "rejected your message" bounces, and FW:/Fwd: forwards.
describe('extractContent — plain text — lone/nested ">" attributions', () => {
    // A LONE ">" line (no run of two) that is itself a quoted attribution must still cut — otherwise the
    // run-of-2 guard leaves a dangling "> On … wrote:" orphan. We de-quote the line and re-test the signals.
    it('cuts a LONE ">"-quoted "On … wrote:" attribution (no run of two)', () => {
        const text = t('Thanks, that works.', '', '> On 2025-09-28T09:08:45.000Z Patrick <patrick@example.com> wrote:')
        expect(ex(text)).toBe('Thanks, that works.')
    })

    // Nested quoting ("> >", "> > >") prefixes the attribution with several ">" markers; de-quoting strips
    // them all so the bare "On … wrote:" fires.
    it('cuts a nested "> > On … wrote:" quoted attribution', () => {
        const text = t('Got it.', '', '> > On Mon, Jun 1, 2026 at 9:00 AM Bob <bob@example.com> wrote:')
        expect(ex(text)).toBe('Got it.')
    })

    // GUARD: the lone-">" protection survives — a single ">" line of plain content de-quotes to a non-signal
    // and stays whole (the original "> 5 years experience" case must not regress).
    it('still keeps a lone ">" plain-content line (de-quoted content is not a signal)', () => {
        const text = t('Here are the requirements:', '', '> 5 years of experience required', '', 'Let me know.')
        expect(ex(text)).toBe(text)
    })

    // GUARD: a lone ">" line that de-quotes to an innocent "On …" sentence (no "wrote:") must NOT cut.
    it('keeps a lone ">" line that de-quotes to an innocent "On …" sentence', () => {
        const text = t('Note:', '', '> On Tuesday we will deploy the build', '', 'Thanks.')
        expect(ex(text)).toBe(text)
    })
})

describe('extractContent — plain text — Turkish / Romanian attributions', () => {
    // Turkish: verb-last "şunu yazdı:" ("wrote"), with a Turkish date lead-in ("… tarihinde …"). Same shape
    // as the French/Spanish batch, so it slots into the foreign-verb signal.
    it('cuts at a Turkish "… yazdı:" attribution line', () => {
        const text = t('Teşekkürler.', '', '12 Nis 2025 Cmt, saat 17:10 tarihinde Hunter <hunter@hunterscouts.co> şunu yazdı:', '> alıntılanan eski içerik')
        expect(ex(text)).toBe('Teşekkürler.')
    })

    // WRAPPED Turkish: a long name/email pushes "yazdı:" onto the next line; the date-led lead-in (year + clock)
    // anchors the wrap window so we cut at the lead-in, not one line too low.
    it('cuts at the lead-in of a WRAPPED Turkish "yazdı:" attribution', () => {
        const text = t('Tamam, görüşürüz.', '', '12 Nis 2025 Cmt, saat 17:10 tarihinde Hunter <hunter@hunterscouts.co>', 'şunu yazdı:', '> eski içerik')
        expect(ex(text)).toBe('Tamam, görüşürüz.')
    })

    // Romanian: verb-last "a scris:" ("wrote").
    it('cuts at a Romanian "… a scris:" attribution line', () => {
        const text = t('Mulțumesc!', '', 'mar., 17 mar. 2026, 23:29 Marian <marianberes@agentmail.to> a scris:', '> conținut citat vechi')
        expect(ex(text)).toBe('Mulțumesc!')
    })
})

describe('extractContent — plain text — Office-365 / inline bounces kept whole', () => {
    // Office-365 bounce that LEADS with a logo-URL line before the headline — isTrueDsn skips the decorative
    // line and matches "Your message TO <addr> couldn't be delivered" (recipient text between message+verb).
    it('keeps an Office-365 bounce whole even when it leads with a logo-URL line', () => {
        const text = t(
            '[https://products.office.com/en-us/CMSImages/Office365Logo_Orange.png?version=abc]',
            "Your message to jyudin@hcg.com couldn't be delivered.",
            'A custom mail flow rule created by an admin has blocked your message.',
            '',
            'From: postmaster@hcg.com',
            'Diagnostic information for administrators:',
            'Generating server: EMP-EXMR.corp.example.com',
        )
        expect(ex(text)).toBe(text)
    })

    // Host-prefixed headline: "<server> rejected your message to the following email addresses:" — matched
    // anywhere on the first meaningful line, so the embedded recipient/From: block is not chopped.
    it('keeps a "<host> rejected your message" bounce whole', () => {
        const text = t(
            'SA2PEPF000015C7.mail.protection.outlook.com rejected your message to the following email addresses:',
            '',
            'omaurer@seic.com<mailto:omaurer@seic.com>',
            'A communication failure occurred during the delivery of this message.',
            'From: postmaster@seic.com',
        )
        expect(ex(text)).toBe(text)
    })

    // SAFETY: a reply that QUOTES a bounce still leads with sender prose, so it is NOT a true DSN and still cuts.
    it('still cuts a reply that quotes an Office-365 bounce below it', () => {
        const text = t(
            'Looks like this one bounced — can you confirm the address?',
            '',
            'On Mon, Jun 1, 2026 at 9:00 AM Mailer Daemon <daemon@example.com> wrote:',
            "> Your message to jyudin@hcg.com couldn't be delivered.",
            '> Final-Recipient: rfc822; jyudin@hcg.com',
        )
        expect(ex(text)).toBe('Looks like this one bounced — can you confirm the address?')
    })
})

describe('extractContent — plain text — FW:/Fwd: pasted forwards kept whole', () => {
    // A body that LEADS with "FW:"/"Fwd:" is a pasted forward; its inner From: header must not chop the
    // forwarded body. CEO decision: forwards are KEPT whole.
    it('keeps a "FW:" pasted forward whole (inner From: header does not cut)', () => {
        const text = t(
            'FW: Re: Quarterly report',
            'From: ceo@example.com',
            'To: team@example.com',
            'Date: Wed, 13 Aug 2025 16:45:12 -0400',
            '',
            'Team — please review the attached report before Monday.',
        )
        expect(ex(text)).toBe(text)
    })

    it('keeps a "Fwd:" pasted forward whole', () => {
        const text = t('Fwd: Invoice #4471', 'From: billing@vendor.com', 'Date: Mon, 1 Jun 2026 09:00 AM', '', 'Please find the invoice below.')
        expect(ex(text)).toBe(text)
    })

    // GUARD: a normal reply that merely MENTIONS forwarding mid-body (not a leading FW: line) still cuts.
    it('does not treat a mid-reply mention of forwarding as a FW: forward', () => {
        const text = t(
            'Sure, I can fwd: that to the team later today.',
            '',
            'On Mon, Jun 1, 2026 at 9:00 AM Bob <bob@example.com> wrote:',
            '> old quoted content',
        )
        expect(ex(text)).toBe('Sure, I can fwd: that to the team later today.')
    })
})

describe('extractContent — new outbound Gmail/UTC attribution format', () => {
    // Our outbound attribution moved from the ISO timestamp to the Gmail-style "On <wkday>, <Mon> <d>, <year>
    // at <h>:<mm> <AM/PM> UTC … wrote:" form. It carries a comma + "wrote:", so Signal 3 catches it directly.
    it('cuts at the new "On … at … UTC … wrote:" outbound attribution (text)', () => {
        const text = t(
            'Hi Adam,',
            '',
            'Sounds great — talk soon.',
            '',
            'On Thu, Apr 3, 2025 at 9:16 PM UTC Adam Goff <adampgoff@gmail.com> wrote:',
            '> some older quoted content',
        )
        expect(ex(text)).toBe(t('Hi Adam,', '', 'Sounds great — talk soon.'))
    })

    it('cuts at the new UTC attribution inside an HTML gmail_quote (html)', () => {
        const html = '<div>Hi Adam,<br><br>Sounds great.</div><div class="gmail_quote"><div>On Thu, Apr 3, 2025 at 9:16 PM UTC Adam Goff &lt;adampgoff@gmail.com&gt; wrote:</div><blockquote>old quoted content</blockquote></div>'
        expect(exHtml(html)).toBe('<div>Hi Adam,<br><br>Sounds great.</div>')
    })
})

// Keep / Guard Tests

describe('extractContent — plain text — keep / guard', () => {
    // No quote signal anywhere -> return the whole email untouched. This is the
    // large "no quote" share of real mail (~55%). Source: our golden-set email [1]
    // ("Hello ... Cheers, Nel").
    it('keeps the whole email when there is no quote signal', () => {
        const text = t('Hello,', '', 'Hope you’re having a great week.', '', 'Cheers,', 'Nel')
        expect(ex(text)).toBe(text)
    })

    // Guard for the foreign-attribution signal: a normal English line can end in a colon
    // ("Here are the next steps:"). Because we key on a foreign VERB (not just the colon),
    // this must stay whole — no foreign verb, no cut. Source: our reasoning.
    it('does not cut an English content line that merely ends with a colon', () => {
        const text = t('Hi team,', '', 'Here are the next steps:', 'Ship the build and tell support.')
        expect(ex(text)).toBe(text)
    })

    // False-positive guard (the #1 trap): a normal sentence STARTING with "On"
    // must NOT trigger a cut. Forces the rule to require "On" AND "wrote:", not
    // just "On". Source: Talon's test_line_starts_with_on.
    it('does not false-positive on a normal sentence beginning with "On"', () => {
        const text = t('Hi team,', '', 'On Tuesday we will deploy the new build.', '', 'Thanks!')
        expect(ex(text)).toBe(text)
    })

    // Window-size guard for the wrapped-attribution lookahead. An innocent sentence
    // starting with "On" sits a few lines ABOVE a real attribution. If the peek window
    // were too big, that innocent line would reach down to the real "wrote:" and wrongly
    // cut the whole email. With a tiny (2-line) window it can't, so the cut lands on the
    // real attribution below — not the innocent line. Source: our reasoning.
    it('keeps an innocent "On ..." sentence and cuts only at the real attribution below', () => {
        const text = t(
            'Hi team,',
            '',
            'On Tuesday we will deploy the new build.',
            'It should go smoothly.',
            'Thanks!',
            '',
            'On Mon, Jun 1, 2026 at 9:00 AM Bob <bob@example.com> wrote:',
            '> old quoted content',
        )
        expect(ex(text)).toBe(t('Hi team,', '', 'On Tuesday we will deploy the new build.', 'It should go smoothly.', 'Thanks!'))
    })

    // SIBLING to the 4-line guard above, at the OTHER distance: just ONE blank line
    // separates an innocent "On ..." sentence from the real attribution below. This is
    // the exact over-cut the corpus audit found (eval_cutmore.ts: "on camera. What should
    // I do?"). A genuine wrapped attribution is contiguous, so the window must STOP at the
    // blank — otherwise it leaps to the real "wrote:" and drops the last sentence.
    // We keep BOTH this and the 4-line test: they guard different distances. Source: our data.
    it('keeps an innocent "On ..." sentence one blank line above the real attribution', () => {
        const text = t(
            'Hi Hunter,',
            '',
            'I have a quick question about the interview setup.',
            'On my laptop the camera is not working.',
            '',
            'On Mon, Jun 1, 2026 at 9:00 AM Hunter <hunter@hunterscouts.co> wrote:',
            '> earlier quoted content',
        )
        expect(ex(text)).toBe(t('Hi Hunter,', '', 'I have a quick question about the interview setup.', 'On my laptop the camera is not working.'))
    })

    // Talon-style glued splitter: some clients collapse the whole thread onto one physical line, so the
    // attribution is mid-line ("reply text On <date> <email> wrote: ..."). We split only strongly-attribution-
    // looking shapes (date-ish + email + wrote/sent) so the usual line-start Signal 3 can cut safely.
    it('cuts at a glued inline "On ... wrote:" attribution on the same physical line', () => {
        const text = 'Reply goes here before the quote On Wed, 27 Aug 2025 at 20:06, AgentMail <hiring@agentmail.to> wrote: > old quoted text > more quote'
        expect(ex(text)).toBe('Reply goes here before the quote')
    })

    // Guard for the glued splitter: content can contain "On ..." mid-sentence. Without the email/date/verb
    // guard, a sentence about "On Roblox..." near other text could be chopped.
    it('does NOT cut an innocent glued "On ..." content sentence', () => {
        const text = 'My work on Roblox is relevant. On Roblox, I shipped systems that scaled to many users and wrote documentation for the team.'
        expect(ex(text)).toBe(text)
    })

    // False-positive guard for the "From:" signal (the #1 trap for it). A content line
    // can legitimately start with "From:" (a footer, a label) with NO email address. The
    // guard must keep this email whole — proving "From:" alone never cuts; it needs the
    // header-like signal (an email address / a Sent: or Date: line). Source: our reasoning.
    it('keeps a content line that starts with "From:" but is not a real header', () => {
        const text = t(
            'Hi,',
            '',
            'Quick note for the newsletter footer — it should read:',
            'From: The Daily Digest Team',
            '',
            'Let me know if that works.',
        )
        expect(ex(text)).toBe(text)
    })

    // CHANGE 3 regression — decision: forwarded content is KEPT whole. The "Forwarded
    // message" divider is NOT a cut signal. This test locks that behavior so it can't
    // silently regress if someone adds a forward-stripping rule later.
    it('keeps a forwarded email whole (forward divider is not a cut signal)', () => {
        const text = t(
            'FYI, see the note below.',
            '',
            '---------- Forwarded message ----------',
            'Subject: Weekly digest',
            '',
            'This is the forwarded content.',
        )
        expect(ex(text)).toBe(text)
    })

    // The real forward case the simple test above misses: a forwarded message brings its OWN
    // From:/Date: header, which would trip Signal 5 and drop the forwarded body. Talon keeps
    // the whole message (its `[te]*f` rule: a forward divider before any quote = keep all).
    // CEO decision: forwards are KEPT. The forward guard must win over the inner From: header.
    it('keeps a forward whole even when the forwarded message has its own From: header', () => {
        const text = t(
            'FYI, sharing the note below.',
            '',
            '---------- Forwarded message ----------',
            'From: Alice <alice@example.com>',
            'Date: Mon, Jun 1, 2026 9:00 AM',
            'Subject: Quarterly report',
            '',
            'Here is the quarterly report body.',
        )
        expect(ex(text)).toBe(text)
    })

    // Apple Mail forward — the divider is "Begin forwarded message:" (no dashes), and the forwarded message
    // carries its own From:/Date: header. Without recognizing the Apple divider, OURS would cut at that From:
    // header and drop the forwarded body; the forward guard now keeps the whole thing. Source: corpus (29 msgs).
    it('keeps an Apple Mail "Begin forwarded message:" forward whole', () => {
        const text = t(
            'Thought you should see this.',
            '',
            'Begin forwarded message:',
            '',
            'From: Alice <alice@example.com>',
            'Subject: Quarterly report',
            'Date: June 1, 2026 at 9:00:00 AM PDT',
            'To: Bob <bob@example.com>',
            '',
            'Here is the quarterly report body.',
        )
        expect(ex(text)).toBe(text)
    })

    // GUARD: "begin forwarded message" only fires as its OWN line — a prose mention must not keep-whole/misfire.
    it('does not treat a prose mention of "forwarded message" as a forward divider', () => {
        const text = t(
            'Please review before I begin forwarding messages to the wider team.',
            '',
            'On Mon, Jun 1, 2026 at 9:00 AM Bob <bob@example.com> wrote:',
            '> old quoted content',
        )
        expect(ex(text)).toBe('Please review before I begin forwarding messages to the wider team.')
    })
})

// DSN (delivery-failure bounce) — CEO decision: a bounce is KEPT WHOLE, never stripped. The detector is
// anchored on the FIRST non-blank line (bounce headline or bare report field), so a real bounce is kept
// whole while a reply that merely QUOTES a bounce (sender prose on top) still strips normally. Source: the
// DSN corpus audit — the answer key wrongly discarded/chopped most true bounces; CEO says keep them whole.
describe('extractContent — plain text — DSN bounces kept whole', () => {
    // Gmail bounce — leads with "** Address not found **". Kept whole (Example B from the audit).
    it('keeps a Gmail "** Address not found **" bounce whole', () => {
        const text = t(
            '** Address not found **',
            '',
            "Your message wasn't delivered to kevin@openai.com because the address couldn't be found.",
            '',
            'Final-Recipient: rfc822; kevin@openai.com',
            'Action: failed',
            'Status: 5.1.1',
        )
        expect(ex(text)).toBe(text)
    })

    // The KEY FIX: an Exchange/Outlook bounce whose body contains the original message's "From:" /
    // "Received:" headers — Signal 5 used to chop here, dropping most of the bounce. Now kept whole.
    it('keeps an Exchange bounce whole even though it embeds a From:/Received: header dump', () => {
        const text = t(
            'Delivery has failed to these recipients or groups:',
            '',
            'someone@company.com',
            "The email address you entered couldn't be found.",
            '',
            'Diagnostic information for administrators:',
            'Generating server: EMP-EXMR1922.corp.example.com',
            'Received: from a.example.com by b.example.com',
            'From: Nikhil K <nikhil@getthehiringagents.com>',
            'To: someone@company.com',
            'Final-Recipient: rfc822; someone@company.com',
            'Status: 5.1.10',
        )
        expect(ex(text)).toBe(text)
    })

    // A bare structured DSN that leads directly with a report field (no human headline). Kept whole.
    it('keeps a structured bounce that leads with a "Reporting-MTA:" field', () => {
        const text = t(
            'Reporting-MTA: dns; googlemail.com',
            'Arrival-Date: Tue, 29 Jul 2025 09:44:28 -0700',
            'Final-Recipient: rfc822; nobody@example.com',
            'Action: failed',
            'Status: 5.1.3',
        )
        expect(ex(text)).toBe(text)
    })

    // SAFETY (the case you flagged): an agent REPLY on top, with the bounce QUOTED below. The first line is
    // sender prose, not a bounce headline, so this is NOT a true DSN — it must still be CUT at the quote.
    it('still CUTS a reply that quotes a bounce below it (prose on top, not a true DSN)', () => {
        const text = t(
            'It looks like there was a problem sending the email. Please check the address and resend.',
            '',
            'On 2025-05-08T05:46:14.000Z MAILER-DAEMON@amazonses.com wrote:',
            '',
            '> An error occurred while trying to deliver the mail to the following recipients:',
            '> Final-Recipient: rfc822; salesmanager@company.com',
            '> Action: failed',
            '> Status: 5.1.1',
        )
        expect(ex(text)).toBe('It looks like there was a problem sending the email. Please check the address and resend.')
    })

    // SAFETY: even if the quoted bounce below has NO ">" markers, the first line is still sender prose, so
    // first-line anchoring keeps it on the cut path (it cuts at the On…wrote: attribution).
    it('still cuts a reply that quotes an UNMARKED bounce below the attribution', () => {
        const text = t(
            'Heads up, the message bounced. Forwarding the details.',
            '',
            'On Mon, Jun 1, 2026 at 9:00 AM Mailer Daemon <daemon@example.com> wrote:',
            'Final-Recipient: rfc822; nobody@example.com',
            'Status: 5.1.1',
        )
        expect(ex(text)).toBe('Heads up, the message bounced. Forwarding the details.')
    })
})

// Trailing-signature rescue (the "split" shape: reply -> quote -> sender's "-- \nName" signature BELOW the
// quote). A plain single cut drops that signature; the rescue scans the dropped tail for the LAST bare
// RFC-3676 "-- " delimiter and reattaches from it to the end — but ONLY when the block between the cut and
// the "-- " is >=50% ">"-prefixed, so the "-- " is unambiguously the SENDER's sig trailing a real quote.
// Source: ~578 text splits in our corpus (verify_cuts.ts) + the gpt-4o key's mislabeling of these.
describe('extractContent — plain text — trailing-signature rescue', () => {
    // CORE: reply, an "On...wrote:" quote, then a bare "-- " signature below it. We keep the reply AND the
    // signature, dropping only the quote in the middle (a hole, not a single trailing cut).
    it('reattaches a trailing "-- " signature below a ">" quote (the split rescue)', () => {
        const text = t(
            'Thanks, that works for me.',
            '',
            'On Mon, Jun 1, 2026 at 9:00 AM Bob <bob@example.com> wrote:',
            '> quoted line one',
            '> quoted line two',
            '',
            '-- ',
            'Jane Doe',
            'Acme CEO',
        )
        expect(ex(text)).toBe(t('Thanks, that works for me.', '', '-- ', 'Jane Doe', 'Acme CEO'))
    })

    // A real reply/quote/signature split keeps the FULL multi-line signature (name, phone, email) to the end.
    it('keeps the full multi-line signature when it rescues', () => {
        const text = t(
            'Yes, I would be interested.',
            '',
            'On Wed, May 14, 2025 at 10:26 AM Hunter <hunter@hunterscouts.co> wrote:',
            '> Thanks for completing the screening interview.',
            '> I will keep you updated on next steps.',
            '',
            '-- ',
            'Mohammad Noman Kazi',
            '+1 (647) 833-6203',
            'nomankazi514@gmail.com',
        )
        expect(ex(text)).toBe(t('Yes, I would be interested.', '', '-- ', 'Mohammad Noman Kazi', '+1 (647) 833-6203', 'nomankazi514@gmail.com'))
    })

    // GUARD: the dropped block is <50% ">"-prefixed (a pasted Outlook header block, no ">" markers), so the
    // "-- " below it is NOT a clear sender signature — fall through to the plain single cut and drop it.
    it('does NOT rescue when the dropped block is <50% ">"-prefixed', () => {
        const text = t(
            'Got it.',
            '',
            'From: Bob <bob@example.com>',
            'Sent: Monday, June 1, 2026 9:00 AM',
            'Subject: Re: interview',
            'Some pasted body with no quote markers',
            '--',
            'Trailing line',
        )
        expect(ex(text)).toBe('Got it.')
    })

    // A ">"-quoted "> --" is part of the QUOTE (the quoted message's own signature delimiter), not the
    // sender's — the regex /^--\s*$/ requires the line to START with "--", so "> --" never matches and we
    // single-cut normally. Trap: reattaching a quoted signature as if it were the sender's.
    it('treats a ">"-quoted "> --" as quote, not a signature delimiter (no reattach)', () => {
        const text = t(
            'Sounds good.',
            '',
            'On Mon Bob wrote:',
            '> Hi there,',
            '> some quoted text',
            '> --',
            '> Bob (quoted signature)',
        )
        expect(ex(text)).toBe('Sounds good.')
    })

    // LAST "-- " wins: when an earlier bare "-- " sits inside the dropped region (e.g. a quoted sig that lost
    // its ">" markers) and the sender's real "-- " is at the very bottom, we reattach from the LAST one — so
    // the inner/quoted sig stays dropped and only the sender's trailing signature is kept.
    it('reattaches from the LAST "-- " (the sender\'s sig), not an earlier one in the quote', () => {
        const text = t(
            'My reply.',
            '',
            'On Mon Bob wrote:',
            '> quoted one',
            '> quoted two',
            '> quoted three',
            '-- ',
            'Inner Name',
            '-- ',
            'Outer Sender',
        )
        expect(ex(text)).toBe(t('My reply.', '', '-- ', 'Outer Sender'))
    })
})

// Inline-reply guard: when the sender's NEW prose is sandwiched BETWEEN ">"-quote blocks (quote -> prose ->
// quote), a single cut would drop the inline prose and lose meaning. So we KEEP THE WHOLE message (like
// forwards/DSN). Safe by construction — keeping whole never loses content. We fire ONLY on ">"-RUN quotes,
// not attribution-only quoting, and only on blank-separated substantial prose (so wrapped quotes don't count).
describe('extractContent — plain text — inline-reply guard', () => {
    // CORE: new prose between two ">"-quote blocks → keep the whole message.
    it('keeps an interleaved reply whole (prose sandwiched between ">" blocks)', () => {
        const text = t(
            'Thanks for the questions, answers below.',
            '',
            '> What is your experience with Rust?',
            '> Please be specific.',
            '',
            'I have used Rust for two years on systems projects.',
            '',
            '> And your availability?',
            '> When can you start?',
        )
        expect(ex(text)).toBe(text)
    })

    // SAFETY property: markdown ">" blockquotes the sender writes in NEW content (prose, quote, prose, quote)
    // are kept whole — keep-whole sidesteps the markdown-vs-quote conflict entirely (no content loss).
    it('keeps the sender\'s own markdown ">" blockquotes whole', () => {
        const text = t(
            'Here is my proposal, with the relevant quotes:',
            '',
            '> First principle we should follow.',
            '> Keep it simple.',
            '',
            'I think this applies directly to our roadmap.',
            '',
            '> Second principle to consider.',
            '> Ship fast.',
            '',
            'Let us discuss this on the call.',
        )
        expect(ex(text)).toBe(text)
    })

    // GUARD: attribution-only quoting (AI Q&A threads quote via "On...wrote:" with NO ">") is NOT treated as
    // interleaved — it still single-cuts at the first attribution. The ">"-run requirement is what excludes it.
    it('does NOT keep whole an attribution-only Q&A thread (no ">" runs) — still single-cuts', () => {
        const text = t(
            'Gabriel, how do you reconcile your Rust interest with your TypeScript focus?',
            '',
            'On 2025-08-14T14:03:18.000Z Gabriel Pérez <gabriel@garox.org> wrote:',
            'I am always on the bleeding edge of technology and learn extremely fast.',
            'On 2025-08-14T13:59:45.000Z AgentMail <hiring@agentmail.to> wrote:',
            'How do you balance innovation with delivery in a startup environment?',
        )
        expect(ex(text)).toBe('Gabriel, how do you reconcile your Rust interest with your TypeScript focus?')
    })

    // GUARD: a wrapped quote line (a long ">" line whose continuation lost its ">" marker) is NOT mistaken for
    // sandwiched new prose — the blank-separation requirement excludes it, so we still single-cut the quote.
    it('does NOT treat a wrapped quote continuation as inline prose', () => {
        const text = t(
            'Reply.',
            '',
            '> Quoted line one of the first block.',
            '> Quoted line two that wraps onto',
            'the next line without a marker.',
            '> Quoted line three.',
            '> Quoted line four.',
        )
        expect(ex(text)).toBe('Reply.')
    })

    // GUARD: prose AFTER a quote with NO quote below it is NOT a sandwich (it is a trailing addition, not
    // interleaved) — we still single-cut. Only the quote -> prose -> quote shape keeps whole.
    it('does NOT keep whole when prose follows a quote but no quote comes after (not sandwiched)', () => {
        const text = t(
            'My main reply.',
            '',
            '> quoted question one',
            '> quoted question two',
            '',
            'This extra thought has no quoted block after it, so it is not sandwiched.',
        )
        expect(ex(text)).toBe('My main reply.')
    })
})

// DSN bounces in HTML — same CEO keep-whole rule, anchored on the first visible (de-tagged) line.
describe('extractContent — HTML — DSN bounces kept whole', () => {
    it('keeps an HTML bounce whole even with an embedded From: header (would otherwise cut at marker #5)', () => {
        const html =
            '<div>Delivery has failed to these recipients or groups:</div>' +
            '<div>someone@company.com</div>' +
            '<div>From: Nikhil K &lt;nikhil@getthehiringagents.com&gt;</div>' +
            '<div>Final-Recipient: rfc822; someone@company.com</div>' +
            '<div>Status: 5.1.10</div>'
        expect(exHtml(html)).toBe(html)
    })

    // SAFETY: an HTML reply that quotes a bounce (prose on top) still cuts at the blockquote.
    it('still cuts an HTML reply that quotes a bounce below it', () => {
        const html =
            '<div>It looks like there was a problem sending the email.</div>' +
            '<blockquote><div>Final-Recipient: rfc822; nobody@example.com</div><div>Status: 5.1.1</div></blockquote>'
        expect(exHtml(html)).toBe('<div>It looks like there was a problem sending the email.</div>')
    })
})

/////////////////////////////////////////////////////////////

// HTML Tests
// Approach: raw-string scan (NO DOM) — find the earliest quote marker and
// slice the HTML before it, then clean (strip comments) + trim. Quotes are overwhelmingly
// TRAILING siblings after the reply, so a slice works for the bulk. We work biggest→smallest
// marker, same as text. Source: our data (scan_html2.js) + Talon's cut_* selectors.
describe('extractContent — HTML — quote markers', () => {
    // #1 by volume (44.5% of HTML mail): Gmail wraps quoted history in a trailing
    // <div class="gmail_quote ..."> (modern Gmail adds "gmail_quote_container", so we match
    // gmail_quote as a SUBSTRING of the class, not the exact class). Slice before that div.
    it('cuts at a gmail_quote div', () => {
        const html =
            '<div dir="ltr">Reply body here.</div>' +
            '<div class="gmail_quote gmail_quote_container">' +
            '<div class="gmail_attr">On Mon, Jun 1, 2026 Bob &lt;bob@x.com&gt; wrote:</div>' +
            '<blockquote>old quoted content</blockquote></div>'
        expect(exHtml(html)).toBe('<div dir="ltr">Reply body here.</div>')
    })

    // HTML split rescue: Gmail can render reply -> gmail_quote -> sender's own Gmail signature. A raw slice
    // at gmail_quote would lose the sender's signature, which is worse for semantic search than over-keeping
    // a little quote noise. Reattach only a trailing Gmail signature after the quote container closes.
    it('reattaches a trailing Gmail HTML signature below a gmail_quote block', () => {
        const html =
            '<div dir="ltr">Yes, we are still interested.</div><br>' +
            '<div class="gmail_quote gmail_quote_container">' +
            '<div class="gmail_attr">On Wed, Jul 16, 2025 Orin &lt;orin@example.com&gt; wrote:</div>' +
            '<blockquote class="gmail_quote">old quoted content</blockquote>' +
            '</div>' +
            '<div><br clear="all"></div><div><br></div>' +
            '<span class="gmail_signature_prefix">-- </span><br>' +
            '<div dir="ltr" class="gmail_signature"><div>Venligste Hilsen</div><div>Lotte Larzen</div></div>'
        expect(exHtml(html)).toBe(
            '<div dir="ltr">Yes, we are still interested.</div><br>' +
            '<span class="gmail_signature_prefix">-- </span><br>' +
            '<div dir="ltr" class="gmail_signature"><div>Venligste Hilsen</div><div>Lotte Larzen</div></div>',
        )
    })

    // If sender-owned HTML content (CTA/button/link block) sits between the quote and Gmail signature, keep
    // the whole trailing sender suffix. The signature proves this layer is not quoted history.
    it('reattaches sender content between a quote block and trailing Gmail signature', () => {
        const html =
            '<div dir="ltr">Please choose any time that works.</div><br>' +
            '<div class="gmail_quote gmail_quote_container"><blockquote class="gmail_quote">old quoted content</blockquote></div>' +
            '<div><a href="https://calendar.example.com">Book A Meeting</a></div>' +
            '<span class="gmail_signature_prefix">-- </span><br>' +
            '<div class="gmail_signature">Anjani Rai</div>'
        expect(exHtml(html)).toBe(
            '<div dir="ltr">Please choose any time that works.</div><br>' +
            '<div><a href="https://calendar.example.com">Book A Meeting</a></div>' +
            '<span class="gmail_signature_prefix">-- </span><br>' +
            '<div class="gmail_signature">Anjani Rai</div>',
        )
    })

    // Guard: a gmail_signature inside the quote container is the quoted sender's signature, not ours.
    it('does NOT reattach a Gmail signature that is inside the quoted block', () => {
        const html =
            '<div dir="ltr">Reply body here.</div><br>' +
            '<div class="gmail_quote gmail_quote_container">' +
            '<div class="gmail_attr">On Wed, Jul 16, 2025 Orin &lt;orin@example.com&gt; wrote:</div>' +
            '<blockquote class="gmail_quote">old quoted content<div class="gmail_signature">Quoted Sender</div></blockquote>' +
            '</div>'
        expect(exHtml(html)).toBe('<div dir="ltr">Reply body here.</div><br>')
    })

    // #2 by volume: a trailing <blockquote> holds the quoted message (non-Gmail
    // clients). We slice at the first blockquote. Source: our data (scan_html2.js).
    it('cuts at a generic blockquote', () => {
        const html =
            '<div>My reply text.</div>' +
            '<blockquote>quoted older message</blockquote>'
        expect(exHtml(html)).toBe('<div>My reply text.</div>')
    })

    // Apple Mail variant of #2: the quote is a <blockquote type="cite">. Same marker
    // catches it (we match "<blockquote" regardless of attributes). Source: scan_html2.js.
    it('cuts at an Apple Mail blockquote (type="cite")', () => {
        const html =
            '<div>Sounds good, talk soon.</div>' +
            '<blockquote type="cite"><div>On Mon, Jun 1, 2026 Bob wrote:</div>old quoted</blockquote>'
        expect(exHtml(html)).toBe('<div>Sounds good, talk soon.</div>')
    })

    // #3 by volume (3.0%) — the gap BOTH Talon libs miss. Modern Outlook (365 / web)
    // brackets the quote with <div id="appendonsend"></div>, an <hr width:98%>, then
    // <div id="divRplyFwdMsg">...headers...</div>. We slice at the earliest of those
    // (appendonsend here). Source: scan_html2.js; Talon only knows Outlook 2003–2013.
    it('cuts at a modern Outlook divRplyFwdMsg block', () => {
        const html =
            '<div dir="ltr">Thanks, that works for me.</div>' +
            '<div id="appendonsend"></div>' +
            '<hr style="display:inline-block;width:98%" tabindex="-1">' +
            '<div id="divRplyFwdMsg" dir="ltr">From: Bob &lt;bob@x.com&gt;<br>Sent: Monday<br><br>old quoted body</div>'
        expect(exHtml(html)).toBe('<div dir="ltr">Thanks, that works for me.</div>')
    })

    // Outlook for Android variant of #3: no appendonsend div — just the <hr width:98%>
    // then the divRplyFwdMsg block. Same markers catch it (the hr fires earliest). Source: scan_html2.js.
    it('cuts at an Outlook (Android) hr + divRplyFwdMsg block', () => {
        const html =
            '<div dir="auto">Sounds good.</div>' +
            '<div>Sent from Outlook for Android</div>' +
            '<hr style="display:inline-block;width:98%" tabindex="-1">' +
            '<div id="divRplyFwdMsg" dir="ltr">From: Hunter &lt;hunter@x.co&gt;</div>'
        // Cutter slices at the hr; stripNoiseHtml then drops the "Sent from Outlook for Android" line.
        expect(exHtml(html)).toBe('<div dir="auto">Sounds good.</div>')
    })

    // #1 sub-bucket of HTML cut-less (92%): the ORPHAN attribution. Non-Gmail clients put
    // the "On...wrote:" line in its own <div> BEFORE the <blockquote>; our structural slice
    // cuts the blockquote but leaves that div dangling — the HTML version of the text orphan
    // bug. The email is a real <a> tag, so the marker must tolerate tags between "On" and
    // "wrote:". We cut at the attribution (earlier than the blockquote). Source: eval_html_cutless.js.
    it('cuts the orphan "On...wrote:" attribution div before a blockquote', () => {
        const html =
            '<div>My reply text.</div>' +
            '<div>On Mon, Jun 1, 2026 at 9:00 AM Bob &lt;<a href="mailto:bob@x.com">bob@x.com</a>&gt; wrote:</div>' +
            '<blockquote>old quoted content</blockquote>'
        expect(exHtml(html)).toBe('<div>My reply text.</div>')
    })

    // Same orphan-attribution shape, but the sender's own Gmail signature is BELOW the quote block. The cut
    // starts at the attribution div, then looks through the immediate blockquote to rescue the trailing sig.
    it('reattaches a trailing Gmail signature below an orphan attribution + blockquote', () => {
        const html =
            '<div>My reply text.</div>' +
            '<div>On Mon, Jun 1, 2026 at 9:00 AM Bob &lt;<a href="mailto:bob@x.com">bob@x.com</a>&gt; wrote:</div>' +
            '<blockquote>old quoted content</blockquote>' +
            '<span class="gmail_signature_prefix">-- </span><br>' +
            '<div class="gmail_signature">Jane Sender</div>'
        expect(exHtml(html)).toBe(
            '<div>My reply text.</div>' +
            '<span class="gmail_signature_prefix">-- </span><br>' +
            '<div class="gmail_signature">Jane Sender</div>',
        )
    })

    // Same orphan, ISO no-comma format — the beachhead, now in HTML. BOTH Talon libs leave
    // this (their text-fallback has the comma bug), so removing it BEATS both. Source: our data.
    it('cuts the orphan ISO "On <ISO>...wrote:" attribution (no comma) in HTML', () => {
        const html =
            '<p>Looking forward to speaking with you.</p>' +
            '<div>On 2025-04-10T03:45:18.000Z Amil &lt;<a href="mailto:a@x.com">a@x.com</a>&gt; wrote:</div>' +
            '<blockquote>old quoted</blockquote>'
        expect(exHtml(html)).toBe('<p>Looking forward to speaking with you.</p>')
    })

    // Guard for the On...wrote: HTML marker: content that mentions "On ..." but has NO
    // "wrote:" (and no structural quote) must stay whole. The marker requires the trailing
    // "wrote:", so a bare "On" sentence can't trip it — same safety property as the text
    // path's "On" guard. Catches an over-greedy marker (e.g. just /\bOn\b/). Source: our reasoning.
    it('does not cut HTML that mentions "On" but has no attribution', () => {
        const html = '<div>Hi team,</div><div>On Tuesday we ship the new build. Thanks!</div>'
        expect(exHtml(html)).toBe(html)
    })

    // Risk 1: the marker must NOT link an innocent "On ..." in one block to an unrelated
    // "wrote:" in a LATER block. The </div> between them is the guard — the marker forbids
    // crossing a closing block tag (the HTML analog of the text path's "same line"). Source: our reasoning.
    it('does NOT link an innocent "On..." to an unrelated "wrote:" within range', () => {
        const html =
            '<div>On Tuesday we ship.</div>' +
            '<div>She wrote: the docs are done.</div>' + // "wrote:" but unrelated to the "On"
            '<p>Thanks!</p>'
        expect(exHtml(html)).toBe(html) // must NOT cut — two separate sentences
    })

    // Risk 2: no regression. When On...wrote: is INSIDE a gmail_quote, earliest-wins must
    // still land on the gmail_quote (it opens before the inner On), not shift the cut.
    // Source: our reasoning.
    it('still cuts correctly when On...wrote: is inside a gmail_quote (no regression)', () => {
        const html =
            '<div dir="ltr">Reply body here.</div>' +
            '<div class="gmail_quote"><div>On Mon, Jun 1 Bob &lt;b@x.com&gt; wrote:</div>' +
            '<blockquote>old</blockquote></div>'
        expect(exHtml(html)).toBe('<div dir="ltr">Reply body here.</div>')
    })

    // #5 by volume — the BIGGEST remaining HTML gap (2.80% of HTML mail, discover_html.ts).
    // The quoted message arrives as a bare header block — From:/Sent:/To:/Subject: — with NO
    // gmail_quote, NO <blockquote>, NO "On...wrote:" to anchor on. None of our markers fire,
    // so the whole quoted history survives into our output. This is the HTML analog of the
    // text path's Signal 5 (From: header, isQuoteSignal in draft_talon.ts). We cut at the
    // From: line and keep everything above it. Source: real corpus (01c701dbb059...@gmail.com,
    // a Yahoo/Outlook reply that survived our cut).
    it('cuts at a From:/Sent: header block with no wrapper (port of text Signal 5)', () => {
        const html =
            '<div dir="ltr">Thanks, will review and revert shortly.</div>' +
            '<div>From: Hunter &lt;hunter@hunterscouts.co&gt;</div>' +
            '<div>Sent: Friday, April 18, 2025 2:02 PM</div>' +
            '<div>Subject: Follow-up on Your Interview Session</div>' +
            '<div>Hi Hunter, here is the older quoted body…</div>'
        expect(exHtml(html)).toBe('<div dir="ltr">Thanks, will review and revert shortly.</div>')
    })

    // Guard for #5 (the SAME trap text Signal 5 has — see the text "keep" test). A content
    // line can legitimately start with "From:" (a footer/label) with NO email and NO
    // following Sent:/Date:. The marker must NOT cut here — it needs the header-like signal
    // (an email address, or a Sent:/Date: line nearby), exactly like the text path. This test
    // should STAY green: it passes now (no marker) and must keep passing after you add one.
    // Source: our reasoning (mirrors the text guard).
    it('keeps HTML whose content merely starts with "From:" but is not a real header', () => {
        const html =
            '<div>Hi,</div>' +
            '<div>Quick note for the newsletter footer — it should read:</div>' +
            '<div>From: The Daily Digest Team</div>' +
            '<div>Let me know if that works.</div>'
        expect(exHtml(html)).toBe(html)
    })

    // #5 follow-up: Outlook DESKTOP buries the email behind a fat inline <span style="…Calibri
    // CSS…"> between the "From:" label and the address, so the email lands ~140 chars after
    // "From:" (still in the SAME block — only inline <span>/<b> tags between them, no closing
    // block tag). This was 96% of the From: blocks that still survived marker #5 with the
    // original {0,100} window (audit_header.ts). We widen the window to {0,250} — the point
    // where recovery plateaus — so the email is reachable. The (?!</block) lookahead is
    // unchanged, so the wider window still can't run away across blocks. Source: our data
    // (real Outlook-desktop headers, e.g. AS8P192MB2090…@OUTLOOK.COM).
    it('cuts an Outlook-desktop From: header whose email is pushed far by inline CSS spans', () => {
        const html =
            '<div dir="ltr">Thanks, will review and revert shortly.</div>' +
            '<div>From: <span style="font-size:11.0pt; font-family:&quot;Calibri&quot;,sans-serif; color:#1F497D; mso-fareast-language:EN-US">Hunter</span> &lt;hunter@hunterscouts.co&gt;</div>' +
            '<div>Sent: Tuesday, May 27, 2025 4:09 PM</div>' +
            '<div>old quoted body</div>'
        expect(exHtml(html)).toBe('<div dir="ltr">Thanks, will review and revert shortly.</div>')
    })

    // Forward guard (HTML) — the analog of the text forward test. A forwarded message carries
    // its own From: header, which marker #5 would otherwise cut, dropping the forwarded body.
    // Talon keeps the whole message when a "Forwarded message" divider precedes any quote; CEO
    // decision: forwards are KEPT. So a divider before the earliest marker keeps the HTML whole.
    it('keeps a forward whole in HTML even when it has its own From: header', () => {
        const html =
            '<div>FYI, sharing the note below.</div>' +
            '<div>---------- Forwarded message ----------</div>' +
            '<div>From: Alice &lt;alice@example.com&gt;</div>' +
            '<div>Date: Mon, Jun 1, 2026</div>' +
            '<div>Here is the quarterly report body.</div>'
        expect(exHtml(html)).toBe(html)
    })

    // Apple Mail forward in HTML — "Begin forwarded message:" divider before the forwarded From: header.
    it('keeps an Apple "Begin forwarded message:" forward whole in HTML', () => {
        const html =
            '<div>Thought you should see this.</div>' +
            '<div>Begin forwarded message:</div>' +
            '<div>From: Alice &lt;alice@example.com&gt;</div>' +
            '<div>Subject: Quarterly report</div>' +
            '<div>Here is the quarterly report body.</div>'
        expect(exHtml(html)).toBe(html)
    })

    // i18n ports — HTML siblings of text Signals 5/6/7. The HTML markers never had these, so
    // localized headers + foreign attributions survived into our output (combined_audit.ts).
    // #5 (localized header): German Von: + email, same shape as English From:.
    it('cuts at a German "Von:" header block in HTML (marker #5 localized label)', () => {
        const html =
            '<div dir="ltr">Danke, das passt.</div>' +
            '<div>Von: Hunter &lt;hunter@hunterscouts.co&gt;</div>' +
            '<div>Gesendet: Montag, 1. Juni 2026</div>' +
            '<div>alter zitierter Inhalt</div>'
        expect(exHtml(html)).toBe('<div dir="ltr">Danke, das passt.</div>')
    })

    // The no-email From:/Sent: HTML branch was tried and reverted (DKIM "h=From:Date:…" over-cut),
    // so a no-email "From: <name>" header with no inline email stays whole — only an email cuts.
    it('keeps a no-email "From:" header in HTML even when a "Sent:" line follows', () => {
        const html =
            '<div dir="ltr">Thanks.</div>' +
            '<div>From: Hunter<br>Sent: Friday, June 5, 2026<br>Subject: Re</div>' +
            '<div>old quoted body</div>'
        expect(exHtml(html)).toBe(html)
    })

    // #6 (Latin attribution): French "… a écrit :" before a blockquote. The verb anchor snaps
    // back to the attribution's <div> (email is entity-encoded, so no inline tag intervenes),
    // and fires EARLIER than the blockquote — so we drop the orphan attribution, not just the quote.
    it('cuts at a French "a écrit :" attribution div in HTML', () => {
        const html =
            '<div>Merci, bien noté.</div>' +
            '<div>Le mardi 22 avril 2025, Hunter &lt;hunter@hunterscouts.co&gt; a écrit :</div>' +
            '<blockquote>ancien contenu cité</blockquote>'
        expect(exHtml(html)).toBe('<div>Merci, bien noté.</div>')
    })

    // #7 (Chinese): 写道 ("wrote") verb-last ending in a full-width colon.
    it('cuts at a Chinese "写道：" attribution in HTML', () => {
        const html =
            '<div>好的，谢谢。</div>' +
            '<div>Hunter &lt;hunter@hunterscouts.co&gt; 于2026年3月2日写道：</div>' +
            '<blockquote>引用的旧内容</blockquote>'
        expect(exHtml(html)).toBe('<div>好的，谢谢。</div>')
    })

    // #7 (Chinese): 发件人 ("From") header.
    it('cuts at a Chinese "发件人：" header in HTML', () => {
        const html =
            '<div>收到，谢谢！</div>' +
            '<div>发件人：Marco &lt;marco@agentmail.to&gt;</div>' +
            '<blockquote>引用内容</blockquote>'
        expect(exHtml(html)).toBe('<div>收到，谢谢！</div>')
    })

    // #7b (Arabic): "… كتب:" attribution before a blockquote.
    it('cuts at an Arabic "كتب:" attribution in HTML', () => {
        const html =
            '<div>شكرا لك.</div>' +
            '<div>Hunter &lt;hunter@hunterscouts.co&gt; كتب:</div>' +
            '<blockquote>المحتوى المقتبس القديم</blockquote>'
        expect(exHtml(html)).toBe('<div>شكرا لك.</div>')
    })

    // #8 (Original Message): localized German separator divider in HTML.
    it('cuts at a German "Ursprüngliche Nachricht" separator in HTML', () => {
        const html =
            '<div dir="ltr">Danke.</div>' +
            '<div>-----Ursprüngliche Nachricht-----</div>' +
            '<div>Von: Bob &lt;bob@example.com&gt;</div>' +
            '<div>alter Inhalt</div>'
        expect(exHtml(html)).toBe('<div dir="ltr">Danke.</div>')
    })

    // GUARD (HTML #5, localized): an innocent "De:" line with NO email in the block stays whole —
    // marker #5 requires a real email near the label, so the common word "de" can't over-cut.
    it('keeps HTML with an innocent "De:" line and no email', () => {
        const html = '<div>Hola,</div><div>De: revisar el informe y enviarlo</div><div>Gracias.</div>'
        expect(exHtml(html)).toBe(html)
    })

    // GUARD (HTML #5, localized): same for the everyday Dutch word "van" — no email, no cut.
    it('keeps HTML with an innocent "Van:" line and no email', () => {
        const html = '<div>Hoi,</div><div>Van: alles wat we vandaag besproken hebben</div><div>Groeten.</div>'
        expect(exHtml(html)).toBe(html)
    })

    // CUT: Danish/Norwegian "Fra:" header + email in HTML (marker #5 Scandinavian label).
    it('cuts at a "Fra:" header block in HTML (email in the block)', () => {
        const html =
            '<div dir="ltr">Tak, det virker.</div>' +
            '<div>Fra: Hunter &lt;hunter@hunterscouts.co&gt;</div>' +
            '<div>Sendt: mandag 1. juni 2026</div>' +
            '<div>gammelt citeret indhold</div>'
        expect(exHtml(html)).toBe('<div dir="ltr">Tak, det virker.</div>')
    })

    // GUARD: innocent "Fra:" content in HTML with no email — "fra" is a common word, must stay whole.
    it('keeps HTML with an innocent "Fra:" line and no email', () => {
        const html = '<div>Hej,</div><div>Fra: og med i morgen er kontoret lukket</div><div>Tak.</div>'
        expect(exHtml(html)).toBe(html)
    })
})

/////////////////////////////////////////////////////////////

// HTML trailing-signature REATTACH — one test per regex/guard in reattachTrailingSignature +
// trimTrailingNoise. The cutter slices at the quote; this rescues a sender signature stranded BELOW
// it (reply → quote → sig) but never a forward/more-quote, and trims any legal footer glued under it.
describe('extractContent — HTML — trailing-signature reattach (guards)', () => {
    // looksLikeSig signal #1 — RFC-3676 "-- " delimiter (non-Gmail clients).
    it('reattaches a sig marked only by a "-- " delimiter', () => {
        expect(exHtml('<div>My reply.</div><blockquote>old quoted</blockquote><div>-- </div><div>John Doe</div><div>john@acme.com</div>'))
            .toBe('<div>My reply.</div><div>-- </div><div>John Doe</div><div>john@acme.com</div>')
    })
    // looksLikeSig signal #2 — a bare email address, no delimiter.
    it('reattaches a sig identified by a trailing email address', () => {
        expect(exHtml('<div>My reply.</div><blockquote>old</blockquote><div>Jane Smith</div><div>jane@acme.com</div>'))
            .toBe('<div>My reply.</div><div>Jane Smith</div><div>jane@acme.com</div>')
    })
    // looksLikeSig signal #3 — a phone number.
    it('reattaches a sig identified by a trailing phone number', () => {
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote><div>Bob</div><div>+1 415 555 1234</div>'))
            .toBe('<div>Reply.</div><div>Bob</div><div>+1 415 555 1234</div>')
    })
    // looksLikeSig signal #4 — a sign-off word ("Regards", "Thanks", ...).
    it('reattaches a sig identified by a sign-off word', () => {
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote><div>Regards,</div><div>Alice</div>'))
            .toBe('<div>Reply.</div><div>Regards,</div><div>Alice</div>')
    })
    // looksLikeSig FALSE — trailing text with no sig signal at all → plain cut, no reattach.
    it('does NOT reattach trailing text that has no signature signal', () => {
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote><div>see you</div>')).toBe('<div>Reply.</div>')
    })
    // hasMoreQuote guard — a trailing "On…wrote:" is MORE quoting, not a signature.
    it('does NOT reattach a trailing "On…wrote:" (more quoting)', () => {
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote><div>On Mon, Jun 1 Bob &lt;bob@x.com&gt; wrote:</div><div>more quoted</div>'))
            .toBe('<div>Reply.</div>')
    })
    // hasMoreQuote guard — a trailing forwarded-message block is not a signature.
    it('does NOT reattach a trailing forwarded-message block', () => {
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote><div>---------- Forwarded message ----------</div><div>From: x@y.com</div>'))
            .toBe('<div>Reply.</div>')
    })
    // trimTrailingNoise — keep the contact-info sig, drop a CONFIDENTIALITY NOTICE glued under it.
    it('reattaches the sig but TRIMS a trailing confidentiality notice', () => {
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote><div>-- </div><div>Jane Doe</div><div>jane@acme.com</div><div>CONFIDENTIALITY NOTICE: This email is confidential.</div>'))
            .toBe('<div>Reply.</div><div>-- </div><div>Jane Doe</div><div>jane@acme.com</div>')
    })
    // trimTrailingNoise — the new "intended only for" disclaimer pattern is trimmed too.
    it('reattaches the sig but TRIMS a trailing "intended only for" disclaimer', () => {
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote><div>Jane Doe</div><div>jane@acme.com</div><div>This information is intended only for the addressee.</div>'))
            .toBe('<div>Reply.</div><div>Jane Doe</div><div>jane@acme.com</div>')
    })
    // The reattach trims TRAILING noise only, so it keeps a sig below leading "Sent from my iPhone" —
    // but the whole-body stripNoiseHtml then cuts at that leading noise line, dropping the sig below it.
    it('strips from a leading noise line, dropping a reattached sig below it', () => {
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote><div>Sent from my iPhone</div><div>-- </div><div>Jane</div><div>jane@acme.com</div>'))
            .toBe('<div>Reply.</div>')
    })
    // After trimming, NOTHING real is left (the tail was pure disclaimer) → bail to a plain cut.
    it('does NOT reattach a tail that is pure disclaimer', () => {
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote><div>This email is confidential, contact legal@acme.com</div>'))
            .toBe('<div>Reply.</div>')
    })
    // size guard — bounded by VISIBLE CHARACTERS, not lines (HTML sigs split into many short lines).
    // A many-line but SHORT sig is reattached...
    it('reattaches a short sig even when split across many HTML lines', () => {
        const tail = '<div>--</div><div>Jane Doe</div><div>Head of Product</div><div>Acme Inc</div>' +
            '<div>jane@acme.com</div><div>+1 555 0100</div><div>linkedin.com/in/jane</div><div>San Francisco</div><div>USA</div>'
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote>' + tail))
            .toBe('<div>Reply.</div>' + tail)
    })
    // ...but a genuinely long trailing block (>400 visible chars) is a new section, not a sign-off → bail.
    it('does NOT reattach a trailing block longer than 400 visible chars', () => {
        const tail = `<div>Regards, Jane. ${'x'.repeat(420)}</div>`
        expect(exHtml('<div>Reply.</div><blockquote>old</blockquote>' + tail)).toBe('<div>Reply.</div>')
    })
    // blockquote gate — when the quote has NO </blockquote> (Outlook divRplyFwdMsg), the reattach
    // can't anchor the quote end, so the stranded sig is (knowingly) dropped. Documents the gap.
    it('does NOT reattach when the quote has no </blockquote> (structural gap)', () => {
        expect(exHtml('<div>Reply.</div><div id="divRplyFwdMsg">quoted</div><div>-- </div><div>Jane</div>'))
            .toBe('<div>Reply.</div>')
    })
})

/////////////////////////////////////////////////////////////

// stripNoise — now applied INSIDE extractContent (text path; stripNoiseHtml for HTML). Tested here in
// isolation. We KEEP meaning (the message + the sender's name/title/company) and drop only trailing
// BOILERPLATE noise: mobile auto-sigs, marketing/list footers, legal disclaimers.
// Keep-by-default. This replaces the old blanket "--" cut, which deleted the meaningful name.
describe('stripNoise', () => {
    // Mobile auto-append — pure noise, cut it.
    it('cuts a "Sent from my iPhone" mobile auto-signature', () => {
        expect(stripNoise(t('Sounds good, see you then.', '', 'Sent from my iPhone'))).toBe('Sounds good, see you then.')
    })

    it('cuts "Sent from Outlook for Android"', () => {
        expect(stripNoise(t('Got it.', '', 'Sent from Outlook for Android'))).toBe('Got it.')
    })

    // THE POINT: keep the meaningful name/title/company (great for search), cut the legal disclaimer.
    it('keeps the name/title/company signature but cuts a trailing legal disclaimer', () => {
        const text = t('Best,', 'Jane Doe', 'CEO, Acme Inc', 'jane@acme.com', '',
            'This email and its attachments are confidential and intended solely for the addressee.',
            'If you are not the intended recipient, please delete it.')
        expect(stripNoise(text)).toBe(t('Best,', 'Jane Doe', 'CEO, Acme Inc', 'jane@acme.com'))
    })

    // NOISE_MARKERS additions — disclaimer phrasings found stranded under HTML quotes.
    it('cuts an "intended only for" disclaimer line', () => {
        expect(stripNoise(t('Best,', 'Jane', 'This is intended only for the recipient.'))).toBe(t('Best,', 'Jane'))
    })
    it('cuts an "email is not secure" disclaimer line', () => {
        expect(stripNoise(t('Thanks', 'This email is not secure.'))).toBe('Thanks')
    })
    it('cuts a "do not print this email" green-footer line', () => {
        expect(stripNoise(t('Cheers', 'Please do not print this email.'))).toBe('Cheers')
    })

    // Marketing / mailing-list footer — cut from the unsubscribe line down.
    it('cuts a marketing unsubscribe footer', () => {
        const text = t('Big sale this week — 20% off!', '',
            'To unsubscribe, click here: http://example.com/u', 'You received this email because you subscribed.')
        expect(stripNoise(text)).toBe('Big sale this week — 20% off!')
    })

    // KEEP-BY-DEFAULT: a meaningful name/title/contact signature with no noise is preserved in full.
    it('keeps a meaningful name/title/contact signature with no noise', () => {
        const text = t('Thanks for the update.', '', 'Best,', 'John Smith', 'Head of Product, Acme', '+1 555 0100')
        expect(stripNoise(text)).toBe(text)
    })

    // The old "--" delimiter no longer blanket-deletes the name — names are meaningful, so KEPT.
    it('keeps a name signature below a "--" delimiter (no longer blanket-cut)', () => {
        const text = t('Thanks!', '', '--', 'Rahul Arora', 'rahul@example.com')
        expect(stripNoise(text)).toBe(text)
    })

    // GUARD: content that merely MENTIONS "confidential"/"unsubscribe" in a sentence is not noise.
    it('keeps content that mentions "confidential" or "unsubscribe" in a sentence', () => {
        expect(stripNoise(t('Please keep this confidential between us.'))).toBe('Please keep this confidential between us.')
        expect(stripNoise(t('Can you unsubscribe me from the weekly list?'))).toBe('Can you unsubscribe me from the weekly list?')
    })

    // extractContent NOW strips noise inline (text path): the quote is cut, then stripNoise removes
    // the trailing mobile auto-signature.
    it('extractContent strips trailing noise from extracted_text', () => {
        const raw = t('Thanks!', '', 'Sent from my iPhone')
        expect(extractContent({ text: raw }).extracted_text).toBe('Thanks!')
    })
    // HTML path strips noise too, via stripNoiseHtml (block-aware).
    it('extractContent strips trailing noise from extracted_html', () => {
        const html = '<div>Thanks!</div><div>Sent from my iPhone</div><div>This email is confidential.</div>'
        expect(extractContent({ html }).extracted_html).toBe('<div>Thanks!</div>')
    })
    // GUARD: noise stripping must NOT eat a real reply that merely mentions a noise word in a sentence.
    it('does not strip a reply that mentions "confidential"/"unsubscribe" in a sentence', () => {
        const html = '<div>Please keep this confidential between us.</div>'
        expect(extractContent({ html }).extracted_html).toBe(html)
        const text = 'Can you unsubscribe me from the weekly list?'
        expect(extractContent({ text }).extracted_text).toBe(text)
    })
})

// Open items — status of past future-work (documentation only; no pending tests).
// DONE — trailing-signature rescue (reply -> quote -> sender signature below the quote). Text path
    // re-attaches from the RFC-3676 "-- " delimiter; HTML path has reattachTrailingSignature (gmail_quote /
    // blockquote, NOISE_MARKERS-trimmed, char-bounded). Validated: 86% of reachable stranded sigs reattached,
    // 0 disclaimer leaks. Covered by the 'trailing-signature rescue' and 'trailing-signature reattach' suites.

    // DONE — interleaved / multi-segment replies. hasInlineReply keeps quote -> prose -> quote messages whole
    // (text path). The HTML inline guard was VALIDATED and deliberately NOT added: a keep-whole guard would
    // wreck ~89% of "interleaved"-flagged HTML (clean cut is correct there); the residual loss is key-noise.

    // DONE — auto-reply / OOO handling: resolved as a downstream pipeline concern, not the extractor's job.

    // WON'T DO — P3 marker-less quoted history: a reply that includes the previous message's text with NO
    // ">", blockquote, "On…wrote:", or header to cut on (0.46%). NOT fixable single-message; needs thread
    // dedup / context. Out of scope for this layer.

    // SCANNED & VERIFIED NEGLIGIBLE — do NOT build dedicated markers (OURS already catches them via
    // existing markers, or they barely exist): Outlook desktop "border-top:solid" microsoft quote (666 msgs
    // but only 8 real misses — blockquote/On…wrote: catch the rest), Yahoo "yahoo_quoted" (143/2 misses),
    // quoted-printable artifacts (79 real, 0 OURS errors), Zimbra "zwchr" (2), Zendesk/Intercom/Front
    // helpdesk reply-above-line markers (~84 total). Re-check only if corpus mix shifts.

    // Tried & REVERTED (do NOT re-attempt): a no-email HTML From:/Sent: branch — without the email
    // guard it matches DKIM "h=From:Date:Subject:…" lists and over-cut 164 messages. Stays residual.

    // AUDITED & SKIPPED — link normalization (Talon _replace_link_brackets): only 65 wrapped <http/mailto>
    // links in 35k text mail, just 4 OURS-disagreements (the 3-line window + 2+-">"-run rule already handle
    // them). Not worth the line-renumber / output-mangle risk. Glued "On…wrote:" (Talon _wrap_splitter) is
    // also tiny (~5 genuine plain-text cases; most are the HTML-in-text artifact above).

    // ADDED this round (absent in our corpus but real client formats → zero-regression defensive coverage):
    //   "On … sent:" (Signal 3, email-guarded) and "Sent from Samsung … <email> wrote:" (Signal 10).
    // Still verified ABSENT and NOT built: Polymail "< mailto: > wrote:", numeric dd.mm.yyyy date splitter,
    // and Korean / Hebrew / Thai / Devanagari attributions (all hits were content, not attributions).
