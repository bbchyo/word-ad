/**
 * EBYÜ Thesis Format Validator v5.0 - Segment-Based Architecture
 * Erzincan Binali Yıldırım Üniversitesi 2022 Tez Yazım Kılavuzu
 */

// ============================================
// CONSTANTS
// ============================================
const RULES = {
    MARGIN_POINTS: 85.05,
    MARGIN_TOP_SPECIAL_POINTS: 198.45,
    MARGIN_TOLERANCE: 3.0,
    FONT_NAME: "Times New Roman",
    FONT_SIZE_BODY: 12,
    FONT_SIZE_HEADING_MAIN: 14,
    FONT_SIZE_HEADING_SUB: 12,
    FONT_SIZE_BLOCK_QUOTE: 11,
    FONT_SIZE_FOOTNOTE: 10,
    FONT_SIZE_TABLE_CONTENT: 11,
    FONT_SIZE_CAPTION_TITLE: 12,
    FONT_SIZE_CAPTION_CONTENT: 11,
    FONT_SIZE_COVER: 16,
    FONT_SIZE_COVER_ABD: 14,
    FONT_SIZE_COVER_SUPPORT: 12,
    FONT_SIZE_EPIGRAPH: 11,
    FONT_SIZE_PAGE_NUMBER: 10,
    FIRST_LINE_INDENT_PT: 35.4,
    BLOCK_QUOTE_INDENT_PT: 35.4,
    BIBLIO_HANGING_INDENT_PT: 28.35,
    INDENT_TOLERANCE: 3.0,
    SPACING_6NK: 6,
    SPACING_3NK: 3,
    SPACING_0NK: 0,
    SPACING_TOLERANCE: 2.0,
    LINE_SPACING_1_5_MIN: 15,
    LINE_SPACING_1_5_MAX: 22,
    LINE_SPACING_SINGLE_MIN: 10,
    LINE_SPACING_SINGLE_MAX: 14,
    MIN_BODY_LENGTH: 30,
    ABSTRACT_MIN_WORDS: 200,
    ABSTRACT_MAX_WORDS: 250,
    ABSTRACT_MIN_KEYWORDS: 3,
    ABSTRACT_MAX_KEYWORDS: 5,
    PAGE_NUMBER_FOOTER_PT: 35.4,
    MIN_PAGES_MASTERS: 50,
    MIN_PAGES_PHD: 80
};

const HIGHLIGHT = { CRITICAL: "Red", FORMAT: "Yellow", FOUND: "Cyan" };

// ============================================
// SEGMENTS - Thesis structure in order
// ============================================
const SEG = {
    OUTER_COVER: 'OUTER_COVER',
    INNER_COVER: 'INNER_COVER',
    ETHICS: 'ETHICS',
    ORIGINALITY: 'ORIGINALITY',
    GUIDE_COMPLIANCE: 'GUIDE_COMPLIANCE',
    APPROVAL: 'APPROVAL',
    PREFACE: 'PREFACE',
    ABSTRACT_TR: 'ABSTRACT_TR',
    ABSTRACT_EN: 'ABSTRACT_EN',
    TOC: 'TOC',
    ABBREVIATIONS: 'ABBREVIATIONS',
    TABLE_LIST: 'TABLE_LIST',
    FIGURE_LIST: 'FIGURE_LIST',
    INTRO: 'INTRO',
    BODY_CHAPTER: 'BODY_CHAPTER',
    CONCLUSION: 'CONCLUSION',
    BIBLIOGRAPHY: 'BIBLIOGRAPHY',
    ETHICS_APPROVAL: 'ETHICS_APPROVAL',
    APPENDIX: 'APPENDIX',
    CV: 'CV'
};

// Paragraph types
const PTYPE = {
    CHAPTER_HEADING: 'CHAPTER_HEADING',
    SUB_HEADING: 'SUB_HEADING',
    BODY_TEXT: 'BODY_TEXT',
    BLOCK_QUOTE: 'BLOCK_QUOTE',
    BIBLIOGRAPHY: 'BIBLIOGRAPHY',
    CAPTION: 'CAPTION',
    EPIGRAPH: 'EPIGRAPH',
    GHOST_HEADING: 'GHOST_HEADING',
    COVER_TEXT: 'COVER_TEXT',
    FRONT_MATTER: 'FRONT_MATTER',
    TOC_ENTRY: 'TOC_ENTRY',
    LIST_ITEM: 'LIST_ITEM',
    EMPTY: 'EMPTY',
    UNKNOWN: 'UNKNOWN'
};
// ============================================
// DETECTION PATTERNS
// ============================================
const PAT = {
    CHAPTER: [
        /^(BİRİNCİ|İKİNCİ|ÜÇÜNCÜ|DÖRDÜNCÜ|BEŞİNCİ|ALTINCI|YEDİNCİ|SEKİZİNCİ|DOKUZUNCU|ONUNCU)\s*BÖLÜM$/i,
        /^BÖLÜM\s*[IVX\d]+/i
    ],
    SPECIAL_HEADING: [
        /^GİRİŞ$/i, /^SONUÇ$/i, /^SONUÇ VE ÖNERİLER$/i, /^TARTIŞMA$/i,
        /^KAYNAKÇA$/i, /^KAYNAKLAR$/i, /^REFERANSLAR$/i
    ],
    SUB_HEADING: [/^\d+\.\d+(\.\d+)*\.?\s+[A-ZÇĞİÖŞÜa-zçğıöşü]/],
    CAPTION_TABLE: /^Tablo\s*(\d+)\.(\d+)\s*[:.]/i,
    CAPTION_FIGURE: /^(Şekil|Grafik|Resim|Harita)\s*(\d+)\.(\d+)\s*[:.]/i,
    TOC_DOTS: /\.{3,}\s*(i{1,4}|v|x{0,3}|\d+)\s*$/i,
    COVER: [
        /^T\.?C\.?$/i, /^ERZİNCAN\s*BİNALİ\s*YILDIRIM/i, /^ÜNİVERSİTESİ$/i,
        /^(FEN|SOSYAL)\s*BİLİMLERİ\s*ENSTİTÜSÜ$/i,
        /^(YÜKSEK\s*LİSANS|DOKTORA)\s*(TEZİ)?$/i,
        /^DANIŞMAN/i, /^Tez\s*Danışmanı/i, /^Hazırlayan/i
    ],
    SEG_DETECT: {
        ETHICS: /^BİLİMSEL\s*ETİ[KĞ]/i,
        ORIGINALITY: /^TEZ\s*ÖZGÜNLÜK/i,
        GUIDE: /^KILAVUZ(A)?\s*UYGUNLUK/i,
        APPROVAL: /^KABUL\s*VE\s*ONAY/i,
        PREFACE: /^(ÖN\s*SÖZ|ÖNSÖZ)$/i,
        ABSTRACT_TR: /^ÖZET$/i,
        ABSTRACT_EN: /^ABSTRACT$/i,
        TOC: /^İÇİNDEKİLER$/i,
        ABBREV: /^(SİMGELER|KISALTMALAR)/i,
        TABLE_LIST: /^TABLOLAR\s*(LİSTESİ|DİZİNİ)?$/i,
        FIGURE_LIST: /^ŞEKİLLER\s*(LİSTESİ|DİZİNİ)?$/i,
        INTRO: /^GİRİŞ$/i,
        CONCLUSION: /^(SONUÇ|SONUÇ\s*VE\s*ÖNERİLER)$/i,
        BIBLIO: /^(KAYNAKÇA|KAYNAKLAR|REFERANSLAR)$/i,
        ETHICS_APPROVAL: /^ETİK\s*KURUL\s*ONAYI$/i,
        APPENDIX: /^EKLER?$/i,
        CV: /^ÖZGEÇMİŞ$/i
    },
    KEYWORDS_TR: /^Anahtar\s*Kelimeler\s*:/i,
    KEYWORDS_EN: /^Keywords\s*:/i,
    EPIGRAPH_QUOTE: /^[""].*[""]$/
};

// ============================================
// STATE
// ============================================
let validationResults = [];
let scanLog = [];
let isScanning = false;

// ============================================
// HELPERS
// ============================================
function logStep(cat, msg, det = null) {
    scanLog.push({ ts: new Date().toISOString(), cat, msg, det });
    console.log(`[${cat}] ${msg}`, det || '');
}

function matchAny(text, patterns) {
    if (!text || !patterns) return false;
    return patterns.some(p => p.test(text.trim()));
}

function isCentered(al) {
    if (al === undefined || al === null) return true;
    if (typeof al === 'string') return al.toLowerCase() === 'centered';
    if (typeof al === 'number') return al === 2;
    try { if (al === Word.Alignment.centered) return true; } catch (e) { }
    return false;
}

function isJustified(al) {
    if (al === undefined || al === null) return true;
    if (typeof al === 'string') return al.toLowerCase() === 'justified';
    if (typeof al === 'number') return al === 3;
    try { if (al === Word.Alignment.justified) return true; } catch (e) { }
    return false;
}

function isRightAligned(al) {
    if (al === undefined || al === null) return false;
    if (typeof al === 'string') return al.toLowerCase() === 'right';
    if (typeof al === 'number') return al === 1;
    return false;
}

function isHeadingStyle(s) {
    if (!s) return false;
    const l = s.toLowerCase();
    return l.includes('heading') || l.includes('başlık') || l.includes('title');
}

function isAllUpperCase(text) {
    const letters = text.replace(/[^a-zA-ZçğıöşüÇĞİÖŞÜ]/g, '');
    if (letters.length === 0) return true;
    return letters === letters.toUpperCase();
}

function isCaption(text) {
    const t = (text || '').trim();
    return !!(t.match(PAT.CAPTION_TABLE) || t.match(PAT.CAPTION_FIGURE));
}

function isTOCEntry(style, text) {
    const s = (style || '').toLowerCase();
    if (s.includes('toc') || s.includes('içindekiler')) return true;
    if (PAT.TOC_DOTS.test(text)) return true;
    return false;
}

function wordCount(text) {
    return (text || '').split(/\s+/).filter(w => w.length > 0).length;
}
// ============================================
// SEGMENT DETECTION - Builds thesis structure map
// ============================================
function detectSegments(paragraphDataList) {
    const segments = [];
    let coverCount = 0; // Track T.C. occurrences for outer/inner cover

    for (let i = 0; i < paragraphDataList.length; i++) {
        const text = (paragraphDataList[i].text || '').trim();
        const upper = text.toUpperCase();

        // T.C. marks cover pages
        if (/^T\.?C\.?$/.test(text)) {
            coverCount++;
            if (coverCount === 1) {
                segments.push({ type: SEG.OUTER_COVER, startIdx: i, endIdx: -1, title: 'Dış Kapak' });
            } else if (coverCount === 2) {
                // Close outer cover
                const outerCover = segments.find(s => s.type === SEG.OUTER_COVER && s.endIdx === -1);
                if (outerCover) outerCover.endIdx = i - 1;
                segments.push({ type: SEG.INNER_COVER, startIdx: i, endIdx: -1, title: 'İç Kapak' });
            }
            continue;
        }

        // Detect segment-starting headings
        const detectors = [
            { pat: PAT.SEG_DETECT.ETHICS, type: SEG.ETHICS, title: 'Bilimsel Etiğe Uygunluk' },
            { pat: PAT.SEG_DETECT.ORIGINALITY, type: SEG.ORIGINALITY, title: 'Tez Özgünlük Sayfası' },
            { pat: PAT.SEG_DETECT.GUIDE, type: SEG.GUIDE_COMPLIANCE, title: 'Kılavuza Uygunluk' },
            { pat: PAT.SEG_DETECT.APPROVAL, type: SEG.APPROVAL, title: 'Kabul ve Onay Tutanağı' },
            { pat: PAT.SEG_DETECT.PREFACE, type: SEG.PREFACE, title: 'Ön Söz' },
            { pat: PAT.SEG_DETECT.ABSTRACT_TR, type: SEG.ABSTRACT_TR, title: 'Özet' },
            { pat: PAT.SEG_DETECT.ABSTRACT_EN, type: SEG.ABSTRACT_EN, title: 'Abstract' },
            { pat: PAT.SEG_DETECT.TOC, type: SEG.TOC, title: 'İçindekiler' },
            { pat: PAT.SEG_DETECT.ABBREV, type: SEG.ABBREVIATIONS, title: 'Simgeler ve Kısaltmalar' },
            { pat: PAT.SEG_DETECT.TABLE_LIST, type: SEG.TABLE_LIST, title: 'Tablolar Listesi' },
            { pat: PAT.SEG_DETECT.FIGURE_LIST, type: SEG.FIGURE_LIST, title: 'Şekiller Listesi' },
            { pat: PAT.SEG_DETECT.INTRO, type: SEG.INTRO, title: 'Giriş' },
            { pat: PAT.SEG_DETECT.CONCLUSION, type: SEG.CONCLUSION, title: 'Sonuç' },
            { pat: PAT.SEG_DETECT.BIBLIO, type: SEG.BIBLIOGRAPHY, title: 'Kaynakça' },
            { pat: PAT.SEG_DETECT.ETHICS_APPROVAL, type: SEG.ETHICS_APPROVAL, title: 'Etik Kurul Onayı' },
            { pat: PAT.SEG_DETECT.APPENDIX, type: SEG.APPENDIX, title: 'Ekler' },
            { pat: PAT.SEG_DETECT.CV, type: SEG.CV, title: 'Özgeçmiş' }
        ];

        // Skip TOC entries (lines with dots and page numbers)
        if (PAT.TOC_DOTS.test(text)) continue;

        for (const det of detectors) {
            if (det.pat.test(text)) {
                // Don't duplicate segments
                if (!segments.find(s => s.type === det.type)) {
                    // Close previous open segment
                    const lastOpen = segments.findLast(s => s.endIdx === -1);
                    if (lastOpen) lastOpen.endIdx = i - 1;
                    segments.push({ type: det.type, startIdx: i, endIdx: -1, title: det.title });
                }
                break;
            }
        }

        // Detect BÖLÜM headings
        if (matchAny(text, PAT.CHAPTER)) {
            const lastOpen = segments.findLast(s => s.endIdx === -1);
            if (lastOpen) lastOpen.endIdx = i - 1;
            segments.push({ type: SEG.BODY_CHAPTER, startIdx: i, endIdx: -1, title: text });
        }
    }

    // Close last segment
    const lastOpen = segments.findLast(s => s.endIdx === -1);
    if (lastOpen) lastOpen.endIdx = paragraphDataList.length - 1;

    return segments;
}

function getSegmentForParagraph(segments, paraIdx) {
    for (const seg of segments) {
        if (paraIdx >= seg.startIdx && paraIdx <= seg.endIdx) return seg;
    }
    return null;
}

// ============================================
// PARAGRAPH TYPE DETECTION
// ============================================
function detectParaType(pd, segment) {
    const text = (pd.text || '').trim();
    const { style, outlineLevel, tableNestingLevel, leftIndent } = pd;

    if (tableNestingLevel > 0) return isCaption(text) ? PTYPE.CAPTION : PTYPE.BODY_TEXT;

    // Empty - check ghost heading
    if (text.length === 0) {
        if (outlineLevel !== undefined && outlineLevel !== null) {
            if (typeof outlineLevel === 'number' && outlineLevel >= 0 && outlineLevel <= 8) return PTYPE.GHOST_HEADING;
            if (typeof outlineLevel === 'string' && !isNaN(parseInt(outlineLevel))) {
                const lv = parseInt(outlineLevel);
                if (lv >= 0 && lv <= 8) return PTYPE.GHOST_HEADING;
            }
        }
        if (isHeadingStyle(style)) return PTYPE.GHOST_HEADING;
        return PTYPE.EMPTY;
    }

    // Cover/front matter segments
    if (segment) {
        const st = segment.type;
        if (st === SEG.OUTER_COVER || st === SEG.INNER_COVER) return PTYPE.COVER_TEXT;
        if ([SEG.ETHICS, SEG.ORIGINALITY, SEG.GUIDE_COMPLIANCE, SEG.APPROVAL, SEG.PREFACE].includes(st)) return PTYPE.FRONT_MATTER;
        if (st === SEG.TOC) return PTYPE.TOC_ENTRY;
    }

    // TOC by style
    if (isTOCEntry(style, text)) return PTYPE.TOC_ENTRY;

    // Captions
    if (isCaption(text)) return PTYPE.CAPTION;

    // Bibliography segment
    if (segment && segment.type === SEG.BIBLIOGRAPHY && text.length > 5) {
        if (!PAT.SEG_DETECT.BIBLIO.test(text)) return PTYPE.BIBLIOGRAPHY;
    }

    // Chapter heading
    if (matchAny(text, PAT.CHAPTER) || matchAny(text, PAT.SPECIAL_HEADING)) return PTYPE.CHAPTER_HEADING;
    if (outlineLevel === 0 || outlineLevel === 1) return PTYPE.CHAPTER_HEADING;
    if (isHeadingStyle(style) && (/heading\s*1/i.test(style) || /başlık\s*1/i.test(style))) return PTYPE.CHAPTER_HEADING;

    // Sub-heading
    if (matchAny(text, PAT.SUB_HEADING)) return PTYPE.SUB_HEADING;
    if (typeof outlineLevel === 'number' && outlineLevel >= 2 && outlineLevel <= 8) return PTYPE.SUB_HEADING;
    if (isHeadingStyle(style)) return PTYPE.SUB_HEADING;

    // Block quote
    if (leftIndent && leftIndent >= 20) return PTYPE.BLOCK_QUOTE;

    // Body text
    if (text.length >= 20) return PTYPE.BODY_TEXT;

    return PTYPE.UNKNOWN;
}
// ============================================
// RESULT MANAGEMENT & UI
// ============================================
function addResult(type, title, desc, location = null, paraIndex = null, severity = null, segment = null) {
    validationResults.push({
        type, title, description: desc,
        location: location || 'Belge Geneli',
        paraIndex, severity: severity || (type === 'error' ? 'CRITICAL' : 'FORMAT'),
        segment: segment || '', timestamp: new Date().toISOString()
    });
}

function clearResults() { validationResults = []; scanLog = []; }

// ============================================
// COVER PAGE VALIDATION
// ============================================
function validateCover(pd, idx, segType) {
    const errs = [];
    const text = (pd.text || '').trim();
    if (text.length === 0) return errs;

    const { font, alignment } = pd;
    const label = segType === SEG.OUTER_COVER ? 'DIŞ KAPAK' : 'İÇ KAPAK';

    // Font must be Times New Roman
    if (font.name && font.name !== RULES.FONT_NAME) {
        errs.push({ type: 'error', title: `${label}: Yazı Tipi`, description: `${RULES.FONT_NAME} olmalı. Mevcut: ${font.name}`, paraIndex: idx, severity: 'CRITICAL' });
    }

    // Determine expected font size based on content
    const isTCLine = /^T\.?C\.?$/i.test(text);
    const isUniLine = /^ERZİNCAN/i.test(text) || /^ÜNİVERSİTESİ$/i.test(text);
    const isInstLine = /ENSTİTÜSÜ$/i.test(text);
    const isABDLine = /ANA\s*BİLİM\s*DALI/i.test(text);
    const isThesisType = /^(YÜKSEK\s*LİSANS|DOKTORA)/i.test(text);
    const isSupportLine = /desteklenmiştir/i.test(text);

    let expectedSize = RULES.FONT_SIZE_COVER; // 16pt default
    if (isABDLine || isThesisType) expectedSize = RULES.FONT_SIZE_COVER_ABD; // 14pt
    if (isSupportLine) expectedSize = RULES.FONT_SIZE_COVER_SUPPORT; // 12pt
    if (segType === SEG.INNER_COVER && isSupportLine) expectedSize = 12;

    if (font.size && Math.abs(font.size - expectedSize) > 1) {
        errs.push({ type: 'error', title: `${label}: Punto`, description: `${expectedSize} punto olmalı. Mevcut: ${font.size} pt`, paraIndex: idx, severity: 'CRITICAL' });
    }

    // Must be centered
    if (!isCentered(alignment)) {
        errs.push({ type: 'warning', title: `${label}: Hizalama`, description: 'Kapak öğeleri ortalanmış olmalı.', paraIndex: idx, severity: 'FORMAT' });
    }

    // Thesis title should be all uppercase
    if (text.length > 20 && !isTCLine && !isUniLine && !isInstLine && !isABDLine && !isThesisType) {
        if (!isAllUpperCase(text)) {
            errs.push({ type: 'warning', title: `${label}: Tez Adı`, description: 'Tez adı tamamı büyük harflerle yazılmalı.', paraIndex: idx, severity: 'FORMAT' });
        }
    }

    return errs;
}

// ============================================
// FRONT MATTER VALIDATION
// ============================================
function validateFrontMatter(pd, idx, segment) {
    const errs = [];
    const text = (pd.text || '').trim();
    if (text.length === 0) return errs;
    const { font } = pd;

    if (font.name && font.name !== RULES.FONT_NAME) {
        errs.push({ type: 'error', title: `${segment.title}: Yazı Tipi`, description: `${RULES.FONT_NAME} olmalı. Mevcut: ${font.name}`, paraIndex: idx, severity: 'CRITICAL' });
    }

    // Section title should be bold and uppercase
    if (idx === segment.startIdx) {
        if (font.bold !== true) {
            errs.push({ type: 'warning', title: `${segment.title}: Başlık`, description: 'Başlık koyu (bold) olmalı.', paraIndex: idx, severity: 'FORMAT' });
        }
        if (!isAllUpperCase(text)) {
            errs.push({ type: 'warning', title: `${segment.title}: Başlık`, description: 'Başlık büyük harflerle yazılmalı.', paraIndex: idx, severity: 'FORMAT' });
        }
    }

    return errs;
}

// ============================================
// CHAPTER HEADING VALIDATION (14pt, bold, centered, UPPERCASE)
// ============================================
function validateChapterHeading(pd, idx) {
    const errs = [];
    const { font, alignment, text } = pd;
    const trimmed = (text || '').trim();

    if (font.size && Math.abs(font.size - RULES.FONT_SIZE_HEADING_MAIN) > 0.5) {
        errs.push({ type: 'warning', title: 'Bölüm Başlığı: Punto', description: `14 punto olmalı. Mevcut: ${font.size} pt`, paraIndex: idx, severity: 'FORMAT' });
    }
    if (font.bold !== true) {
        errs.push({ type: 'warning', title: 'Bölüm Başlığı: Kalın Yazı', description: 'Bölüm başlıkları koyu (bold) olmalı.', paraIndex: idx, severity: 'FORMAT' });
    }
    if (!isCentered(alignment)) {
        errs.push({ type: 'warning', title: 'Bölüm Başlığı: Hizalama', description: 'Bölüm başlıkları ortalanmış olmalı.', paraIndex: idx, severity: 'FORMAT' });
    }
    if (font.name && font.name !== RULES.FONT_NAME) {
        errs.push({ type: 'error', title: 'Bölüm Başlığı: Yazı Tipi', description: `${RULES.FONT_NAME} olmalı. Mevcut: ${font.name}`, paraIndex: idx, severity: 'CRITICAL' });
    }
    // Must be uppercase
    if (!isAllUpperCase(trimmed) && matchAny(trimmed, PAT.CHAPTER)) {
        errs.push({ type: 'warning', title: 'Bölüm Başlığı: Büyük Harf', description: 'Bölüm başlıkları tamamı büyük harflerle yazılmalı.', paraIndex: idx, severity: 'FORMAT' });
    }

    return errs;
}

// ============================================
// SUB-HEADING VALIDATION (12pt, bold, 1.25cm Tab indent)
// ============================================
function validateSubHeading(pd, idx) {
    const errs = [];
    const { font, firstLineIndent } = pd;

    if (font.size && Math.abs(font.size - RULES.FONT_SIZE_HEADING_SUB) > 0.5) {
        errs.push({ type: 'warning', title: 'Alt Başlık: Punto', description: `12 punto olmalı. Mevcut: ${font.size} pt`, paraIndex: idx, severity: 'FORMAT' });
    }
    if (font.bold !== true) {
        errs.push({ type: 'warning', title: 'Alt Başlık: Kalın Yazı', description: 'Alt başlıklar koyu (bold) olmalı.', paraIndex: idx, severity: 'FORMAT' });
    }
    if (font.name && font.name !== RULES.FONT_NAME) {
        errs.push({ type: 'error', title: 'Alt Başlık: Yazı Tipi', description: `${RULES.FONT_NAME} olmalı. Mevcut: ${font.name}`, paraIndex: idx, severity: 'CRITICAL' });
    }
    // 1.25cm indent (same as paragraph indent)
    if (firstLineIndent !== undefined && Math.abs(firstLineIndent - RULES.FIRST_LINE_INDENT_PT) > RULES.INDENT_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Alt Başlık: Girinti', description: `1.25 cm girinti olmalı. Mevcut: ${(firstLineIndent / 28.35).toFixed(2)} cm`, paraIndex: idx, severity: 'FORMAT' });
    }

    return errs;
}

// ============================================
// BODY TEXT VALIDATION
// ============================================
function validateBodyText(pd, idx) {
    const errs = [];
    const { font, firstLineIndent, lineSpacing, spaceBefore, spaceAfter, text, leftIndent, isListItem } = pd;
    const trimmed = (text || '').trim();

    if (trimmed.length < RULES.MIN_BODY_LENGTH) return errs;
    if (isListItem) return errs; // Skip list items

    if (font.name && font.name !== RULES.FONT_NAME) {
        errs.push({ type: 'error', title: 'Metin: Yazı Tipi', description: `${RULES.FONT_NAME} olmalı. Mevcut: ${font.name}`, paraIndex: idx, severity: 'CRITICAL' });
    }
    if (font.size && Math.abs(font.size - RULES.FONT_SIZE_BODY) > 0.5) {
        errs.push({ type: 'warning', title: 'Metin: Punto', description: `12 punto olmalı. Mevcut: ${font.size} pt`, paraIndex: idx, severity: 'FORMAT' });
    }
    // First line indent 1.25cm
    if (firstLineIndent !== undefined && Math.abs(firstLineIndent - RULES.FIRST_LINE_INDENT_PT) > RULES.INDENT_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Metin: İlk Satır Girintisi', description: `1.25 cm olmalı. Mevcut: ${(firstLineIndent / 28.35).toFixed(2)} cm`, paraIndex: idx, severity: 'FORMAT' });
    }
    // Line spacing 1.5
    if (lineSpacing !== undefined && lineSpacing !== null && (lineSpacing < RULES.LINE_SPACING_1_5_MIN || lineSpacing > RULES.LINE_SPACING_1_5_MAX)) {
        errs.push({ type: 'warning', title: 'Metin: Satır Aralığı', description: `1.5 satır aralığı olmalı. Mevcut: ${lineSpacing.toFixed(1)} pt`, paraIndex: idx, severity: 'FORMAT' });
    }
    // Paragraph spacing 6nk before/after
    if (spaceBefore !== undefined && spaceBefore !== null && Math.abs(spaceBefore - RULES.SPACING_6NK) > RULES.SPACING_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Metin: Paragraf Öncesi', description: `6 nk olmalı. Mevcut: ${spaceBefore.toFixed(1)} nk`, paraIndex: idx, severity: 'FORMAT' });
    }
    if (spaceAfter !== undefined && spaceAfter !== null && Math.abs(spaceAfter - RULES.SPACING_6NK) > RULES.SPACING_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Metin: Paragraf Sonrası', description: `6 nk olmalı. Mevcut: ${spaceAfter.toFixed(1)} nk`, paraIndex: idx, severity: 'FORMAT' });
    }
    // Manual tab/space warning
    if (text && text.startsWith('\t')) {
        errs.push({ type: 'warning', title: 'Metin: Manuel Tab', description: 'Girintiyi Tab tuşuyla değil Paragraf ayarlarından 1.25 cm olarak ayarlayın.', paraIndex: idx, severity: 'FORMAT' });
    }

    return errs;
}
// ============================================
// BLOCK QUOTE VALIDATION (11pt, italic, 1.25cm indent both sides)
// ============================================
function validateBlockQuote(pd, idx) {
    const errs = [];
    const { font, leftIndent, rightIndent } = pd;

    if (font.size && Math.abs(font.size - RULES.FONT_SIZE_BLOCK_QUOTE) > 0.5) {
        errs.push({ type: 'warning', title: 'Blok Alıntı: Punto', description: `11 punto olmalı. Mevcut: ${font.size} pt`, paraIndex: idx, severity: 'FORMAT' });
    }
    if (font.italic !== true) {
        errs.push({ type: 'warning', title: 'Blok Alıntı: İtalik', description: 'Blok alıntı italik olmalı.', paraIndex: idx, severity: 'FORMAT' });
    }
    if (leftIndent !== undefined && Math.abs(leftIndent - RULES.BLOCK_QUOTE_INDENT_PT) > RULES.INDENT_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Blok Alıntı: Sol Girinti', description: `1.25 cm olmalı. Mevcut: ${(leftIndent / 28.35).toFixed(2)} cm`, paraIndex: idx, severity: 'FORMAT' });
    }
    if (rightIndent !== undefined && Math.abs(rightIndent - RULES.BLOCK_QUOTE_INDENT_PT) > RULES.INDENT_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Blok Alıntı: Sağ Girinti', description: `1.25 cm olmalı. Mevcut: ${(rightIndent / 28.35).toFixed(2)} cm`, paraIndex: idx, severity: 'FORMAT' });
    }
    return errs;
}

// ============================================
// BIBLIOGRAPHY VALIDATION (12pt, hanging 1cm, single spacing, 3nk)
// ============================================
function validateBibliography(pd, idx) {
    const errs = [];
    const { font, leftIndent, firstLineIndent, lineSpacing, spaceBefore, spaceAfter } = pd;

    if (font.name && font.name !== RULES.FONT_NAME) {
        errs.push({ type: 'error', title: 'Kaynakça: Yazı Tipi', description: `${RULES.FONT_NAME} olmalı.`, paraIndex: idx, severity: 'CRITICAL' });
    }
    if (font.size && Math.abs(font.size - RULES.FONT_SIZE_BODY) > 0.5) {
        errs.push({ type: 'warning', title: 'Kaynakça: Punto', description: `12 punto olmalı. Mevcut: ${font.size} pt`, paraIndex: idx, severity: 'FORMAT' });
    }
    // Hanging indent 1cm
    if (leftIndent !== undefined && firstLineIndent !== undefined) {
        const hanging = leftIndent - firstLineIndent;
        if (Math.abs(hanging - RULES.BIBLIO_HANGING_INDENT_PT) > RULES.INDENT_TOLERANCE) {
            errs.push({ type: 'warning', title: 'Kaynakça: Asılı Girinti', description: '1 cm asılı girinti olmalı.', paraIndex: idx, severity: 'FORMAT' });
        }
    }
    // Single spacing
    if (lineSpacing !== undefined && (lineSpacing < RULES.LINE_SPACING_SINGLE_MIN || lineSpacing > RULES.LINE_SPACING_SINGLE_MAX)) {
        errs.push({ type: 'warning', title: 'Kaynakça: Satır Aralığı', description: `Tek satır olmalı. Mevcut: ${lineSpacing.toFixed(1)} pt`, paraIndex: idx, severity: 'FORMAT' });
    }
    // 3nk before
    if (spaceBefore !== undefined && Math.abs(spaceBefore - RULES.SPACING_3NK) > RULES.SPACING_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Kaynakça: Paragraf Öncesi', description: `3 nk olmalı. Mevcut: ${spaceBefore.toFixed(1)} nk`, paraIndex: idx, severity: 'FORMAT' });
    }
    // 3nk after
    if (spaceAfter !== undefined && Math.abs(spaceAfter - RULES.SPACING_3NK) > RULES.SPACING_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Kaynakça: Paragraf Sonrası', description: `3 nk olmalı. Mevcut: ${spaceAfter.toFixed(1)} nk`, paraIndex: idx, severity: 'FORMAT' });
    }
    return errs;
}

// ============================================
// CAPTION VALIDATION (12pt title, centered, bold numbering)
// ============================================
function validateCaption(pd, idx) {
    const errs = [];
    const { font, alignment, spaceBefore, spaceAfter } = pd;

    if (font.size && Math.abs(font.size - RULES.FONT_SIZE_CAPTION_TITLE) > 0.5) {
        errs.push({ type: 'warning', title: 'Tablo/Şekil Başlığı: Punto', description: `12 punto olmalı. Mevcut: ${font.size} pt`, paraIndex: idx, severity: 'FORMAT' });
    }
    if (!isCentered(alignment)) {
        errs.push({ type: 'warning', title: 'Tablo/Şekil Başlığı: Hizalama', description: 'Ortalanmış olmalı.', paraIndex: idx, severity: 'FORMAT' });
    }
    if (spaceBefore !== undefined && spaceBefore > RULES.SPACING_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Tablo/Şekil Başlığı: Öncesi', description: `0 nk olmalı. Mevcut: ${spaceBefore.toFixed(1)} nk`, paraIndex: idx, severity: 'FORMAT' });
    }
    if (spaceAfter !== undefined && spaceAfter > RULES.SPACING_TOLERANCE) {
        errs.push({ type: 'warning', title: 'Tablo/Şekil Başlığı: Sonrası', description: `0 nk olmalı. Mevcut: ${spaceAfter.toFixed(1)} nk`, paraIndex: idx, severity: 'FORMAT' });
    }
    return errs;
}

// ============================================
// GHOST HEADING VALIDATION
// ============================================
function validateGhostHeading(pd, idx) {
    const { style, outlineLevel } = pd;
    let reason = '';
    if (typeof outlineLevel === 'number' && outlineLevel >= 0 && outlineLevel <= 8) {
        reason = `Taslak düzeyi ${outlineLevel + 1} olarak ayarlanmış`;
    } else if (isHeadingStyle(style)) {
        reason = `"${style}" başlık stili uygulanmış`;
    }
    return [{ type: 'error', title: 'BOŞ BAŞLIK (Ghost Heading)', description: `Bu boş satıra ${reason}. İçindekiler tablosunda hatalı boş satır oluşturur! Silin veya "Normal" stiline dönüştürün.`, paraIndex: idx, severity: 'CRITICAL' }];
}

// ============================================
// ABSTRACT/ÖZET VALIDATION (200-250 words, single page, 3-5 keywords)
// ============================================
function validateAbstractSection(paragraphDataList, startIdx, isEnglish) {
    const errs = [];
    const label = isEnglish ? 'ABSTRACT' : 'ÖZET';
    const keywordPat = isEnglish ? /keywords/i : /anahtar\s*kelime/i;

    let abstractWords = 0;
    let keywordLine = '';
    let keywordFound = false;
    let paragraphCount = 0;

    // Collect abstract text (skip heading line, stop at keywords)
    for (let i = startIdx + 1; i < paragraphDataList.length && i < startIdx + 40; i++) {
        const text = (paragraphDataList[i].text || '').trim();
        if (text.length === 0) continue;

        // Stop at keywords
        if (keywordPat.test(text)) {
            keywordFound = true;
            keywordLine = text;
            break;
        }
        // Stop at next major section
        if (/^(ABSTRACT|ÖZET|İÇİNDEKİLER|GİRİŞ|BİRİNCİ\s*BÖLÜM)$/i.test(text)) break;
        // Stop if we detect TOC entry
        if (PAT.TOC_DOTS.test(text)) break;

        // Skip metadata lines (university name, thesis info etc)
        if (/^(Erzincan|Sosyal|Doktora|Yüksek|Danışman|Supervisor)/i.test(text)) continue;
        if (/Üniversitesi|University|Enstitüsü|Institute/i.test(text)) continue;

        // Count actual abstract content words
        if (text.length > 10) {
            abstractWords += wordCount(text);
            paragraphCount++;
        }
    }

    // Word count check: 200-250
    if (abstractWords < RULES.ABSTRACT_MIN_WORDS) {
        errs.push({ type: 'warning', title: `${label}: Kelime Sayısı Az`, description: `En az 200 kelime olmalı. Mevcut: ${abstractWords} kelime`, severity: 'FORMAT' });
    } else if (abstractWords > RULES.ABSTRACT_MAX_WORDS) {
        errs.push({ type: 'warning', title: `${label}: Kelime Sayısı Fazla`, description: `En fazla 250 kelime olmalı. Mevcut: ${abstractWords} kelime`, severity: 'FORMAT' });
    } else {
        errs.push({ type: 'success', title: `${label}: Kelime Sayısı ✓`, description: `${abstractWords} kelime (200-250 arası).`, severity: 'FORMAT' });
    }

    // Keywords check
    if (keywordFound && keywordLine) {
        const kwPart = keywordLine.replace(/^(Anahtar\s*Kelimeler|Keywords)\s*:\s*/i, '');
        const keywords = kwPart.split(/[,;]/).map(k => k.trim()).filter(k => k.length > 0);
        if (keywords.length < RULES.ABSTRACT_MIN_KEYWORDS || keywords.length > RULES.ABSTRACT_MAX_KEYWORDS) {
            errs.push({ type: 'warning', title: `${label}: Anahtar Kelime Sayısı`, description: `3-5 anahtar kelime olmalı. Mevcut: ${keywords.length}`, severity: 'FORMAT' });
        }
    } else {
        errs.push({ type: 'error', title: `${label}: Anahtar Kelimeler Eksik`, description: `"${isEnglish ? 'Keywords' : 'Anahtar Kelimeler'}" satırı bulunamadı.`, severity: 'CRITICAL' });
    }

    // Check font italic (should NOT be italic in abstract - YÖK rule)
    const headingPd = paragraphDataList[startIdx];
    if (headingPd && headingPd.font.italic === true) {
        errs.push({ type: 'warning', title: `${label}: İtalik Kullanılmamalı`, description: 'Özet sayfasında italik yazı tipi kullanılmamalıdır (YÖK kuralı).', severity: 'FORMAT' });
    }

    return errs;
}
// ============================================
// TABLE/IMAGE/PAGE NUMBER VALIDATION
// ============================================
async function validateTables(context) {
    const errs = [];
    try {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();
        for (let i = 0; i < tables.items.length; i++) {
            const table = tables.items[i];
            table.load(['alignment', 'font/size', 'font/name']);
            await context.sync();
            if (table.alignment && !isCentered(table.alignment) && table.alignment !== 'Mixed' && table.alignment !== 'Unknown') {
                errs.push({ type: 'warning', title: `Tablo ${i + 1}: Hizalama`, description: `Tablolar ortalanmış olmalı. Mevcut: ${table.alignment}`, severity: 'FORMAT' });
                table.font.highlightColor = HIGHLIGHT.FORMAT;
            }
            if (table.font.size && Math.abs(table.font.size - RULES.FONT_SIZE_TABLE_CONTENT) > 0.5) {
                errs.push({ type: 'warning', title: `Tablo ${i + 1}: Punto`, description: `Tablo içeriği 11 punto olmalı. Mevcut: ${table.font.size} pt`, severity: 'FORMAT' });
            }
        }
    } catch (e) { logStep('TABLES', `Error: ${e.message}`); }
    return errs;
}

async function validateImages(context) {
    const errs = [];
    try {
        const pics = context.document.body.inlinePictures;
        pics.load('items');
        await context.sync();
        for (let i = 0; i < pics.items.length; i++) {
            const pic = pics.items[i];
            pic.paragraph.load('alignment');
            await context.sync();
            if (pic.paragraph.alignment && !isCentered(pic.paragraph.alignment)) {
                errs.push({ type: 'warning', title: `Resim ${i + 1}: Hizalama`, description: `Resimler ortalanmış olmalı.`, severity: 'FORMAT' });
                pic.paragraph.font.highlightColor = HIGHLIGHT.FORMAT;
            }
        }
    } catch (e) { logStep('IMAGES', `Error: ${e.message}`); }
    return errs;
}

async function validatePageNumbers(context, sections) {
    const errs = [];
    try {
        for (let i = 0; i < sections.items.length; i++) {
            try {
                sections.items[i].pageSetup.load('footerDistance');
                await context.sync();
                const fd = sections.items[i].pageSetup.footerDistance;
                if (fd !== undefined && Math.abs(fd - RULES.PAGE_NUMBER_FOOTER_PT) > RULES.MARGIN_TOLERANCE) {
                    errs.push({ type: 'warning', title: `Bölüm ${i + 1}: Sayfa No Konumu`, description: `Alt kenardan 1.25 cm yukarıda olmalı. Mevcut: ${(fd / 28.35).toFixed(2)} cm`, location: `Bölüm ${i + 1}`, severity: 'FORMAT' });
                }
                const footer = sections.items[i].getFooter("Primary");
                footer.load('text');
                await context.sync();
                if (footer.text && footer.text.trim().length === 0) {
                    errs.push({ type: 'warning', title: `Bölüm ${i + 1}: Sayfa No Eksik`, description: 'Alt bilgide sayfa numarası bulunamadı.', location: `Bölüm ${i + 1}`, severity: 'FORMAT' });
                }
            } catch (e) { logStep('PAGE_NUM', `Section ${i + 1} error: ${e.message}`); }
        }
    } catch (e) { logStep('PAGE_NUM', `Error: ${e.message}`); }
    return errs;
}

async function validateSectionMargins(context, sections) {
    const errs = [];
    try {
        for (let i = 0; i < sections.items.length; i++) {
            try {
                const ps = sections.items[i].getPageSetup();
                ps.load('topMargin, bottomMargin, leftMargin, rightMargin');
                const body = sections.items[i].body;
                body.paragraphs.load('items');
                await context.sync();
                let firstText = '';
                if (body.paragraphs.items && body.paragraphs.items.length > 0) {
                    body.paragraphs.items[0].load('text');
                    await context.sync();
                    firstText = (body.paragraphs.items[0].text || '').trim();
                }
                const isChapter = matchAny(firstText, PAT.CHAPTER) || matchAny(firstText, PAT.SPECIAL_HEADING);
                const expectedTop = isChapter ? RULES.MARGIN_TOP_SPECIAL_POINTS : RULES.MARGIN_POINTS;
                const tol = RULES.MARGIN_TOLERANCE;
                if (ps.topMargin !== undefined && Math.abs(ps.topMargin - expectedTop) > tol) {
                    errs.push({ type: 'error', title: `Bölüm ${i + 1}: Üst Kenar`, description: `${isChapter ? 7 : 3} cm olmalı. Mevcut: ${(ps.topMargin / 28.35).toFixed(2)} cm`, location: `Bölüm ${i + 1}`, severity: 'CRITICAL' });
                }
                if (ps.bottomMargin !== undefined && Math.abs(ps.bottomMargin - RULES.MARGIN_POINTS) > tol) {
                    errs.push({ type: 'error', title: `Bölüm ${i + 1}: Alt Kenar`, description: `3 cm olmalı. Mevcut: ${(ps.bottomMargin / 28.35).toFixed(2)} cm`, location: `Bölüm ${i + 1}`, severity: 'CRITICAL' });
                }
                if (ps.leftMargin !== undefined && Math.abs(ps.leftMargin - RULES.MARGIN_POINTS) > tol) {
                    errs.push({ type: 'error', title: `Bölüm ${i + 1}: Sol Kenar`, description: `3 cm olmalı. Mevcut: ${(ps.leftMargin / 28.35).toFixed(2)} cm`, location: `Bölüm ${i + 1}`, severity: 'CRITICAL' });
                }
                if (ps.rightMargin !== undefined && Math.abs(ps.rightMargin - RULES.MARGIN_POINTS) > tol) {
                    errs.push({ type: 'error', title: `Bölüm ${i + 1}: Sağ Kenar`, description: `3 cm olmalı. Mevcut: ${(ps.rightMargin / 28.35).toFixed(2)} cm`, location: `Bölüm ${i + 1}`, severity: 'CRITICAL' });
                }
            } catch (e) { logStep('MARGIN', `Section ${i + 1}: ${e.message}`); }
        }
    } catch (e) {
        logStep('MARGIN', `Error: ${e.message}`);
        addResult('warning', 'Kenar Boşlukları (Manuel Kontrol)', 'Otomatik kontrol başarısız. Lütfen manuel kontrol edin.');
    }
    return errs;
}

// ============================================
// STRUCTURE VALIDATION - Required segments check
// ============================================
function validateStructure(segments) {
    const errs = [];
    const required = [
        { type: SEG.OUTER_COVER, name: 'Dış Kapak' },
        { type: SEG.INNER_COVER, name: 'İç Kapak' },
        { type: SEG.ETHICS, name: 'Bilimsel Etiğe Uygunluk' },
        { type: SEG.ORIGINALITY, name: 'Tez Özgünlük Sayfası' },
        { type: SEG.GUIDE_COMPLIANCE, name: 'Kılavuza Uygunluk' },
        { type: SEG.APPROVAL, name: 'Kabul ve Onay Tutanağı' },
        { type: SEG.ABSTRACT_TR, name: 'Özet (Türkçe)' },
        { type: SEG.ABSTRACT_EN, name: 'Abstract (İngilizce)' },
        { type: SEG.TOC, name: 'İçindekiler' },
        { type: SEG.BIBLIOGRAPHY, name: 'Kaynakça' }
    ];

    for (const req of required) {
        if (!segments.find(s => s.type === req.type)) {
            errs.push({ type: 'error', title: `Eksik Bölüm: ${req.name}`, description: `Tezde "${req.name}" bölümü bulunamadı.`, severity: 'CRITICAL' });
        }
    }

    // Check segment order
    const order = [SEG.OUTER_COVER, SEG.INNER_COVER, SEG.ETHICS, SEG.ORIGINALITY, SEG.GUIDE_COMPLIANCE, SEG.APPROVAL, SEG.ABSTRACT_TR, SEG.ABSTRACT_EN, SEG.TOC];
    let lastIdx = -1;
    for (const expectedType of order) {
        const seg = segments.find(s => s.type === expectedType);
        if (seg) {
            if (seg.startIdx < lastIdx) {
                errs.push({ type: 'warning', title: `Sıralama: ${seg.title}`, description: `"${seg.title}" beklenen sırada değil.`, severity: 'FORMAT' });
            }
            lastIdx = seg.startIdx;
        }
    }

    return errs;
}
// ============================================
// UI FUNCTIONS
// ============================================
function initializeUI() {
    const btn = document.getElementById('scanBtn');
    if (btn) btn.onclick = scanDocument;
    logStep('UI', 'v5.0 initialized');
}

function setButtonState(enabled) {
    const btn = document.getElementById('scanBtn');
    if (btn) {
        btn.disabled = !enabled;
        const span = btn.querySelector('span');
        if (span) span.textContent = enabled ? 'DÖKÜMAN TARA' : 'ARANIYOR...';
    }
}

function updateProgress(pct, msg) {
    const c = document.getElementById('progressSection');
    const b = document.getElementById('progressFill');
    const t = document.getElementById('progressText');
    if (c) c.classList.remove('hidden');
    if (b) b.style.width = `${pct}%`;
    if (t) t.textContent = msg;
}

function hideProgress() {
    const c = document.getElementById('progressSection');
    if (c) c.classList.add('hidden');
}

function displayResults() {
    const container = document.getElementById('resultsList');
    const summary = document.getElementById('summarySection');
    const errEl = document.getElementById('errorCount');
    const warnEl = document.getElementById('warningCount');
    const succEl = document.getElementById('successCount');
    const filterTabs = document.getElementById('filterTabs');

    if (!container) return;

    if (validationResults.length === 0) {
        container.innerHTML = '<div class="empty-state"><p>✅ Hiçbir hata bulunamadı.</p></div>';
        if (summary) summary.classList.add('hidden');
        return;
    }

    if (summary) summary.classList.remove('hidden');
    if (filterTabs) filterTabs.classList.remove('hidden');

    const errors = validationResults.filter(r => r.type === 'error');
    const warnings = validationResults.filter(r => r.type === 'warning');
    const successes = validationResults.filter(r => r.type === 'success');

    if (errEl) errEl.textContent = errors.length;
    if (warnEl) warnEl.textContent = warnings.length;
    if (succEl) succEl.textContent = successes.length;

    let html = '';
    for (const r of errors) html += createResultItem(r, 'error');
    for (const r of warnings) html += createResultItem(r, 'warning');
    for (const r of successes) html += createResultItem(r, 'success');
    container.innerHTML = html;

    // Setup filter tabs
    if (filterTabs) {
        filterTabs.querySelectorAll('.filter-tab').forEach(tab => {
            tab.onclick = function () {
                filterTabs.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
                this.classList.add('active');
                const filter = this.dataset.filter;
                container.querySelectorAll('.result-item').forEach(item => {
                    if (filter === 'all') { item.style.display = ''; return; }
                    item.style.display = item.classList.contains(filter) ? '' : 'none';
                });
            };
        });
    }
}

function createResultItem(result, type) {
    const paraMatch = result.location ? result.location.match(/Paragraf\s*(\d+)/i) : null;
    const paraIdx = paraMatch ? parseInt(paraMatch[1]) - 1 : null;
    const showBtn = paraIdx !== null ? `<button class="show-error-btn" onclick="goToError(${paraIdx})">GÖSTER</button>` : '';
    const segLabel = result.segment ? `<span class="segment-label">${result.segment}</span>` : '';

    return `<div class="result-item ${type}">
<div class="result-header">${segLabel}<span class="result-title">${result.title}</span>${showBtn}</div>
<div class="result-description">${result.description}</div>
<div class="result-location">${result.location || ''}</div>
</div>`;
}

async function goToError(pidx) {
    try {
        await Word.run(async (ctx) => {
            const paras = ctx.document.body.paragraphs;
            paras.load('items');
            await ctx.sync();
            if (pidx >= 0 && pidx < paras.items.length) {
                paras.items[pidx].select();
                await ctx.sync();
            }
        });
    } catch (e) { console.error('Nav error:', e); }
}

async function clearHighlightsAndResults() {
    clearResults();
    try {
        await Word.run(async (ctx) => {
            ctx.document.body.font.highlightColor = null;
            await ctx.sync();
        });
        displayResults();
    } catch (e) { console.error('Clear error:', e); }
}

// ============================================
// MAIN SCAN FUNCTION
// ============================================
async function scanDocument() {
    if (isScanning) return;
    isScanning = true;
    const t0 = performance.now();
    clearResults();
    setButtonState(false);
    updateProgress(0, 'Tarama başlatılıyor...');

    try {
        await Word.run(async (context) => {
            // Step 1: Clear highlights
            updateProgress(5, 'Önceki işaretler temizleniyor...');
            context.document.body.font.highlightColor = null;
            await context.sync();

            // Step 2: Load everything
            updateProgress(10, 'Belge yapısı yükleniyor...');
            const sections = context.document.sections;
            sections.load('items');
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load([
                'items/text', 'items/style', 'items/outlineLevel', 'items/tableNestingLevel',
                'items/font/name', 'items/font/size', 'items/font/bold', 'items/font/italic',
                'items/alignment', 'items/firstLineIndent', 'items/leftIndent', 'items/rightIndent',
                'items/lineSpacing', 'items/spaceBefore', 'items/spaceAfter',
                'items/isListItem', 'items/listItemOrNullObject/listString', 'items/listItemOrNullObject/level'
            ].join(','));
            await context.sync();
            logStep('LOAD', `${paragraphs.items.length} paragraf, ${sections.items.length} bölüm yüklendi`);

            // Step 3: Build paragraph data
            updateProgress(20, 'Paragraf verileri hazırlanıyor...');
            const pdList = [];
            for (let i = 0; i < paragraphs.items.length; i++) {
                const p = paragraphs.items[i];
                const f = p.font || {};
                const li = p.listItemOrNullObject;
                let ls = '', ll = null;
                if (li && !li.isNullObject) { ls = li.listString || ''; ll = li.level; }
                pdList.push({
                    index: i, text: p.text || '', style: p.style || '', outlineLevel: p.outlineLevel,
                    tableNestingLevel: p.tableNestingLevel || 0,
                    font: { name: f.name, size: f.size, bold: f.bold, italic: f.italic },
                    alignment: p.alignment, firstLineIndent: p.firstLineIndent,
                    leftIndent: p.leftIndent, rightIndent: p.rightIndent,
                    lineSpacing: p.lineSpacing, spaceBefore: p.spaceBefore, spaceAfter: p.spaceAfter,
                    isListItem: p.isListItem || false, listString: ls, listLevel: ll, paragraph: p
                });
            }

            // Step 4: Detect segments
            updateProgress(30, 'Tez yapısı analiz ediliyor...');
            const segments = detectSegments(pdList);
            logStep('SEGMENTS', `${segments.length} segment tespit edildi`, segments.map(s => `${s.title}[${s.startIdx}-${s.endIdx}]`));

            // Step 5: Validate structure
            const structErrors = validateStructure(segments);
            for (const e of structErrors) addResult(e.type, e.title, e.description, 'Yapı Kontrolü', null, e.severity, 'Yapı');

            // Step 6: Validate margins
            updateProgress(35, 'Kenar boşlukları kontrol ediliyor...');
            const marginErrs = await validateSectionMargins(context, sections);
            for (const e of marginErrs) addResult(e.type, e.title, e.description, e.location, null, e.severity, 'Kenar Boşlukları');

            // Step 7: Paragraph-by-paragraph validation
            updateProgress(40, 'Paragraflar doğrulanıyor...');
            let ghostCount = 0;

            for (let i = 0; i < pdList.length; i++) {
                if (i % 50 === 0) updateProgress(40 + Math.floor((i / pdList.length) * 40), `Paragraf ${i + 1} / ${pdList.length}`);

                const pd = pdList[i];
                const seg = getSegmentForParagraph(segments, i);
                const ptype = detectParaType(pd, seg);
                let errors = [];
                const segTitle = seg ? seg.title : '';

                switch (ptype) {
                    case PTYPE.GHOST_HEADING:
                        errors = validateGhostHeading(pd, i);
                        pd.paragraph.font.highlightColor = HIGHLIGHT.CRITICAL;
                        ghostCount++;
                        break;
                    case PTYPE.COVER_TEXT:
                        if (seg) errors = validateCover(pd, i, seg.type);
                        break;
                    case PTYPE.FRONT_MATTER:
                        if (seg) errors = validateFrontMatter(pd, i, seg);
                        break;
                    case PTYPE.CHAPTER_HEADING:
                        errors = validateChapterHeading(pd, i);
                        break;
                    case PTYPE.SUB_HEADING:
                        errors = validateSubHeading(pd, i);
                        break;
                    case PTYPE.BODY_TEXT:
                        if (!seg || ![SEG.OUTER_COVER, SEG.INNER_COVER, SEG.ETHICS, SEG.ORIGINALITY, SEG.GUIDE_COMPLIANCE, SEG.APPROVAL, SEG.TOC].includes(seg.type)) {
                            errors = validateBodyText(pd, i);
                        }
                        break;
                    case PTYPE.BLOCK_QUOTE:
                        errors = validateBlockQuote(pd, i);
                        break;
                    case PTYPE.BIBLIOGRAPHY:
                        errors = validateBibliography(pd, i);
                        break;
                    case PTYPE.CAPTION:
                        errors = validateCaption(pd, i);
                        break;
                    case PTYPE.TOC_ENTRY:
                    case PTYPE.EMPTY:
                    case PTYPE.UNKNOWN:
                        break;
                }

                // Add errors and highlight
                let hasCritical = false;
                for (const err of errors) {
                    addResult(err.type, err.title, err.description, `Paragraf ${i + 1}`, err.paraIndex, err.severity, segTitle);
                    if (err.severity === 'CRITICAL' || err.type === 'error') hasCritical = true;
                }
                if (errors.length > 0) {
                    pd.paragraph.font.highlightColor = hasCritical ? HIGHLIGHT.CRITICAL : HIGHLIGHT.FORMAT;
                }
            }

            // Step 8: Tables
            updateProgress(82, 'Tablolar kontrol ediliyor...');
            const tblErrs = await validateTables(context);
            for (const e of tblErrs) addResult(e.type, e.title, e.description, `Tablo`, null, e.severity, 'Tablolar');

            // Step 9: Images
            updateProgress(86, 'Resimler kontrol ediliyor...');
            const imgErrs = await validateImages(context);
            for (const e of imgErrs) addResult(e.type, e.title, e.description, `Resim`, null, e.severity, 'Resimler');

            // Step 10: Page numbers
            updateProgress(89, 'Sayfa numaraları kontrol ediliyor...');
            const pnErrs = await validatePageNumbers(context, sections);
            for (const e of pnErrs) addResult(e.type, e.title, e.description, e.location, null, e.severity, 'Sayfa No');

            // Step 11: Abstract validation
            updateProgress(92, 'Özet/Abstract kontrol ediliyor...');
            const ozetSeg = segments.find(s => s.type === SEG.ABSTRACT_TR);
            if (ozetSeg) {
                const ozetErrs = validateAbstractSection(pdList, ozetSeg.startIdx, false);
                for (const e of ozetErrs) addResult(e.type, e.title, e.description, 'Özet Sayfası', null, e.severity, 'Özet');
            }
            const absSeg = segments.find(s => s.type === SEG.ABSTRACT_EN);
            if (absSeg) {
                const absErrs = validateAbstractSection(pdList, absSeg.startIdx, true);
                for (const e of absErrs) addResult(e.type, e.title, e.description, 'Abstract Sayfası', null, e.severity, 'Abstract');
            }

            // Step 12: Apply & summary
            updateProgress(96, 'İşaretler uygulanıyor...');
            await context.sync();

            if (ghostCount > 0) {
                addResult('error', `${ghostCount} Boş Başlık (Ghost Heading)`, 'İçindekiler tablosunda hatalı satırlara neden olur.', 'Belge', null, 'CRITICAL', 'Yapı');
            }

            const errs = validationResults.filter(r => r.type === 'error').length;
            const warns = validationResults.filter(r => r.type === 'warning').length;
            const succs = validationResults.filter(r => r.type === 'success').length;

            if (errs === 0 && warns === 0) {
                addResult('success', '✅ Tebrikler!', 'Belge EBYÜ 2022 Tez Yazım Kılavuzu formatına uygun görünüyor.');
            } else {
                addResult(errs > 0 ? 'error' : 'warning', 'Tarama Özeti',
                    `🔴 Kritik: ${errs} | 🟡 Format: ${warns} | ✅ Başarılı: ${succs}`);
            }

            updateProgress(100, 'Tarama tamamlandı!');
            logStep('DONE', `${((performance.now() - t0) / 1000).toFixed(2)} saniye`);
        });
    } catch (error) {
        logStep('ERROR', error.message);
        addResult('error', 'Tarama Hatası', `Hata: ${error.message}. Lütfen tekrar deneyin.`);
    } finally {
        isScanning = false;
        setButtonState(true);
        hideProgress();
        displayResults();
    }
}

// ============================================
// OFFICE.JS INIT
// ============================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log('EBYÜ Thesis Validator v5.0 (Segment-Based): Office.js initialized');
        initializeUI();
    } else {
        console.error('This add-in only works with Microsoft Word.');
    }
});
