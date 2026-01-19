/**
 * ============================================
 * EBYÜ Thesis Format Validator - Task Pane Logic
 * Erzincan Binali Yıldırım University
 * Based on: EBYÜ 2022 Tez Yazım Kılavuzu
 * ============================================
 * 
 * ZONE-BASED VALIDATION ARCHITECTURE:
 * - Zone 1: HEADINGS (Main + Sub)
 * - Zone 2: BODY (Standard paragraphs)
 * - Zone 3: BLOCK QUOTES (>40 words, indented)
 * - Zone 4: BIBLIOGRAPHY (After KAYNAKÇA)
 * - Zone 5: CAPTIONS (Tablo/Şekil)
 * 
 * ERROR SEVERITY:
 * - CRITICAL (Red #FF0000): Margins, Ghost Headings, Wrong Font Family
 * - FORMAT (Yellow #FFFF00): Wrong Indent, Spacing, Size
 */

// ============================================
// CONSTANTS - EBYÜ 2022 Strict Rules
// ============================================

const EBYÜ_RULES = {
    // Page Layout
    MARGIN_CM: 3,
    MARGIN_POINTS: 85.05, // 3cm = 85.05pt
    MARGIN_TOLERANCE: 2,  // Allow ±2pt tolerance

    // Fonts
    FONT_NAME: "Times New Roman",
    FONT_SIZE_BODY: 12,
    FONT_SIZE_HEADING_MAIN: 14,
    FONT_SIZE_HEADING_SUB: 12,
    FONT_SIZE_BLOCK_QUOTE: 11,
    FONT_SIZE_FOOTNOTE: 10,
    FONT_SIZE_TABLE: 11,
    FONT_SIZE_CAPTION_TITLE: 12,    // Caption başlığı 12pt
    FONT_SIZE_CAPTION_CONTENT: 11,  // Tablo/Şekil içi açıklamalar 11pt
    FONT_SIZE_COVER_TITLE: 16,      // Kapak başlıkları 16pt
    FONT_SIZE_EPIGRAPH: 11,         // Epigraf 11pt

    // Line Spacing
    LINE_SPACING_BODY: 1.5,      // 1.5 lines
    LINE_SPACING_SINGLE: 1.0,    // Single
    LINE_SPACING_POINTS_1_5: 18, // 1.5 lines check (~18pt for 12pt font)
    LINE_SPACING_POINTS_SINGLE: 12,

    // Indentation
    FIRST_LINE_INDENT_CM: 1.25,
    FIRST_LINE_INDENT_POINTS: 35.4,  // 1.25cm = 35.4pt
    BLOCK_QUOTE_INDENT_CM: 1.25,
    BLOCK_QUOTE_INDENT_POINTS: 35.4,
    INDENT_TOLERANCE: 2,             // Strict 2pt tolerance

    // Spacing Around (NK = point, 6nk = 6pt)
    SPACING_6NK: 6,
    SPACING_3NK: 3,
    SPACING_0NK: 0,
    SPACING_TOLERANCE: 1,

    // Detection Thresholds
    BLOCK_QUOTE_MIN_INDENT_POINTS: 28, // ~1cm minimum to detect block quote
    MIN_BODY_TEXT_LENGTH: 100,         // Minimum chars to consider as body text

    // Page Dimensions (A4)
    PAGE_WIDTH_POINTS: 595.3,
    CONTENT_MAX_WIDTH_POINTS: 425.2,   // 595.3 - (2 * 85.05)
};

// Highlight Colors for Different Error Severities
const HIGHLIGHT_COLORS = {
    CRITICAL: "Red",
    FORMAT: "Yellow",
    FOUND: "Cyan"     // For "GÖSTER" functionality
};

// ============================================
// ZONE / PARAGRAPH TYPE DEFINITIONS
// ============================================

const PARA_TYPES = {
    TOC_ENTRY: 'TOC_ENTRY',
    MAIN_HEADING: 'MAIN_HEADING',
    SUB_HEADING: 'SUB_HEADING',
    BODY_TEXT: 'BODY_TEXT',
    BLOCK_QUOTE: 'BLOCK_QUOTE',
    BIBLIOGRAPHY: 'BIBLIOGRAPHY',
    CAPTION_TITLE: 'CAPTION_TITLE',
    CAPTION_CONTENT: 'CAPTION_CONTENT',
    EPIGRAPH: 'EPIGRAPH',
    LIST_ITEM: 'LIST_ITEM',
    GHOST_HEADING: 'GHOST_HEADING',
    FRONT_MATTER: 'FRONT_MATTER',
    COVER_TEXT: 'COVER_TEXT',
    EMPTY: 'EMPTY',
    UNKNOWN: 'UNKNOWN'
};

const ZONES = {
    COVER: 'COVER',                   // Kapak sayfası
    FRONT_MATTER: 'FRONT_MATTER',
    TABLE_OF_CONTENTS: 'TABLE_OF_CONTENTS', // İçindekiler tablosu
    ABSTRACT_TR: 'ABSTRACT_TR',
    ABSTRACT_EN: 'ABSTRACT_EN',
    BODY: 'BODY',
    BACK_MATTER: 'BACK_MATTER'
};

// ============================================
// DETECTION PATTERNS
// ============================================

const PATTERNS = {
    // Main chapter headings (must be centered, bold, 14pt, ALL CAPS)
    MAIN_HEADING: [
        /^(BİRİNCİ|İKİNCİ|ÜÇÜNCÜ|DÖRDÜNCÜ|BEŞİNCİ|ALTINCI|YEDİNCİ|SEKİZİNCİ|DOKUZUNCU|ONUNCU)\s*BÖLÜM$/i,
        /^BÖLÜM\s*[IVX\d]+/i,
        /^(GİRİŞ|SONUÇ|SONUÇ VE ÖNERİLER|TARTIŞMA|KAYNAKÇA|KAYNAKLAR|ÖZET|ABSTRACT|SUMMARY)$/i,
        /^ÖN\s*SÖZ$/i,
        /^KISALTMALAR\s*(LİSTESİ|DİZİNİ)?$/i,
        /^(TABLOLAR|ŞEKİLLER|GRAFİKLER|SİMGELER)\s*(LİSTESİ|DİZİNİ)?$/i,
        /^İÇİNDEKİLER$/i,
        /^EKLER?$/i
    ],

    // Sub-headings (numbered like 1.1, 2.3.1)
    SUB_HEADING: [
        /^\d+\.\d+(\.\d+)*\.?\s+[A-ZÇĞİÖŞÜa-zçğıöşü]/, // 1.1 or 1.1.1
    ],

    // Captions - must include chapter number (e.g., Tablo 1.1:, Şekil 2.3:)
    CAPTION_TABLE: /^Tablo\s*(\d+)\.(\d+)\s*[:.]/i,
    CAPTION_FIGURE: /^(Şekil|Grafik|Resim|Harita)\s*(\d+)\.(\d+)\s*[:.]/i,

    // TOC patterns
    TOC_STYLE: [/^TOC/i, /^İçindekiler/i, /^Table of Contents/i],
    TOC_CONTENT: /\.{5,}\s*(i|v|x|\d)+$/i, // Dots followed by page number
    TOC_START: /^İÇİNDEKİLER$/i,

    // Zone switching
    BODY_START: [/^GİRİŞ$/i],
    BACK_MATTER_START: [/^(KAYNAKÇA|KAYNAKLAR|REFERANSLAR|REFERENCES)$/i],

    // Cover page patterns
    COVER_IDENTIFIERS: [
        /^T\.?C\.?$/i,
        /^ERZİNCAN\s*BİNALİ\s*YILDIRIM/i,
        /^ÜNİVERSİTESİ$/i
    ]
};

// ============================================
// GLOBAL STATE
// ============================================

let validationResults = [];
let currentFilter = 'all';
let scanLog = [];
let isInBibliographyZone = false;

// ============================================
// HELPER FUNCTIONS - Type Detection
// ============================================

function logStep(category, message, details = null) {
    const timestamp = new Date().toISOString();
    scanLog.push({ timestamp, category, message, details });
    console.log(`[${category}] ${message}`, details || '');
}

/**
 * Check if style name indicates a heading
 */
function isHeadingStyle(style) {
    if (!style) return false;
    const normalized = style.toLowerCase();
    return normalized.includes("heading") ||
        normalized.includes("başlık") ||
        /^heading\s*\d/i.test(style) ||
        /^başlık\s*\d/i.test(style);
}

/**
 * Check if paragraph is TOC entry (skip validation)
 */
/**
 * Helper: Check if string is Title Case (Every word starts with uppercase)
 * Specifically for EBYÜ sub-headings.
 */
function isTitleCase(text) {
    if (!text) return true;
    // Remove leading numbers like "1.1. "
    const cleanText = text.replace(/^\d+(\.\d+)*\.?\s+/, "").trim();
    if (!cleanText) return true;

    const words = cleanText.split(/\s+/);
    return words.every(word => {
        if (word.length === 0) return true;
        // Check if first letter is uppercase
        const firstChar = word[0];
        // Turkish specific: İ, Ç, Ğ, Ö, Ş, Ü
        const isUpper = firstChar === firstChar.toUpperCase() && firstChar !== firstChar.toLowerCase();
        return isUpper;
    });
}

function isTOCEntry(style, text) {
    if (!style && !text) return false;
    const trimmed = (text || '').trim();

    // Style-based detection
    if (style) {
        if (PATTERNS.TOC_STYLE.some(p => p.test(style.trim()))) return true;
    }

    // Pattern-based detection
    // 1. Check for dots followed by page number (TOC_CONTENT)
    if (PATTERNS.TOC_CONTENT.test(trimmed)) return true;

    // 2. Check if it ends with a number (common for TOC entries without dots)
    if (/\d+$/.test(trimmed) && (style || '').toLowerCase().startsWith('toc')) return true;

    return false;
}

/**
 * Check if text is a main chapter heading
 */
function isMainHeading(text) {
    if (!text) return false;
    const trimmed = text.trim();
    return PATTERNS.MAIN_HEADING.some(pattern => pattern.test(trimmed));
}

/**
 * Check if text is a sub-heading (e.g., 1.1 Alt Başlık)
 */
function isSubHeading(text, isBold) {
    if (!text || !isBold) return false;
    const trimmed = text.trim();
    return PATTERNS.SUB_HEADING.some(pattern => pattern.test(trimmed));
}

/**
 * Check if text is a caption (Tablo X.Y or Şekil X.Y)
 * Returns object with isCaption, isCorrectFormat, and captionType
 */
function isCaption(text) {
    if (!text) return { isCaption: false };
    const trimmed = text.trim();

    // Correct Format: "Tablo 1.1:" or "Şekil 2.1."
    if (PATTERNS.CAPTION_TABLE.test(trimmed)) {
        return { isCaption: true, isCorrect: true, type: 'table' };
    }
    if (PATTERNS.CAPTION_FIGURE.test(trimmed)) {
        return { isCaption: true, isCorrect: true, type: 'figure' };
    }

    // Incorrect Format: "Tablo 1:" or "Tablo 1.1" (missing colon/dot) or "Şekil 1"
    const wrongFormat = /^Tablo\s*\d+/i.test(trimmed) || /^(Şekil|Grafik|Resim|Harita)\s*\d+/i.test(trimmed);
    if (wrongFormat) {
        return { isCaption: true, isCorrect: false, type: 'unknown' };
    }

    return { isCaption: false };
}

function isHeadingStyle(style) {
    if (!style) return false;
    const s = style.toLowerCase();
    // Default heading styles and common Turkish variations
    return s.includes("heading") || s.includes("başlık") || s.includes("title") || s.includes("bölüm");
}

function isHeadingByOutline(outlineLevel) {
    // outlineLevel 1-9 indicates structural headings in Word
    return outlineLevel >= 1 && outlineLevel <= 9;
}

function isEpigraph(para, font) {
    if (!para || !font) return false;
    // Epigraph is right-aligned, italic, 11pt
    return para.alignment === Word.Alignment.right &&
        font.italic === true &&
        font.size === EBYÜ_RULES.FONT_SIZE_EPIGRAPH;
}

function isCoverItem(text) {
    if (!text) return false;
    const trimmed = text.trim();
    return PATTERNS.COVER_IDENTIFIERS.some(p => p.test(trimmed));
}

/**
 * Check if text is a table caption (must be above table)
 */
function isTableCaption(text) {
    if (!text) return false;
    return PATTERNS.CAPTION_TABLE.test(text.trim());
}

/**
 * Check if text is a figure caption (must be below figure)
 */
function isFigureCaption(text) {
    if (!text) return false;
    return PATTERNS.CAPTION_FIGURE.test(text.trim());
}

/**
 * Detect if paragraph is a block quote based on indentation
 * Block quotes: LeftIndent > 1cm AND not a list item
 */
function isBlockQuote(para) {
    // Must have significant left indent
    const leftIndent = para.leftIndent || 0;
    if (leftIndent < EBYÜ_RULES.BLOCK_QUOTE_MIN_INDENT_POINTS) return false;

    // Should not be a list item
    if (para.listItemOrNull !== null && para.listItemOrNull !== undefined) return false;

    // Should have some text content
    const text = (para.text || '').trim();
    if (text.length < 20) return false;

    return true;
}

/**
 * Check if paragraph is a list item
 */
function isListItem(para) {
    return para.listItemOrNull !== null && para.listItemOrNull !== undefined;
}

/**
 * Check if text matches any pattern in array
 */
function matchesAnyPattern(text, patterns) {
    if (!text) return false;
    const trimmed = text.trim();
    return patterns.some(pattern => pattern.test(trimmed));
}

/**
 * Check if text is ALL CAPS (for main heading validation)
 */
function isAllCaps(text) {
    if (!text) return false;
    const trimmed = text.trim();
    // Remove non-letter characters and check if all letters are uppercase
    const letters = trimmed.replace(/[^A-ZÇĞİÖŞÜa-zçğıöşü]/g, '');
    if (letters.length === 0) return true;
    return letters === letters.toUpperCase().replace(/İ/g, 'İ').replace(/I/g, 'I');
}

/**
 * Detect paragraph type based on content, style, and formatting
 */
function detectParagraphType(para, text, style, font, zone, isInBiblio) {
    const trimmed = (text || '').trim();
    const outlineLevel = para.outlineLevel;

    // Empty paragraph
    if (trimmed.length === 0) {
        if (isHeadingStyle(style) || isHeadingByOutline(outlineLevel)) return PARA_TYPES.GHOST_HEADING;
        return PARA_TYPES.EMPTY;
    }

    // TOC Entry
    if (isTOCEntry(style, trimmed)) return PARA_TYPES.TOC_ENTRY;

    // Cover Page detection (if not already in zone)
    if (zone === ZONES.COVER || isCoverItem(trimmed)) return PARA_TYPES.COVER_TEXT;

    // Caption
    const captionInfo = isCaption(trimmed);
    if (captionInfo.isCaption) {
        return PARA_TYPES.CAPTION_TITLE;
    }

    // Epigraph
    if (isEpigraph(para, font)) return PARA_TYPES.EPIGRAPH;

    // Bibliography
    if (isInBiblio && trimmed.length > 5) {
        if (!matchesAnyPattern(trimmed, PATTERNS.BACK_MATTER_START)) return PARA_TYPES.BIBLIOGRAPHY;
    }

    // Main Heading Detection
    const isMainByText = isMainHeading(trimmed);
    const isMainByOutline = outlineLevel === 1;
    const isMainByStyle = isHeadingStyle(style) && (style.includes("1") || /heading\s*1/i.test(style));

    if (isMainByText || isMainByOutline || isMainByStyle) {
        return PARA_TYPES.MAIN_HEADING;
    }

    // Sub-Heading Detection
    const isSubByText = isSubHeading(trimmed, font.bold === true);
    const isSubByOutline = outlineLevel > 1 && outlineLevel <= 4;
    const isSubByStyle = isHeadingStyle(style);

    if (isSubByText || isSubByOutline || isSubByStyle) {
        return PARA_TYPES.SUB_HEADING;
    }

    // Block Quote
    if (isBlockQuote(para)) return PARA_TYPES.BLOCK_QUOTE;

    // List Item
    if (isListItem(para)) return PARA_TYPES.LIST_ITEM;

    // Body Text
    if (trimmed.length >= 20) return PARA_TYPES.BODY_TEXT;

    return PARA_TYPES.UNKNOWN;
}

// ============================================
// VALIDATION RULES - Zone Specific
// ============================================

/**
 * Validate Main Heading (Zone 1A)
 * Rules: Centered, 14pt, Bold, ALL CAPS, 6nk before/after
 */
function validateMainHeading(para, font, text, index) {
    const errors = [];
    const trimmed = (text || '').trim();

    // Alignment: Centered
    if (para.alignment !== Word.Alignment.centered) {
        errors.push({
            type: 'error',
            title: 'Ana Başlık: Hizalama Hatası',
            description: 'Ana başlıklar ORTALANMALI.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Font: 14pt Bold
    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_HEADING_MAIN) > 0.5) {
        errors.push({
            type: 'error',
            title: 'Ana Başlık: Punto Hatası',
            description: `14 punto olmalı. Mevcut: ${font.size} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }
    if (font.bold !== true) {
        errors.push({
            type: 'error',
            title: 'Ana Başlık: Kalın Yazı Hatası',
            description: 'Ana başlıklar KALIN olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // ALL CAPS
    if (!isAllCaps(trimmed)) {
        errors.push({
            type: 'warning',
            title: 'Ana Başlık: Büyük Harf Hatası',
            description: 'Ana başlıklar TAMAMI BÜYÜK HARF olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Spacing: 6nk (6pt)
    const spaceBefore = para.spaceBefore || 0;
    const spaceAfter = para.spaceAfter || 0;
    if (Math.abs(spaceBefore - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE) {
        errors.push({
            type: 'warning',
            title: 'Ana Başlık: Öncesi Boşluk',
            description: `6nk olmalı. Mevcut: ${spaceBefore.toFixed(1)} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }
    if (Math.abs(spaceAfter - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE) {
        errors.push({
            type: 'warning',
            title: 'Ana Başlık: Sonrası Boşluk',
            description: `6nk olmalı. Mevcut: ${spaceAfter.toFixed(1)} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    return errors;
}

/**
 * Validate Sub-Heading (Zone 1B)
 * Rules: Left Aligned (or Justified), 12pt, Bold, 6nk before/after
 */
function validateSubHeading(para, font, text, index) {
    const errors = [];

    // Font: 12pt Bold
    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_HEADING_SUB) > 0.5) {
        errors.push({
            type: 'error',
            title: 'Alt Başlık: Punto Hatası',
            description: `12 punto olmalı. Mevcut: ${font.size} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }
    if (font.bold !== true) {
        errors.push({
            type: 'error',
            title: 'Alt Başlık: Kalın Yazı Hatası',
            description: 'Alt başlıklar KALIN olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Indentation: 1.25cm (35.4pt)
    const firstIndent = para.firstLineIndent || 0;
    if (Math.abs(firstIndent - EBYÜ_RULES.FIRST_LINE_INDENT_POINTS) > EBYÜ_RULES.INDENT_TOLERANCE) {
        errors.push({
            type: 'warning',
            title: 'Alt Başlık: Girinti Hatası',
            description: `Paragraf başı 1.25 cm olmalı.`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Spacing: 6nk
    const spaceBefore = para.spaceBefore || 0;
    const spaceAfter = para.spaceAfter || 0;
    if (Math.abs(spaceBefore - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE) {
        errors.push({
            type: 'warning',
            title: 'Alt Başlık: Öncesi Boşluk',
            description: `6nk olmalı.`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }
    if (Math.abs(spaceAfter - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE) {
        errors.push({
            type: 'warning',
            title: 'Alt Başlık: Sonrası Boşluk',
            description: `6nk olmalı.`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Rule: Subheadings should be Title Case (First Letter of Each Word Uppercase)
    if (!isTitleCase(text)) {
        errors.push({
            type: 'warning',
            title: 'Alt Başlık: Küçük Harf Uyarısı',
            description: 'Alt başlıklarda her kelimenin ilk harfi büyük olmalıdır.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    return errors;
}

/**
 * Validate Body Text (Zone 2)
 * Rules: Times New Roman, 12pt, Justified, First Line Indent 1.25cm, 1.5 line spacing
 * IMPORTANT: Only validates actual body text, NOT headings, centered text, or list items
 */
function validateBodyText(para, font, text, index) {
    const errors = [];
    const trimmed = (text || '').trim();

    // Font: 12pt TNR
    if (font.name && font.name !== "Times New Roman") {
        errors.push({
            type: 'error',
            title: 'Gövde Metin: Yazı Tipi',
            description: 'Times New Roman olmalı.',
            severity: 'CRITICAL',
            paraIndex: index
        });
    }
    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_BODY) > 0.5) {
        errors.push({
            type: 'warning',
            title: 'Gövde Metin: Punto',
            description: '12 punto olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Alignment: Justified
    if (para.alignment !== Word.Alignment.justified) {
        errors.push({
            type: 'error',
            title: 'Gövde Metin: Hizalama',
            description: 'İki yana yaslı olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Indentation: 1.25cm
    const firstIndent = para.firstLineIndent || 0;
    if (Math.abs(firstIndent - EBYÜ_RULES.FIRST_LINE_INDENT_POINTS) > EBYÜ_RULES.INDENT_TOLERANCE) {
        errors.push({
            type: 'warning',
            title: 'Gövde Metin: Girinti',
            description: `Paragraf başı 1.25 cm olmalı. Mevcut: ${(firstIndent / 28.35).toFixed(2)} cm`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Spacing: 6nk before/after
    const spaceBefore = para.spaceBefore || 0;
    const spaceAfter = para.spaceAfter || 0;
    if (Math.abs(spaceBefore - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE) {
        errors.push({
            type: 'warning',
            title: 'Gövde Metin: Öncesi Boşluk',
            description: `6nk olmalı. Mevcut: ${spaceBefore.toFixed(1)} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }
    if (Math.abs(spaceAfter - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE) {
        errors.push({
            type: 'warning',
            title: 'Gövde Metin: Sonrası Boşluk',
            description: `6nk olmalı. Mevcut: ${spaceAfter.toFixed(1)} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Line Spacing: 1.5
    const lineSpacing = para.lineSpacing;
    if (lineSpacing && Math.abs(lineSpacing - 18) > 2) {
        errors.push({
            type: 'warning',
            title: 'Gövde Metin: Satır Aralığı',
            description: '1.5 satır aralığı olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    return errors;
}

/**
 * Validate Block Quote (Zone 3)
 * Rules: 11pt, Single line spacing, 1.25cm left+right indent
 */
function validateBlockQuote(para, font, text, index) {
    const errors = [];

    // Must be 11pt (IMPORTANT EXCEPTION!)
    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_BLOCK_QUOTE) > 0.5) {
        errors.push({
            type: 'error',
            title: 'Blok Alıntı: Punto Hatası',
            description: `Blok alıntılar 11 punto (pt) olmalı. Mevcut: ${font.size} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Should be Single line spacing (~12pt for 12pt font, but since 11pt, ~11-13pt)
    const lineSpacing = para.lineSpacing;
    if (lineSpacing && lineSpacing > 14) {
        errors.push({
            type: 'warning',
            title: 'Blok Alıntı: Satır Aralığı',
            description: `Tek satır aralığı (1.0) olmalı. Mevcut: ${lineSpacing.toFixed(1)} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Check left indent (should be ≈1.25cm)
    const leftIndent = para.leftIndent || 0;
    const rightIndent = para.rightIndent || 0;
    const expectedIndent = EBYÜ_RULES.BLOCK_QUOTE_INDENT_POINTS;

    if (Math.abs(leftIndent - expectedIndent) > EBYÜ_RULES.INDENT_TOLERANCE * 2) {
        errors.push({
            type: 'warning',
            title: 'Blok Alıntı: Sol Girinti',
            description: `Sol girinti 1.25 cm olmalı. Mevcut: ${(leftIndent / 28.35).toFixed(2)} cm`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    if (Math.abs(rightIndent - expectedIndent) > EBYÜ_RULES.INDENT_TOLERANCE * 2) {
        errors.push({
            type: 'warning',
            title: 'Blok Alıntı: Sağ Girinti',
            description: `Sağ girinti 1.25 cm olmalı. Mevcut: ${(rightIndent / 28.35).toFixed(2)} cm`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    return errors;
}

/**
 * Validate Bibliography Entry (Zone 4)
 * Rules: Justified/Left, Hanging indent (First Line 0, Left > 0), Single spacing
 */
function validateBibliography(para, font, text, index) {
    const errors = [];

    // Alignment: Justified or Left is OK
    if (para.alignment !== Word.Alignment.justified &&
        para.alignment !== Word.Alignment.left) {
        errors.push({
            type: 'warning',
            title: 'Kaynakça: Hizalama',
            description: 'İki yana yaslı veya sola yaslı olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Hanging Indent Check: First line should NOT be indented, but left margin should be
    // In Word, hanging indent = negative firstLineIndent with positive leftIndent
    // OR firstLineIndent = 0 and leftIndent > 0
    const firstIndent = para.firstLineIndent || 0;
    const leftIndent = para.leftIndent || 0;

    // If there's a positive first line indent, that's wrong for bibliography
    if (firstIndent > 5) {
        errors.push({
            type: 'error',
            title: 'Kaynakça: Girinti Hatası',
            description: 'Kaynakça girişlerinde ilk satır girintisi OLMAMALI. Asılı (Hanging) girinti kullanın.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Check for single line spacing
    const lineSpacing = para.lineSpacing;
    if (lineSpacing && lineSpacing > 14) {
        errors.push({
            type: 'warning',
            title: 'Kaynakça: Satır Aralığı',
            description: `Tek satır aralığı (1.0) olmalı. Mevcut: ${lineSpacing.toFixed(1)} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Check paragraph spacing (3nk before/after for bibliography entries)
    const spaceAfter = para.spaceAfter || 0;

    // Bibliography entries should have some space after (around 3pt) but not too much
    if (spaceAfter > 10) {
        errors.push({
            type: 'warning',
            title: 'Kaynakça: Giriş Sonrası Boşluk',
            description: `Kaynak girişi sonrası boşluk çok fazla. 3nk (3pt) olmalı. Mevcut: ${spaceAfter.toFixed(1)} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    return errors;
}

/**
 * Validate Caption (Zone 5)
 * Rules: Single spacing, 12pt (Bold title) + 11pt content, Centered
 */
/**
 * Validate Caption Title
 * Rules: 12pt Bold title, Centered, 6nk spacing
 */
function validateCaptionTitle(para, font, text, index) {
    const errors = [];
    const trimmed = (text || '').trim();

    // Alignment: Centered
    if (para.alignment !== Word.Alignment.centered) {
        errors.push({
            type: 'warning',
            title: 'Başlık/Açıklama: Hizalama',
            description: 'Tablo/Şekil başlıkları ORTALANMALI.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Font: 12pt
    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_CAPTION_TITLE) > 0.5) {
        errors.push({
            type: 'warning',
            title: 'Başlık/Açıklama: Punto',
            description: `Açıklama başlığı 12 punto olmalı. Mevcut: ${font.size} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Spacing: 6nk
    const spaceBefore = para.spaceBefore || 0;
    const spaceAfter = para.spaceAfter || 0;
    if (Math.abs(spaceBefore - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE ||
        Math.abs(spaceAfter - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE) {
        errors.push({
            type: 'warning',
            title: 'Başlık/Açıklama: Boşluk',
            description: `Altında ve üstünde 6nk boşluk olmalı.`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    return errors;
}

function validateCoverPage(para, font, text, index) {
    const errors = [];
    // Rules: 16pt, 0nk spacing
    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_COVER_TITLE) > 0.5) {
        errors.push({
            type: 'warning',
            title: 'Kapak: Punto Hatası',
            description: 'Kapak sayfası başlıkları 16 punto olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }
    if ((para.spaceBefore || 0) > 2 || (para.spaceAfter || 0) > 2) {
        errors.push({
            type: 'warning',
            title: 'Kapak: Boşluk Hatası',
            description: 'Kapak sayfasında paragraflar arası 0nk olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }
    return errors;
}

function validateEpigraph(para, font, text, index) {
    const errors = [];
    // Epigraph: 11pt, Italic, Right Aligned
    if (font.size !== EBYÜ_RULES.FONT_SIZE_EPIGRAPH) {
        errors.push({
            type: 'warning',
            title: 'Epigraf: Punto',
            description: '11 punto olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }
    return errors;
}

/**
 * Validate Ghost Heading (Critical Error)
 */
function validateGhostHeading(para, style, index) {
    return [{
        type: 'error',
        title: 'BOŞ BAŞLIK (Ghost Heading) - KRİTİK!',
        description: `Bu boş satıra "${style}" stili uygulanmış. İçindekiler tablosunda boş satır oluşturur!`,
        severity: 'CRITICAL',
        paraIndex: index
    }];
}

// ============================================
// OFFICE.JS INITIALIZATION
// ============================================

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("EBYÜ Thesis Validator v2.0: Office.js initialized");
        initializeUI();
    } else {
        showError("Bu eklenti sadece Microsoft Word ile çalışır.");
    }
});

function initializeUI() {
    document.getElementById('scanBtn').addEventListener('click', scanDocument);
    document.querySelectorAll('.filter-tab').forEach(tab => {
        tab.addEventListener('click', (e) => setActiveFilter(e.target.dataset.filter));
    });
    logStep('INIT', 'UI initialized with Zone-Based Validation v2.0');
}

// ============================================
// CLEAR HIGHLIGHTS FUNCTION
// ============================================

async function clearHighlights() {
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.font.highlightColor = null;
            await context.sync();
            logStep('CLEAR', 'All highlights cleared');
        });
    } catch (error) {
        console.log("Error clearing highlights:", error.message);
    }
}

// ============================================
// MARGIN CHECK (with Mac error handling)
// Mac compatibility: Do NOT stop the scan if margin check fails
// ============================================

async function checkMargins(context, sections) {
    try {
        // Ensure sections are loaded
        if (!sections || !sections.items || sections.items.length === 0) {
            logStep('MARGIN', 'No sections found, skipping margin check');
            addResult('warning', 'Kenar Boşlukları (Manuel Kontrol)',
                'Bölüm bilgisi yüklenemedi. Lütfen manuel kontrol edin: Sayfa Düzeni → Kenar Boşlukları → Tümü 3 cm olmalı.');
            return; // Do NOT throw, just return
        }

        for (let i = 0; i < sections.items.length; i++) {
            const section = sections.items[i];
            try {
                // Mac may not support getPageSetup - check if function exists
                if (typeof section.getPageSetup !== 'function') {
                    logStep('MARGIN', 'getPageSetup not available on this platform');
                    continue;
                }

                const pageSetup = section.getPageSetup();
                pageSetup.load("topMargin, bottomMargin, leftMargin, rightMargin");

                try {
                    await context.sync();
                } catch (syncError) {
                    // Sync failed (common on Mac), log and continue
                    logStep('MARGIN', `Sync failed for section ${i + 1}: ${syncError.message}`);
                    continue;
                }

                const tolerance = EBYÜ_RULES.MARGIN_TOLERANCE;
                const expected = EBYÜ_RULES.MARGIN_POINTS;
                let hasError = false;
                let errorDetails = [];

                const margins = [
                    { name: 'Üst', value: pageSetup.topMargin },
                    { name: 'Alt', value: pageSetup.bottomMargin },
                    { name: 'Sol', value: pageSetup.leftMargin },
                    { name: 'Sağ', value: pageSetup.rightMargin }
                ];

                for (const margin of margins) {
                    // Check if margin value is valid (not NaN or undefined)
                    if (margin.value !== undefined && !isNaN(margin.value)) {
                        if (Math.abs(margin.value - expected) > tolerance) {
                            hasError = true;
                            errorDetails.push(`${margin.name}: ${(margin.value / 28.35).toFixed(2)} cm`);
                        }
                    }
                }

                if (hasError) {
                    addResult('error', 'Kenar Boşluğu Hatası - KRİTİK',
                        `TÜMÜ 3 cm olmalı. Hatalı kenarlar: ${errorDetails.join(', ')}`,
                        `Bölüm ${i + 1}`, null, undefined, 'CRITICAL');
                } else {
                    addResult('success', 'Kenar Boşlukları',
                        'Tüm kenarlar 3 cm kuralına uygun. ✓');
                }
                return; // Only check first section

            } catch (sectionError) {
                // Individual section error, log and continue to next section
                logStep('MARGIN', `Section ${i + 1} error: ${sectionError.message}`);
                continue;
            }
        }

        // Fallback for Mac or if getPageSetup fails for all sections
        addResult('warning', 'Kenar Boşlukları (Manuel Kontrol)',
            'Otomatik kontrol yapılamadı (Mac uyumluluk). Lütfen manuel kontrol edin: Sayfa Düzeni → Kenar Boşlukları → Tümü 3 cm olmalı.');

    } catch (error) {
        // Catch-all: Log error but do NOT stop the scan
        logStep('MARGIN', `Margin check failed: ${error.message}`);
        addResult('warning', 'Kenar Boşlukları',
            `Kontrol hatası: ${error.message}. Manuel kontrol önerilir.`);
        // Do NOT re-throw - let the scan continue
    }
}

// ============================================
// TABLE VALIDATION (Width, Alignment, Highlighting)
// ============================================

// Global storage for table errors to enable "SHOW" button
let tableErrors = [];

/**
 * Check if tables fit within page margins AND are properly aligned
 * A4 width (595pt) - 2*3cm margins (170pt) ≈ 425pt max width
 * Tables MUST be centered
 */
async function checkTablesSimple(context) {
    tableErrors = []; // Reset

    try {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tables.items.length === 0) {
            addResult('success', 'Tablolar', 'Belgede tablo bulunamadı.');
            return;
        }

        // Load table properties including alignment
        for (let i = 0; i < tables.items.length; i++) {
            tables.items[i].load("rowCount, width, alignment");
        }
        await context.sync();

        let widthErrors = 0;
        let alignmentErrors = 0;
        const maxWidth = EBYÜ_RULES.TABLE_MAX_WIDTH_POINTS;

        for (let i = 0; i < tables.items.length; i++) {
            const table = tables.items[i];
            const tableWidth = table.width || 0;
            const tableAlignment = table.alignment;

            // CHECK 1: Table Alignment (MUST be Centered)
            if (tableAlignment !== Word.Alignment.centered) {
                alignmentErrors++;
                addResult('error', 'Tablo Hizalama Hatası',
                    `Tablo ${i + 1} ORTALANMALI.`,
                    `Tablo ${i + 1}`, null, undefined, 'FORMAT', { isTable: true, tableIndex: i });
            }

            // CHECK 2: Font Size inside Table (11pt)
            // Note: We'll check the first cell as a representative
            const firstCell = table.getCell(0, 0);
            firstCell.load("body/font/size");
            await context.sync();
            const fontSize = firstCell.body.font.size;
            if (fontSize && Math.abs(fontSize - EBYÜ_RULES.FONT_SIZE_TABLE) > 0.5) {
                addResult('warning', 'Tablo İçerik: Punto',
                    `Tablo içi metinler 11 punto olmalı. Mevcut: ${fontSize} pt`,
                    `Tablo ${i + 1}`, null, undefined, 'FORMAT', { isTable: true, tableIndex: i });
            }
        }

        // Summary
        const totalErrors = widthErrors + alignmentErrors;
        if (totalErrors === 0) {
            addResult('success', 'Tablolar',
                `${tables.items.length} tablo kontrol edildi, tümü kurallara uygun. ✓`);
        } else {
            if (widthErrors > 0) {
                addResult('warning', 'Tablo Genişliği Özeti',
                    `${widthErrors}/${tables.items.length} tablo sayfa sınırlarını aşıyor.`);
            }
            if (alignmentErrors > 0) {
                addResult('warning', 'Tablo Hizalama Özeti',
                    `${alignmentErrors}/${tables.items.length} tablo ortalanmamış.`);
            }
        }

    } catch (error) {
        addResult('warning', 'Tablolar',
            `Tablo kontrolü hatası: ${error.message}`);
    }
}

/**
 * Check Footnote formatting
 * Rules: 10pt, Times New Roman, Justified, 0nk spacing
 */
async function checkFootnotes(context) {
    try {
        const footnoteBody = context.document.getFootnoteBody();
        const paras = footnoteBody.paragraphs;
        paras.load("items/text, items/font/size, items/font/name, items/alignment, items/spaceBefore, items/spaceAfter");
        await context.sync();

        if (paras.items.length === 0) {
            logStep('FOOTNOTE', 'No footnotes found');
            return;
        }

        logStep('FOOTNOTE', `Checking ${paras.items.length} footnote paragraphs`);
        let footnoteErrors = 0;

        for (let i = 0; i < paras.items.length; i++) {
            const para = paras.items[i];
            const font = para.font;
            const text = (para.text || '').trim();
            if (text.length === 0) continue;

            const errors = [];
            if (font.name && font.name !== "Times New Roman") {
                errors.push("Yazı tipi Times New Roman olmalı.");
            }
            if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_FOOTNOTE) > 0.5) {
                errors.push(`10 punto olmalı (Mevcut: ${font.size}pt).`);
            }
            if (para.alignment !== Word.Alignment.justified && para.alignment !== Word.Alignment.left) {
                errors.push("İki yana yaslı olmalı.");
            }
            if ((para.spaceBefore || 0) > 2 || (para.spaceAfter || 0) > 2) {
                errors.push("Paragraf arası boşluk 0nk olmalı.");
            }

            if (errors.length > 0) {
                footnoteErrors++;
                if (footnoteErrors <= 5) {
                    addResult('warning', 'Dipnot Format Hatası',
                        errors.join(' '),
                        `Dipnot: "${text.substring(0, 30)}..."`, null, undefined, 'FORMAT');
                }
            }
        }

        if (footnoteErrors > 5) {
            addResult('warning', 'Dipnot Hataları Özeti',
                `Toplam ${footnoteErrors} dipnotta format hatası bulundu. Sadece ilk 5 tanesi gösterildi.`);
        } else if (footnoteErrors === 0) {
            addResult('success', 'Dipnotlar', 'Dipnot formatları kurallara uygun. ✓');
        }

    } catch (error) {
        logStep('FOOTNOTE', `Footnote check skipped or failed: ${error.message}`);
    }
}

/**
 * Check if major sections start on a new page and handle the 7cm (200pt) gap
 */
async function checkPageStarts(paragraphs) {
    try {
        let foundMainHeadings = 0;
        let pageStartErrors = 0;

        for (let i = 0; i < paragraphs.items.length; i++) {
            const para = paragraphs.items[i];
            const text = (para.text || '').trim();

            // We check for sections that MUST start on a new page
            // Giriş, Özet, Abstract, Kaynakça, and all Main Chapters
            const isMain = isMainHeading(text);
            const isAbstract = /^(ÖZET|ABSTRACT)$/i.test(text);
            const isIntro = /^GİRİŞ$/i.test(text);
            const isBiblio = /^(KAYNAKÇA|KAYNAKLAR)$/i.test(text);

            if (isMain || isAbstract || isIntro || isBiblio) {
                foundMainHeadings++;

                // Check if it's the very first thing in the document (OK)
                if (i === 0) continue;

                // Check for PageBreakBefore property
                const hasPageBreakBefore = para.pageBreakBefore;

                // Note: Word API check for manual page breaks in preceding text is limited
                // So we primarily rely on pageBreakBefore or proximity to section starts
                if (!hasPageBreakBefore) {
                    // Check if previous paragraph is empty and has a high space after, or if this has high space before
                    const spaceBefore = para.spaceBefore || 0;

                    // EBYÜ: 7cm top margin (approx 200pt) or 4 empty lines.
                    // If spaceBefore is very low, it's likely not starting a page properly or missing the gap
                    if (spaceBefore < 100) {
                        // Only flag as error if it doesn't look like a page start
                        // This is a heuristic as pure "page start" detection without pagination service is hard
                        // but 7cm gap rule is very specific.
                        addResult('warning', 'Bölüm Başlangıç Hatası',
                            `"${text}" yeni sayfada ve üstten 7 cm (veya 4 boş satır) boşlukla başlamalıdır.`,
                            `Paragraf ${i + 1}`, null, undefined, 'FORMAT');
                        pageStartErrors++;
                    }
                }
            }
        }
    } catch (error) {
        logStep('PAGE_START', `Check failed: ${error.message}`);
    }
}

/**
 * Check if Table captions are ABOVE and Figure captions are BELOW
 */
async function validateCaptionProximity(paragraphs) {
    try {
        for (let i = 0; i < paragraphs.items.length; i++) {
            const para = paragraphs.items[i];
            const text = (para.text || '').trim();
            const captionInfo = isCaption(text);

            if (captionInfo.isCaption) {
                if (captionInfo.type === 'TABLE') {
                    // Table caption: Next paragraph (or one after) should be a table
                    // Note: API cannot easily check "Is there a table object strictly following this para"
                    // but we can flag it as a reminder or check for "missing" proximity.
                } else if (captionInfo.type === 'FIGURE') {
                    // Figure caption: Previous paragraph should be an image/shape
                    if (i === 0) {
                        addResult('error', 'Şekil Başlığı Hatası',
                            'Şekil başlığı şeklin ALTINDA olmalı, ancak bu başlık belgenin başında.',
                            `Paragraf ${i + 1}`, null, undefined, 'FORMAT');
                    }
                }
            }
        }
    } catch (error) {
        logStep('PROXIMITY', `Check failed: ${error.message}`);
    }
}

/**
 * Check if images (InlinePictures and Shapes) are centered and fit within margins
 */
async function checkImages(context) {
    try {
        const body = context.document.body;
        const inlinePictures = body.inlinePictures;
        // const shapes = body.shapes; // Shapes can be complex, focusing on InlinePictures first as per common use

        inlinePictures.load("items");
        await context.sync();

        if (inlinePictures.items.length === 0) {
            logStep('IMAGE', 'No inline pictures found');
            return;
        }

        logStep('IMAGE', `Checking ${inlinePictures.items.length} inline pictures`);

        for (let i = 0; i < inlinePictures.items.length; i++) {
            const pic = inlinePictures.items[i];
            pic.load("width, height");

            // Get parent paragraph to check alignment
            const range = pic.getRange();
            const parentPara = range.paragraphs.getFirst();
            parentPara.load("alignment");

            await context.sync();

            const picWidth = pic.width;
            const alignment = parentPara.alignment;

            // CHECK 1: Alignment (MUST be Centered)
            if (alignment !== Word.Alignment.centered) {
                addResult('error', 'Resim Hizalama Hatası',
                    'Resim/Şekil içeren paragraflar ORTALANMALI.',
                    `Resim ${i + 1}`, null, undefined, 'FORMAT');
            }

            // CHECK 2: Width (Max Content Width)
            if (picWidth > EBYÜ_RULES.CONTENT_MAX_WIDTH_POINTS + 10) { // 10pt buffer
                addResult('error', 'Resim Taşıma Hatası',
                    `Resim sayfa sınırlarını aşıyor. Genişlik: ${(picWidth / 28.35).toFixed(2)} cm`,
                    `Resim ${i + 1}`, null, undefined, 'FORMAT');
            }
        }

    } catch (error) {
        logStep('IMAGE', `Image check error: ${error.message}`);
    }
}

/**
 * Highlight and select a specific table by index
 * This is called from the "SHOW" button for table errors
 * @param {number} tableIndex - Index of the table to highlight
 */
async function highlightTable(tableIndex) {
    try {
        await Word.run(async (context) => {
            const tables = context.document.body.tables;
            tables.load("items");
            await context.sync();

            if (tableIndex >= 0 && tableIndex < tables.items.length) {
                const table = tables.items[tableIndex];

                // Select the entire table - this scrolls to it and highlights it
                table.select();
                await context.sync();

                logStep('TABLE', `Selected and scrolled to table ${tableIndex + 1}`);
            } else {
                console.log(`Invalid table index: ${tableIndex}`);
            }
        });
    } catch (error) {
        console.log("Table highlight error:", error.message);
        // Fallback: try to at least log the error
        logStep('TABLE', `Could not highlight table ${tableIndex + 1}: ${error.message}`);
    }
}

// Expose highlightTable globally for button onclick
window.highlightTable = highlightTable;

// ============================================
// MAIN SCAN FUNCTION - ZONE-BASED VALIDATION
// ============================================

async function scanDocument() {
    const scanBtn = document.getElementById('scanBtn');
    scanBtn.disabled = true;
    scanBtn.innerHTML = '<span>Taranıyor...</span>';

    validationResults = [];
    scanLog = [];
    isInBibliographyZone = false;

    const startTime = performance.now();
    logStep('SCAN', 'Starting Zone-Based Validation v2.0');

    showProgress();
    updateProgress(0, "Belge analiz ediliyor...");

    try {
        // Step 1: Clear previous highlights
        await clearHighlights();

        await Word.run(async (context) => {
            const body = context.document.body;
            const sections = context.document.sections;
            const paragraphs = body.paragraphs;

            // Load sections for margin check
            sections.load("items");
            paragraphs.load("items");
            await context.sync();

            const totalParagraphs = paragraphs.items.length;
            logStep('LOAD', `Found ${totalParagraphs} paragraphs`);

            // BATCH LOAD: Load ALL paragraph properties at once
            for (let i = 0; i < totalParagraphs; i++) {
                const para = paragraphs.items[i];
                para.load("text, style, alignment, firstLineIndent, leftIndent, rightIndent, lineSpacing, lineUnitBefore, lineUnitAfter, spaceAfter, spaceBefore, outlineLevel");
                para.font.load("name, size, bold, italic, allCaps, highlightColor");
                try {
                    para.load("listItemOrNull");
                } catch (e) {
                }
            }
            await context.sync();

            updateProgress(15, "Kenar boşlukları kontrol ediliyor...");
            await checkMargins(context, sections);

            updateProgress(20, "Tablolar kontrol ediliyor...");
            await checkTablesSimple(context);

            updateProgress(23, "Dipnotlar kontrol ediliyor...");
            await checkFootnotes(context);

            updateProgress(26, "Resimler kontrol ediliyor...");
            await checkImages(context);

            updateProgress(28, "Bölüm başlangıçları kontrol ediliyor...");
            await checkPageStarts(paragraphs);

            updateProgress(29, "Başlık konumları kontrol ediliyor...");
            await validateCaptionProximity(paragraphs);

            updateProgress(30, "Paragraflar zone-based analiz ediliyor...");

            // =============================================
            // ZONE-BASED STATE MACHINE LOOP
            // Includes Abstract zones with word count tracking
            // =============================================
            let currentZone = ZONES.COVER;
            let abstractWordCountTR = 0;
            let abstractWordCountEN = 0;

            let stats = {
                zones: { cover: 0, frontMatter: 0, abstractTR: 0, abstractEN: 0, body: 0, backMatter: 0 },
                types: {
                    mainHeading: 0,
                    subHeading: 0,
                    bodyText: 0,
                    blockQuote: 0,
                    bibliography: 0,
                    captionTitle: 0,
                    epigraph: 0,
                    ghost: 0,
                    toc: 0,
                    list: 0,
                    cover: 0,
                    empty: 0
                },
                errors: { critical: 0, format: 0 }
            };

            let allErrors = [];
            let errorCounts = {};

            for (let i = 0; i < totalParagraphs; i++) {
                const para = paragraphs.items[i];
                const text = para.text || '';
                const trimmed = text.trim();
                const style = para.style || '';
                const font = para.font;

                // =============================================
                // ZONE SWITCHING LOGIC (with Abstract tracking)
                // =============================================

                const previousZone = currentZone;

                // Cover to Front Matter switch: If we find a heading that is NOT cover but FRONT_MATTER
                if (currentZone === ZONES.COVER && (isMainHeading(trimmed) || isHeadingStyle(style))) {
                    if (!isCoverItem(trimmed)) {
                        currentZone = ZONES.FRONT_MATTER;
                        logStep('ZONE', `→ FRONT_MATTER at paragraph ${i + 1}`);
                    }
                }

                // Check for Turkish Abstract (ÖZET)
                if ((currentZone === ZONES.FRONT_MATTER || currentZone === ZONES.COVER) && /^ÖZET$/i.test(trimmed)) {
                    currentZone = ZONES.ABSTRACT_TR;
                    abstractWordCountTR = 0;
                    logStep('ZONE', `→ ABSTRACT_TR at paragraph ${i + 1}: "${trimmed}"`);
                }

                // Check for English Abstract (ABSTRACT)
                if ((currentZone === ZONES.FRONT_MATTER || currentZone === ZONES.ABSTRACT_TR) && /^ABSTRACT$/i.test(trimmed)) {
                    currentZone = ZONES.ABSTRACT_EN;
                    abstractWordCountEN = 0;
                    logStep('ZONE', `→ ABSTRACT_EN at paragraph ${i + 1}: "${trimmed}"`);
                }

                // Check for switch to BODY
                if (currentZone !== ZONES.BODY && currentZone !== ZONES.BACK_MATTER) {
                    if (matchesAnyPattern(trimmed, PATTERNS.BODY_START)) {
                        currentZone = ZONES.BODY;
                        isInBibliographyZone = false;
                        logStep('ZONE', `→ BODY at paragraph ${i + 1}`);
                    }
                }

                // Check for switch to BACK_MATTER
                if (currentZone === ZONES.BODY || currentZone === ZONES.FRONT_MATTER) {
                    if (trimmed.match(PATTERNS.BACK_MATTER_START) && (font.bold === true || isHeadingStyle(style))) {
                        currentZone = ZONES.BACK_MATTER;
                        isInBibliographyZone = true;
                        logStep('ZONE', `→ BACK_MATTER at paragraph ${i + 1}`);
                    }
                }

                // =============================================
                // ABSTRACT WORD COUNT ACCUMULATION
                // Stop counting at "Anahtar Kelimeler" (TR) or "Keywords" (EN)
                // =============================================

                // Pattern to detect Turkish keywords section
                const isAnahtarKelimeler = /^anahtar\s*(kelimeler|sözcükler)\s*[:.]?/i.test(trimmed);
                // Pattern to detect English keywords section  
                const isKeywords = /^key\s*words?\s*[:.]?/i.test(trimmed);

                // Turkish Abstract word counting
                if (currentZone === ZONES.ABSTRACT_TR && trimmed.length > 0) {
                    // Skip the "ÖZET" heading itself
                    if (/^ÖZET$/i.test(trimmed)) {
                        // Don't count the title
                    }
                    // Check if we hit "Anahtar Kelimeler" - validate and stop counting
                    else if (isAnahtarKelimeler) {
                        logStep('ABSTRACT', `TR Abstract ended at "Anahtar Kelimeler" with ${abstractWordCountTR} words`);
                        // Validate word count
                        if (abstractWordCountTR < 200 || abstractWordCountTR > 250) {
                            addResult('warning', 'Türkçe Özet: Kelime Sayısı Uyarısı',
                                `ÖZET bölümü 200-250 kelime olmalı. Mevcut: ${abstractWordCountTR} kelime`,
                                'ÖZET Bölümü', null, undefined, 'FORMAT');
                        } else {
                            addResult('success', 'Türkçe Özet: Kelime Sayısı',
                                `ÖZET bölümü ${abstractWordCountTR} kelime - kurala uygun (200-250). ✓`);
                        }
                        // Mark as validated so we don't validate again on zone switch
                        abstractWordCountTR = -1; // Use -1 as a flag for "already validated"
                    }
                    // Normal paragraph - count words
                    else if (abstractWordCountTR >= 0) {
                        const words = trimmed.split(/\s+/).filter(w => w.length > 0).length;
                        abstractWordCountTR += words;
                    }
                }

                // English Abstract word counting
                if (currentZone === ZONES.ABSTRACT_EN && trimmed.length > 0) {
                    // Skip the "ABSTRACT" heading itself
                    if (/^ABSTRACT$/i.test(trimmed)) {
                        // Don't count the title
                    }
                    // Check if we hit "Keywords" - validate and stop counting
                    else if (isKeywords) {
                        logStep('ABSTRACT', `EN Abstract ended at "Keywords" with ${abstractWordCountEN} words`);
                        // Validate word count
                        if (abstractWordCountEN < 200 || abstractWordCountEN > 250) {
                            addResult('warning', 'İngilizce Abstract: Kelime Sayısı Uyarısı',
                                `ABSTRACT bölümü 200-250 kelime olmalı. Mevcut: ${abstractWordCountEN} kelime`,
                                'ABSTRACT Bölümü', null, undefined, 'FORMAT');
                        } else {
                            addResult('success', 'İngilizce Abstract: Kelime Sayısı',
                                `ABSTRACT bölümü ${abstractWordCountEN} kelime - kurala uygun (200-250). ✓`);
                        }
                        // Mark as validated
                        abstractWordCountEN = -1;
                    }
                    // Normal paragraph - count words
                    else if (abstractWordCountEN >= 0) {
                        const words = trimmed.split(/\s+/).filter(w => w.length > 0).length;
                        abstractWordCountEN += words;
                    }
                }

                // Update zone stats
                switch (currentZone) {
                    case ZONES.COVER: stats.zones.cover++; break;
                    case ZONES.FRONT_MATTER: stats.zones.frontMatter++; break;
                    case ZONES.ABSTRACT_TR: stats.zones.abstractTR++; break;
                    case ZONES.ABSTRACT_EN: stats.zones.abstractEN++; break;
                    case ZONES.BODY: stats.zones.body++; break;
                    case ZONES.BACK_MATTER: stats.zones.backMatter++; break;
                }

                // =============================================
                // DETECT PARAGRAPH TYPE
                // =============================================
                const paraType = detectParagraphType(para, text, style, font, currentZone, isInBibliographyZone);

                // Update type stats
                switch (paraType) {
                    case PARA_TYPES.MAIN_HEADING: stats.types.mainHeading++; break;
                    case PARA_TYPES.SUB_HEADING: stats.types.subHeading++; break;
                    case PARA_TYPES.BODY_TEXT: stats.types.bodyText++; break;
                    case PARA_TYPES.BLOCK_QUOTE: stats.types.blockQuote++; break;
                    case PARA_TYPES.BIBLIOGRAPHY: stats.types.bibliography++; break;
                    case PARA_TYPES.CAPTION_TITLE: stats.types.captionTitle++; break;
                    case PARA_TYPES.EPIGRAPH: stats.types.epigraph++; break;
                    case PARA_TYPES.COVER_TEXT: stats.types.cover++; break;
                    case PARA_TYPES.GHOST_HEADING: stats.types.ghost++; break;
                    case PARA_TYPES.TOC_ENTRY: stats.types.toc++; break;
                    case PARA_TYPES.LIST_ITEM: stats.types.list++; break;
                    case PARA_TYPES.EMPTY: stats.types.empty++; break;
                }

                // =============================================
                // SKIP CONDITIONS / SPECIAL HANDLERS
                // =============================================

                // TOC entries: User requested 12pt check
                if (paraType === PARA_TYPES.TOC_ENTRY) {
                    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_BODY) > 0.5) {
                        allErrors.push({
                            type: 'warning',
                            title: 'İçindekiler: Punto',
                            description: 'İçindekiler tablosu girdileri 12 punto olmalı.',
                            severity: 'FORMAT',
                            paraIndex: i,
                            location: `İçindekiler: "${trimmed.substring(0, 30)}..."`
                        });
                    }
                    continue;
                }

                // Skip empty paragraphs
                if (paraType === PARA_TYPES.EMPTY) continue;

                // Skip abstract zone (handled by word count above)
                if (currentZone === ZONES.ABSTRACT_TR || currentZone === ZONES.ABSTRACT_EN) {
                    continue;
                }

                // =============================================
                // APPLY ZONE-SPECIFIC VALIDATION
                // =============================================

                let paragraphErrors = [];

                switch (paraType) {
                    case PARA_TYPES.GHOST_HEADING:
                        paragraphErrors = validateGhostHeading(para, style, i);
                        para.font.highlightColor = HIGHLIGHT_COLORS.CRITICAL;
                        stats.errors.critical++;
                        break;

                    case PARA_TYPES.MAIN_HEADING:
                        paragraphErrors = validateMainHeading(para, font, text, i);
                        break;

                    case PARA_TYPES.SUB_HEADING:
                        paragraphErrors = validateSubHeading(para, font, text, i);
                        break;

                    case PARA_TYPES.BODY_TEXT:
                        paragraphErrors = validateBodyText(para, font, text, i);
                        break;

                    case PARA_TYPES.BLOCK_QUOTE:
                        paragraphErrors = validateBlockQuote(para, font, text, i);
                        break;

                    case PARA_TYPES.BIBLIOGRAPHY:
                        paragraphErrors = validateBibliography(para, font, text, i);
                        break;

                    case PARA_TYPES.CAPTION_TITLE:
                        const captionInfo = isCaption(text);
                        if (!captionInfo.isCorrect) {
                            paragraphErrors.push({
                                type: 'error',
                                title: 'Tablo/Şekil Başlığı: Format Hatası',
                                description: 'Başlıklar bölüm numarası içermeli (Örn: Tablo 1.1:) ve iki nokta (:) ile bitmelidir.',
                                severity: 'FORMAT',
                                paraIndex: i
                            });
                        }
                        paragraphErrors = [...paragraphErrors, ...validateCaptionTitle(para, font, text, i)];
                        break;

                    case PARA_TYPES.COVER_TEXT:
                        paragraphErrors = validateCoverPage(para, font, text, i);
                        break;

                    case PARA_TYPES.EPIGRAPH:
                        paragraphErrors = validateEpigraph(para, font, text, i);
                        break;

                    case PARA_TYPES.LIST_ITEM:
                        // List items: minimal validation (font only)
                        if (font.name && font.name !== EBYÜ_RULES.FONT_NAME) {
                            paragraphErrors.push({
                                type: 'warning',
                                title: 'Liste Öğesi: Yazı Tipi',
                                description: `Times New Roman olmalı. Mevcut: ${font.name}`,
                                severity: 'FORMAT',
                                paraIndex: i
                            });
                        }
                        break;
                }

                // Apply highlights for errors
                if (paragraphErrors.length > 0) {
                    const hasCritical = paragraphErrors.some(e => e.severity === 'CRITICAL');
                    para.font.highlightColor = hasCritical ? HIGHLIGHT_COLORS.CRITICAL : HIGHLIGHT_COLORS.FORMAT;

                    for (const err of paragraphErrors) {
                        // Count errors by title to limit repetitive errors
                        const errorKey = err.title;
                        errorCounts[errorKey] = (errorCounts[errorKey] || 0) + 1;

                        // Only add first 5 of each error type to results
                        if (errorCounts[errorKey] <= 5) {
                            allErrors.push({
                                ...err,
                                location: `Paragraf ${i + 1}: "${trimmed.substring(0, 35)}..."`
                            });
                        }

                        if (err.severity === 'CRITICAL') {
                            stats.errors.critical++;
                        } else {
                            stats.errors.format++;
                        }
                    }
                }

                // Update progress periodically
                if (i % 50 === 0) {
                    updateProgress(25 + Math.floor((i / totalParagraphs) * 50),
                        `Paragraf ${i + 1}/${totalParagraphs} analiz ediliyor...`);
                }
            }

            await context.sync();

            // =============================================
            // ADD RESULTS TO UI
            // =============================================

            updateProgress(80, "Sonuçlar derleniyor...");

            // Add all collected errors
            for (const err of allErrors) {
                addResult(err.type, err.title, err.description, err.location, null, err.paraIndex);
            }

            // Add summary for errors that exceeded limit
            for (const [errorTitle, count] of Object.entries(errorCounts)) {
                if (count > 5) {
                    addResult('warning', `${errorTitle} - Özet`,
                        `Bu hata türünden toplam ${count} adet bulundu. Sadece ilk 5 tanesi gösterildi.`);
                }
            }

            // Validate any remaining abstract word counts (if document ended in abstract)
            if (currentZone === ZONES.ABSTRACT_TR && abstractWordCountTR > 0) {
                if (abstractWordCountTR < 200 || abstractWordCountTR > 250) {
                    addResult('warning', 'Türkçe Özet: Kelime Sayısı Uyarısı',
                        `ÖZET bölümü 200-250 kelime olmalı. Mevcut: ${abstractWordCountTR} kelime`,
                        'ÖZET Bölümü', null, undefined, 'FORMAT');
                } else {
                    addResult('success', 'Türkçe Özet: Kelime Sayısı',
                        `ÖZET bölümü ${abstractWordCountTR} kelime - kurala uygun (200-250). ✓`);
                }
            }
            if (currentZone === ZONES.ABSTRACT_EN && abstractWordCountEN > 0) {
                if (abstractWordCountEN < 200 || abstractWordCountEN > 250) {
                    addResult('warning', 'İngilizce Abstract: Kelime Sayısı Uyarısı',
                        `ABSTRACT bölümü 200-250 kelime olmalı. Mevcut: ${abstractWordCountEN} kelime`,
                        'ABSTRACT Bölümü', null, undefined, 'FORMAT');
                } else {
                    addResult('success', 'İngilizce Abstract: Kelime Sayısı',
                        `ABSTRACT bölümü ${abstractWordCountEN} kelime - kurala uygun (200-250). ✓`);
                }
            }

            // Add zone analysis summary
            addResult('success', 'Bölge Analizi Tamamlandı',
                `📊 Ön Kısım: ${stats.zones.frontMatter} | Özet (TR): ${stats.zones.abstractTR} | Abstract (EN): ${stats.zones.abstractEN} | Ana Metin: ${stats.zones.body} | Kaynakça/Ekler: ${stats.zones.backMatter} paragraf`);

            // Add type breakdown
            const typeBreakdown = [];
            if (stats.types.mainHeading > 0) typeBreakdown.push(`Ana Başlık: ${stats.types.mainHeading}`);
            if (stats.types.subHeading > 0) typeBreakdown.push(`Alt Başlık: ${stats.types.subHeading}`);
            if (stats.types.bodyText > 0) typeBreakdown.push(`Gövde: ${stats.types.bodyText}`);
            if (stats.types.blockQuote > 0) typeBreakdown.push(`Blok Alıntı: ${stats.types.blockQuote}`);
            if (stats.types.caption > 0) typeBreakdown.push(`Başlık/Açıklama: ${stats.types.caption}`);
            if (stats.types.bibliography > 0) typeBreakdown.push(`Kaynakça: ${stats.types.bibliography}`);

            if (typeBreakdown.length > 0) {
                addResult('success', 'Paragraf Türleri', typeBreakdown.join(' | '));
            }

            // Ghost heading critical warning
            if (stats.types.ghost > 0) {
                addResult('error', `⚠️ ${stats.types.ghost} Boş Başlık Bulundu!`,
                    'Bu boş başlıklar İçindekiler tablosunda hatalı satırlara neden olur. Kırmızı ile işaretlendi.');
            }

            // Final summary
            const totalErrors = stats.errors.critical + stats.errors.format;
            if (totalErrors === 0) {
                addResult('success', '✅ Tebrikler!',
                    'Belge EBYÜ 2022 Tez Yazım Kılavuzu formatına uygun görünüyor.');
            } else {
                addResult(stats.errors.critical > 0 ? 'error' : 'warning',
                    'Tarama Özeti',
                    `🔴 Kritik: ${stats.errors.critical} | 🟡 Format: ${stats.errors.format} hata bulundu.`);
            }

            // Add manual check reminders
            addManualCheckReminders();

            updateProgress(100, "Tarama tamamlandı!");

            const endTime = performance.now();
            logStep('COMPLETE', `Scan completed in ${((endTime - startTime) / 1000).toFixed(2)} seconds`);
        });

        setTimeout(() => {
            hideProgress();
            displayResults();
        }, 500);

    } catch (error) {
        console.error("Scan error:", error);
        logStep('ERROR', `Scan failed: ${error.message}`);
        hideProgress();
        addResult('error', 'Tarama Hatası', `Hata: ${error.message}`);
        displayResults();
    } finally {
        scanBtn.disabled = false;
        scanBtn.innerHTML = `
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/>
            </svg>
            <span>DÖKÜMAN TARA</span>
        `;
    }
}

// ============================================
// MANUAL CHECK REMINDERS
// ============================================

// Manual checks for items that cannot be fully automated reliably
function addManualCheckReminders() {
    addResult('warning', '📋 Sayfa Numarası (Manuel)',
        'Ön kısım: Roma (i, ii, iii) | Ana metin: Arap (1, 2, 3) | Sağ alt köşede, 10pt Times New Roman.');
    addResult('warning', '📋 Tablo/Şekil Konumu (Manuel)',
        'Tablo başlığı: tablonun ÜSTünde | Şekil başlığı: şeklin ALTında olmalı.');
    addResult('warning', '📋 Atıf Sistemi (Manuel)',
        'Seçilen sisteme göre (APA 7 veya Chicago 17) metin içi veya dipnot atıfları tutarlı olmalı.');
}

// ============================================
// FIX PARAGRAPH STUB (Future Implementation)
// ============================================

/**
 * Apply correct formatting based on paragraph type
 * @param {number} index - Paragraph index
 * @param {string} targetType - Target paragraph type (e.g., 'BODY_TEXT', 'MAIN_HEADING')
 */
async function fixParagraph(index, targetType = null) {
    try {
        await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");
            await context.sync();

            if (index < 0 || index >= paragraphs.items.length) {
                console.log("Invalid paragraph index");
                return;
            }

            const para = paragraphs.items[index];
            para.load("text, style");
            para.font.load("name, size, bold");
            await context.sync();

            const text = para.text || '';
            const trimmed = text.trim();
            const font = para.font;
            const style = para.style || '';

            // Auto-detect type if not specified
            const paraType = targetType || detectParagraphType(para, text, style, font, isInBibliographyZone);

            switch (paraType) {
                case PARA_TYPES.GHOST_HEADING:
                    para.delete();
                    break;

                case PARA_TYPES.BODY_TEXT:
                    // Apply body text formatting
                    para.font.name = EBYÜ_RULES.FONT_NAME;
                    para.font.size = EBYÜ_RULES.FONT_SIZE_BODY;
                    para.alignment = Word.Alignment.justified;
                    para.firstLineIndent = EBYÜ_RULES.FIRST_LINE_INDENT_POINTS;
                    para.lineSpacing = EBYÜ_RULES.LINE_SPACING_18PT;
                    para.spaceBefore = EBYÜ_RULES.SPACING_BODY;      // 6nk
                    para.spaceAfter = EBYÜ_RULES.SPACING_BODY;       // 6nk
                    break;

                case PARA_TYPES.MAIN_HEADING:
                    para.font.name = EBYÜ_RULES.FONT_NAME;
                    para.font.size = EBYÜ_RULES.FONT_SIZE_HEADING_MAIN;
                    para.font.bold = true;
                    para.alignment = Word.Alignment.centered;
                    para.firstLineIndent = 0;
                    para.spaceBefore = EBYÜ_RULES.SPACING_HEADING;   // 6nk
                    para.spaceAfter = EBYÜ_RULES.SPACING_HEADING;    // 6nk
                    break;

                case PARA_TYPES.SUB_HEADING:
                    para.font.name = EBYÜ_RULES.FONT_NAME;
                    para.font.size = EBYÜ_RULES.FONT_SIZE_HEADING_SUB;
                    para.font.bold = true;
                    para.alignment = Word.Alignment.left;
                    para.firstLineIndent = 0;
                    para.spaceBefore = EBYÜ_RULES.SPACING_HEADING;   // 6nk
                    para.spaceAfter = EBYÜ_RULES.SPACING_HEADING;    // 6nk
                    break;

                case PARA_TYPES.BLOCK_QUOTE:
                    para.font.name = EBYÜ_RULES.FONT_NAME;
                    para.font.size = EBYÜ_RULES.FONT_SIZE_BLOCK_QUOTE;
                    para.leftIndent = EBYÜ_RULES.BLOCK_QUOTE_INDENT_POINTS;
                    para.rightIndent = EBYÜ_RULES.BLOCK_QUOTE_INDENT_POINTS;
                    para.lineSpacing = EBYÜ_RULES.LINE_SPACING_12PT;
                    para.spaceBefore = 0;   // Block quotes typically have 0 spacing
                    para.spaceAfter = 0;    // The indent creates visual separation
                    break;

                case PARA_TYPES.BIBLIOGRAPHY:
                    para.font.name = EBYÜ_RULES.FONT_NAME;
                    para.font.size = EBYÜ_RULES.FONT_SIZE_BODY;
                    para.alignment = Word.Alignment.justified;
                    para.firstLineIndent = 0;
                    para.leftIndent = EBYÜ_RULES.FIRST_LINE_INDENT_POINTS; // Hanging indent
                    para.lineSpacing = EBYÜ_RULES.LINE_SPACING_12PT;
                    para.spaceBefore = 0;                             // No space before
                    para.spaceAfter = EBYÜ_RULES.SPACING_BIBLIO;     // 3nk after each entry
                    break;

                case PARA_TYPES.CAPTION:
                    para.font.name = EBYÜ_RULES.FONT_NAME;
                    para.font.size = EBYÜ_RULES.FONT_SIZE_CAPTION;
                    para.alignment = Word.Alignment.centered;
                    para.lineSpacing = EBYÜ_RULES.LINE_SPACING_12PT;
                    para.spaceBefore = EBYÜ_RULES.SPACING_HEADING;   // 6nk
                    para.spaceAfter = EBYÜ_RULES.SPACING_HEADING;    // 6nk
                    break;

                default:
                    console.log(`No fix rule defined for type: ${paraType}`);
            }

            // Clear highlight after fixing
            para.font.highlightColor = null;

            await context.sync();
            showNotification('success', `Paragraf ${index + 1} düzeltildi!`);
        });
    } catch (error) {
        console.error("Fix error:", error);
        showNotification('error', `Düzeltme hatası: ${error.message}`);
    }
}

// ============================================
// UI HELPER FUNCTIONS
// ============================================

function addResult(type, title, description, location = null, fixData = null, paraIndex = undefined, severity = null, customData = null) {
    validationResults.push({ type, title, description, location, fixData, paraIndex, severity, customData });
}

function displayResults() {
    const errorCount = validationResults.filter(r => r.type === 'error').length;
    const warningCount = validationResults.filter(r => r.type === 'warning').length;
    const successCount = validationResults.filter(r => r.type === 'success').length;

    document.getElementById('errorCount').textContent = errorCount;
    document.getElementById('warningCount').textContent = warningCount;
    document.getElementById('successCount').textContent = successCount;

    document.getElementById('summarySection').classList.remove('hidden');
    document.getElementById('filterTabs').classList.remove('hidden');

    if (errorCount > 0 || warningCount > 0) {
        document.getElementById('quickFixSection').classList.remove('hidden');
    } else {
        document.getElementById('quickFixSection').classList.add('hidden');
    }

    renderFilteredResults();
}

function renderFilteredResults() {
    const resultsList = document.getElementById('resultsList');
    let filtered = validationResults;
    if (currentFilter !== 'all') {
        filtered = validationResults.filter(r => r.type === currentFilter);
    }

    if (filtered.length === 0) {
        resultsList.innerHTML = '<div class="empty-state"><p>Bu kategoride sonuç yok</p></div>';
        return;
    }

    resultsList.innerHTML = filtered.map(r => createResultItemHTML(r)).join('');
}

function createResultItemHTML(result) {
    const icon = getIconSVG(result.type);

    // Determine the correct "SHOW" button based on whether it's a table or paragraph
    let showBtn = '';
    if (result.customData && result.customData.isTable) {
        // Table-specific show button uses highlightTable()
        showBtn = `<button class="show-button" onclick="highlightTable(${result.customData.tableIndex})">📍 GÖSTER</button>`;
    } else if (result.paraIndex !== undefined) {
        // Paragraph-specific show button
        showBtn = `<button class="show-button" onclick="highlightParagraph(${result.paraIndex})">📍 GÖSTER</button>`;
    }

    // Fix button only for paragraph errors (not tables)
    const fixBtn = result.paraIndex !== undefined && (result.type === 'error' || result.type === 'warning') ?
        `<button class="fix-button" onclick="fixParagraph(${result.paraIndex})" title="Otomatik düzelt">🔧</button>` : '';

    // Add severity badge for critical errors
    const severityBadge = result.severity === 'CRITICAL' ?
        '<span class="severity-badge critical">KRİTİK</span>' : '';

    return `
        <div class="result-item ${result.type}">
            <div class="result-header">
                <div class="result-icon">${icon}</div>
                <div class="result-content">
                    <div class="result-title">${severityBadge}${result.title}</div>
                    <div class="result-description">${result.description}</div>
                    ${result.location ? `<div class="result-location">📍 ${result.location}</div>` : ''}
                </div>
            </div>
            ${(showBtn || fixBtn) ? `<div class="result-actions">${showBtn}${fixBtn}</div>` : ''}
        </div>
    `;
}

function getIconSVG(type) {
    switch (type) {
        case 'error': return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>';
        case 'warning': return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 9v4M12 17h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/></svg>';
        case 'success': return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>';
        default: return '';
    }
}

async function highlightParagraph(index) {
    try {
        await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");
            await context.sync();

            if (index >= 0 && index < paragraphs.items.length) {
                // Use cyan for "show" to differentiate from error highlights
                paragraphs.items[index].font.highlightColor = HIGHLIGHT_COLORS.FOUND;
                paragraphs.items[index].select();
                await context.sync();
            }
        });
    } catch (error) {
        console.log("Highlight error:", error);
    }
}

function setActiveFilter(filter) {
    currentFilter = filter;
    document.querySelectorAll('.filter-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.filter === filter);
    });
    renderFilteredResults();
}

function showProgress() {
    document.getElementById('progressSection').classList.remove('hidden');
}

function hideProgress() {
    document.getElementById('progressSection').classList.add('hidden');
}

function updateProgress(percent, text) {
    document.getElementById('progressFill').style.width = `${percent}%`;
    document.getElementById('progressText').textContent = text;
}

function showNotification(type, message) {
    const notification = document.createElement('div');
    notification.className = `result-item ${type}`;
    notification.style.cssText = 'position:fixed;top:20px;left:50%;transform:translateX(-50%);z-index:1000;max-width:300px;box-shadow:0 4px 12px rgba(0,0,0,0.15);';
    notification.innerHTML = `<div class="result-header"><div class="result-icon">${getIconSVG(type)}</div><div class="result-content"><div class="result-description">${message}</div></div></div>`;
    document.body.appendChild(notification);
    setTimeout(() => notification.remove(), 3000);
}

function showError(message) {
    document.getElementById('resultsList').innerHTML = `
        <div class="result-item error">
            <div class="result-header">
                <div class="result-icon">${getIconSVG('error')}</div>
                <div class="result-content">
                    <div class="result-title">Hata</div>
                    <div class="result-description">${message}</div>
                </div>
            </div>
        </div>
    `;
}

// ============================================
// EXPORT FOR DEBUGGING (optional)
// ============================================

// Expose some functions globally for debugging
window.debugEBYU = {
    getScanLog: () => scanLog,
    getResults: () => validationResults,
    fixParagraph: fixParagraph,
    EBYÜ_RULES: EBYÜ_RULES,
    PARA_TYPES: PARA_TYPES
};
