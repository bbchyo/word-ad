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
    FONT_SIZE_CAPTION: 11,

    // Line Spacing
    LINE_SPACING_BODY: 1.5,      // 1.5 lines ≈ 18pt for 12pt font
    LINE_SPACING_SINGLE: 1.0,   // Single ≈ 12pt for 12pt font
    LINE_SPACING_18PT: 18,
    LINE_SPACING_12PT: 12,

    // Indentation
    FIRST_LINE_INDENT_CM: 1.25,
    FIRST_LINE_INDENT_POINTS: 35.4,  // 1.25cm = 35.4pt
    BLOCK_QUOTE_INDENT_CM: 1.25,
    BLOCK_QUOTE_INDENT_POINTS: 35.4,
    INDENT_TOLERANCE: 3,

    // Spacing Around (NK = point, 6nk = 6pt)
    SPACING_HEADING: 6,           // 6nk before AND after headings
    SPACING_BODY: 6,              // 6nk before AND after body paragraphs
    SPACING_BIBLIO: 3,            // 3nk for bibliography entries
    SPACING_TOLERANCE: 2,         // Allow ±2pt tolerance (5-8pt is OK for 6pt)

    // Table Validation
    TABLE_MAX_WIDTH_POINTS: 425,  // A4 width (595pt) - 2*3cm margins (170pt) ≈ 425pt

    // Detection Thresholds
    BLOCK_QUOTE_MIN_INDENT_POINTS: 28, // ~1cm minimum to detect block quote
    MIN_BODY_TEXT_LENGTH: 30,           // Minimum chars to consider as body text
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
    CAPTION: 'CAPTION',
    LIST_ITEM: 'LIST_ITEM',
    GHOST_HEADING: 'GHOST_HEADING',
    FRONT_MATTER: 'FRONT_MATTER',
    EMPTY: 'EMPTY',
    UNKNOWN: 'UNKNOWN'
};

const ZONES = {
    FRONT_MATTER: 'FRONT_MATTER',
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
        /^\d+\.\d+\.?\s+\S/,     // 1.1 or 1.1.
        /^\d+\.\d+\.\d+\.?\s+\S/, // 1.1.1 or 1.1.1.
        /^\d+\.\s+[A-ZÇĞİÖŞÜ]/   // 1. Başlık (numbered, starts with capital)
    ],

    // Captions
    CAPTION_TABLE: /^Tablo\s*\d+[\.:]/i,
    CAPTION_FIGURE: /^(Şekil|Grafik|Resim|Harita)\s*\d+[\.:]/i,

    // Zone switching patterns
    BODY_START: [
        /^GİRİŞ$/i,
        /^1\.\s*GİRİŞ/i,
        /^BİRİNCİ\s*BÖLÜM/i,
        /^BÖLÜM\s*[1I]/i
    ],
    BACK_MATTER_START: [
        /^KAYNAKÇA$/i,
        /^KAYNAKLAR$/i,
        /^REFERANSLAR$/i,
        /^REFERENCES$/i,
        /^BIBLIOGRAPHY$/i
    ],
    APPENDIX_START: [
        /^EKLER?$/i,
        /^EK\s*\d/i,
        /^APPENDIX/i
    ],

    // TOC patterns
    TOC_STYLE: [/^TOC/i, /İçindekiler/i],
    TOC_CONTENT: /\.{3,}/   // Lines with multiple dots
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
function isTOCEntry(style, text) {
    if (!style && !text) return false;

    // Style-based detection
    if (style) {
        if (PATTERNS.TOC_STYLE.some(p => p.test(style))) return true;
    }

    // Content-based detection (multiple dots)
    if (text && PATTERNS.TOC_CONTENT.test(text)) return true;

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
 * Check if text is a caption (Tablo X or Şekil Y)
 */
function isCaption(text) {
    if (!text) return false;
    const trimmed = text.trim();
    return PATTERNS.CAPTION_TABLE.test(trimmed) || PATTERNS.CAPTION_FIGURE.test(trimmed);
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
function detectParagraphType(para, text, style, font, isInBiblio) {
    const trimmed = (text || '').trim();

    // Empty paragraph
    if (trimmed.length === 0) {
        // Check for Ghost Heading (empty with heading style)
        if (isHeadingStyle(style)) {
            return PARA_TYPES.GHOST_HEADING;
        }
        return PARA_TYPES.EMPTY;
    }

    // TOC Entry (skip validation)
    if (isTOCEntry(style, trimmed)) {
        return PARA_TYPES.TOC_ENTRY;
    }

    // Caption (Tablo X, Şekil Y)
    if (isCaption(trimmed)) {
        return PARA_TYPES.CAPTION;
    }

    // Bibliography entry (after KAYNAKÇA heading)
    if (isInBiblio && trimmed.length > 10) {
        // Skip the KAYNAKÇA heading itself
        if (!matchesAnyPattern(trimmed, PATTERNS.BACK_MATTER_START)) {
            return PARA_TYPES.BIBLIOGRAPHY;
        }
    }

    // Main Heading (GİRİŞ, BÖLÜM, KAYNAKÇA, etc.)
    if (isMainHeading(trimmed)) {
        return PARA_TYPES.MAIN_HEADING;
    }

    // Sub-Heading (1.1 Alt Başlık) - must be bold
    if (isSubHeading(trimmed, font.bold === true)) {
        return PARA_TYPES.SUB_HEADING;
    }

    // Heading detected via style
    if (isHeadingStyle(style)) {
        // Determine if main or sub based on style level
        if (style.toLowerCase().includes("1") || /heading\s*1/i.test(style) || /başlık\s*1/i.test(style)) {
            return PARA_TYPES.MAIN_HEADING;
        }
        return PARA_TYPES.SUB_HEADING;
    }

    // Block Quote (indented paragraph, not a list)
    if (isBlockQuote(para)) {
        return PARA_TYPES.BLOCK_QUOTE;
    }

    // List Item
    if (isListItem(para)) {
        return PARA_TYPES.LIST_ITEM;
    }

    // Body Text (default for substantial paragraphs)
    if (trimmed.length >= EBYÜ_RULES.MIN_BODY_TEXT_LENGTH) {
        return PARA_TYPES.BODY_TEXT;
    }

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

    // Must be Centered
    if (para.alignment !== Word.Alignment.centered) {
        errors.push({
            type: 'error',
            title: 'Ana Başlık: Hizalama Hatası',
            description: 'Ana başlıklar ORTALANMALI (Centered).',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Must be 14pt
    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_HEADING_MAIN) > 0.5) {
        errors.push({
            type: 'error',
            title: 'Ana Başlık: Punto Hatası',
            description: `14 punto (pt) olmalı. Mevcut: ${font.size} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Must be Bold
    if (font.bold !== true) {
        errors.push({
            type: 'error',
            title: 'Ana Başlık: Kalın Yazı Hatası',
            description: 'Ana başlıklar KALIN (Bold) olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Must be ALL CAPS
    if (!isAllCaps(trimmed)) {
        errors.push({
            type: 'warning',
            title: 'Ana Başlık: Büyük Harf Uyarısı',
            description: 'Ana başlıklar TAMAMI BÜYÜK HARF olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Check spacing (6nk before/after)
    // Word API returns spacing in points: spaceBefore and spaceAfter
    const spaceBefore = para.spaceBefore || 0;
    const spaceAfter = para.spaceAfter || 0;
    const expectedSpacing = EBYÜ_RULES.SPACING_HEADING;
    const tolerance = EBYÜ_RULES.SPACING_TOLERANCE;

    // Check spaceBefore (should be ~6pt, allow 4-8pt)
    if (spaceBefore < (expectedSpacing - tolerance) || spaceBefore > (expectedSpacing + tolerance + 2)) {
        if (spaceBefore === 0 || spaceBefore > 12) {
            errors.push({
                type: 'warning',
                title: 'Ana Başlık: Öncesi Boşluk (NK)',
                description: `Başlık öncesi 6nk (6pt) olmalı. Mevcut: ${spaceBefore.toFixed(1)} pt`,
                severity: 'FORMAT',
                paraIndex: index
            });
        }
    }

    // Check spaceAfter (should be ~6pt, allow 4-8pt)
    if (spaceAfter < (expectedSpacing - tolerance) || spaceAfter > (expectedSpacing + tolerance + 2)) {
        if (spaceAfter === 0 || spaceAfter > 12) {
            errors.push({
                type: 'warning',
                title: 'Ana Başlık: Sonrası Boşluk (NK)',
                description: `Başlık sonrası 6nk (6pt) olmalı. Mevcut: ${spaceAfter.toFixed(1)} pt`,
                severity: 'FORMAT',
                paraIndex: index
            });
        }
    }

    return errors;
}

/**
 * Validate Sub-Heading (Zone 1B)
 * Rules: Left Aligned (or Justified), 12pt, Bold, 6nk before/after
 */
function validateSubHeading(para, font, text, index) {
    const errors = [];

    // Must be Left Aligned or Justified
    if (para.alignment !== Word.Alignment.left &&
        para.alignment !== Word.Alignment.justified) {
        errors.push({
            type: 'error',
            title: 'Alt Başlık: Hizalama Hatası',
            description: 'Alt başlıklar SOLA YASLI veya İKİ YANA YASLI olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Must be 12pt
    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_HEADING_SUB) > 0.5) {
        errors.push({
            type: 'error',
            title: 'Alt Başlık: Punto Hatası',
            description: `12 punto (pt) olmalı. Mevcut: ${font.size} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Must be Bold
    if (font.bold !== true) {
        errors.push({
            type: 'error',
            title: 'Alt Başlık: Kalın Yazı Hatası',
            description: 'Alt başlıklar KALIN (Bold) olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Check spacing (6nk before/after)
    const spaceBefore = para.spaceBefore || 0;
    const spaceAfter = para.spaceAfter || 0;
    const expectedSpacing = EBYÜ_RULES.SPACING_HEADING;
    const tolerance = EBYÜ_RULES.SPACING_TOLERANCE;

    // Check spaceBefore (should be ~6pt, allow 4-8pt)
    if (spaceBefore < (expectedSpacing - tolerance) || spaceBefore > (expectedSpacing + tolerance + 2)) {
        if (spaceBefore === 0 || spaceBefore > 12) {
            errors.push({
                type: 'warning',
                title: 'Alt Başlık: Öncesi Boşluk (NK)',
                description: `Başlık öncesi 6nk (6pt) olmalı. Mevcut: ${spaceBefore.toFixed(1)} pt`,
                severity: 'FORMAT',
                paraIndex: index
            });
        }
    }

    // Check spaceAfter (should be ~6pt, allow 4-8pt)
    if (spaceAfter < (expectedSpacing - tolerance) || spaceAfter > (expectedSpacing + tolerance + 2)) {
        if (spaceAfter === 0 || spaceAfter > 12) {
            errors.push({
                type: 'warning',
                title: 'Alt Başlık: Sonrası Boşluk (NK)',
                description: `Başlık sonrası 6nk (6pt) olmalı. Mevcut: ${spaceAfter.toFixed(1)} pt`,
                severity: 'FORMAT',
                paraIndex: index
            });
        }
    }

    return errors;
}

/**
 * Validate Body Text (Zone 2)
 * Rules: Times New Roman, 12pt, Justified, First Line Indent 1.25cm, 1.5 line spacing
 */
function validateBodyText(para, font, text, index) {
    const errors = [];

    // Must be Times New Roman
    if (font.name && font.name !== EBYÜ_RULES.FONT_NAME && font.name !== "Times New Roman") {
        errors.push({
            type: 'error',
            title: 'Gövde Metin: Yazı Tipi Hatası',
            description: `"${font.name}" yerine Times New Roman kullanılmalı.`,
            severity: 'CRITICAL',
            paraIndex: index
        });
    }

    // Must be 12pt
    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_BODY) > 0.5) {
        errors.push({
            type: 'warning',
            title: 'Gövde Metin: Punto Hatası',
            description: `12 punto (pt) olmalı. Mevcut: ${font.size} pt`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Must be Justified
    if (para.alignment !== Word.Alignment.justified) {
        errors.push({
            type: 'error',
            title: 'Gövde Metin: Hizalama Hatası',
            description: 'Paragraf İKİ YANA YASLI (Justified) olmalı.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Must have First Line Indent of 1.25cm (~35.4pt)
    const firstIndent = para.firstLineIndent || 0;
    // Only check if paragraph is long enough (not a short intro line)
    if ((text || '').trim().length > 50) {
        const expectedIndent = EBYÜ_RULES.FIRST_LINE_INDENT_POINTS;
        const diff = Math.abs(firstIndent - expectedIndent);

        // If indent is 0 or significantly different, flag it
        if (firstIndent < 5 || diff > EBYÜ_RULES.INDENT_TOLERANCE) {
            errors.push({
                type: 'warning',
                title: 'Gövde Metin: Girinti Hatası',
                description: `İlk satır girintisi 1.25 cm olmalı. Mevcut: ${(firstIndent / 28.35).toFixed(2)} cm`,
                severity: 'FORMAT',
                paraIndex: index
            });
        }
    }

    // Check line spacing (should be 1.5 lines ≈ 18pt for 12pt font)
    const lineSpacing = para.lineSpacing;
    if (lineSpacing && lineSpacing > 0) {
        // 1.5 line spacing for 12pt font ≈ 18pt
        // Allow range 17-20pt as valid
        const isValid15 = (lineSpacing >= 16 && lineSpacing <= 21);
        if (!isValid15 && lineSpacing < 16) {
            errors.push({
                type: 'warning',
                title: 'Gövde Metin: Satır Aralığı',
                description: `1.5 satır aralığı olmalı (~18pt). Mevcut: ${lineSpacing.toFixed(1)} pt`,
                severity: 'FORMAT',
                paraIndex: index
            });
        }
    }

    // Check paragraph spacing (6nk before/after)
    const spaceBefore = para.spaceBefore || 0;
    const spaceAfter = para.spaceAfter || 0;
    const expectedSpacing = EBYÜ_RULES.SPACING_BODY;
    const tolerance = EBYÜ_RULES.SPACING_TOLERANCE;

    // Only flag if spacing is 0 or excessively large (>12pt)
    // This avoids too many false positives for paragraphs without explicit spacing
    if (spaceBefore === 0 && (text || '').trim().length > 100) {
        // Only warn for substantial paragraphs without any spacing
        // Skip this for now to reduce noise - can be enabled for strict mode
    }

    if (spaceAfter > 12) {
        errors.push({
            type: 'warning',
            title: 'Gövde Metin: Sonrası Boşluk (NK)',
            description: `Paragraf sonrası boşluk çok fazla. 6nk (6pt) olmalı. Mevcut: ${spaceAfter.toFixed(1)} pt`,
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
    const expectedIndent = EBYÜ_RULES.BLOCK_QUOTE_INDENT_POINTS;
    const diff = Math.abs(leftIndent - expectedIndent);

    if (diff > EBYÜ_RULES.INDENT_TOLERANCE * 2) {
        errors.push({
            type: 'warning',
            title: 'Blok Alıntı: Girinti Uyarısı',
            description: `Sol girinti 1.25 cm olmalı. Mevcut: ${(leftIndent / 28.35).toFixed(2)} cm`,
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
function validateCaption(para, font, text, index) {
    const errors = [];
    const trimmed = (text || '').trim();

    // Must be Centered
    if (para.alignment !== Word.Alignment.centered) {
        errors.push({
            type: 'warning',
            title: 'Başlık/Açıklama: Hizalama',
            description: 'Tablo/Şekil başlıkları ORTALANMALI.',
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Check single line spacing
    const lineSpacing = para.lineSpacing;
    if (lineSpacing && lineSpacing > 14) {
        errors.push({
            type: 'warning',
            title: 'Başlık/Açıklama: Satır Aralığı',
            description: `Tek satır aralığı (1.0) olmalı.`,
            severity: 'FORMAT',
            paraIndex: index
        });
    }

    // Font check: 11pt or 12pt is acceptable for captions
    if (font.size && font.size < 10.5) {
        errors.push({
            type: 'warning',
            title: 'Başlık/Açıklama: Punto',
            description: `En az 11 punto olmalı. Mevcut: ${font.size} pt`,
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
// ============================================

async function checkMargins(context, sections) {
    try {
        for (let i = 0; i < sections.items.length; i++) {
            const section = sections.items[i];
            try {
                // Mac may not support getPageSetup
                if (typeof section.getPageSetup === 'function') {
                    const pageSetup = section.getPageSetup();
                    pageSetup.load("topMargin, bottomMargin, leftMargin, rightMargin");
                    await context.sync();

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
                        if (Math.abs(margin.value - expected) > tolerance) {
                            hasError = true;
                            errorDetails.push(`${margin.name}: ${(margin.value / 28.35).toFixed(2)} cm`);
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
                }
            } catch (e) {
                // getPageSetup not available (Mac compatibility)
                logStep('MARGIN', `getPageSetup not available: ${e.message}`);
            }
        }

        // Fallback for Mac or if getPageSetup fails
        addResult('warning', 'Kenar Boşlukları (Manuel Kontrol)',
            'Otomatik kontrol yapılamadı (Mac). Lütfen manuel kontrol edin: Sayfa Düzeni → Kenar Boşlukları → Tümü 3 cm olmalı.');

    } catch (error) {
        addResult('warning', 'Kenar Boşlukları',
            `Kontrol hatası: ${error.message}. Manuel kontrol önerilir.`);
    }
}

// ============================================
// TABLE WIDTH CHECK
// ============================================

/**
 * Check if tables fit within page margins
 * A4 width (595pt) - 2*3cm margins (170pt) ≈ 425pt max width
 */
async function checkTablesSimple(context) {
    try {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tables.items.length === 0) {
            addResult('success', 'Tablolar', 'Belgede tablo bulunamadı.');
            return;
        }

        // Load table properties
        for (let i = 0; i < tables.items.length; i++) {
            tables.items[i].load("rowCount, width");
        }
        await context.sync();

        let widthErrors = 0;
        const maxWidth = EBYÜ_RULES.TABLE_MAX_WIDTH_POINTS;

        for (let i = 0; i < tables.items.length; i++) {
            const table = tables.items[i];
            const tableWidth = table.width || 0;

            if (tableWidth > maxWidth + 10) { // Allow 10pt tolerance
                widthErrors++;

                // Try to highlight the table content
                try {
                    const firstCell = table.getCell(0, 0);
                    firstCell.body.font.highlightColor = HIGHLIGHT_COLORS.FORMAT;
                } catch (e) {
                    // Cell access may fail, just log error
                    logStep('TABLE', `Could not highlight table ${i + 1}: ${e.message}`);
                }

                addResult('error', 'Tablo Genişliği Hatası',
                    `Tablo ${i + 1} sayfa sınırlarını aşıyor. Genişlik: ${(tableWidth / 28.35).toFixed(2)} cm (Max: ${(maxWidth / 28.35).toFixed(2)} cm)`,
                    `Tablo ${i + 1} (${table.rowCount} satır)`, null, undefined, 'FORMAT');
            }
        }

        if (widthErrors === 0) {
            addResult('success', 'Tablolar',
                `${tables.items.length} tablo kontrol edildi, tümü sayfa sınırları içinde. ✓`);
        } else {
            addResult('warning', 'Tablo Genişliği Özeti',
                `${widthErrors}/${tables.items.length} tablo sayfa sınırlarını aşıyor.`);
        }

    } catch (error) {
        addResult('warning', 'Tablolar',
            `Tablo kontrolü hatası: ${error.message}`);
    }
}

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
                para.load("text, style, alignment, firstLineIndent, leftIndent, rightIndent, lineSpacing, lineUnitBefore, lineUnitAfter, spaceAfter, spaceBefore");
                para.font.load("name, size, bold, italic, allCaps, highlightColor");
                try {
                    para.load("listItemOrNull");
                } catch (e) {
                    // listItemOrNull not available on all platforms
                }
            }
            await context.sync();

            updateProgress(15, "Kenar boşlukları kontrol ediliyor...");
            await checkMargins(context, sections);

            updateProgress(20, "Tablolar kontrol ediliyor...");
            await checkTablesSimple(context);

            updateProgress(30, "Paragraflar zone-based analiz ediliyor...");

            // =============================================
            // ZONE-BASED STATE MACHINE LOOP
            // =============================================
            let currentZone = ZONES.FRONT_MATTER;
            let stats = {
                zones: { frontMatter: 0, body: 0, backMatter: 0 },
                types: {
                    mainHeading: 0,
                    subHeading: 0,
                    bodyText: 0,
                    blockQuote: 0,
                    bibliography: 0,
                    caption: 0,
                    ghost: 0,
                    toc: 0,
                    list: 0,
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
                // ZONE SWITCHING LOGIC
                // =============================================

                // Check for switch to BODY zone
                if (currentZone === ZONES.FRONT_MATTER) {
                    if (matchesAnyPattern(trimmed, PATTERNS.BODY_START) &&
                        (font.bold === true || (font.size && font.size >= 13))) {
                        currentZone = ZONES.BODY;
                        isInBibliographyZone = false;
                        logStep('ZONE', `→ BODY at paragraph ${i + 1}: "${trimmed.substring(0, 40)}"`);
                    }
                }

                // Check for switch to BACK_MATTER zone (Bibliography)
                if (currentZone === ZONES.BODY || currentZone === ZONES.FRONT_MATTER) {
                    if (matchesAnyPattern(trimmed, PATTERNS.BACK_MATTER_START) &&
                        (font.bold === true || isHeadingStyle(style))) {
                        currentZone = ZONES.BACK_MATTER;
                        isInBibliographyZone = true;
                        logStep('ZONE', `→ BACK_MATTER (Bibliography) at paragraph ${i + 1}: "${trimmed}"`);
                    }
                }

                // Check for Appendix (still back matter but different rules)
                if (matchesAnyPattern(trimmed, PATTERNS.APPENDIX_START)) {
                    isInBibliographyZone = false; // Appendix entries don't follow biblio rules
                    logStep('ZONE', `→ APPENDIX at paragraph ${i + 1}`);
                }

                // Update zone stats
                switch (currentZone) {
                    case ZONES.FRONT_MATTER: stats.zones.frontMatter++; break;
                    case ZONES.BODY: stats.zones.body++; break;
                    case ZONES.BACK_MATTER: stats.zones.backMatter++; break;
                }

                // =============================================
                // DETECT PARAGRAPH TYPE
                // =============================================
                const paraType = detectParagraphType(para, text, style, font, isInBibliographyZone);

                // Update type stats
                switch (paraType) {
                    case PARA_TYPES.MAIN_HEADING: stats.types.mainHeading++; break;
                    case PARA_TYPES.SUB_HEADING: stats.types.subHeading++; break;
                    case PARA_TYPES.BODY_TEXT: stats.types.bodyText++; break;
                    case PARA_TYPES.BLOCK_QUOTE: stats.types.blockQuote++; break;
                    case PARA_TYPES.BIBLIOGRAPHY: stats.types.bibliography++; break;
                    case PARA_TYPES.CAPTION: stats.types.caption++; break;
                    case PARA_TYPES.GHOST_HEADING: stats.types.ghost++; break;
                    case PARA_TYPES.TOC_ENTRY: stats.types.toc++; break;
                    case PARA_TYPES.LIST_ITEM: stats.types.list++; break;
                    case PARA_TYPES.EMPTY: stats.types.empty++; break;
                }

                // =============================================
                // SKIP CONDITIONS
                // =============================================

                // Skip TOC entries entirely
                if (paraType === PARA_TYPES.TOC_ENTRY) continue;

                // Skip empty paragraphs (unless ghost heading - handled below)
                if (paraType === PARA_TYPES.EMPTY) continue;

                // Skip front matter paragraphs (except ghost headings)
                if (currentZone === ZONES.FRONT_MATTER && paraType !== PARA_TYPES.GHOST_HEADING) {
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

                    case PARA_TYPES.CAPTION:
                        paragraphErrors = validateCaption(para, font, text, i);
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

            // Add zone analysis summary
            addResult('success', 'Bölge Analizi Tamamlandı',
                `📊 Ön Kısım: ${stats.zones.frontMatter} | Ana Metin: ${stats.zones.body} | Kaynakça/Ekler: ${stats.zones.backMatter} paragraf`);

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

function addManualCheckReminders() {
    addResult('warning', '📋 Dipnot Kontrolü (Manuel)',
        '10pt, Times New Roman, tek satır aralığı, iki yana yaslı, 0nk boşluk olmalı.');
    addResult('warning', '📋 Sayfa Numarası (Manuel)',
        'Ön kısım: Roma (i, ii, iii) | Ana metin: Arap (1, 2, 3) | Sağ alt köşede');
    addResult('warning', '📋 Tablo/Şekil Konumu (Manuel)',
        'Tablo başlığı: tablonun ÜSTünde | Şekil başlığı: şeklin ALTında');
    addResult('warning', '📋 Başlık Sayfası (Manuel)',
        'Üst kenar boşluğu 7 cm, başlık ile alt metin arasında 4 satır boşluk');
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

function addResult(type, title, description, location = null, fixData = null, paraIndex = undefined, severity = null) {
    validationResults.push({ type, title, description, location, fixData, paraIndex, severity });
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
    const showBtn = result.paraIndex !== undefined ?
        `<button class="show-button" onclick="highlightParagraph(${result.paraIndex})">📍 GÖSTER</button>` : '';
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
