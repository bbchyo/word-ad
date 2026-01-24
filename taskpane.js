/**

* ============================================

* EBYÜ Thesis Format Validator - Task Pane Logic

* Erzincan Binali Yıldırım University

* Based on: EBYÜ 2022 Tez Yazım Kılavuzu

* ============================================

*

* OPTIMIZED VERSION with:

* - Batch Loading (no sync inside loops)

* - Section-Based Margin Validation

* - Robust Ghost Heading Detection via outlineLevel

* - Document Architecture approach

*

* ERROR SEVERITY:

* - CRITICAL (Red): Margins, Ghost Headings, Wrong Font Family

* - FORMAT (Yellow): Wrong Indent, Spacing, Size

*/



// ============================================

// CONSTANTS - EBYÜ 2022 Strict Rules

// ============================================



const EBYÜ_RULES = {

    // Page Layout

    MARGIN_CM: 3,

    MARGIN_POINTS: 85.05, // 3cm

    MARGIN_TOP_SPECIAL_CM: 7,

    MARGIN_TOP_SPECIAL_POINTS: 198.45, // 7cm for Main Chapter Starts

    MARGIN_TOLERANCE: 2.5, // Tolerance for floating point diffs



    // Fonts

    FONT_NAME: "Times New Roman",

    FONT_SIZE_BODY: 12,

    FONT_SIZE_HEADING_MAIN: 14,

    FONT_SIZE_HEADING_SUB: 12,

    FONT_SIZE_BLOCK_QUOTE: 11,

    FONT_SIZE_FOOTNOTE: 10,

    FONT_SIZE_TABLE: 11,

    FONT_SIZE_CAPTION_TITLE: 12,

    FONT_SIZE_CAPTION_CONTENT: 11,

    FONT_SIZE_COVER_TITLE: 16,

    FONT_SIZE_EPIGRAPH: 11,



    // Spacing

    FIRST_LINE_INDENT_CM: 1.25,

    FIRST_LINE_INDENT_POINTS: 35.4, // 1.25cm

    BLOCK_QUOTE_INDENT_POINTS: 35.4,

    BIBLIOGRAPHY_HANGING_INDENT_POINTS: 28.35, // 1cm

    INDENT_TOLERANCE: 2.5,



    // Paragraph Spacing (nk = points)

    SPACING_6NK: 6,

    SPACING_3NK: 3,

    SPACING_0NK: 0,

    SPACING_TOLERANCE: 1.5,



    // Line Spacing

    LINE_SPACING_1_5_MIN: 15,

    LINE_SPACING_1_5_MAX: 22,

    LINE_SPACING_SINGLE_MIN: 10,

    LINE_SPACING_SINGLE_MAX: 14,



    // Detection

    MIN_BODY_TEXT_LENGTH: 30,

    BLOCK_QUOTE_MIN_INDENT: 20,



    // Page Dimensions (A4)

    PAGE_WIDTH_POINTS: 595.3,

    PAGE_HEIGHT_POINTS: 841.9,



    // Page Number Rules (Sayfa Numaralandırma)

    PAGE_NUMBER_FOOTER_DISTANCE_POINTS: 35.4, // 1.25 cm

    PAGE_NUMBER_SIZE: 10,



    // Abstract Rules (Özet Sayfası)

    ABSTRACT_MIN_WORDS: 200,

    ABSTRACT_MAX_WORDS: 250,

    ABSTRACT_MIN_KEYWORDS: 3,

    ABSTRACT_MAX_KEYWORDS: 5,



    // Thesis Length (Tez Uzunluğu)

    MIN_PAGES_MASTERS: 50,

    MIN_PAGES_PHD: 80,

    MAX_PAGES_TOTAL: 500,



    // Table Content Font Size

    TABLE_CONTENT_SIZE: 11

};



// Highlight Colors

const HIGHLIGHT_COLORS = {

    CRITICAL: "Red",

    FORMAT: "Yellow",

    FOUND: "Cyan"

};



// Paragraph Types

const PARA_TYPES = {

    TOC_ENTRY: 'TOC_ENTRY',

    MAIN_HEADING: 'MAIN_HEADING',

    SUB_HEADING: 'SUB_HEADING',

    BODY_TEXT: 'BODY_TEXT',

    BLOCK_QUOTE: 'BLOCK_QUOTE',

    BIBLIOGRAPHY: 'BIBLIOGRAPHY',

    CAPTION_TITLE: 'CAPTION_TITLE',

    EPIGRAPH: 'EPIGRAPH',

    LIST_ITEM: 'LIST_ITEM',

    GHOST_HEADING: 'GHOST_HEADING',

    FRONT_MATTER: 'FRONT_MATTER',

    COVER_TEXT: 'COVER_TEXT',

    EMPTY: 'EMPTY',

    UNKNOWN: 'UNKNOWN'

};



// Document Zones

const ZONES = {

    COVER: 'COVER',

    FRONT_MATTER: 'FRONT_MATTER',

    TABLE_OF_CONTENTS: 'TABLE_OF_CONTENTS',

    ABSTRACT_TR: 'ABSTRACT_TR',

    ABSTRACT_EN: 'ABSTRACT_EN',

    BODY: 'BODY',

    BACK_MATTER: 'BACK_MATTER'

};



// ============================================

// DETECTION PATTERNS

// ============================================



const PATTERNS = {

    // Main chapter headings (7cm top margin triggers)

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

        /^\d+\.\d+(\.\d+)*\.?\s+[A-ZÇĞİÖŞÜa-zçğıöşü]/

    ],



    // Captions

    CAPTION_TABLE: /^Tablo\s*(\d+)\.(\d+)\s*[:.]/i,

    CAPTION_FIGURE: /^(Şekil|Grafik|Resim|Harita)\s*(\d+)\.(\d+)\s*[:.]/i,



    // TOC patterns

    TOC_STYLE: [/^TOC/i, /^İçindekiler/i, /^Table of Contents/i],

    TOC_CONTENT: /\.{5,}\s*(i|v|x|\d)+$/i,

    TOC_START: /^İÇİNDEKİLER$/i,



    // Zone switching

    BODY_START: [/^GİRİŞ$/i],

    BACK_MATTER_START: [/^(KAYNAKÇA|KAYNAKLAR|REFERANSLAR|REFERENCES)$/i],



    // Cover page patterns

    COVER_IDENTIFIERS: [

        /^T\.?C\.?$/i,

        /^ERZİNCAN\s*BİNALİ\s*YILDIRIM/i,

        /^ÜNİVERSİTESİ$/i,

        /^(FEN|SOSYAL)\s*BİLİMLERİ\s*ENSTİTÜSÜ$/i,

        /^(YÜKSEK\s*LİSANS|DOKTORA)\s*TEZİ$/i,

        /^DANIŞMAN/i,

        /^Tez\s*Danışmanı/i

    ],



    // Front matter (Roma rakamları - ÖN KISIM)

    FRONT_MATTER_IDENTIFIERS: [

        /^İÇİNDEKİLER$/i,

        /^ÖN\s*SÖZ$/i,

        /^ÖNSÖZ$/i,

        /^TEŞEKKÜR$/i,

        /^KISALTMALAR/i,

        /^SİMGELER/i,

        /^TABLOLAR\s*(LİSTESİ|DİZİNİ)?$/i,

        /^ŞEKİLLER\s*(LİSTESİ|DİZİNİ)?$/i,

        /^ÖZET$/i,

        /^ABSTRACT$/i

    ],



    // Abstract patterns

    ABSTRACT_TR: /^ÖZET$/i,

    ABSTRACT_EN: /^ABSTRACT$/i,

    KEYWORDS_TR: /^Anahtar\s*Kelimeler\s*:/i,

    KEYWORDS_EN: /^Keywords\s*:/i

};



// ============================================

// GLOBAL STATE

// ============================================



let validationResults = [];

let scanLog = [];

let isScanning = false;



// ============================================

// LOGGING UTILITY

// ============================================



function logStep(category, message, details = null) {

    const timestamp = new Date().toISOString();

    scanLog.push({ timestamp, category, message, details });

    console.log(`[${category}] ${message}`, details || '');

}



// ============================================

// HELPER FUNCTIONS - Pattern Matching

// ============================================



function matchesAnyPattern(text, patterns) {

    if (!text || !patterns) return false;

    return patterns.some(p => p.test(text.trim()));

}



function isMainHeadingText(text) {

    return matchesAnyPattern(text, PATTERNS.MAIN_HEADING);

}



function isSubHeadingText(text) {

    return matchesAnyPattern(text, PATTERNS.SUB_HEADING);

}



function isTOCEntry(style, text) {

    const styleLower = (style || '').toLowerCase();

    if (styleLower.includes('toc') || styleLower.includes('içindekiler')) return true;

    if (PATTERNS.TOC_CONTENT.test(text)) return true;

    return false;

}



function isCaption(text) {

    const trimmed = (text || '').trim();

    const tableMatch = trimmed.match(PATTERNS.CAPTION_TABLE);

    const figureMatch = trimmed.match(PATTERNS.CAPTION_FIGURE);

    return {

        isCaption: !!(tableMatch || figureMatch),

        type: tableMatch ? 'table' : (figureMatch ? 'figure' : null),

        isCorrect: !!(tableMatch || figureMatch)

    };

}



function isCoverItem(text) {

    return matchesAnyPattern(text, PATTERNS.COVER_IDENTIFIERS);

}



/**

* Check if alignment is centered - handles Word.Alignment enum variations

* Word.Alignment.centered can be: "Centered", "centered", 2, or the enum value

*/

function isCentered(alignment) {

    if (alignment === undefined || alignment === null) return true; // Skip check if undefined



    // String comparison (case-insensitive)

    if (typeof alignment === 'string') {

        return alignment.toLowerCase() === 'centered';

    }



    // Numeric comparison (Word.Alignment.centered = 2)

    if (typeof alignment === 'number') {

        return alignment === 2;

    }



    // Direct enum comparison

    if (alignment === Word.Alignment.centered) {

        return true;

    }



    return false;

}



function isHeadingStyle(style) {

    if (!style) return false;

    const s = style.toLowerCase();

    return s.includes('heading') || s.includes('başlık') || s.includes('title');

}



// ============================================

// PARAGRAPH TYPE DETECTION (Priority-Based)

// ============================================



/**

* Detect paragraph type using outlineLevel as primary indicator

* @param {Object} paraData - Pre-loaded paragraph data

* @param {string} zone - Current document zone

* @param {boolean} isInBiblio - Currently in bibliography section

* @returns {string} - PARA_TYPES value

*/

function detectParagraphType(paraData, zone, isInBiblio) {

    const { text, style, outlineLevel, tableNestingLevel, leftIndent, font } = paraData;

    const trimmed = (text || '').trim();



    // Priority 1: Inside table = table content

    if (tableNestingLevel > 0) {

        const captionInfo = isCaption(trimmed);

        return captionInfo.isCaption ? PARA_TYPES.CAPTION_TITLE : PARA_TYPES.BODY_TEXT;

    }



    // Priority 2: Empty paragraph - check for Ghost Heading

    if (trimmed.length === 0) {

        // Ghost Heading: empty text but has outline level (not body text level)

        // In Word, outlineLevel is a number 0-8 for headings, or "BodyText" for normal text

        if (outlineLevel !== undefined && outlineLevel !== null) {

            // If outlineLevel is a number (0-8), it's a heading level

            if (typeof outlineLevel === 'number' || (typeof outlineLevel === 'string' && !isNaN(parseInt(outlineLevel)))) {

                const level = typeof outlineLevel === 'number' ? outlineLevel : parseInt(outlineLevel);

                if (level >= 0 && level <= 8) {

                    return PARA_TYPES.GHOST_HEADING;

                }

            }

        }

        // Also check heading style

        if (isHeadingStyle(style)) {

            return PARA_TYPES.GHOST_HEADING;

        }

        return PARA_TYPES.EMPTY;

    }



    // Priority 3: TOC entries

    if (isTOCEntry(style, trimmed)) {

        return PARA_TYPES.TOC_ENTRY;

    }



    // Priority 4: Cover page

    if (zone === ZONES.COVER || isCoverItem(trimmed)) {

        return PARA_TYPES.COVER_TEXT;

    }



    // Priority 5: Captions

    const captionInfo = isCaption(trimmed);

    if (captionInfo.isCaption) {

        return PARA_TYPES.CAPTION_TITLE;

    }



    // Priority 6: Bibliography zone

    if (isInBiblio && trimmed.length > 5) {

        if (!matchesAnyPattern(trimmed, PATTERNS.BACK_MATTER_START)) {

            return PARA_TYPES.BIBLIOGRAPHY;

        }

    }



    // Priority 7: Main Heading (outlineLevel first, then text pattern)

    const isMainByOutline = outlineLevel === 0 || outlineLevel === 1;

    const isMainByText = isMainHeadingText(trimmed);

    const isMainByStyle = isHeadingStyle(style) && (/heading\s*1/i.test(style) || /başlık\s*1/i.test(style));



    if (isMainByOutline || isMainByText || isMainByStyle) {

        return PARA_TYPES.MAIN_HEADING;

    }



    // Priority 8: Sub-Heading

    const isSubByOutline = typeof outlineLevel === 'number' && outlineLevel >= 2 && outlineLevel <= 8;

    const isSubByText = isSubHeadingText(trimmed);

    const isSubByStyle = isHeadingStyle(style) && !isMainByStyle;



    if (isSubByOutline || isSubByText || isSubByStyle) {

        return PARA_TYPES.SUB_HEADING;

    }



    // Priority 9: Block quote (significant left indent)

    if (leftIndent && leftIndent >= EBYÜ_RULES.BLOCK_QUOTE_MIN_INDENT) {

        return PARA_TYPES.BLOCK_QUOTE;

    }



    // Priority 10: Body text

    if (trimmed.length >= 20) {

        return PARA_TYPES.BODY_TEXT;

    }



    return PARA_TYPES.UNKNOWN;

}



// ============================================

// RESULT MANAGEMENT

// ============================================



function addResult(type, title, description, location = null, paraIndex = null, severity = null) {

    validationResults.push({

        type,

        title,

        description,

        location: location || 'Belge Geneli',

        paraIndex,

        severity: severity || (type === 'error' ? 'CRITICAL' : 'FORMAT'),

        timestamp: new Date().toISOString()

    });

}



function clearResults() {

    validationResults = [];

    scanLog = [];

}



// ============================================

// UI FUNCTIONS

// ============================================



function initializeUI() {

    // Match index.html: id="scanBtn"

    const scanBtn = document.getElementById('scanBtn');

    if (scanBtn) {

        scanBtn.onclick = scanDocument;

    } else {

        console.error('Scan button (scanBtn) not found!');

    }

    logStep('UI', 'User interface initialized');

}



function setButtonState(enabled) {

    const btn = document.getElementById('scanBtn');

    if (btn) {

        btn.disabled = !enabled;

        // Button has SVG + span structure in index.html

        const textSpan = btn.querySelector('span');

        if (textSpan) {

            textSpan.textContent = enabled ? 'DÖKÜMAN TARA' : 'ARANIYOR...';

        }

    }

}



function updateProgress(percent, message) {

    // Match index.html: progressSection, progressFill, progressText

    const progressContainer = document.getElementById('progressSection');

    const progressBar = document.getElementById('progressFill');

    const progressText = document.getElementById('progressText');



    if (progressContainer) {

        progressContainer.classList.remove('hidden');

    }

    if (progressBar) progressBar.style.width = `${percent}%`;

    if (progressText) progressText.textContent = message;

}



function hideProgress() {

    const progressContainer = document.getElementById('progressSection');

    if (progressContainer) {

        progressContainer.classList.add('hidden');

    }

}



function displayResults() {

    // Match index.html: resultsList, summarySection, errorCount, warningCount, successCount

    const resultsContainer = document.getElementById('resultsList');

    const summarySection = document.getElementById('summarySection');

    const errorCountEl = document.getElementById('errorCount');

    const warningCountEl = document.getElementById('warningCount');

    const successCountEl = document.getElementById('successCount');



    if (!resultsContainer) return;



    if (validationResults.length === 0) {

        resultsContainer.innerHTML = '<div class="empty-state"><p>✅ Hiçbir hata bulunamadı.</p></div>';

        if (summarySection) summarySection.classList.add('hidden');

        return;

    }



    // Show summary section

    if (summarySection) summarySection.classList.remove('hidden');



    let html = '';

    const errors = validationResults.filter(r => r.type === 'error');

    const warnings = validationResults.filter(r => r.type === 'warning');

    const successes = validationResults.filter(r => r.type === 'success');



    // Update stat cards

    if (errorCountEl) errorCountEl.textContent = errors.length;

    if (warningCountEl) warningCountEl.textContent = warnings.length;

    if (successCountEl) successCountEl.textContent = successes.length;



    // Errors first

    for (const result of errors) {

        html += createResultItem(result, 'error');

    }

    // Then warnings

    for (const result of warnings) {

        html += createResultItem(result, 'warning');

    }

    // Then successes

    for (const result of successes) {

        html += createResultItem(result, 'success');

    }



    resultsContainer.innerHTML = html;

}



function createResultItem(result, type) {

    // Extract paragraph number from location for navigation

    const paraMatch = result.location ? result.location.match(/Paragraf\s*(\d+)/i) : null;

    const paraIndex = paraMatch ? parseInt(paraMatch[1]) - 1 : null;



    const showButton = paraIndex !== null

        ? `<button class="show-error-btn" onclick="goToError(${paraIndex})">GÖSTER</button>`

        : '';



    return `

<div class="result-item ${type}">

<div class="result-header">

<span class="result-title">${result.title}</span>

${showButton}

</div>

<div class="result-description">${result.description}</div>

<div class="result-location">${result.location || ''}</div>

</div>

`;

}



// ============================================

// NAVIGATE TO ERROR (Hataya Git)

// ============================================



async function goToError(paragraphIndex) {

    try {

        await Word.run(async (context) => {

            const paragraphs = context.document.body.paragraphs;

            paragraphs.load('items');

            await context.sync();



            if (paragraphIndex >= 0 && paragraphIndex < paragraphs.items.length) {

                const para = paragraphs.items[paragraphIndex];

                para.select();

                await context.sync();

                logStep('NAVIGATE', `Navigated to paragraph ${paragraphIndex + 1}`);

            }

        });

    } catch (error) {

        console.error('Navigation error:', error);

    }

}



// ============================================

// CLEAR HIGHLIGHTS

// ============================================



async function clearHighlightsAndResults() {

    clearResults();



    try {

        await Word.run(async (context) => {

            context.document.body.font.highlightColor = null;

            await context.sync();

        });

        displayResults();

        logStep('CLEAR', 'Highlights and results cleared');

    } catch (error) {

        console.error('Clear error:', error);

    }

}



// ============================================

// SECTION-BASED MARGIN VALIDATION (7cm Rule)

// ============================================



/**

* Validate section margins using batch loading

* Checks for 7cm top margin on main chapter starts

*/

async function validateSectionMargins(context, sections) {

    const marginErrors = [];



    try {

        // Batch load all section data

        for (let i = 0; i < sections.items.length; i++) {

            const section = sections.items[i];



            try {

                const pageSetup = section.getPageSetup();

                pageSetup.load('topMargin, bottomMargin, leftMargin, rightMargin');



                const body = section.body;

                body.paragraphs.load('items');

            } catch (e) {

                // Mac compatibility - getPageSetup may not be available

                logStep('MARGIN', `Section ${i + 1}: getPageSetup not available`);

            }

        }



        await context.sync();



        // Now validate each section

        for (let i = 0; i < sections.items.length; i++) {

            const section = sections.items[i];



            try {

                const pageSetup = section.getPageSetup();

                const body = section.body;



                // Get first paragraph to check if it's a main heading

                let firstParaText = '';

                if (body.paragraphs.items && body.paragraphs.items.length > 0) {

                    const firstPara = body.paragraphs.items[0];

                    firstPara.load('text');

                    await context.sync();

                    firstParaText = (firstPara.text || '').trim();

                }



                // Determine expected top margin

                const isMainChapterStart = isMainHeadingText(firstParaText);

                const expectedTopMargin = isMainChapterStart

                    ? EBYÜ_RULES.MARGIN_TOP_SPECIAL_POINTS

                    : EBYÜ_RULES.MARGIN_POINTS;



                const tolerance = EBYÜ_RULES.MARGIN_TOLERANCE;



                // Check top margin

                if (pageSetup.topMargin !== undefined) {

                    if (Math.abs(pageSetup.topMargin - expectedTopMargin) > tolerance) {

                        const expectedCm = isMainChapterStart ? 7 : 3;

                        marginErrors.push({

                            type: 'error',

                            title: `Bölüm ${i + 1}: Üst Kenar Boşluğu`,

                            description: `Üst kenar ${expectedCm} cm olmalı. Mevcut: ${(pageSetup.topMargin / 28.35).toFixed(2)} cm`,

                            location: `Bölüm ${i + 1}`,

                            severity: 'CRITICAL'

                        });

                    }



                    // Check other margins (should always be 3cm)

                    if (Math.abs(pageSetup.bottomMargin - EBYÜ_RULES.MARGIN_POINTS) > tolerance) {

                        marginErrors.push({

                            type: 'error',

                            title: `Bölüm ${i + 1}: Alt Kenar Boşluğu`,

                            description: `Alt kenar 3 cm olmalı. Mevcut: ${(pageSetup.bottomMargin / 28.35).toFixed(2)} cm`,

                            location: `Bölüm ${i + 1}`,

                            severity: 'CRITICAL'

                        });

                    }



                    if (Math.abs(pageSetup.leftMargin - EBYÜ_RULES.MARGIN_POINTS) > tolerance) {

                        marginErrors.push({

                            type: 'error',

                            title: `Bölüm ${i + 1}: Sol Kenar Boşluğu`,

                            description: `Sol kenar 3 cm olmalı. Mevcut: ${(pageSetup.leftMargin / 28.35).toFixed(2)} cm`,

                            location: `Bölüm ${i + 1}`,

                            severity: 'CRITICAL'

                        });

                    }



                    if (Math.abs(pageSetup.rightMargin - EBYÜ_RULES.MARGIN_POINTS) > tolerance) {

                        marginErrors.push({

                            type: 'error',

                            title: `Bölüm ${i + 1}: Sağ Kenar Boşluğu`,

                            description: `Sağ kenar 3 cm olmalı. Mevcut: ${(pageSetup.rightMargin / 28.35).toFixed(2)} cm`,

                            location: `Bölüm ${i + 1}`,

                            severity: 'CRITICAL'

                        });

                    }

                }

            } catch (sectionError) {

                logStep('MARGIN', `Section ${i + 1} check failed: ${sectionError.message}`);

            }

        }



        if (marginErrors.length === 0) {

            logStep('MARGIN', 'All section margins validated successfully');

        }



    } catch (error) {

        logStep('MARGIN', `Margin validation failed: ${error.message}`);

        addResult('warning', 'Kenar Boşlukları (Manuel Kontrol)',

            'Otomatik kontrol başarısız. Lütfen manuel kontrol edin: Tümü 3 cm, ana bölüm başlangıçları 7 cm üst kenar.');

    }



    return marginErrors;

}



// ============================================

// PARAGRAPH VALIDATION FUNCTIONS

// ============================================



function validateMainHeading(paraData, index) {

    const errors = [];

    const { font, alignment, text, isListItem, listString } = paraData;



    // === NEW: Başlık Numaralandırma (List String) Kontrolü ===

    // Ana başlıklar numaralandırılmalı: "1.", "2.", "3." formatında

    // Not: GİRİŞ, SONUÇ, ÖZET gibi özel başlıklar numaralandırılmaz

    const trimmedText = (text || '').trim();

    const isSpecialHeading = /^(GİRİŞ|SONUÇ|SONUÇ VE ÖNERİLER|TARTIŞMA|KAYNAKÇA|KAYNAKLAR|ÖZET|ABSTRACT|SUMMARY|ÖN\s*SÖZ|İÇİNDEKİLER|KISALTMALAR|TABLOLAR|ŞEKİLLER|GRAFİKLER|SİMGELER|EKLER?)$/i.test(trimmedText);



    if (!isSpecialHeading) {

        // Check if paragraph is a list item (has automatic numbering)

        if (!isListItem) {

            errors.push({

                type: 'error',

                title: 'Ana Başlık: Numaralandırılmamış',

                description: 'Ana başlık otomatik numaralandırma listesi ile numaralandırılmalı (Örn: "1.", "2."). Word\'de Çok Düzeyli Liste kullanın.',

                paraIndex: index,

                severity: 'CRITICAL'

            });

        } else {

            // Validate listString format: should be like "1.", "2.", "3." etc.

            const mainHeadingPattern = /^\d+\.?$/;

            const cleanListString = (listString || '').trim();



            if (!cleanListString) {

                errors.push({

                    type: 'error',

                    title: 'Ana Başlık: Numaralandırma Eksik',

                    description: 'Ana başlık için otomatik numara (listString) alınamadı. Word\'de Çok Düzeyli Liste ile "1.", "2." formatında numaralandırın.',

                    paraIndex: index,

                    severity: 'CRITICAL'

                });

            } else if (!mainHeadingPattern.test(cleanListString)) {

                errors.push({

                    type: 'warning',

                    title: 'Ana Başlık: Numaralandırma Formatı',

                    description: `Ana başlık numarası "1.", "2." formatında olmalı. Mevcut: "${cleanListString}"`,

                    paraIndex: index,

                    severity: 'FORMAT'

                });

            }

        }

    }

    // === END NEW ===



    // Font size: 14pt

    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_HEADING_MAIN) > 0.5) {

        errors.push({

            type: 'warning',

            title: 'Ana Başlık: Punto Hatası',

            description: `Ana başlık 14 punto olmalı. Mevcut: ${font.size} pt`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Bold

    if (font.bold !== true) {

        errors.push({

            type: 'warning',

            title: 'Ana Başlık: Kalın Yazı',

            description: 'Ana başlık kalın (bold) olmalı.',

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Centered

    if (!isCentered(alignment)) {

        errors.push({

            type: 'warning',

            title: 'Ana Başlık: Hizalama',

            description: 'Ana başlık ortalanmış olmalı.',

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Font name

    if (font.name && font.name !== EBYÜ_RULES.FONT_NAME) {

        errors.push({

            type: 'error',

            title: 'Ana Başlık: Yazı Tipi',

            description: `${EBYÜ_RULES.FONT_NAME} olmalı. Mevcut: ${font.name}`,

            paraIndex: index,

            severity: 'CRITICAL'

        });

    }



    return errors;

}



function validateSubHeading(paraData, index) {

    const errors = [];

    const { font, alignment, text, isListItem, listString } = paraData;



    // === NEW: Alt Başlık Numaralandırma (List String) Kontrolü ===

    // Alt başlıklar hiyerarşik numaralandırılmalı: "1.1.", "1.1.2.", "2.3.1." formatında

    // Check if paragraph is a list item (has automatic numbering)

    if (!isListItem) {

        errors.push({

            type: 'error',

            title: 'Alt Başlık: Numaralandırılmamış',

            description: 'Alt başlık otomatik numaralandırma listesi ile numaralandırılmalı (Örn: "1.1.", "2.3.1."). Word\'de Çok Düzeyli Liste kullanın.',

            paraIndex: index,

            severity: 'CRITICAL'

        });

    } else {

        // Validate listString format: should be hierarchical like "1.1.", "1.2.", "1.1.1.", "2.3.1."

        // Pattern: at least two numbers separated by periods, optionally ending with period

        const subHeadingPattern = /^\d+\.\d+(\.\d+)*\.?$/;

        const cleanListString = (listString || '').trim();



        if (!cleanListString) {

            errors.push({

                type: 'error',

                title: 'Alt Başlık: Numaralandırma Eksik',

                description: 'Alt başlık için otomatik numara (listString) alınamadı. Word\'de Çok Düzeyli Liste ile "1.1.", "1.2." formatında numaralandırın.',

                paraIndex: index,

                severity: 'CRITICAL'

            });

        } else if (!subHeadingPattern.test(cleanListString)) {

            errors.push({

                type: 'warning',

                title: 'Alt Başlık: Numaralandırma Formatı',

                description: `Alt başlık numarası "1.1.", "1.2.", "1.1.1." formatında olmalı. Mevcut: "${cleanListString}"`,

                paraIndex: index,

                severity: 'FORMAT'

            });

        }

    }

    // === END NEW ===



    // Font size: 12pt

    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_HEADING_SUB) > 0.5) {

        errors.push({

            type: 'warning',

            title: 'Alt Başlık: Punto Hatası',

            description: `Alt başlık 12 punto olmalı. Mevcut: ${font.size} pt`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Bold

    if (font.bold !== true) {

        errors.push({

            type: 'warning',

            title: 'Alt Başlık: Kalın Yazı',

            description: 'Alt başlık kalın (bold) olmalı.',

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Font name

    if (font.name && font.name !== EBYÜ_RULES.FONT_NAME) {

        errors.push({

            type: 'error',

            title: 'Alt Başlık: Yazı Tipi',

            description: `${EBYÜ_RULES.FONT_NAME} olmalı. Mevcut: ${font.name}`,

            paraIndex: index,

            severity: 'CRITICAL'

        });

    }



    return errors;

}



function validateBodyText(paraData, index) {

    const errors = [];

    const { font, firstLineIndent, lineSpacing, spaceBefore, spaceAfter, text } = paraData;



    // Skip short paragraphs

    if ((text || '').trim().length < EBYÜ_RULES.MIN_BODY_TEXT_LENGTH) {

        return errors;

    }



    // === DEBUG: Log paragraph spacing values ===

    logStep('BODY_TEXT_DEBUG', `Paragraf ${index + 1}: lineSpacing=${lineSpacing}, spaceBefore=${spaceBefore}, spaceAfter=${spaceAfter}, firstLineIndent=${firstLineIndent}, leftIndent=${paraData.leftIndent}, text="${(text || '').substring(0, 50)}..."`);

    // === END DEBUG ===



    // === NEW: Manuel Tab/Boşluk Kontrolü ===

    // Check if text starts with Tab character or multiple spaces (manual indentation)

    if (text) {

        // Check for Tab character at the beginning

        if (text.startsWith('\t')) {

            errors.push({

                type: 'warning',

                title: 'UYARI: Manuel Tab Kullanmayın',

                description: 'Girintiyi "Tab" tuşuyla değil, Cetvel veya Paragraf ayarlarından 1.25 cm olarak ayarlayın.',

                paraIndex: index,

                severity: 'FORMAT'

            });

        }

        // Check for multiple spaces at the beginning (2 or more spaces)

        else if (/^\s{2,}/.test(text) && !text.startsWith('\t')) {

            errors.push({

                type: 'warning',

                title: 'UYARI: Manuel Boşluk Kullanmayın',

                description: 'Girintiyi boşluk tuşuyla değil, Cetvel veya Paragraf ayarlarından 1.25 cm olarak ayarlayın.',

                paraIndex: index,

                severity: 'FORMAT'

            });

        }

    }

    // === END NEW ===



    // Font name

    if (font.name && font.name !== EBYÜ_RULES.FONT_NAME) {

        errors.push({

            type: 'error',

            title: 'Metin: Yazı Tipi',

            description: `${EBYÜ_RULES.FONT_NAME} olmalı. Mevcut: ${font.name}`,

            paraIndex: index,

            severity: 'CRITICAL'

        });

    }



    // Font size: 12pt

    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_BODY) > 0.5) {

        errors.push({

            type: 'warning',

            title: 'Metin: Punto Hatası',

            description: `Metin 12 punto olmalı. Mevcut: ${font.size} pt`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // First line indent: 1.25cm (35.4pt)

    if (firstLineIndent !== undefined && Math.abs(firstLineIndent - EBYÜ_RULES.FIRST_LINE_INDENT_POINTS) > EBYÜ_RULES.INDENT_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'Metin: İlk Satır Girintisi',

            description: `1.25 cm olmalı. Mevcut: ${(firstLineIndent / 28.35).toFixed(2)} cm`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Line spacing: 1.5 (15-22pt range)

    if (lineSpacing !== undefined && lineSpacing !== null && (lineSpacing < EBYÜ_RULES.LINE_SPACING_1_5_MIN || lineSpacing > EBYÜ_RULES.LINE_SPACING_1_5_MAX)) {

        errors.push({

            type: 'warning',

            title: 'Metin: Satır Aralığı',

            description: `1.5 satır aralığı (15-22 pt) olmalı. Mevcut: ${lineSpacing.toFixed(1)} pt`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Paragraph spacing: 6pt before and after

    if (spaceBefore !== undefined && spaceBefore !== null && Math.abs(spaceBefore - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'Metin: Paragraf Öncesi',

            description: `6 nk olmalı. Mevcut: ${spaceBefore.toFixed(1)} nk`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    if (spaceAfter !== undefined && spaceAfter !== null && Math.abs(spaceAfter - EBYÜ_RULES.SPACING_6NK) > EBYÜ_RULES.SPACING_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'Metin: Paragraf Sonrası',

            description: `6 nk olmalı. Mevcut: ${spaceAfter.toFixed(1)} nk`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // === NEW: Block indent (Sol Girinti) detection ===

    // If the paragraph has significant left indent (not first line indent), it might be improperly formatted

    const leftIndent = paraData.leftIndent;

    if (leftIndent !== undefined && leftIndent !== null && leftIndent >= EBYÜ_RULES.BLOCK_QUOTE_MIN_INDENT) {

        // This looks like a block quote or indented section - check if it's properly styled

        errors.push({

            type: 'warning',

            title: 'Metin: Blok Girinti Tespit Edildi',

            description: `Bu paragrafta ${(leftIndent / 28.35).toFixed(2)} cm sol girinti var. Blok alıntı ise 11pt/italik olmalı, aksi halde girintiyi kaldırın.`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }

    // === END NEW ===



    return errors;

}



function validateGhostHeading(paraData, index) {

    const { style, outlineLevel } = paraData;



    let reason = '';

    if (typeof outlineLevel === 'number' && outlineLevel >= 0 && outlineLevel <= 8) {

        reason = `Taslak düzeyi ${outlineLevel + 1} olarak ayarlanmış`;

    } else if (isHeadingStyle(style)) {

        reason = `"${style}" başlık stili uygulanmış`;

    }



    return [{

        type: 'error',

        title: 'BOŞ BAŞLIK (Ghost Heading) - KRİTİK!',

        description: `Bu boş satıra ${reason}. İçindekiler tablosunda hatalı boş satır oluşturur! Satırı silin veya "Normal" stiline dönüştürün.`,

        paraIndex: index,

        severity: 'CRITICAL'

    }];

}



function validateBibliography(paraData, index) {

    const errors = [];

    const { font, leftIndent, firstLineIndent, lineSpacing, spaceBefore, spaceAfter } = paraData;



    // Font name

    if (font.name && font.name !== EBYÜ_RULES.FONT_NAME) {

        errors.push({

            type: 'error',

            title: 'Kaynakça: Yazı Tipi',

            description: `${EBYÜ_RULES.FONT_NAME} olmalı.`,

            paraIndex: index,

            severity: 'CRITICAL'

        });

    }



    // 12pt

    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_BODY) > 0.5) {

        errors.push({

            type: 'warning',

            title: 'Kaynakça: Punto',

            description: `12 punto olmalı. Mevcut: ${font.size} pt`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Hanging indent (1cm = 28.35pt)

    const hangingIndent = leftIndent - firstLineIndent;

    if (Math.abs(hangingIndent - EBYÜ_RULES.BIBLIOGRAPHY_HANGING_INDENT_POINTS) > EBYÜ_RULES.INDENT_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'Kaynakça: Asılı Girinti',

            description: `1 cm asılı girinti olmalı.`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Single line spacing

    if (lineSpacing !== undefined && (lineSpacing < EBYÜ_RULES.LINE_SPACING_SINGLE_MIN || lineSpacing > EBYÜ_RULES.LINE_SPACING_SINGLE_MAX)) {

        errors.push({

            type: 'warning',

            title: 'Kaynakça: Satır Aralığı',

            description: `Tek satır olmalı. Mevcut: ${lineSpacing.toFixed(1)} pt`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // 3pt paragraph spacing

    if (spaceBefore !== undefined && Math.abs(spaceBefore - EBYÜ_RULES.SPACING_3NK) > EBYÜ_RULES.SPACING_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'Kaynakça: Paragraf Öncesi',

            description: `3 nk olmalı. Mevcut: ${spaceBefore.toFixed(1)} nk`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    return errors;

}



function validateBlockQuote(paraData, index) {

    const errors = [];

    const { font, leftIndent, rightIndent, lineSpacing } = paraData;



    // 11pt

    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_BLOCK_QUOTE) > 0.5) {

        errors.push({

            type: 'warning',

            title: 'Blok Alıntı: Punto',

            description: `11 punto olmalı. Mevcut: ${font.size} pt`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Italic

    if (font.italic !== true) {

        errors.push({

            type: 'warning',

            title: 'Blok Alıntı: İtalik',

            description: 'Blok alıntı italik olmalı.',

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // 1.25cm left and right indent

    if (leftIndent !== undefined && Math.abs(leftIndent - EBYÜ_RULES.BLOCK_QUOTE_INDENT_POINTS) > EBYÜ_RULES.INDENT_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'Blok Alıntı: Sol Girinti',

            description: `1.25 cm olmalı. Mevcut: ${(leftIndent / 28.35).toFixed(2)} cm`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    if (rightIndent !== undefined && Math.abs(rightIndent - EBYÜ_RULES.BLOCK_QUOTE_INDENT_POINTS) > EBYÜ_RULES.INDENT_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'Blok Alıntı: Sağ Girinti',

            description: `1.25 cm olmalı. Mevcut: ${(rightIndent / 28.35).toFixed(2)} cm`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    return errors;

}



function validateCaption(paraData, index) {

    const errors = [];

    const { font, alignment, text } = paraData;



    // 12pt for caption title

    if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_CAPTION_TITLE) > 0.5) {

        errors.push({

            type: 'warning',

            title: 'Başlık: Punto',

            description: `Tablo/Şekil başlığı 12 punto olmalı. Mevcut: ${font.size} pt`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Centered

    if (!isCentered(alignment)) {

        errors.push({

            type: 'warning',

            title: 'Başlık: Hizalama',

            description: 'Tablo/Şekil başlığı ortalanmış olmalı.',

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Caption spacing: 0nk before and after

    if (paraData.spaceBefore !== undefined && paraData.spaceBefore > EBYÜ_RULES.SPACING_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'Tablo/Şekil Başlığı: Paragraf Öncesi',

            description: `Şekil/Tablo başlıklarında 0 nk olmalı. Mevcut: ${paraData.spaceBefore.toFixed(1)} nk`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    if (paraData.spaceAfter !== undefined && paraData.spaceAfter > EBYÜ_RULES.SPACING_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'Tablo/Şekil Başlığı: Paragraf Sonrası',

            description: `Şekil/Tablo başlıklarında 0 nk olmalı. Mevcut: ${paraData.spaceAfter.toFixed(1)} nk`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    return errors;

}



// ============================================

// COVER PAGE VALIDATION (Kapak Sayfası - 16pt, 0nk)

// ============================================



function validateCoverPage(paraData, index) {

    const errors = [];

    const { font, alignment, spaceBefore, spaceAfter, text } = paraData;

    const trimmed = (text || '').trim();



    // Skip empty paragraphs

    if (trimmed.length === 0) return errors;



    // Cover title should be 16pt (main titles on cover)

    const isMainCoverTitle = /^(T\.?C\.?|ERZİNCAN|ÜNİVERSİTESİ|ENSTİTÜSÜ|TEZİ)$/i.test(trimmed) ||

        trimmed.length > 20; // Thesis title



    if (isMainCoverTitle && font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_COVER_TITLE) > 1) {

        errors.push({

            type: 'error',

            title: 'KAPAK: Punto Hatası',

            description: `Kapak başlıkları 16 punto olmalı. Mevcut: ${font.size} pt`,

            paraIndex: index,

            severity: 'CRITICAL'

        });

    }



    // Cover should be centered

    if (!isCentered(alignment)) {

        errors.push({

            type: 'warning',

            title: 'KAPAK: Hizalama',

            description: 'Kapak öğeleri ortalanmış olmalı.',

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Cover spacing should be 0nk

    if (spaceBefore !== undefined && spaceBefore > EBYÜ_RULES.SPACING_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'KAPAK: Paragraf Öncesi Boşluk',

            description: `Kapakta 0 nk olmalı. Mevcut: ${spaceBefore.toFixed(1)} nk`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    if (spaceAfter !== undefined && spaceAfter > EBYÜ_RULES.SPACING_TOLERANCE) {

        errors.push({

            type: 'warning',

            title: 'KAPAK: Paragraf Sonrası Boşluk',

            description: `Kapakta 0 nk olmalı. Mevcut: ${spaceAfter.toFixed(1)} nk`,

            paraIndex: index,

            severity: 'FORMAT'

        });

    }



    // Font must be Times New Roman

    if (font.name && font.name !== EBYÜ_RULES.FONT_NAME) {

        errors.push({

            type: 'error',

            title: 'KAPAK: Yazı Tipi',

            description: `${EBYÜ_RULES.FONT_NAME} olmalı. Mevcut: ${font.name}`,

            paraIndex: index,

            severity: 'CRITICAL'

        });

    }



    return errors;

}



// ============================================

// TABLE ALIGNMENT VALIDATION (Tablo Hizalama)

// ============================================



async function validateTables(context) {

    const errors = [];



    try {

        const tables = context.document.body.tables;

        tables.load('items');

        await context.sync();



        for (let i = 0; i < tables.items.length; i++) {

            const table = tables.items[i];

            table.load(['alignment', 'font/size', 'font/name']);

            await context.sync();



            // Check alignment - must be centered

            if (table.alignment && !isCentered(table.alignment) &&

                table.alignment !== 'Mixed' &&

                table.alignment !== 'Unknown') {



                errors.push({

                    type: 'warning',

                    title: `Tablo ${i + 1}: Hizalama Hatası`,

                    description: `Tablolar ortalanmış olmalı. Mevcut: ${table.alignment}`,

                    severity: 'FORMAT',

                    tableIndex: i

                });



                // Highlight table

                table.font.highlightColor = HIGHLIGHT_COLORS.FORMAT;

            }



            // Check font size - table content should be 11pt

            if (table.font.size && Math.abs(table.font.size - EBYÜ_RULES.TABLE_CONTENT_SIZE) > 0.5) {

                errors.push({

                    type: 'warning',

                    title: `Tablo ${i + 1}: Punto Hatası`,

                    description: `Tablo içeriği 11 punto olmalı. Mevcut: ${table.font.size} pt`,

                    severity: 'FORMAT',

                    tableIndex: i

                });

            }

        }



        logStep('TABLES', `Validated ${tables.items.length} tables, found ${errors.length} errors`);

    } catch (error) {

        logStep('TABLES', `Table validation error: ${error.message}`);

    }



    return errors;

}



// ============================================

// IMAGE ALIGNMENT VALIDATION (Resim Hizalama)

// ============================================



async function validateImages(context) {

    const errors = [];



    try {

        const pictures = context.document.body.inlinePictures;

        pictures.load('items');

        await context.sync();



        for (let i = 0; i < pictures.items.length; i++) {

            const pic = pictures.items[i];

            pic.paragraph.load('alignment');

            await context.sync();



            const alignment = pic.paragraph.alignment;



            // Images should be centered

            if (alignment && !isCentered(alignment)) {



                errors.push({

                    type: 'warning',

                    title: `Resim ${i + 1}: Hizalama Hatası`,

                    description: `Resimler ortalanmış olmalı. Mevcut: ${alignment}`,

                    severity: 'FORMAT',

                    pictureIndex: i

                });



                // Highlight the paragraph containing image

                pic.paragraph.font.highlightColor = HIGHLIGHT_COLORS.FORMAT;

            }

        }



        logStep('IMAGES', `Validated ${pictures.items.length} images, found ${errors.length} errors`);

    } catch (error) {

        logStep('IMAGES', `Image validation error: ${error.message}`);

    }



    return errors;

}



// ============================================

// PAGE NUMBER VALIDATION (Sayfa No Kontrolü)

// ============================================



async function validatePageNumbers(context, sections) {

    const errors = [];



    try {

        for (let i = 0; i < sections.items.length; i++) {

            const section = sections.items[i];



            try {

                section.pageSetup.load('footerDistance');

                await context.sync();



                const footerDistance = section.pageSetup.footerDistance;



                // Footer distance should be 1.25 cm (35.4 pt)

                if (footerDistance !== undefined &&

                    Math.abs(footerDistance - EBYÜ_RULES.PAGE_NUMBER_FOOTER_DISTANCE_POINTS) > EBYÜ_RULES.MARGIN_TOLERANCE) {



                    errors.push({

                        type: 'warning',

                        title: `Bölüm ${i + 1}: Sayfa No Konumu`,

                        description: `Sayfa numarası alt kenardan 1.25 cm yukarıda olmalı. Mevcut: ${(footerDistance / 28.35).toFixed(2)} cm`,

                        location: `Bölüm ${i + 1}`,

                        severity: 'FORMAT'

                    });

                }



                // Check footer content for page number

                const footer = section.getFooter("Primary");

                footer.load('text');

                await context.sync();



                // Footer should contain page number (numeric content)

                if (footer.text && footer.text.trim().length === 0) {

                    errors.push({

                        type: 'warning',

                        title: `Bölüm ${i + 1}: Sayfa Numarası Eksik`,

                        description: 'Alt bilgide sayfa numarası bulunamadı.',

                        location: `Bölüm ${i + 1}`,

                        severity: 'FORMAT'

                    });

                }

            } catch (sectionError) {

                logStep('PAGE_NUM', `Section ${i + 1} page number check failed: ${sectionError.message}`);

            }

        }



        logStep('PAGE_NUM', `Validated page numbers, found ${errors.length} errors`);

    } catch (error) {

        logStep('PAGE_NUM', `Page number validation error: ${error.message}`);

    }



    return errors;

}



// ============================================

// ABSTRACT VALIDATION (Özet Sayfası)

// ============================================



function validateAbstract(paraData, abstractParagraphs) {

    const errors = [];



    // Count total words in abstract

    let totalWords = 0;

    let abstractText = '';



    for (const para of abstractParagraphs) {

        const text = (para.text || '').trim();

        if (text.length > 0 && !text.match(/^(ÖZET|ABSTRACT|Anahtar Kelimeler|Keywords)/i)) {

            const words = text.split(/\s+/).filter(w => w.length > 0);

            totalWords += words.length;

            abstractText += text + ' ';

        }

    }



    // Word count check: 200-250 words

    if (totalWords < EBYÜ_RULES.ABSTRACT_MIN_WORDS) {

        errors.push({

            type: 'warning',

            title: 'Özet: Kelime Sayısı Az',

            description: `Özet en az ${EBYÜ_RULES.ABSTRACT_MIN_WORDS} kelime olmalı. Mevcut: ${totalWords} kelime`,

            severity: 'FORMAT'

        });

    } else if (totalWords > EBYÜ_RULES.ABSTRACT_MAX_WORDS) {

        errors.push({

            type: 'warning',

            title: 'Özet: Kelime Sayısı Fazla',

            description: `Özet en fazla ${EBYÜ_RULES.ABSTRACT_MAX_WORDS} kelime olmalı. Mevcut: ${totalWords} kelime`,

            severity: 'FORMAT'

        });

    }



    return errors;

}



// ============================================

// ADD ERROR COMMENT (Hata Yorumu Ekleme)

// ============================================



async function addErrorComment(context, paragraph, errorMessage) {

    try {

        const range = paragraph.getRange();

        range.insertComment(errorMessage);

        logStep('COMMENT', `Added comment: ${errorMessage.substring(0, 50)}...`);

    } catch (error) {

        // Comments may not be supported in all environments

        logStep('COMMENT', `Could not add comment: ${error.message}`);

    }

}



// ============================================

// MAIN SCAN FUNCTION (Batch Loading Optimized)

// ============================================



async function scanDocument() {

    if (isScanning) return;

    isScanning = true;



    const startTime = performance.now();

    clearResults();

    setButtonState(false);

    updateProgress(0, 'Tarama başlatılıyor...');

    logStep('START', 'Document scan initiated');



    try {

        await Word.run(async (context) => {

            // Step 1: Clear previous highlights

            updateProgress(5, 'Önceki işaretler temizleniyor...');

            context.document.body.font.highlightColor = null;

            await context.sync();



            // Step 2: Load document structure

            updateProgress(10, 'Belge yapısı yükleniyor...');



            const sections = context.document.sections;

            sections.load('items');



            const paragraphs = context.document.body.paragraphs;



            // BATCH LOAD: Load all paragraph properties at once

            // FIX: Use direct paragraph properties instead of paragraphFormat/* for Mac Word compatibility

            // Word.Paragraph has leftIndent, lineSpacing, spaceBefore, spaceAfter, firstLineIndent, alignment directly

            paragraphs.load([

                'items/text',

                'items/style',

                'items/outlineLevel',

                'items/tableNestingLevel',

                'items/font/name',

                'items/font/size',

                'items/font/bold',

                'items/font/italic',

                'items/alignment',

                'items/firstLineIndent',

                'items/leftIndent',

                'items/rightIndent',

                'items/lineSpacing',

                'items/spaceBefore',

                'items/spaceAfter',

                'items/isListItem',

                'items/listItemOrNullObject/listString',

                'items/listItemOrNullObject/level'

            ].join(','));



            await context.sync();

            logStep('LOAD', `Loaded ${paragraphs.items.length} paragraphs, ${sections.items.length} sections`);



            // Step 3: Validate section margins (7cm rule)

            updateProgress(20, 'Kenar boşlukları kontrol ediliyor...');

            const marginErrors = await validateSectionMargins(context, sections);

            for (const err of marginErrors) {

                addResult(err.type, err.title, err.description, err.location, null, err.severity);

            }



            // Step 4: Prepare paragraph data objects (no sync needed)

            updateProgress(30, 'Paragraf verileri hazırlanıyor...');

            const paragraphDataList = [];



            for (let i = 0; i < paragraphs.items.length; i++) {

                const p = paragraphs.items[i];



                // Defensive null checks

                const pFont = p.font || {};



                // Extract listItem properties for heading numbering validation

                const listItem = p.listItemOrNullObject;

                let listString = '';

                let listLevel = null;



                // Check if listItem was loaded and is not a null object

                if (listItem && !listItem.isNullObject) {

                    listString = listItem.listString || '';

                    listLevel = listItem.level;

                }



                // FIX: Use direct paragraph properties instead of paragraphFormat for Mac Word

                paragraphDataList.push({

                    index: i,

                    text: p.text || '',

                    style: p.style || '',

                    outlineLevel: p.outlineLevel,

                    tableNestingLevel: p.tableNestingLevel || 0,

                    font: {

                        name: pFont.name,

                        size: pFont.size,

                        bold: pFont.bold,

                        italic: pFont.italic

                    },

                    alignment: p.alignment,

                    firstLineIndent: p.firstLineIndent,

                    leftIndent: p.leftIndent,

                    rightIndent: p.rightIndent,

                    lineSpacing: p.lineSpacing,

                    spaceBefore: p.spaceBefore,

                    spaceAfter: p.spaceAfter,

                    // List item properties for heading numbering

                    isListItem: p.isListItem || false,

                    listString: listString,

                    listLevel: listLevel,

                    paragraph: p // Keep reference for highlighting

                });

            }



            // Step 5: Zone-based validation

            updateProgress(40, 'Paragraflar analiz ediliyor...');



            let currentZone = ZONES.COVER;

            let isInBiblio = false;

            let ghostCount = 0;

            let errorCount = 0;

            let warningCount = 0;



            for (let i = 0; i < paragraphDataList.length; i++) {

                const paraData = paragraphDataList[i];

                const text = paraData.text.trim();



                // Update progress periodically

                if (i % 50 === 0) {

                    const progressPercent = 40 + Math.floor((i / paragraphDataList.length) * 50);

                    updateProgress(progressPercent, `Paragraf ${i + 1} / ${paragraphDataList.length}`);

                }



                // Zone switching - KAPAK -> ÖN KISIM (Roma) -> ANA METİN (Normal) -> KAYNAKÇA

                // Cover is only the first 15 paragraphs OR until ÖZET/front matter

                const textUpper = text.toUpperCase();



                // Cover ends at ÖZET, İÇİNDEKİLER, ÖNSÖZ, or after 15 paragraphs

                if (currentZone === ZONES.COVER) {

                    if (textUpper === 'ÖZET' || textUpper.startsWith('ÖZET') ||

                        textUpper === 'ÖNSÖZ' || textUpper.startsWith('ÖNSÖZ') ||

                        matchesAnyPattern(text, PATTERNS.FRONT_MATTER_IDENTIFIERS) ||

                        i >= 15) {

                        currentZone = ZONES.FRONT_MATTER;

                        logStep('ZONE', `Switched to FRONT_MATTER at paragraph ${i + 1}: "${text.substring(0, 30)}..."`);

                    }

                }



                // Body starts with GİRİŞ - Normal rakamlar başlar

                if (matchesAnyPattern(text, PATTERNS.BODY_START)) {

                    currentZone = ZONES.BODY;

                    logStep('ZONE', `Switched to BODY at paragraph ${i + 1}: "${text.substring(0, 30)}..."`);

                }



                // Back matter starts with KAYNAKÇA

                if (matchesAnyPattern(text, PATTERNS.BACK_MATTER_START)) {

                    currentZone = ZONES.BACK_MATTER;

                    isInBiblio = true;

                    logStep('ZONE', `Switched to BACK_MATTER at paragraph ${i + 1}: "${text.substring(0, 30)}..."`);

                }



                // Detect paragraph type

                const paraType = detectParagraphType(paraData, currentZone, isInBiblio);



                // Validate based on type

                let errors = [];



                switch (paraType) {

                    case PARA_TYPES.GHOST_HEADING:

                        errors = validateGhostHeading(paraData, i);

                        // Highlight ghost heading

                        paraData.paragraph.font.highlightColor = HIGHLIGHT_COLORS.CRITICAL;

                        ghostCount++;

                        break;



                    case PARA_TYPES.MAIN_HEADING:

                        errors = validateMainHeading(paraData, i);

                        break;



                    case PARA_TYPES.SUB_HEADING:

                        errors = validateSubHeading(paraData, i);

                        break;



                    case PARA_TYPES.BODY_TEXT:

                        errors = validateBodyText(paraData, i);

                        break;



                    case PARA_TYPES.BLOCK_QUOTE:

                        errors = validateBlockQuote(paraData, i);

                        break;



                    case PARA_TYPES.BIBLIOGRAPHY:

                        errors = validateBibliography(paraData, i);

                        break;



                    case PARA_TYPES.CAPTION_TITLE:

                        errors = validateCaption(paraData, i);

                        break;



                    case PARA_TYPES.COVER_TEXT:

                        errors = validateCoverPage(paraData, i);

                        break;

                }



                // Add errors and apply highlights + comments

                let hasCriticalError = false;



                for (const err of errors) {

                    addResult(err.type, err.title, err.description, `Paragraf ${i + 1}`, err.paraIndex, err.severity);



                    // Track if we have a critical error

                    if (err.severity === 'CRITICAL' || err.type === 'error') {

                        hasCriticalError = true;

                    }



                    if (err.type === 'error') {

                        errorCount++;

                    } else if (err.type === 'warning') {

                        warningCount++;

                    }

                }



                // Apply highlight ONCE after all errors checked (avoid reading highlightColor)

                if (errors.length > 0) {

                    const highlightColor = hasCriticalError

                        ? HIGHLIGHT_COLORS.CRITICAL

                        : HIGHLIGHT_COLORS.FORMAT;

                    paraData.paragraph.font.highlightColor = highlightColor;

                }

            }



            // Step 6: Validate Tables (Tablo Hizalama)

            updateProgress(85, 'Tablolar kontrol ediliyor...');

            const tableErrors = await validateTables(context);

            for (const err of tableErrors) {

                addResult(err.type, err.title, err.description, `Tablo ${err.tableIndex + 1}`, null, err.severity);

                if (err.type === 'error') errorCount++;

                else warningCount++;

            }



            // Step 7: Validate Images (Resim Hizalama)

            updateProgress(88, 'Resimler kontrol ediliyor...');

            const imageErrors = await validateImages(context);

            for (const err of imageErrors) {

                addResult(err.type, err.title, err.description, `Resim ${err.pictureIndex + 1}`, null, err.severity);

                if (err.type === 'error') errorCount++;

                else warningCount++;

            }



            // Step 8: Validate Page Numbers (Sayfa No Konumu)

            updateProgress(91, 'Sayfa numaraları kontrol ediliyor...');

            const pageNumErrors = await validatePageNumbers(context, sections);

            for (const err of pageNumErrors) {

                addResult(err.type, err.title, err.description, err.location, null, err.severity);

                if (err.type === 'error') errorCount++;

                else warningCount++;

            }



            // Step 8.5: Validate Özet and Abstract (Özet Sayfaları)

            updateProgress(93, 'Özet/Abstract kontrol ediliyor...');



            // Find ÖZET section paragraphs

            let ozetStartIndex = -1;

            let abstractStartIndex = -1;



            for (let i = 0; i < paragraphDataList.length; i++) {

                const text = paragraphDataList[i].text.trim().toUpperCase();

                if (text === 'ÖZET' || text.startsWith('ÖZET')) {

                    ozetStartIndex = i;

                }

                if (text === 'ABSTRACT' || text.startsWith('ABSTRACT')) {

                    abstractStartIndex = i;

                }

            }



            // Validate ÖZET (Turkish abstract) - words between ÖZET and Anahtar Kelimeler

            if (ozetStartIndex >= 0) {

                const ozetParagraphs = [];

                for (let i = ozetStartIndex + 1; i < paragraphDataList.length && i < ozetStartIndex + 30; i++) {

                    const text = paragraphDataList[i].text.trim();

                    const textUpper = text.toUpperCase();

                    // Stop at "Anahtar Kelimeler" which marks end of ÖZET

                    if (textUpper.startsWith('ANAHTAR KELİMELER') || textUpper.startsWith('ANAHTAR SÖZCÜKLER')) {

                        break;

                    }

                    // Skip empty or very short paragraphs

                    if (text.length > 10) {

                        ozetParagraphs.push(paragraphDataList[i]);

                    }

                }



                const ozetErrors = validateAbstract(null, ozetParagraphs);

                for (const err of ozetErrors) {

                    err.title = err.title.replace('Özet:', 'ÖZET:');

                    addResult(err.type, err.title, err.description, 'Özet Sayfası', null, err.severity);

                    warningCount++;

                }

            }



            // Validate ABSTRACT (English abstract) - words between ABSTRACT and Keywords

            if (abstractStartIndex >= 0) {

                const abstractParagraphs = [];

                for (let i = abstractStartIndex + 1; i < paragraphDataList.length && i < abstractStartIndex + 30; i++) {

                    const text = paragraphDataList[i].text.trim();

                    const textUpper = text.toUpperCase();

                    // Stop at "Keywords" which marks end of ABSTRACT

                    if (textUpper.startsWith('KEYWORDS') || textUpper.startsWith('KEY WORDS')) {

                        break;

                    }

                    // Skip empty or very short paragraphs

                    if (text.length > 10) {

                        abstractParagraphs.push(paragraphDataList[i]);

                    }

                }



                const abstractErrors = validateAbstract(null, abstractParagraphs);

                for (const err of abstractErrors) {

                    err.title = err.title.replace('Özet:', 'ABSTRACT:');

                    addResult(err.type, err.title, err.description, 'Abstract Sayfası', null, err.severity);

                    warningCount++;

                }

            }



            // Step 9: Apply highlights

            updateProgress(95, 'İşaretler uygulanıyor...');

            await context.sync();



            // Step 10: Summary

            updateProgress(98, 'Özet hazırlanıyor...');



            if (ghostCount > 0) {

                addResult('error', `${ghostCount} Boş Başlık (Ghost Heading) Bulundu`,

                    'Bu boş başlıklar İçindekiler tablosunda hatalı satırlara neden olur. Kırmızı ile işaretlendi.',

                    'Belge Geneli', null, 'CRITICAL');

            }



            // DEBUG section removed - spacing values confirmed working



            const totalErrors = errorCount + marginErrors.filter(e => e.type === 'error').length;

            const totalWarnings = warningCount + marginErrors.filter(e => e.type === 'warning').length +

                tableErrors.length + imageErrors.length + pageNumErrors.length;



            if (totalErrors === 0 && totalWarnings === 0) {

                addResult('success', '✅ Tebrikler!',

                    'Belge EBYÜ 2022 Tez Yazım Kılavuzu formatına uygun görünüyor.');

            } else {

                addResult(totalErrors > 0 ? 'error' : 'warning', 'Tarama Özeti',

                    `🔴 Kritik: ${totalErrors} | 🟡 Format: ${totalWarnings} hata bulundu. Hatalı yerler belgede işaretlendi ve yorum eklendi.`);

            }



            updateProgress(100, 'Tarama tamamlandı!');



            const endTime = performance.now();

            logStep('COMPLETE', `Scan completed in ${((endTime - startTime) / 1000).toFixed(2)} seconds`);

        });



    } catch (error) {

        logStep('ERROR', `Scan failed: ${error.message}`);

        addResult('error', 'Tarama Hatası', `Hata: ${error.message}. Lütfen tekrar deneyin.`);

    } finally {

        isScanning = false;

        setButtonState(true);

        hideProgress();

        displayResults();

    }

}



// ============================================

// OFFICE.JS INITIALIZATION

// ============================================



Office.onReady((info) => {

    if (info.host === Office.HostType.Word) {

        console.log('EBYÜ Thesis Validator v4.0 (Enhanced with Tables/Images/PageNum): Office.js initialized');

        initializeUI();

    } else {

        console.error('This add-in only works with Microsoft Word.');

    }

});
