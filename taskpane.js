/**
 * ============================================
 * EBYÜ Thesis Format Validator - Task Pane Logic
 * Erzincan Binali Yıldırım University
 * Based on: EBYÜ 2022 Tez Yazım Kılavuzu
 * ============================================
 * 
 * VALIDATION RULES (from University Guide):
 * 
 * 1. PAGE LAYOUT & MARGINS:
 *    - All margins: 3 cm (approx 85 points)
 * 
 * 2. FONT & TYPOGRAPHY:
 *    - General Text: Times New Roman, 12 pt
 *    - Footnotes: Times New Roman, 10 pt
 *    - Headings (Level 1): 14 pt, Bold, All Caps, Centered
 *    - Sub-headings: 12 pt, Bold, Indented 1.25 cm
 * 
 * 3. PARAGRAPH FORMATTING:
 *    - Alignment: Justified
 *    - Line Spacing: 1.5 lines for body text
 *    - First Line Indent: 1.25 cm (approx 35.4 points)
 *    - Spacing: Before 6pt, After 6pt
 * 
 * 4. IMAGES & TABLES:
 *    - Wrapping: Must be "In Line with Text" (not floating)
 *    - Alignment: Centered
 */

// ============================================
// CONSTANTS - EBYÜ Thesis Formatting Rules
// ============================================

const EBYÜ_RULES = {
    // Margin in points (3 cm = 85.04 points, 1 cm ≈ 28.35 points)
    MARGIN_CM: 3,
    MARGIN_POINTS: 85,
    MARGIN_TOLERANCE: 2, // Allow 2 point tolerance

    // Font settings
    FONT_NAME: "Times New Roman",
    FONT_SIZE_BODY: 12,
    FONT_SIZE_HEADING1: 14,
    FONT_SIZE_FOOTNOTE: 10,
    FONT_SIZE_TABLE: 11, // Table text should be 11pt (smaller than body)

    // Paragraph settings
    LINE_SPACING: 1.5,
    LINE_SPACING_TABLE: 1.0, // Single line spacing for tables
    FIRST_LINE_INDENT_CM: 1.25,
    FIRST_LINE_INDENT_POINTS: 35.4,
    INDENT_TOLERANCE: 2,

    // Spacing (in points)
    SPACING_BEFORE: 6,
    SPACING_AFTER: 6,

    // Alignment
    ALIGNMENT_BODY: "Justify",
    ALIGNMENT_HEADING: "Center",

    // Table settings
    TABLE_CAPTION_POSITION: "above", // Caption must be above table
    TABLE_LINE_SPACING_POINTS: 12 // Single spacing for 12pt = 12 points
};

// ============================================
// GLOBAL STATE
// ============================================

let validationResults = [];
let currentFilter = 'all';
let scanLog = []; // Debug log for analysis

// Section patterns to EXCLUDE from analysis (Turkish thesis sections)
const EXCLUDED_SECTIONS = {
    // Cover pages and front matter
    COVER_PATTERNS: [
        /^T\.C\./i,
        /^ERZİNCAN BİNALİ YILDIRIM/i,
        /^ÜNİVERSİTESİ/i,
        /^SOSYAL BİLİMLER ENSTİTÜSÜ/i,
        /^FEN BİLİMLERİ ENSTİTÜSÜ/i,
        /^(Yüksek Lisans|Doktora)\s*(Tezi)?/i,
        /^Hazırlayan/i,
        /^Danışman/i,
        /^(Prof\.|Doç\.|Dr\.|Öğr\.)/i
    ],
    // Back matter sections
    BIBLIOGRAPHY_PATTERNS: [
        /^KAYNAKÇA$/i,
        /^KAYNAKLAR$/i,
        /^REFERENCES$/i,
        /^BIBLIOGRAPHY$/i
    ],
    APPENDIX_PATTERNS: [
        /^EKLER$/i,
        /^EK\s*\d+/i,
        /^APPENDIX/i,
        /^ÖZGEÇMİŞ$/i
    ],
    // Other excluded sections
    TOC_PATTERNS: [
        /^İÇİNDEKİLER$/i,
        /^TABLOLAR LİSTESİ$/i,
        /^ŞEKİLLER LİSTESİ$/i,
        /^SİMGELER/i,
        /^KISALTMALAR/i
    ]
};

// Helper function to log analysis steps
function logStep(category, message, details = null) {
    const timestamp = new Date().toISOString();
    const logEntry = { timestamp, category, message, details };
    scanLog.push(logEntry);
    console.log(`[${category}] ${message}`, details || '');
}

// Helper function to check if paragraph should be excluded
function shouldExcludeParagraph(text) {
    if (!text || text.trim().length === 0) return true;

    const trimmedText = text.trim();

    // Check all exclusion patterns
    for (const pattern of EXCLUDED_SECTIONS.COVER_PATTERNS) {
        if (pattern.test(trimmedText)) return true;
    }
    for (const pattern of EXCLUDED_SECTIONS.BIBLIOGRAPHY_PATTERNS) {
        if (pattern.test(trimmedText)) return true;
    }
    for (const pattern of EXCLUDED_SECTIONS.APPENDIX_PATTERNS) {
        if (pattern.test(trimmedText)) return true;
    }
    for (const pattern of EXCLUDED_SECTIONS.TOC_PATTERNS) {
        if (pattern.test(trimmedText)) return true;
    }

    return false;
}

// ============================================
// OFFICE.JS INITIALIZATION
// ============================================

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("EBYÜ Thesis Validator: Office.js initialized for Word");
        initializeUI();
    } else {
        console.error("This add-in only works with Microsoft Word");
        showError("Bu eklenti sadece Microsoft Word ile çalışır.");
    }
});

// ============================================
// UI INITIALIZATION
// ============================================

function initializeUI() {
    // Scan button
    document.getElementById('scanBtn').addEventListener('click', scanDocument);

    // Filter tabs
    document.querySelectorAll('.filter-tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            const filter = e.target.dataset.filter;
            setActiveFilter(filter);
        });
    });

    logStep('INIT', 'UI initialized');
}

// ============================================
// MAIN SCAN FUNCTION
// ============================================

async function scanDocument() {
    const scanBtn = document.getElementById('scanBtn');
    scanBtn.disabled = true;
    scanBtn.innerHTML = '<span>Taranıyor...</span>';

    // Reset results and logs
    validationResults = [];
    scanLog = [];
    const startTime = performance.now();
    logStep('SCAN', 'Tarama başlatıldı');

    // Show progress
    showProgress();
    updateProgress(0, "Belge analiz ediliyor...");

    try {
        await Word.run(async (context) => {
            // Load document body
            const body = context.document.body;
            body.load("text");

            // Load sections for margin check
            const sections = context.document.sections;
            sections.load("items");

            // Load paragraphs
            const paragraphs = body.paragraphs;
            paragraphs.load("items");

            await context.sync();

            logStep('LOAD', `Belge yüklendi`, {
                paragraphCount: paragraphs.items.length,
                sectionCount: sections.items.length
            });

            // Step 1: Check Margins (20%)
            updateProgress(10, "Kenar boşlukları kontrol ediliyor...");
            await checkMargins(context, sections);
            await context.sync();

            // Step 1.5: Check Cover Page (25%)
            updateProgress(20, "Kapak sayfası kontrol ediliyor...");
            await checkCoverPage(context, paragraphs);
            await context.sync();

            // Step 2: Check Fonts (40%)
            updateProgress(30, "Yazı tipleri kontrol ediliyor...");
            await checkFonts(context, paragraphs);
            await context.sync();

            // Step 3: Check Headings (50%)
            updateProgress(40, "Başlıklar kontrol ediliyor...");
            await checkHeadings(context, paragraphs);
            await context.sync();

            // Step 4: Check Paragraph Formatting (60%)
            updateProgress(50, "Paragraf formatları kontrol ediliyor...");
            await checkParagraphFormatting(context, paragraphs);
            await context.sync();

            // Step 5: Check Line Spacing (80%)
            updateProgress(70, "Satır aralıkları kontrol ediliyor...");
            await checkLineSpacing(context, paragraphs);
            await context.sync();

            // Step 5: Check Images (85%)
            updateProgress(80, "Görseller kontrol ediliyor...");
            await checkImages(context);
            await context.sync();

            // Step 6: Check Tables (90%)
            updateProgress(85, "Tablolar kontrol ediliyor...");
            await checkTables(context);
            await context.sync();

            // Step 7: Add manual check reminders (95%)
            updateProgress(95, "Ek kontroller ekleniyor...");
            addManualCheckReminders();

            updateProgress(100, "Tarama tamamlandı!");
        });

        // Display results
        setTimeout(() => {
            hideProgress();
            displayResults();
        }, 500);

    } catch (error) {
        console.error("Scan error:", error);
        hideProgress();
        addResult('error', 'Tarama Hatası', `Belge taranırken bir hata oluştu: ${error.message}`);
        displayResults();
    } finally {
        scanBtn.disabled = false;
        scanBtn.innerHTML = `
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="11" cy="11" r="8"/>
                <path d="M21 21l-4.35-4.35"/>
            </svg>
            <span>DÖKÜMAN TARA</span>
        `;
    }
}

// ============================================
// VALIDATION FUNCTIONS
// ============================================

/**
 * Check page margins
 * EBYÜ Rule: All margins must be 3 cm (85 points)
 */
async function checkMargins(context, sections) {
    try {
        for (let i = 0; i < sections.items.length; i++) {
            const section = sections.items[i];

            // Load section properties directly - Mac compatible approach
            section.load("headerFooterDistance");
            await context.sync();

            // Try to get page setup using body properties
            const sectionBody = section.body;
            sectionBody.load("*");
            await context.sync();

            // Use the section's getNext/getFirst approach to access page properties
            // For Mac compatibility, we access properties differently
            let margins = null;

            try {
                // First, try the standard approach (works on Windows)
                if (typeof section.getPageSetup === 'function') {
                    const pageSetup = section.getPageSetup();
                    pageSetup.load("topMargin, bottomMargin, leftMargin, rightMargin");
                    await context.sync();

                    margins = {
                        top: pageSetup.topMargin,
                        bottom: pageSetup.bottomMargin,
                        left: pageSetup.leftMargin,
                        right: pageSetup.rightMargin
                    };
                }
            } catch (pageSetupError) {
                console.log("getPageSetup not available, using alternative method");
            }

            // If we couldn't get margins via getPageSetup, try alternative
            if (!margins) {
                try {
                    // Alternative: Access via document sections properties
                    // Load section properties that might contain margin info
                    const sectionProps = context.document.sections.getFirst();
                    sectionProps.load("body");
                    await context.sync();

                    // Use ContentControl or Range-based approach
                    const range = sectionBody.getRange();
                    range.load("*");
                    await context.sync();

                    // If still no margins, show a generic message
                    addResult(
                        'warning',
                        'Kenar Boşlukları',
                        'Kenar boşlukları otomatik olarak kontrol edilemedi (Mac sürümü). Lütfen manuel olarak kontrol edin: Sayfa Düzeni > Kenar Boşlukları > Tüm kenarlar 3 cm olmalıdır.',
                        `Bölüm ${i + 1}`,
                        { type: 'margin', sectionIndex: i }
                    );
                    return;
                } catch (altError) {
                    console.log("Alternative margin check failed:", altError);
                    addResult(
                        'warning',
                        'Kenar Boşlukları',
                        'Kenar boşlukları kontrol edilemedi. Lütfen manuel olarak kontrol edin: Sayfa Düzeni > Kenar Boşlukları > 3 cm.',
                        `Bölüm ${i + 1}`,
                        { type: 'margin', sectionIndex: i }
                    );
                    return;
                }
            }

            const tolerance = EBYÜ_RULES.MARGIN_TOLERANCE;
            const expected = EBYÜ_RULES.MARGIN_POINTS;

            let hasError = false;
            let errorDetails = [];

            if (Math.abs(margins.top - expected) > tolerance) {
                hasError = true;
                errorDetails.push(`Üst: ${(margins.top / 28.35).toFixed(2)} cm`);
            }
            if (Math.abs(margins.bottom - expected) > tolerance) {
                hasError = true;
                errorDetails.push(`Alt: ${(margins.bottom / 28.35).toFixed(2)} cm`);
            }
            if (Math.abs(margins.left - expected) > tolerance) {
                hasError = true;
                errorDetails.push(`Sol: ${(margins.left / 28.35).toFixed(2)} cm`);
            }
            if (Math.abs(margins.right - expected) > tolerance) {
                hasError = true;
                errorDetails.push(`Sağ: ${(margins.right / 28.35).toFixed(2)} cm`);
            }

            if (hasError) {
                addResult(
                    'error',
                    'Kenar Boşluğu Hatası',
                    `Kenar boşlukları 3 cm olmalıdır. Mevcut: ${errorDetails.join(', ')}`,
                    `Bölüm ${i + 1}`,
                    { type: 'margin', sectionIndex: i }
                );
            } else {
                addResult(
                    'success',
                    'Kenar Boşlukları',
                    'Tüm kenar boşlukları 3 cm kuralına uygun.',
                    `Bölüm ${i + 1}`
                );
            }
        }
    } catch (error) {
        console.error("Margin check error:", error);
        addResult(
            'warning',
            'Kenar Boşlukları',
            'Kenar boşlukları kontrol edilemedi. Lütfen manuel olarak kontrol edin: Sayfa Düzeni > Kenar Boşlukları > 3 cm.',
            null,
            { type: 'margin' }
        );
    }
}

/**
 * Check cover page formatting
 * EBYÜ Rules:
 * - T.C., University name, Institute, Thesis type must be present
 * - All elements must be centered
 * - Title/Institute name should be 16pt (or 14pt for some elements)
 */
async function checkCoverPage(context, paragraphs) {
    logStep('COVER', 'Kapak sayfası kontrolü başladı');

    // Cover page elements to look for (first ~20 paragraphs)
    const coverElements = {
        tc: false,           // T.C.
        university: false,   // Erzincan Binali Yıldırım Üniversitesi
        institute: false,    // Sosyal Bilimler Enstitüsü / Fen Bilimleri Enstitüsü
        thesisType: false,   // Yüksek Lisans Tezi / Doktora Tezi
        title: false,        // Thesis title
        author: false,       // Hazırlayan / Author name
        advisor: false,      // Danışman / Prof. Dr. etc.
        year: false          // Year (e.g., 2024, 2025)
    };

    let alignmentIssues = 0;
    let coverParagraphCount = 0;

    // Check first 30 paragraphs (cover page area)
    const checkCount = Math.min(paragraphs.items.length, 30);

    for (let i = 0; i < checkCount; i++) {
        const para = paragraphs.items[i];
        const text = (para.text || '').trim();

        if (text === '') continue;
        coverParagraphCount++;

        // Check for cover elements
        if (/^T\.?\s*C\.?$/i.test(text)) coverElements.tc = true;
        if (/ERZİNCAN BİNALİ YILDIRIM/i.test(text) || /ÜNİVERSİTESİ/i.test(text)) coverElements.university = true;
        if (/ENSTİTÜSÜ/i.test(text)) coverElements.institute = true;
        if (/YÜKSEK LİSANS|DOKTORA|TEZİ/i.test(text)) coverElements.thesisType = true;
        if (/HAZIRLAYAN/i.test(text)) coverElements.author = true;
        if (/DANIŞMAN/i.test(text) || /Prof\.|Doç\.|Dr\.|Öğr\./i.test(text)) coverElements.advisor = true;
        if (/^20(2[0-9]|3[0-9])$/.test(text)) coverElements.year = true;

        // Check if cover elements are centered
        if (para.alignment !== Word.Alignment.centered) {
            // Only count if it looks like a cover element
            if (/^T\.?\s*C\.?$|ERZİNCAN|ENSTİTÜSÜ|TEZİ|HAZIRLAYAN|DANIŞMAN/i.test(text)) {
                alignmentIssues++;
            }
        }
    }

    const foundElements = Object.values(coverElements).filter(v => v).length;
    const totalElements = Object.keys(coverElements).length;

    logStep('COVER', `Kontrol tamamlandı`, {
        found: foundElements,
        total: totalElements,
        alignmentIssues
    });

    // Report results
    if (foundElements >= 5) {
        // Cover page detected
        if (alignmentIssues > 0) {
            addResult('warning', 'Kapak Sayfası Hizalama',
                `${alignmentIssues} kapak öğesi ortalanmamış. Kapak sayfasındaki tüm öğeler ortalı olmalıdır.`,
                null, { type: 'coverAlignment' });
        }

        const missingElements = [];
        if (!coverElements.tc) missingElements.push('T.C.');
        if (!coverElements.university) missingElements.push('Üniversite adı');
        if (!coverElements.institute) missingElements.push('Enstitü');
        if (!coverElements.thesisType) missingElements.push('Tez türü');
        if (!coverElements.year) missingElements.push('Yıl');

        if (missingElements.length > 0) {
            addResult('warning', 'Kapak Sayfası Eksikleri',
                `Kapak sayfasında eksik öğeler: ${missingElements.join(', ')}`,
                null, { type: 'coverMissing' });
        } else {
            addResult('success', 'Kapak Sayfası',
                'Kapak sayfası öğeleri tespit edildi ve kontrol edildi.');
        }
    } else {
        addResult('warning', 'Kapak Sayfası',
            'Kapak sayfası tam olarak tespit edilemedi. Lütfen kapak sayfasının EBYÜ kurallarına uygun olduğunu manuel kontrol edin.',
            'Beklenen öğeler: T.C., Üniversite adı, Enstitü, Tez türü, Hazırlayan, Danışman, Yıl',
            { type: 'coverNotFound' });
    }
}

/**
 * Check font settings (BATCH OPTIMIZED + SECTION FILTERING)
 * EBYÜ Rule: Times New Roman, 12pt for body text, 14pt for headings
 * Performance: Single sync call for all paragraphs
 */
async function checkFonts(context, paragraphs) {
    logStep('FONT', 'Yazı tipi kontrolü başladı');

    let nonTNRCount = 0;
    let wrongSizeCount = 0;
    let checkedCount = 0;
    let excludedCount = 0;
    let headingCount = 0;
    const fontErrors = []; // Store first 5 errors for display

    // BATCH LOAD: Load all paragraph properties in ONE sync call
    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("text");
        para.font.load("name, size, bold, allCaps");
    }

    // Single sync for all paragraphs
    await context.sync();

    logStep('FONT', `${paragraphs.items.length} paragraf yüklendi`);

    // Now analyze all paragraphs (no more sync calls needed)
    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const font = para.font;
        const text = para.text || '';

        // Skip empty paragraphs
        if (text.trim() === '') continue;

        // Skip excluded sections (cover, bibliography, appendix, TOC)
        if (shouldExcludeParagraph(text)) {
            excludedCount++;
            continue;
        }

        checkedCount++;

        // Detect if this is a heading
        const isHeading = detectHeading(text, font);
        if (isHeading) {
            headingCount++;
            // Headings should be 14pt, bold, Times New Roman
            if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_HEADING1) > 0.5) {
                // Only report if significantly wrong
                if (Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_HEADING1) > 2) {
                    if (fontErrors.length < 5) {
                        fontErrors.push({
                            index: i,
                            type: 'heading_size',
                            fontName: font.name,
                            fontSize: font.size,
                            text: text.substring(0, 40)
                        });
                    }
                }
            }
            continue; // Skip body text checks for headings
        }

        // Check font name for body text
        if (font.name && font.name !== EBYÜ_RULES.FONT_NAME) {
            nonTNRCount++;
            if (fontErrors.length < 5) {
                fontErrors.push({
                    index: i,
                    type: 'font',
                    fontName: font.name,
                    text: text.substring(0, 40)
                });
            }
        }

        // Check font size for body text (should be 12pt)
        if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_BODY) > 0.5) {
            // Allow 10pt (footnotes), 11pt (tables), 12pt (body), 14pt (headings)
            const allowedSizes = [10, 11, 12, 14, 16];
            const isAllowed = allowedSizes.some(s => Math.abs(font.size - s) < 0.5);
            if (!isAllowed) {
                wrongSizeCount++;
            }
        }
    }

    logStep('FONT', `Kontrol tamamlandı`, {
        checked: checkedCount,
        excluded: excludedCount,
        headings: headingCount,
        fontErrors: nonTNRCount,
        sizeErrors: wrongSizeCount
    });

    // Display font errors with examples
    for (const err of fontErrors) {
        if (err.type === 'font') {
            addResult(
                'error',
                'Yazı Tipi Hatası',
                `"${err.fontName}" yerine "Times New Roman" kullanılmalıdır.`,
                `Paragraf ${err.index + 1}: "${err.text}..."`,
                { type: 'font', paragraphIndex: err.index }
            );
        } else if (err.type === 'heading_size') {
            addResult(
                'warning',
                'Başlık Boyutu',
                `Başlık ${err.fontSize}pt yerine 14pt olmalıdır.`,
                `"${err.text}..."`,
                { type: 'headingSize', paragraphIndex: err.index }
            );
        }
    }

    // Summary for font issues
    if (nonTNRCount > 5) {
        addResult(
            'warning',
            'Çoklu Yazı Tipi Hatası',
            `Toplamda ${nonTNRCount} paragrafta Times New Roman dışında yazı tipi kullanılmış.`,
            null,
            { type: 'fontAll' }
        );
    }

    if (wrongSizeCount > 0) {
        addResult(
            'warning',
            'Yazı Boyutu Uyarısı',
            `${wrongSizeCount} paragrafta standart dışı yazı boyutu tespit edildi.`,
            `İzin verilen: 10pt (dipnot), 11pt (tablo), 12pt (metin), 14pt (başlık)`,
            { type: 'fontSize' }
        );
    }

    if (nonTNRCount === 0 && wrongSizeCount === 0) {
        addResult(
            'success',
            'Yazı Tipi Kontrolü',
            `${checkedCount} paragraf kontrol edildi. Tümü kurallara uygun. (${excludedCount} paragraf hariç tutuldu)`
        );
    }
}

/**
 * Detect if a paragraph is a heading
 * FIXED: Requires minimum 5 characters and stricter bold requirement
 */
function detectHeading(text, font) {
    if (!text) return false;

    const trimmed = text.trim();

    // MINIMUM 5 characters to be a heading (avoid false positives from empty/short text)
    if (trimmed.length < 5) return false;

    // Skip if it's just whitespace or line breaks
    if (/^[\s\r\n]+$/.test(text)) return false;

    // Chapter headings: "BİRİNCİ BÖLÜM", "İKİNCİ BÖLÜM" etc. - these MUST be bold
    const chapterPatterns = [
        /^(BİRİNCİ|İKİNCİ|ÜÇÜNCÜ|DÖRDÜNCÜ|BEŞİNCİ|ALTINCI|YEDİNCİ|SEKİZİNCİ|DOKUZUNCU|ONUNCU)\s*BÖLÜM/i,
        /^(GİRİŞ|SONUÇ|KAYNAKÇA|ÖZET|ABSTRACT|İÇİNDEKİLER|TABLOLAR|ŞEKİLLER|KISALTMALAR|SİMGELER)$/i,
        /^BÖLÜM\s*\d+/i,
        /^ÖN\s*SÖZ$/i,
        /^TEŞEKKÜR$/i
    ];

    for (const pattern of chapterPatterns) {
        if (pattern.test(trimmed)) {
            // Chapter headings must be bold
            return font.bold === true;
        }
    }

    // Section headings: numbered like "1.1.", "2.3.1." etc. - MUST be bold
    if (/^\d+\.\d+\.?\s+\S/.test(trimmed)) {
        return font.bold === true && trimmed.length < 200;
    }

    // Single number heading like "1. Giriş", "2. Yöntem" - MUST be bold
    if (/^\d+\.\s+[A-ZÇĞİÖŞÜa-zçğıöşü]/.test(trimmed)) {
        return font.bold === true && trimmed.length < 150;
    }

    // Letter heading like "A. Başlık", "B. Alt Başlık" - MUST be bold
    if (/^[A-Z]\.\s+\S/.test(trimmed)) {
        return font.bold === true && trimmed.length < 150;
    }

    // All caps + bold + reasonably short = likely a heading
    if (font.bold === true && font.allCaps === true && trimmed.length >= 5 && trimmed.length < 80) {
        return true;
    }

    return false;
}

/**
 * Check heading formatting
 * EBYÜ Rules:
 * - Chapter headings: 14pt, bold, centered, all caps
 * - Section headings (1.1, 1.2): 12pt, bold, left-aligned
 * NOTE: detectHeading now REQUIRES bold, so bold errors should not occur
 */
async function checkHeadings(context, paragraphs) {
    logStep('HEADING', 'Başlık kontrolü başladı');

    let headingCount = 0;
    let sizeErrors = 0;
    let alignmentErrors = 0;
    const headingDetails = [];

    // Load required properties
    for (let i = 0; i < paragraphs.items.length; i++) {
        paragraphs.items[i].load("text, alignment");
        paragraphs.items[i].font.load("size, bold, allCaps");
    }
    await context.sync();

    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const text = para.text || '';
        const font = para.font;

        // AGGRESSIVE FILTERING: Skip empty, whitespace-only, or very short paragraphs
        const trimmed = text.trim();
        if (trimmed.length < 5) continue;
        if (/^[\s\r\n\t]+$/.test(text)) continue;

        // Only count real text paragraphs
        // Skip paragraphs that are just punctuation or numbers
        if (/^[\d\s.,;:!?]+$/.test(trimmed)) continue;

        // Check if this is a heading (detectHeading now requires bold)
        if (!detectHeading(text, font)) continue;

        headingCount++;

        // Determine heading level for size/alignment checks
        const isChapterHeading = /^(BİRİNCİ|İKİNCİ|ÜÇÜNCÜ|DÖRDÜNCÜ|BEŞİNCİ|ALTINCI|YEDİNCİ|SEKİZİNCİ|DOKUZUNCU|ONUNCU)\s*BÖLÜM/i.test(trimmed) ||
            /^(GİRİŞ|SONUÇ|KAYNAKÇA|ÖZET|ABSTRACT|İÇİNDEKİLER|ÖN\s*SÖZ|TEŞEKKÜR)$/i.test(trimmed) ||
            /^BÖLÜM\s*\d+/i.test(trimmed);

        // Check font size for chapter headings (14pt)
        if (isChapterHeading) {
            if (font.size && Math.abs(font.size - 14) > 1) {
                sizeErrors++;
                if (headingDetails.length < 3) {
                    headingDetails.push({ text: trimmed.substring(0, 40), issue: 'size', actual: font.size });
                }
            }
            // Chapter headings should be centered
            if (para.alignment !== Word.Alignment.centered) {
                alignmentErrors++;
            }
        }
    }

    logStep('HEADING', `Kontrol tamamlandı`, { headings: headingCount, sizeErrors, alignmentErrors });

    // Report results - NO LONGER REPORT BOLD ERRORS (detectHeading handles this)
    if (alignmentErrors > 0) {
        addResult('warning', 'Başlık Hizalama',
            `${alignmentErrors} bölüm başlığı ortalanmamış. Ana bölüm başlıkları ortalı olmalıdır.`,
            null, { type: 'headingAlignment' });
    }

    if (sizeErrors > 0) {
        addResult('warning', 'Başlık Boyutu',
            `${sizeErrors} başlıkta yazı boyutu uygun değil. Bölüm başlıkları 14pt olmalıdır.`,
            headingDetails.map(h => `"${h.text}..." (${h.actual}pt)`).join(', '),
            { type: 'headingSize' });
    }

    if (headingCount > 0 && sizeErrors === 0 && alignmentErrors === 0) {
        addResult('success', 'Başlık Kontrolü', `${headingCount} başlık kontrol edildi. Tümü kurallara uygun.`);
    } else if (headingCount === 0) {
        // Don't show warning for no headings - might be intentional
        logStep('HEADING', 'Standart başlık tespit edilemedi');
    }
}

/**
 * Check paragraph formatting (BATCH OPTIMIZED + SECTION FILTERING)
 * EBYÜ Rule: Justified alignment, 1.25cm first line indent for body text
 * Headings can be centered, so they're excluded from justify check
 */
async function checkParagraphFormatting(context, paragraphs) {
    logStep('PARA', 'Paragraf formatı kontrolü başladı');

    let alignmentErrors = 0;
    let indentErrors = 0;
    let checkedCount = 0;
    let excludedCount = 0;

    // BATCH LOAD: Load all paragraph properties in ONE sync call
    for (let i = 0; i < paragraphs.items.length; i++) {
        paragraphs.items[i].load("text, alignment, firstLineIndent");
        paragraphs.items[i].font.load("bold, allCaps");
    }

    // Single sync for all paragraphs
    await context.sync();

    // Now analyze all paragraphs (no more sync calls needed)
    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const text = para.text || '';
        const font = para.font;

        // Skip empty paragraphs
        if (text.trim() === '') continue;

        // Skip excluded sections (cover, bibliography, appendix, TOC)
        if (shouldExcludeParagraph(text)) {
            excludedCount++;
            continue;
        }

        // Skip headings - they may be centered
        if (detectHeading(text, font)) {
            excludedCount++;
            continue;
        }

        // Skip very short paragraphs (likely captions, labels)
        if (text.length < 50) {
            excludedCount++;
            continue;
        }

        checkedCount++;

        // Check alignment - body text should be justified
        if (para.alignment !== Word.Alignment.justified) {
            alignmentErrors++;
        }

        // Check first line indent (only for body paragraphs)
        const expectedIndent = EBYÜ_RULES.FIRST_LINE_INDENT_POINTS;
        if (Math.abs(para.firstLineIndent - expectedIndent) > EBYÜ_RULES.INDENT_TOLERANCE) {
            // Allow 0 indent (block quotes, lists) and expected indent
            if (para.firstLineIndent !== 0) {
                indentErrors++;
            }
        }
    }

    logStep('PARA', `Kontrol tamamlandı`, {
        checked: checkedCount,
        excluded: excludedCount,
        alignmentErrors,
        indentErrors
    });

    if (alignmentErrors > 0) {
        addResult(
            'error',
            'Hizalama Hatası',
            `${alignmentErrors} paragraf iki yana yaslanmamış (Justify).`,
            `Kontrol edilen: ${checkedCount} paragraf (${excludedCount} hariç tutuldu)`,
            { type: 'alignment' }
        );
    } else {
        addResult('success', 'Paragraf Hizalama', `${checkedCount} paragraf kontrol edildi. Tümü iki yana yaslı.`);
    }

    if (indentErrors > 0) {
        addResult(
            'warning',
            'Girinti Uyarısı',
            `${indentErrors} paragrafta ilk satır girintisi 1.25 cm değil.`,
            `Not: Bazı öğeler (alıntı, liste) girinti olmadan da olabilir.`,
            { type: 'indent' }
        );
    }
}

/**
 * Check line spacing (BATCH OPTIMIZED + SECTION FILTERING)
 * EBYÜ Rule: 1.5 line spacing for body text
 * Tables, footnotes, and captions use single spacing
 */
async function checkLineSpacing(context, paragraphs) {
    logStep('SPACING', 'Satır aralığı kontrolü başladı');

    let spacingErrors = 0;
    let checkedCount = 0;
    let excludedCount = 0;

    // BATCH LOAD: Load all paragraph properties in ONE sync call
    for (let i = 0; i < paragraphs.items.length; i++) {
        paragraphs.items[i].load("text, lineSpacing, lineUnitAfter, lineUnitBefore");
        paragraphs.items[i].font.load("bold, allCaps, size");
    }

    // Single sync for all paragraphs
    await context.sync();

    // Now analyze all paragraphs (no more sync calls needed)
    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const text = para.text || '';
        const font = para.font;

        // Skip empty paragraphs
        if (text.trim() === '') continue;

        // Skip excluded sections (cover, bibliography, appendix, TOC)
        if (shouldExcludeParagraph(text)) {
            excludedCount++;
            continue;
        }

        // Skip headings - they may have different spacing
        if (detectHeading(text, font)) {
            excludedCount++;
            continue;
        }

        // Skip short paragraphs (likely captions, labels, table content)
        if (text.length < 50) {
            excludedCount++;
            continue;
        }

        // Skip if font size is not 12pt (likely table content at 11pt or footnotes at 10pt)
        if (font.size && (font.size < 11.5 || font.size > 12.5)) {
            excludedCount++;
            continue;
        }

        checkedCount++;

        // Check line spacing for body text
        // 1.5 line spacing for 12pt font = approximately 18 points
        // Allow some tolerance (17-19 points)
        if (para.lineSpacing) {
            const spacing = para.lineSpacing;
            // Valid 1.5 spacing values: 18 (points), or values between 17-20
            const isValid15 = (spacing >= 17 && spacing <= 20) || spacing === 1.5;
            if (!isValid15) {
                spacingErrors++;
            }
        }
    }

    logStep('SPACING', `Kontrol tamamlandı`, {
        checked: checkedCount,
        excluded: excludedCount,
        errors: spacingErrors
    });

    if (spacingErrors > 0) {
        addResult(
            'error',
            'Satır Aralığı Hatası',
            `${spacingErrors} paragrafta satır aralığı 1.5 satır değil.`,
            `Kontrol edilen: ${checkedCount} paragraf (${excludedCount} hariç tutuldu - tablo, dipnot, başlık vb.)`,
            { type: 'lineSpacing' }
        );
    } else {
        addResult('success', 'Satır Aralığı', `${checkedCount} paragraf kontrol edildi. Tümü 1.5 satır aralığında.`);
    }
}

/**
 * Check images and shapes
 * EBYÜ Rule: Images must be inline (not floating) to prevent layout shifts
 */
async function checkImages(context) {
    try {
        const inlineShapes = context.document.body.inlinePictures;
        inlineShapes.load("items");

        await context.sync();

        const inlineCount = inlineShapes.items.length;

        if (inlineCount > 0) {
            addResult(
                'success',
                'Görsel Kontrolü',
                `${inlineCount} adet satır içi (inline) görsel tespit edildi. Bu görseller kayma sorununa yol açmaz.`
            );
        }

        // Note: Floating shapes are harder to detect with Office.js
        // We add a general warning about floating images
        addResult(
            'warning',
            'Görsel Uyarısı',
            'Kayan (floating) görseller varsa, bunları "Metinle Aynı Hizada" olarak ayarlayın. Kayan görseller sayfa kaymasına neden olabilir.',
            null,
            { type: 'imageWarning' }
        );

    } catch (error) {
        console.log("Image check info:", error.message);
        addResult(
            'warning',
            'Görsel Kontrolü',
            'Görseller kontrol edilemedi. Lütfen görsellerin "Metinle Aynı Hizada" olduğundan emin olun.'
        );
    }
}

/**
 * Check tables for formatting issues (BATCH OPTIMIZED)
 * EBYÜ Rules:
 * - Table captions must be above the table
 * - Tables must fit within page margins
 * Performance: Maximum 5 sync calls regardless of table count
 */
async function checkTables(context) {
    logStep('TABLE', 'Tablo kontrolü başladı');
    const tableStartTime = performance.now();

    try {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        const tableCount = tables.items.length;
        logStep('TABLE', `${tableCount} tablo bulundu`);

        if (tableCount === 0) {
            addResult(
                'success',
                'Tablo Kontrolü',
                'Belgede tablo bulunamadı.'
            );
            return;
        }

        // Batch load all table properties at once
        for (let i = 0; i < tableCount; i++) {
            tables.items[i].load("rowCount, width");
        }
        await context.sync();

        logStep('TABLE', 'Tablo özellikleri yüklendi');

        let validTableCount = 0;
        let widthErrors = 0;
        const pageWidthAvailable = 425; // A4 - 3cm margins = 425pt

        // Analyze tables (no more sync calls needed)
        for (let i = 0; i < tableCount; i++) {
            const table = tables.items[i];
            validTableCount++;

            // Check table width
            if (table.width && table.width > pageWidthAvailable + 10) {
                widthErrors++;
                addResult(
                    'error',
                    'Tablo Genişliği Hatası',
                    `Tablo ${i + 1}: Tablo sayfa kenar boşluklarını aşıyor. Genişlik: ${(table.width / 28.35).toFixed(2)} cm`,
                    `Tablo ${i + 1}`,
                    { type: 'tableWidth', tableIndex: i }
                );
            }
        }

        // Summary
        const tableEndTime = performance.now();
        logStep('TABLE', `Tablo kontrolü tamamlandı`, {
            duration: `${(tableEndTime - tableStartTime).toFixed(0)}ms`,
            tableCount: validTableCount,
            errors: widthErrors
        });

        if (widthErrors === 0) {
            addResult(
                'success',
                'Tablo Kontrolü',
                `${validTableCount} tablo kontrol edildi. Tüm tablolar sayfa sınırlarına uygun.`
            );
        } else {
            addResult(
                'warning',
                'Tablo Kontrolü Tamamlandı',
                `${validTableCount} tablo kontrol edildi. ${widthErrors} genişlik hatası bulundu.`
            );
        }

    } catch (error) {
        console.error("Table check error:", error);
        logStep('TABLE', `Hata: ${error.message}`);
        addResult(
            'warning',
            'Tablo Kontrolü',
            `Tablolar kontrol edilirken bir sorun oluştu: ${error.message}`
        );
    }
}

/**
 * Add reminders for checks that cannot be automated via API
 * These require manual verification by the user
 */
function addManualCheckReminders() {
    logStep('MANUAL', 'Manuel kontrol uyarıları ekleniyor');

    // Footnote reminder (API cannot access footnote content reliably)
    addResult(
        'warning',
        'Dipnot Kontrolü (Manuel)',
        'Dipnotların 10pt, Times New Roman, tek satır aralığı ve iki yana yaslı olduğunu kontrol edin.',
        'Dipnotlar: Ekle → Dipnot → Format ayarları',
        { type: 'manualFootnote' }
    );

    // Page number reminder (API cannot access headers/footers reliably on Mac)
    addResult(
        'warning',
        'Sayfa Numarası Kontrolü (Manuel)',
        'Sayfa numaralarının alt ortada, doğru formatta (Roma/Arap) olduğunu kontrol edin.',
        'Ön kısım: Roma (i, ii, iii...), Ana metin: Arap (1, 2, 3...)',
        { type: 'manualPageNumber' }
    );

    // TOC/Lists reminder
    addResult(
        'warning',
        'İçindekiler/Listeler (Manuel)',
        'İçindekiler, Tablolar Listesi ve Şekiller Listesinin güncel olduğunu kontrol edin.',
        'Sağ tıklayın → "Alanı Güncelle" veya F9 tuşuna basın',
        { type: 'manualTOC' }
    );

    // Figure/Table caption reminder
    addResult(
        'warning',
        'Tablo/Şekil Başlıkları (Manuel)',
        'Tablo başlıkları tablonun ÜSTünde, Şekil başlıkları şeklin ALTında olmalıdır.',
        'Format: "Tablo 1. Açıklama" veya "Şekil 1. Açıklama"',
        { type: 'manualCaptions' }
    );

    logStep('MANUAL', 'Uyarılar eklendi');
}

// ============================================
// FIX FUNCTIONS
// ============================================

/**
 * Fix all fonts to Times New Roman
 */
async function fixAllFonts() {
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.font.name = EBYÜ_RULES.FONT_NAME;
            body.font.size = EBYÜ_RULES.FONT_SIZE_BODY;

            await context.sync();

            showNotification('success', 'Yazı tipleri Times New Roman 12pt olarak düzeltildi.');
        });
    } catch (error) {
        showNotification('error', `Hata: ${error.message}`);
    }
}

/**
 * Fix all margins to 3 cm
 */
async function fixAllMargins() {
    try {
        await Word.run(async (context) => {
            const sections = context.document.sections;
            sections.load("items");

            await context.sync();

            for (let i = 0; i < sections.items.length; i++) {
                const pageSetup = sections.items[i].getPageSetup();
                pageSetup.topMargin = EBYÜ_RULES.MARGIN_POINTS;
                pageSetup.bottomMargin = EBYÜ_RULES.MARGIN_POINTS;
                pageSetup.leftMargin = EBYÜ_RULES.MARGIN_POINTS;
                pageSetup.rightMargin = EBYÜ_RULES.MARGIN_POINTS;
            }

            await context.sync();

            showNotification('success', 'Tüm kenar boşlukları 3 cm olarak düzeltildi.');
        });
    } catch (error) {
        showNotification('error', `Hata: ${error.message}`);
    }
}

/**
 * Fix all line spacing to 1.5
 */
async function fixAllSpacing() {
    try {
        await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");

            await context.sync();

            for (let i = 0; i < paragraphs.items.length; i++) {
                paragraphs.items[i].lineSpacing = 18; // 1.5 line spacing for 12pt font
                paragraphs.items[i].spaceAfter = EBYÜ_RULES.SPACING_AFTER;
                paragraphs.items[i].spaceBefore = EBYÜ_RULES.SPACING_BEFORE;
            }

            await context.sync();

            showNotification('success', 'Tüm satır aralıkları 1.5 satır olarak düzeltildi.');
        });
    } catch (error) {
        showNotification('error', `Hata: ${error.message}`);
    }
}

// ============================================
// UI HELPER FUNCTIONS
// ============================================

function addResult(type, title, description, location = null, fixData = null) {
    validationResults.push({
        type,
        title,
        description,
        location,
        fixData
    });
}

function displayResults() {
    const resultsList = document.getElementById('resultsList');
    const summarySection = document.getElementById('summarySection');
    const filterTabs = document.getElementById('filterTabs');
    const quickFixSection = document.getElementById('quickFixSection');

    // Calculate counts
    const errorCount = validationResults.filter(r => r.type === 'error').length;
    const warningCount = validationResults.filter(r => r.type === 'warning').length;
    const successCount = validationResults.filter(r => r.type === 'success').length;

    // Update summary
    document.getElementById('errorCount').textContent = errorCount;
    document.getElementById('warningCount').textContent = warningCount;
    document.getElementById('successCount').textContent = successCount;

    summarySection.classList.remove('hidden');
    filterTabs.classList.remove('hidden');

    // Show quick fix section if there are errors
    if (errorCount > 0 || warningCount > 0) {
        quickFixSection.classList.remove('hidden');
    } else {
        quickFixSection.classList.add('hidden');
    }

    // Filter and display results
    renderFilteredResults();
}

function renderFilteredResults() {
    const resultsList = document.getElementById('resultsList');

    let filteredResults = validationResults;
    if (currentFilter !== 'all') {
        filteredResults = validationResults.filter(r => r.type === currentFilter);
    }

    if (filteredResults.length === 0) {
        resultsList.innerHTML = `
            <div class="empty-state">
                <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1">
                    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                    <polyline points="22 4 12 14.01 9 11.01"/>
                </svg>
                <p>Bu kategoride sonuç bulunamadı</p>
            </div>
        `;
        return;
    }

    resultsList.innerHTML = filteredResults.map(result => createResultItemHTML(result)).join('');

    // Add event listeners to fix buttons
    resultsList.querySelectorAll('.fix-button').forEach(btn => {
        btn.addEventListener('click', handleFixClick);
    });
}

function createResultItemHTML(result) {
    const iconSVG = getIconSVG(result.type);
    const fixButton = result.fixData ? `
        <button class="fix-button" data-fix-type="${result.fixData.type}">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <path d="M22 4l-8 8-4-4"/>
            </svg>
            Düzelt
        </button>
    ` : '';

    return `
        <div class="result-item ${result.type}">
            <div class="result-header">
                <div class="result-icon">${iconSVG}</div>
                <div class="result-content">
                    <div class="result-title">${result.title}</div>
                    <div class="result-description">${result.description}</div>
                    ${result.location ? `<div class="result-location">📍 ${result.location}</div>` : ''}
                </div>
            </div>
            ${fixButton ? `<div class="result-actions">${fixButton}</div>` : ''}
        </div>
    `;
}

function getIconSVG(type) {
    switch (type) {
        case 'error':
            return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>';
        case 'warning':
            return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 9v4M12 17h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/></svg>';
        case 'success':
            return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>';
        default:
            return '';
    }
}

function handleFixClick(e) {
    const fixType = e.currentTarget.dataset.fixType;

    switch (fixType) {
        case 'margin':
            fixAllMargins();
            break;
        case 'font':
        case 'fontAll':
            fixAllFonts();
            break;
        case 'lineSpacing':
            fixAllSpacing();
            break;
        case 'alignment':
            fixAlignment();
            break;
        default:
            showNotification('warning', 'Bu hata için otomatik düzeltme mevcut değil.');
    }
}

async function fixAlignment() {
    try {
        await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");

            await context.sync();

            for (let i = 0; i < paragraphs.items.length; i++) {
                paragraphs.items[i].alignment = Word.Alignment.justified;
            }

            await context.sync();

            showNotification('success', 'Tüm paragraflar iki yana yaslandı.');
        });
    } catch (error) {
        showNotification('error', `Hata: ${error.message}`);
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
    // Create a simple notification
    const notification = document.createElement('div');
    notification.className = `result-item ${type}`;
    notification.style.cssText = 'position: fixed; top: 20px; left: 50%; transform: translateX(-50%); z-index: 1000; max-width: 300px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);';
    notification.innerHTML = `
        <div class="result-header">
            <div class="result-icon">${getIconSVG(type)}</div>
            <div class="result-content">
                <div class="result-description">${message}</div>
            </div>
        </div>
    `;

    document.body.appendChild(notification);

    setTimeout(() => {
        notification.remove();
    }, 3000);
}

function showError(message) {
    const resultsList = document.getElementById('resultsList');
    resultsList.innerHTML = `
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
