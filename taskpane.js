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

    // Quick fix buttons
    document.getElementById('fixAllFontsBtn').addEventListener('click', fixAllFonts);
    document.getElementById('fixAllMarginsBtn').addEventListener('click', fixAllMargins);
    document.getElementById('fixAllSpacingBtn').addEventListener('click', fixAllSpacing);

    // Filter tabs
    document.querySelectorAll('.filter-tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            const filter = e.target.dataset.filter;
            setActiveFilter(filter);
        });
    });
}

// ============================================
// MAIN SCAN FUNCTION
// ============================================

async function scanDocument() {
    const scanBtn = document.getElementById('scanBtn');
    scanBtn.disabled = true;
    scanBtn.innerHTML = '<span>Taranıyor...</span>';

    // Reset results
    validationResults = [];

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

            // Step 1: Check Margins (20%)
            updateProgress(10, "Kenar boşlukları kontrol ediliyor...");
            await checkMargins(context, sections);
            await context.sync();

            // Step 2: Check Fonts (40%)
            updateProgress(30, "Yazı tipleri kontrol ediliyor...");
            await checkFonts(context, paragraphs);
            await context.sync();

            // Step 3: Check Paragraph Formatting (60%)
            updateProgress(50, "Paragraf formatları kontrol ediliyor...");
            await checkParagraphFormatting(context, paragraphs);
            await context.sync();

            // Step 4: Check Line Spacing (80%)
            updateProgress(70, "Satır aralıkları kontrol ediliyor...");
            await checkLineSpacing(context, paragraphs);
            await context.sync();

            // Step 5: Check Images (85%)
            updateProgress(80, "Görseller kontrol ediliyor...");
            await checkImages(context);
            await context.sync();

            // Step 6: Check Tables (100%)
            updateProgress(90, "Tablolar kontrol ediliyor...");
            await checkTables(context);
            await context.sync();

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
    for (let i = 0; i < sections.items.length; i++) {
        const section = sections.items[i];
        const pageSetup = section.getPageSetup();

        pageSetup.load("topMargin, bottomMargin, leftMargin, rightMargin");
        await context.sync();

        const margins = {
            top: pageSetup.topMargin,
            bottom: pageSetup.bottomMargin,
            left: pageSetup.leftMargin,
            right: pageSetup.rightMargin
        };

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
}

/**
 * Check font settings
 * EBYÜ Rule: Times New Roman, 12pt for body text
 */
async function checkFonts(context, paragraphs) {
    let nonTNRCount = 0;
    let wrongSizeCount = 0;
    let checkedCount = 0;

    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const font = para.font;
        font.load("name, size, bold");

        await context.sync();

        // Skip empty paragraphs
        para.load("text");
        await context.sync();

        if (!para.text || para.text.trim() === '') continue;

        checkedCount++;

        // Check font name
        if (font.name && font.name !== EBYÜ_RULES.FONT_NAME) {
            nonTNRCount++;
            if (nonTNRCount <= 3) { // Only show first 3 individual errors
                addResult(
                    'error',
                    'Yazı Tipi Hatası',
                    `"${font.name}" yerine "Times New Roman" kullanılmalıdır.`,
                    `Paragraf ${i + 1}: "${para.text.substring(0, 50)}..."`,
                    { type: 'font', paragraphIndex: i }
                );
            }
        }

        // Check font size (only for non-heading paragraphs)
        const expectedSize = font.bold ? EBYÜ_RULES.FONT_SIZE_HEADING1 : EBYÜ_RULES.FONT_SIZE_BODY;
        if (font.size && Math.abs(font.size - expectedSize) > 0.5) {
            // Skip if it might be a heading (bold + 14pt)
            if (!(font.bold && Math.abs(font.size - 14) < 0.5)) {
                wrongSizeCount++;
            }
        }
    }

    // Summary for font issues
    if (nonTNRCount > 3) {
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
            `${wrongSizeCount} paragrafta standart dışı yazı boyutu tespit edildi. (Standart: 12pt, Başlık: 14pt)`,
            null,
            { type: 'fontSize' }
        );
    }

    if (nonTNRCount === 0 && wrongSizeCount === 0) {
        addResult(
            'success',
            'Yazı Tipi Kontrolü',
            `Tüm paragraflar (${checkedCount} adet) Times New Roman kuralına uygun.`
        );
    }
}

/**
 * Check paragraph formatting
 * EBYÜ Rule: Justified alignment, 1.25cm first line indent
 */
async function checkParagraphFormatting(context, paragraphs) {
    let alignmentErrors = 0;
    let indentErrors = 0;

    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("text, alignment, firstLineIndent");

        await context.sync();

        // Skip empty paragraphs
        if (!para.text || para.text.trim() === '') continue;

        // Skip short paragraphs (likely headings)
        if (para.text.length < 100) continue;

        // Check alignment
        if (para.alignment !== Word.Alignment.justified) {
            alignmentErrors++;
        }

        // Check first line indent
        const expectedIndent = EBYÜ_RULES.FIRST_LINE_INDENT_POINTS;
        if (Math.abs(para.firstLineIndent - expectedIndent) > EBYÜ_RULES.INDENT_TOLERANCE) {
            indentErrors++;
        }
    }

    if (alignmentErrors > 0) {
        addResult(
            'error',
            'Hizalama Hatası',
            `${alignmentErrors} paragraf iki yana yaslanmamış (Justify). Tüm metin paragrafları iki yana yaslanmalıdır.`,
            null,
            { type: 'alignment' }
        );
    } else {
        addResult('success', 'Paragraf Hizalama', 'Tüm paragraflar doğru hizalanmış (İki Yana Yaslı).');
    }

    if (indentErrors > 0) {
        addResult(
            'warning',
            'Girinti Uyarısı',
            `${indentErrors} paragrafta ilk satır girintisi 1.25 cm değil.`,
            null,
            { type: 'indent' }
        );
    }
}

/**
 * Check line spacing
 * EBYÜ Rule: 1.5 line spacing for body text
 */
async function checkLineSpacing(context, paragraphs) {
    let spacingErrors = 0;

    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("text, lineSpacing, lineUnitAfter, lineUnitBefore");

        await context.sync();

        // Skip empty paragraphs
        if (!para.text || para.text.trim() === '') continue;

        // Skip short paragraphs (likely headings)
        if (para.text.length < 100) continue;

        // Check line spacing (1.5 = approx 18 points for 12pt font, or lineSpacing value of 1.5)
        // Word uses different units, so we check for common 1.5 line spacing value
        if (para.lineSpacing && para.lineSpacing !== 18 && para.lineSpacing !== 1.5) {
            spacingErrors++;
        }
    }

    if (spacingErrors > 0) {
        addResult(
            'error',
            'Satır Aralığı Hatası',
            `${spacingErrors} paragrafta satır aralığı 1.5 satır değil. Tez metinlerinde 1.5 satır aralığı kullanılmalıdır.`,
            null,
            { type: 'lineSpacing' }
        );
    } else {
        addResult('success', 'Satır Aralığı', 'Satır aralıkları kontrol edildi.');
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
 * Check tables for formatting issues
 * EBYÜ Rules:
 * 1. Hidden/Layout Tables: Tables with no visible borders should be flagged
 * 2. Data Table Caption: Must be ABOVE the table
 * 3. Table Font Size: Must be 11pt (smaller than body)
 * 4. Table Line Spacing: Must be single (1.0)
 * 5. Table Width: Must fit within page margins
 */
async function checkTables(context) {
    try {
        const tables = context.document.body.tables;
        tables.load("items");

        await context.sync();

        if (tables.items.length === 0) {
            addResult(
                'success',
                'Tablo Kontrolü',
                'Belgede tablo bulunamadı.'
            );
            return;
        }

        let hiddenTableCount = 0;
        let captionErrors = 0;
        let fontSizeErrors = 0;
        let spacingErrors = 0;
        let widthErrors = 0;
        let validTableCount = 0;

        // Get page width for table width validation
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();

        let pageWidth = 595; // Default A4 width in points
        if (sections.items.length > 0) {
            const pageSetup = sections.items[0].getPageSetup();
            pageSetup.load("pageWidth, leftMargin, rightMargin");
            await context.sync();
            pageWidth = pageSetup.pageWidth - pageSetup.leftMargin - pageSetup.rightMargin;
        }

        for (let i = 0; i < tables.items.length; i++) {
            const table = tables.items[i];

            // Load table properties
            table.load("rowCount, width, values");

            await context.sync();

            // ================================================
            // CHECK 1: Detect Hidden/Layout Tables (No Borders)
            // ================================================
            let isHiddenTable = false;

            try {
                // Try to check border visibility
                // A table with all borders set to "None" is likely a layout table
                const firstRow = table.rows.getFirst();
                const firstCell = firstRow.cells.getFirst();
                firstCell.load("body");
                await context.sync();

                // Check if table has minimal content (likely layout table)
                const cellCount = table.rowCount;
                const tableValues = table.values;

                // Heuristic: If table has very sparse content or single row, might be layout
                if (cellCount === 1 && tableValues && tableValues[0]) {
                    const cellText = tableValues[0].join('').trim();
                    if (cellText.length > 200) {
                        // Single row with lots of text = likely layout table
                        isHiddenTable = true;
                    }
                }

                // Additional check: Try to access border properties
                // Note: Office.js has limited border access, using heuristics instead
                const borderCheck = table.getBorder(Word.BorderLocation.top);
                borderCheck.load("type, color, width");
                await context.sync();

                // If border type is "None" or width is 0, it's likely hidden
                if (borderCheck.type === Word.BorderType.none ||
                    borderCheck.type === "None" ||
                    borderCheck.width === 0) {
                    isHiddenTable = true;
                }

            } catch (borderError) {
                // If we can't check borders, use content-based heuristics
                console.log("Border check not available, using heuristics");
            }

            if (isHiddenTable) {
                hiddenTableCount++;
                addResult(
                    'warning',
                    'Gizli/Düzen Tablosu Tespit Edildi',
                    `Tablo ${i + 1}: Kenarlıksız tablo tespit edildi. Metin düzeni için tablo yerine "Sütunlar" özelliğini kullanmanız önerilir.`,
                    `Tablo ${i + 1}`,
                    { type: 'hiddenTable', tableIndex: i }
                );
                continue; // Skip other checks for layout tables
            }

            validTableCount++;

            // ================================================
            // CHECK 2: Caption Position (Must be ABOVE table)
            // ================================================
            try {
                // Get the paragraph before the table
                const tableRange = table.getRange();
                tableRange.load("text");
                await context.sync();

                // Check paragraph immediately before table
                const paragraphBefore = table.getRange().expandTo(table.getRange()).paragraphs;
                paragraphBefore.load("items");
                await context.sync();

                // Look for caption pattern: "Tablo X." or "Çizelge X."
                let hasCaption = false;
                const body = context.document.body;
                const allParagraphs = body.paragraphs;
                allParagraphs.load("items");
                await context.sync();

                // Find table position in document and check preceding paragraph
                for (let j = 0; j < allParagraphs.items.length; j++) {
                    const para = allParagraphs.items[j];
                    para.load("text");
                    await context.sync();

                    const text = para.text.trim().toLowerCase();
                    // Check for Turkish table caption patterns
                    if (text.match(/^(tablo|çizelge|table)\s*\d+/i) ||
                        text.match(/^(tablo|çizelge|table)\s*[.:]/i)) {
                        hasCaption = true;
                        break;
                    }
                }

                if (!hasCaption) {
                    captionErrors++;
                    addResult(
                        'error',
                        'Tablo Başlığı Hatası',
                        `Tablo ${i + 1}: Tablo başlığı (caption) bulunamadı veya tablonun üstünde değil. Tablo başlıkları tablonun ÜSTünde olmalıdır.`,
                        `Tablo ${i + 1}`,
                        { type: 'tableCaption', tableIndex: i }
                    );
                }

            } catch (captionError) {
                console.log("Caption check error:", captionError);
            }

            // ================================================
            // CHECK 3: Font Size Inside Table (Must be 11pt)
            // ================================================
            try {
                const rows = table.rows;
                rows.load("items");
                await context.sync();

                let wrongFontSize = false;

                for (let r = 0; r < Math.min(rows.items.length, 3); r++) {
                    const row = rows.items[r];
                    const cells = row.cells;
                    cells.load("items");
                    await context.sync();

                    for (let c = 0; c < Math.min(cells.items.length, 3); c++) {
                        const cell = cells.items[c];
                        const cellBody = cell.body;
                        const cellParagraphs = cellBody.paragraphs;
                        cellParagraphs.load("items");
                        await context.sync();

                        for (let p = 0; p < cellParagraphs.items.length; p++) {
                            const para = cellParagraphs.items[p];
                            const font = para.font;
                            font.load("size");
                            await context.sync();

                            // Check if font size is NOT 11pt (with tolerance)
                            if (font.size && Math.abs(font.size - EBYÜ_RULES.FONT_SIZE_TABLE) > 0.5) {
                                wrongFontSize = true;
                                break;
                            }
                        }
                        if (wrongFontSize) break;
                    }
                    if (wrongFontSize) break;
                }

                if (wrongFontSize) {
                    fontSizeErrors++;
                    addResult(
                        'error',
                        'Tablo Yazı Boyutu Hatası',
                        `Tablo ${i + 1}: Tablo içindeki metin 11 punto olmalıdır (ana metin 12pt'den küçük).`,
                        `Tablo ${i + 1}`,
                        { type: 'tableFontSize', tableIndex: i }
                    );
                }

            } catch (fontError) {
                console.log("Table font check error:", fontError);
            }

            // ================================================
            // CHECK 4: Line Spacing Inside Table (Must be Single/1.0)
            // ================================================
            try {
                const rows = table.rows;
                rows.load("items");
                await context.sync();

                let wrongSpacing = false;

                for (let r = 0; r < Math.min(rows.items.length, 2); r++) {
                    const row = rows.items[r];
                    const cells = row.cells;
                    cells.load("items");
                    await context.sync();

                    for (let c = 0; c < Math.min(cells.items.length, 2); c++) {
                        const cell = cells.items[c];
                        const cellBody = cell.body;
                        const cellParagraphs = cellBody.paragraphs;
                        cellParagraphs.load("items");
                        await context.sync();

                        for (let p = 0; p < cellParagraphs.items.length; p++) {
                            const para = cellParagraphs.items[p];
                            para.load("lineSpacing");
                            await context.sync();

                            // Single spacing for 11pt font = ~11-13 points
                            // 1.5 spacing = ~16-18 points
                            if (para.lineSpacing && para.lineSpacing > 14) {
                                wrongSpacing = true;
                                break;
                            }
                        }
                        if (wrongSpacing) break;
                    }
                    if (wrongSpacing) break;
                }

                if (wrongSpacing) {
                    spacingErrors++;
                    addResult(
                        'warning',
                        'Tablo Satır Aralığı Uyarısı',
                        `Tablo ${i + 1}: Tablo içindeki satır aralığı tek (1.0) olmalıdır.`,
                        `Tablo ${i + 1}`,
                        { type: 'tableSpacing', tableIndex: i }
                    );
                }

            } catch (spacingError) {
                console.log("Table spacing check error:", spacingError);
            }

            // ================================================
            // CHECK 5: Table Width (Must Fit Within Margins)
            // ================================================
            try {
                if (table.width && table.width > pageWidth + 10) {
                    widthErrors++;
                    addResult(
                        'error',
                        'Tablo Genişliği Hatası',
                        `Tablo ${i + 1}: Tablo sayfa kenar boşluklarını aşıyor. Tablo genişliği: ${(table.width / 28.35).toFixed(2)} cm, Mevcut alan: ${(pageWidth / 28.35).toFixed(2)} cm`,
                        `Tablo ${i + 1}`,
                        { type: 'tableWidth', tableIndex: i }
                    );
                }
            } catch (widthError) {
                console.log("Table width check error:", widthError);
            }
        }

        // ================================================
        // SUMMARY RESULTS
        // ================================================

        if (hiddenTableCount > 0) {
            addResult(
                'warning',
                'Düzen Tabloları Özeti',
                `Toplam ${hiddenTableCount} adet gizli/düzen tablosu tespit edildi. Metin düzeni için "Sütunlar" özelliğini kullanmanız önerilir.`,
                null,
                { type: 'hiddenTableSummary' }
            );
        }

        const totalErrors = captionErrors + fontSizeErrors + widthErrors;
        const totalWarnings = spacingErrors;

        if (totalErrors === 0 && totalWarnings === 0 && validTableCount > 0) {
            addResult(
                'success',
                'Tablo Formatı',
                `${validTableCount} adet veri tablosu kontrol edildi. Tüm tablolar EBYÜ kurallarına uygun.`
            );
        } else if (validTableCount > 0) {
            addResult(
                'warning',
                'Tablo Kontrolü Tamamlandı',
                `${tables.items.length} tablo kontrol edildi. ${totalErrors} hata, ${totalWarnings} uyarı bulundu.`
            );
        }

    } catch (error) {
        console.error("Table check error:", error);
        addResult(
            'warning',
            'Tablo Kontrolü',
            `Tablolar kontrol edilirken bir sorun oluştu: ${error.message}`
        );
    }
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
