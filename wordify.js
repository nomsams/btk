/**
 * wordify.js - Ultimate Robust Version
 * Includes: Unified Image Fetcher (URLs + Base64), Aspect Ratio lock, 
 * Info/Image section rendering, Table rendering, Page Breaks, and Total Sum.
 */

(function () {
    const TAG = "[Wordify]";
    console.log(`${TAG} 🚀 Script loaded and executing...`);

    if (typeof docx === 'undefined') {
        console.error(`${TAG} ❌ ERROR: docx library not found. Ensure you are using index.umd.js in index.html.`);
        return;
    }

    const { 
        Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
        ImageRun, WidthType, BorderStyle, AlignmentType, VerticalAlign, 
        HeadingLevel 
    } = docx;

    // --- Utility: Clean UI Artifacts ---
    function stripUiArtifacts(htmlStr) {
        if (!htmlStr) return "";
        return String(htmlStr)
            .replace(/✖/g, '')
            .replace(/👨‍🍳/g, '')
            .replace(/📝/g, '')
            .replace(/<span[^>]*class="[^"]*screen-only[^"]*"[^>]*>.*?<\/span>/gi, '')
            .replace(/<br\s*\/?>/gi, '\n')
            .trim();
    }

    // --- Utility: Parse HTML to Word Runs ---
    function parseHtmlToRuns(htmlStr, defaultItalics = false, defaultBold = false) {
        const cleanHtml = stripUiArtifacts(htmlStr);
        if (!cleanHtml) return [new TextRun({ text: "" })];
        
        const parser = new DOMParser();
        const docNode = parser.parseFromString(cleanHtml, 'text/html');
        const runs = [];

        function traverse(node, isBold, isItalic) {
            if (node.nodeType === Node.TEXT_NODE) {
                const text = node.textContent;
                if (!text) return;
                const lines = text.split('\n');
                lines.forEach((line, index) => {
                    const runOpts = { text: line, bold: isBold, italics: isItalic };
                    if (index > 0) runOpts.break = 1;
                    runs.push(new TextRun(runOpts));
                });
            } else if (node.nodeType === Node.ELEMENT_NODE) {
                const tag = node.tagName.toLowerCase();
                const nextBold = isBold || tag === 'b' || tag === 'strong';
                const nextItalic = isItalic || tag === 'i' || tag === 'em';
                Array.from(node.childNodes).forEach(child => traverse(child, nextBold, nextItalic));
            }
        }
        Array.from(docNode.body.childNodes).forEach(child => traverse(child, defaultBold, defaultItalics));
        return runs.length > 0 ? runs : [new TextRun({ text: "" })];
    }

    // --- Utility: Image Dimension Loader ---
    async function getImageDimensions(src) {
        return new Promise(resolve => {
            const img = new Image();
            const timeout = setTimeout(() => resolve({ width: 200, height: 200 }), 3000); 
            img.onload = () => { clearTimeout(timeout); resolve({ width: img.width || 200, height: img.height || 200 }); };
            img.onerror = () => { clearTimeout(timeout); resolve({ width: 200, height: 200 }); };
            img.src = src;
        });
    }

    // --- Utility: Universal Image Fetcher (Handles Base64 AND standard URLs) ---
    async function getDocxImageData(src) {
        if (!src) return null;
        try {
            let uint8, mime = 'png', dimensions;

            if (src.startsWith('data:image')) {
                // Handle Base64
                const parts = src.split(',');
                const match = parts[0].match(/data:image\/(png|jpeg|jpg|gif|bmp)/i);
                mime = match ? match[1].toLowerCase().replace('jpg', 'jpeg') : 'png';
                const binary = window.atob(parts[1]);
                uint8 = new Uint8Array(binary.length);
                for(let i = 0; i < binary.length; i++) uint8[i] = binary.charCodeAt(i);
                dimensions = await getImageDimensions(src);
            } else {
                // Handle standard URL (e.g., "logo.png")
                const response = await fetch(src, { mode: 'cors', cache: 'no-cache' });
                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                const blob = await response.blob();
                const arrayBuffer = await blob.arrayBuffer();
                uint8 = new Uint8Array(arrayBuffer);
                
                const match = blob.type.match(/image\/(png|jpeg|jpg|gif|bmp)/i);
                mime = match ? match[1].toLowerCase().replace('jpg', 'jpeg') : 'png';
                
                const objUrl = URL.createObjectURL(blob);
                dimensions = await getImageDimensions(objUrl);
                URL.revokeObjectURL(objUrl);
            }
            return { uint8, mime, width: dimensions.width, height: dimensions.height };
        } catch (e) {
            console.warn(`${TAG} Failed to process image: ${src.substring(0, 40)}...`, e);
            return null; // Fail gracefully so export doesn't crash
        }
    }

    // --- Utility: Safe Decryption ---
    function decryptCaesar(text, shift) {
        const mangleMap = { '€': 128, '‚': 130, 'ƒ': 131, '„': 132, '…': 133, '†': 134, '‡': 135, 'ˆ': 136, '‰': 137, 'Š': 138, '‹': 139, 'Œ': 140, 'Ž': 142, '‘': 145, '’': 146, '“': 147, '”': 148, '•': 149, '–': 150, '—': 151, '˜': 152, '™': 153, 'š': 154, '›': 155, 'œ': 156, 'ž': 158, 'Ÿ': 159 };
        return text.split('').map(char => {
            let code = char.charCodeAt(0);
            if (mangleMap[char] !== undefined) code = mangleMap[char];
            return String.fromCharCode(code - shift);
        }).join('');
    }

    // --- Layout Definitions ---
    const INVISIBLE_BORDERS = { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 }, insideHorizontal: { style: BorderStyle.NONE, size: 0 }, insideVertical: { style: BorderStyle.NONE, size: 0 } };
    const LIGHT_BORDER = { style: BorderStyle.SINGLE, size: 1, color: "E0E0E0" };
    const TABLE_BORDERS = { top: LIGHT_BORDER, bottom: LIGHT_BORDER, insideHorizontal: LIGHT_BORDER, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 }, insideVertical: { style: BorderStyle.NONE, size: 0 } };
    
    // Explicit Widths for Google Docs Compatibility
    const COL_WIDTHS_DXA = [900, 4950, 1350, 1800]; // 9000 twips (standard A4)
    const HALF_WIDTH_DXA = [4500, 4500];

    async function generateWordDocument() {
        try {
            console.log(`${TAG} 🎬 EXPORT STARTED`);

            // Fetch Data
            let data = window.jsonData || (typeof jsonData !== 'undefined' ? jsonData : null);
            if (!data || !data.quote) {
                console.warn(`${TAG} Global jsonData missing. Trying LocalStorage...`);
                const localData = localStorage.getItem('quoteData');
                if (localData) data = JSON.parse(decryptCaesar(localData, 17));
            }
            if (!data || !data.quote) {
                alert("Export misslyckades: Ingen data hittades.");
                return;
            }

            const lang = data.quote.language || 'sv';
            const t = (typeof translations !== 'undefined' && translations[lang]) || {};
            const labels = data.quote.labels || {};
            const visibility = data.quote.visibility || { optional: true, info: true, terms: true };
            const docChildren = [];
            const emptyParagraph = () => new Paragraph({ children: [new TextRun("")] });

            // 1. Process Logo
            let logoSrc = localStorage.getItem('companyLogo');
            if (!logoSrc) {
                // Fallback to DOM img source if no local storage (e.g., default 'logo.png')
                const domLogo = document.getElementById('companyLogo');
                if (domLogo && domLogo.getAttribute('src')) logoSrc = domLogo.src;
            }

            const logoRuns = [];
            if (logoSrc) {
                const imgData = await getDocxImageData(logoSrc);
                if (imgData) {
                    const targetWidth = data.quote.defaultLogoWidth || 200;
                    const targetHeight = Math.round((targetWidth / imgData.width) * imgData.height);
                    logoRuns.push(new ImageRun({ 
                        data: imgData.uint8, 
                        transformation: { width: Math.round(targetWidth), height: targetHeight },
                        type: imgData.mime 
                    }));
                }
            }

            docChildren.push(new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                columnWidths: HALF_WIDTH_DXA,
                borders: INVISIBLE_BORDERS,
                rows: [new TableRow({
                    children: [
                        new TableCell({ 
                            width: { size: 50, type: WidthType.PERCENTAGE }, 
                            children: [new Paragraph({ children: logoRuns.length > 0 ? logoRuns : [new TextRun("")] })], 
                            verticalAlign: VerticalAlign.CENTER 
                        }),
                        new TableCell({ 
                            width: { size: 50, type: WidthType.PERCENTAGE }, 
                            children: [
                                new Paragraph({ children: [new TextRun({ text: stripUiArtifacts(labels.quoteTitle || t.quoteTitle || "Offert"), size: 32, bold: true })], heading: HeadingLevel.HEADING_1, alignment: AlignmentType.RIGHT }),
                                new Paragraph({ children: [new TextRun({ text: `${stripUiArtifacts(t.quoteNumberLabel || "Nr:")} ${stripUiArtifacts(data.quote.quoteNumber || "")}` })], alignment: AlignmentType.RIGHT }),
                                new Paragraph({ children: [new TextRun({ text: `${stripUiArtifacts(t.dateLabel || "Datum:")} ${stripUiArtifacts(data.quote.date || "")}` })], alignment: AlignmentType.RIGHT })
                            ], 
                            verticalAlign: VerticalAlign.CENTER 
                        })
                    ]
                })]
            }));

            docChildren.push(emptyParagraph());

            // 2. Addresses
            const addressLines = (comp) => Object.values(comp || {}).filter(v => !!v).map(l => new Paragraph({ children: parseHtmlToRuns(l) }));
            const compA = addressLines(data.companyA);
            const compB = addressLines(data.companyB);

            docChildren.push(new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                columnWidths: HALF_WIDTH_DXA,
                borders: INVISIBLE_BORDERS,
                rows: [new TableRow({
                    children: [
                        new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: stripUiArtifacts(t.toLabel || "Till:"), bold: true })] }), ...(compA.length ? compA : [emptyParagraph()])] }),
                        new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: stripUiArtifacts(t.fromLabel || "Från:"), bold: true })] }), ...(compB.length ? compB : [emptyParagraph()])] })
                    ]
                })]
            }));

            docChildren.push(emptyParagraph());

            const safeFormatPrice = (val) => (typeof formatPrice === 'function' ? formatPrice(val) : String(val || 0));
            
            // 3. Build Items Table
            const buildTable = (items, includeTotal = false) => {
                if (!items || items.length === 0) return null;
                const rows = [new TableRow({
                    tableHeader: true,
                    children: [
                        new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: "Nr", bold: true })] })] }),
                        new TableCell({ width: { size: 55, type: WidthType.PERCENTAGE }, shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: "Namn", bold: true })] })] }),
                        new TableCell({ width: { size: 15, type: WidthType.PERCENTAGE }, shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: "Antal", bold: true })], alignment: AlignmentType.CENTER })] }),
                        new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: "Pris", bold: true })], alignment: AlignmentType.RIGHT })] })
                    ]
                })];

                let totalSum = 0;

                items.forEach(item => {
                    if (item.type === 'separator') return;
                    if (item.isHiddenFromPrint) return;
                    
                    totalSum += item.targetPrice || 0;

                    const descRuns = item.itemDescription ? parseHtmlToRuns(item.itemDescription) : [];
                    const cellChildren = [new Paragraph({ children: [new TextRun({ text: stripUiArtifacts(item.name || ""), bold: true })] })];
                    if (descRuns.length > 0) cellChildren.push(new Paragraph({ children: descRuns }));

                    rows.push(new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ text: String(stripUiArtifacts(item.itemNumber || "")) })] }),
                            new TableCell({ children: cellChildren }),
                            new TableCell({ children: [new Paragraph({ text: String(stripUiArtifacts(item.quantity || "")), alignment: AlignmentType.CENTER })] }),
                            new TableCell({ children: [new Paragraph({ text: item.isPriceBakedIn ? "" : safeFormatPrice(item.targetPrice), alignment: AlignmentType.RIGHT })] })
                        ]
                    }));

                    if (item.subItems && item.subItems.length > 0) {
                        item.subItems.forEach(sub => {
                            if (sub.isHiddenFromPrint) return;
                            if (!sub.isPriceBakedIn) totalSum += sub.subItemTargetPrice || 0;
                            
                            const subDescRuns = sub.subItemDescription ? parseHtmlToRuns(sub.subItemDescription) : [];
                            const subCellChildren = [new Paragraph({ children: [new TextRun({ text: stripUiArtifacts(sub.subItemName || "") })] })];
                            if (subDescRuns.length > 0) subCellChildren.push(new Paragraph({ children: subDescRuns }));

                            rows.push(new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: String(stripUiArtifacts(sub.subItemNumber || "")) })] }),
                                    new TableCell({ children: subCellChildren }),
                                    new TableCell({ children: [new Paragraph({ text: String(stripUiArtifacts(sub.subItemQuantity || "")), alignment: AlignmentType.CENTER })] }),
                                    new TableCell({ children: [new Paragraph({ text: sub.isPriceBakedIn ? "" : safeFormatPrice(sub.subItemTargetPrice), alignment: AlignmentType.RIGHT })] })
                                ]
                            }));
                        });
                    }
                });

                // Add Total Row if requested
                if (includeTotal && !data.quote.removeTotal) {
                    const currencyLabel = data.quote.useCustomCurrency ? data.quote.customCurrency : "SEK";
                    rows.push(new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("")] }),
                            new TableCell({ children: [new Paragraph("")] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Total:", bold: true })], alignment: AlignmentType.RIGHT })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${safeFormatPrice(totalSum)} ${currencyLabel}`, bold: true })], alignment: AlignmentType.RIGHT })] })
                        ]
                    }));
                }

                return new Table({ 
                    width: { size: 100, type: WidthType.PERCENTAGE }, 
                    columnWidths: COL_WIDTHS_DXA, 
                    borders: TABLE_BORDERS, 
                    rows 
                });
            };

            const mainTbl = buildTable(data.items, true);
            if (mainTbl) docChildren.push(mainTbl);

            // 4. Optional Items
            if (visibility.optional && data.optionalItems && data.optionalItems.length > 0) {
                docChildren.push(emptyParagraph());
                docChildren.push(new Paragraph({ children: [new TextRun({ text: stripUiArtifacts(labels.optionalItemsHeading || t.optionalItemsHeading || "Alternativ"), bold: true, size: 28 })], heading: HeadingLevel.HEADING_2 }));
                docChildren.push(emptyParagraph());
                const optTbl = buildTable(data.optionalItems, false);
                if (optTbl) docChildren.push(optTbl);
            }

            // 5. Info / Images Section
            if (visibility.info && data.infoImages && data.infoImages.length > 0) {
                docChildren.push(emptyParagraph());
                docChildren.push(new Paragraph({ children: [new TextRun({ text: stripUiArtifacts(labels.infoImagesHeading || t.infoImagesHeading || "Info / Bilder"), bold: true, size: 28 })], heading: HeadingLevel.HEADING_2 }));
                docChildren.push(emptyParagraph());

                for (const infoItem of data.infoImages) {
                    if (infoItem.type === 'page-break') {
                        docChildren.push(new Paragraph({ pageBreakBefore: true }));
                    } 
                    else if (infoItem.type === 'text') {
                        let alignment = AlignmentType.LEFT;
                        if (infoItem.centering === 'center') alignment = AlignmentType.CENTER;
                        
                        docChildren.push(new Paragraph({ 
                            children: parseHtmlToRuns(infoItem.content),
                            alignment: alignment
                        }));
                    } 
                    else if (infoItem.type === 'image') {
                        const imgData = await getDocxImageData(infoItem.src);
                        if (imgData) {
                            // Scale dimensions safely. Cap max width to 600px to prevent going off-page.
                            let targetW = parseFloat(infoItem.width) || imgData.width;
                            if (targetW > 600) targetW = 600; 
                            const targetH = Math.round((targetW / imgData.width) * imgData.height);

                            let alignment = AlignmentType.CENTER;
                            if (infoItem.centering === 'left') alignment = AlignmentType.LEFT;

                            docChildren.push(new Paragraph({
                                alignment: alignment,
                                children: [
                                    new ImageRun({
                                        data: imgData.uint8,
                                        transformation: { width: targetW, height: targetH },
                                        type: imgData.mime
                                    })
                                ]
                            }));
                        }
                    }
                    else if (infoItem.type === 'table') {
                        const tRows = [];
                        for (let r = 0; r < infoItem.rows; r++) {
                            const cells = [];
                            for (let c = 0; c < infoItem.cols; c++) {
                                const cellData = (infoItem.data && infoItem.data[r]) ? infoItem.data[r][c] : "";
                                cells.push(new TableCell({ children: [new Paragraph({ children: parseHtmlToRuns(cellData) })] }));
                            }
                            tRows.push(new TableRow({ children: cells }));
                        }
                        docChildren.push(new Table({ 
                            rows: tRows, 
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            borders: TABLE_BORDERS 
                        }));
                    }
                    docChildren.push(emptyParagraph());
                }
            }

            // 6. Terms
            if (visibility.terms && data.terms && data.terms.length > 0) {
                docChildren.push(emptyParagraph());
                docChildren.push(new Paragraph({ children: [new TextRun({ text: stripUiArtifacts(labels.termsHeading || t.termsHeading || "Villkor"), bold: true, size: 28 })], heading: HeadingLevel.HEADING_2 }));
                data.terms.forEach(term => {
                    docChildren.push(new Paragraph({ children: parseHtmlToRuns(term) }));
                });
            }

            // --- Save / Export ---
            const doc = new Document({ sections: [{ children: docChildren }] });
            const blob = await Packer.toBlob(doc);
            const fileName = `Offert_${stripUiArtifacts(data.quote.quoteNumber) || "Draft"}.docx`;
            
            if (typeof saveAs !== 'undefined') {
                saveAs(blob, fileName);
            } else {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = fileName; a.click();
                URL.revokeObjectURL(url);
            }
            console.log(`${TAG} ✅ Export Success`);

        } catch (err) {
            console.error(`${TAG} ❌ Detailed Error:`, err);
            alert("Ett fel uppstod vid skapandet av filen. Kontrollera webbläsarens konsol (F12) för detaljer.");
        }
    }

    // --- Binding ---
    function attach() {
        const btn = document.getElementById('exportWordBtn');
        if (btn) {
            btn.removeEventListener('click', generateWordDocument);
            btn.addEventListener('click', generateWordDocument);
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', attach);
    } else {
        attach();
    }
})();
