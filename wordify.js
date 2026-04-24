/**
 * wordify.js - Robust, Failsafe, and Aligned Word Document Exporter
 * Requires docx and FileSaver.js to be loaded in the global scope.
 */

(function () {
    // --- 1. Library Extraction & Failsafe Check ---
    if (typeof docx === 'undefined') {
        console.error("Error: docx library is not loaded. Please ensure it is included via CDN.");
        return;
    }
    
    const { 
        Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
        ImageRun, WidthType, BorderStyle, AlignmentType, VerticalAlign, 
        HeadingLevel 
    } = docx;

    // --- 2. Robust Helpers ---

    // Safely convert Data URL/Base64 to ArrayBuffer
    function base64ToArrayBuffer(dataUrl) {
        try {
            const base64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
            const binaryString = window.atob(base64);
            const len = binaryString.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            return bytes.buffer;
        } catch (e) {
            console.error("Failed to decode base64 image data.", e);
            return new ArrayBuffer(0); // Return empty buffer to prevent total crash
        }
    }

    // Advanced HTML to TextRun parser (preserves Bold, Italics, and Newlines)
    function parseHtmlToRuns(htmlStr, defaultItalics = false, defaultBold = false) {
        if (!htmlStr) return [new TextRun({ text: "", italics: defaultItalics, bold: defaultBold })];
        
        const cleanHtml = String(htmlStr).replace(/<br\s*\/?>/gi, '\n');
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
                    if (index > 0) runOpts.break = 1; // Strict docx v8 API
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
        return runs.length > 0 ? runs : [new TextRun({ text: "", italics: defaultItalics, bold: defaultBold })];
    }

    // Get image dimensions safely with a timeout
    async function getImageDimensions(dataUrl) {
        return new Promise(resolve => {
            const img = new Image();
            const timeout = setTimeout(() => resolve({ width: 200, height: 200 }), 3000); 
            
            img.onload = () => {
                clearTimeout(timeout);
                resolve({ width: img.width, height: img.height });
            };
            img.onerror = () => {
                clearTimeout(timeout);
                resolve({ width: 200, height: 200 }); 
            };
            img.src = dataUrl;
        });
    }

    // Caesar Shift Decrypt for LocalStorage Rescue
    function decryptCaesar(text, shift) {
        const mangleMap = {
            '€': 128, '‚': 130, 'ƒ': 131, '„': 132, '…': 133, '†': 134, '‡': 135,
            'ˆ': 136, '‰': 137, 'Š': 138, '‹': 139, 'Œ': 140, 'Ž': 142, '‘': 145,
            '’': 146, '“': 147, '”': 148, '•': 149, '–': 150, '—': 151, '˜': 152,
            '™': 153, 'š': 154, '›': 155, 'œ': 156, 'ž': 158, 'Ÿ': 159
        };
        return text.split('').map(char => {
            let code = char.charCodeAt(0);
            if (mangleMap[char] !== undefined) code = mangleMap[char];
            return String.fromCharCode(code - shift);
        }).join('');
    }

    // Formatting Constants
    const INVISIBLE_BORDERS = {
        top: { style: BorderStyle.NONE, size: 0, color: "auto" },
        bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
        left: { style: BorderStyle.NONE, size: 0, color: "auto" },
        right: { style: BorderStyle.NONE, size: 0, color: "auto" },
        insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "auto" },
        insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" }
    };

    const LIGHT_BORDER = { style: BorderStyle.SINGLE, size: 1, color: "E0E0E0" };
    const TABLE_BORDERS = {
        top: LIGHT_BORDER, bottom: LIGHT_BORDER, 
        insideHorizontal: LIGHT_BORDER,
        left: { style: BorderStyle.NONE, size: 0, color: "auto" },
        right: { style: BorderStyle.NONE, size: 0, color: "auto" },
        insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" }
    };

    const COL_WIDTHS = [
        { size: "10%", type: WidthType.PERCENTAGE },
        { size: "55%", type: WidthType.PERCENTAGE },
        { size: "15%", type: WidthType.PERCENTAGE },
        { size: "20%", type: WidthType.PERCENTAGE } 
    ];

    // --- 3. Main Document Generator ---

    async function generateWordDocument() {
        // Bulletproof scope retrieval 
        let data = typeof window !== 'undefined' && window.jsonData ? window.jsonData : (typeof jsonData !== 'undefined' ? jsonData : null);
        
        // Failsafe: Rescue from LocalStorage if JS scope is detached
        if (!data || !data.quote) {
            try {
                const localData = localStorage.getItem('quoteData');
                if (localData) data = JSON.parse(decryptCaesar(localData, 17));
            } catch (e) {
                console.warn("Could not rescue data from localStorage.", e);
            }
        }

        if (!data || !data.quote) {
            alert("Export misslyckades: Offertdatan saknas eller är korrupt.");
            return;
        }

        const lang = typeof currentLanguage !== 'undefined' ? currentLanguage : (data.quote.language || 'sv');
        const t = (typeof translations !== 'undefined' && translations[lang]) || {};
        const labels = data.quote?.labels || {};
        const visibility = data.quote?.visibility || { optional: true, info: true, terms: true };
        
        const docChildren = [];
        const emptyParagraph = () => new Paragraph({ children: [new TextRun("")] });

        // --- Header Layout (Logo & Metadata) ---
        const logoData = localStorage.getItem('companyLogo');
        const logoRuns = [];
        
        if (logoData) {
            try {
                const dims = await getImageDimensions(logoData);
                const targetWidth = data.quote?.defaultLogoWidth || 200;
                // Strictly round integers to prevent XML crash
                const targetHeight = Math.round((targetWidth / (dims.width || 1)) * (dims.height || targetWidth));
                const buffer = base64ToArrayBuffer(logoData);
                
                if (buffer.byteLength > 0) {
                    logoRuns.push(new ImageRun({
                        data: buffer,
                        transformation: { width: Math.round(targetWidth), height: targetHeight }
                    }));
                }
            } catch (e) { console.warn("Could not embed logo.", e); }
        }

        const headerTable = new Table({
            width: { size: "100%", type: WidthType.PERCENTAGE },
            borders: INVISIBLE_BORDERS,
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            width: { size: "50%", type: WidthType.PERCENTAGE },
                            children: [new Paragraph({ children: logoRuns.length > 0 ? logoRuns : [new TextRun("")] })],
                            verticalAlign: VerticalAlign.CENTER
                        }),
                        new TableCell({
                            width: { size: "50%", type: WidthType.PERCENTAGE },
                            children: [
                                new Paragraph({
                                    children: [new TextRun({ text: labels.quoteTitle || t.quoteTitle || "Quote", size: 32, bold: true })],
                                    heading: HeadingLevel.HEADING_1,
                                    alignment: AlignmentType.RIGHT
                                }),
                                new Paragraph({
                                    children: [
                                        new TextRun({ text: `${t.quoteNumberLabel || "No:"} `, bold: true }),
                                        new TextRun({ text: data.quote?.quoteNumber || "" })
                                    ],
                                    alignment: AlignmentType.RIGHT
                                }),
                                new Paragraph({
                                    children: [
                                        new TextRun({ text: `${t.dateLabel || "Date:"} `, bold: true }),
                                        new TextRun({ text: data.quote?.date || "" })
                                    ],
                                    alignment: AlignmentType.RIGHT
                                })
                            ],
                            verticalAlign: VerticalAlign.CENTER
                        })
                    ]
                })
            ]
        });
        docChildren.push(headerTable);
        docChildren.push(emptyParagraph()); // Spacing

        // --- Addresses Layout (To & From) ---
        const formatAddressLines = (companyData) => {
            if (!companyData || typeof companyData !== 'object') return [emptyParagraph()];
            const lines = Object.values(companyData).filter(val => val && String(val).trim() !== "");
            if (lines.length === 0) return [emptyParagraph()];
            return lines.map(line => new Paragraph({ children: parseHtmlToRuns(line) }));
        };

        const toChildren = [new Paragraph({ children: [new TextRun({ text: t.toLabel || "To:", bold: true })] })];
        if (data.quote?.showTillSection !== false) toChildren.push(...formatAddressLines(data.companyA));

        const fromChildren = [new Paragraph({ children: [new TextRun({ text: t.fromLabel || "From:", bold: true })] })];
        fromChildren.push(...formatAddressLines(data.companyB));

        const addressTable = new Table({
            width: { size: "100%", type: WidthType.PERCENTAGE },
            borders: INVISIBLE_BORDERS,
            rows: [
                new TableRow({
                    children: [
                        new TableCell({ width: { size: "50%", type: WidthType.PERCENTAGE }, children: toChildren }),
                        new TableCell({ width: { size: "50%", type: WidthType.PERCENTAGE }, children: fromChildren })
                    ]
                })
            ]
        });
        docChildren.push(addressTable);
        docChildren.push(emptyParagraph()); 

        // Reference Line
        if (data.quote?.reference) {
            docChildren.push(new Paragraph({
                children: [
                    new TextRun({ text: `${t.refLabel || "Ref:"} `, bold: true }),
                    new TextRun({ text: data.quote.reference })
                ],
                spacing: { after: 200 }
            }));
        }

        // --- Unified Table Builder Logic ---
        const safeFormatPrice = (val) => {
            if (val === undefined || val === null) return "0";
            return typeof formatPrice === 'function' ? formatPrice(val) : String(val);
        };

        const buildItemTable = (itemsArray, isOptional = false) => {
            if (!Array.isArray(itemsArray) || itemsArray.length === 0) return null;
            const rows = [];
            
            rows.push(new TableRow({
                tableHeader: true,
                children: [
                    new TableCell({ width: COL_WIDTHS[0], shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: (isOptional ? labels.optNrHeader : labels.nrHeader) || t.nrHeader || "Nr", bold: true })] })], margins: { top: 100, bottom: 100, left: 100, right: 100 } }),
                    new TableCell({ width: COL_WIDTHS[1], shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: (isOptional ? labels.optArticleHeader : labels.articleNameHeader) || t.articleNameHeader || "Name", bold: true })] })], margins: { top: 100, bottom: 100, left: 100, right: 100 } }),
                    new TableCell({ width: COL_WIDTHS[2], shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: (isOptional ? labels.optQuantityHeader : labels.quantityHeader) || t.quantityHeader || "Qty", bold: true })], alignment: AlignmentType.CENTER })], margins: { top: 100, bottom: 100, left: 100, right: 100 } }),
                    new TableCell({ width: COL_WIDTHS[3], shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: (isOptional ? labels.optPriceHeader : labels.priceHeader) || t.priceHeader || "Price", bold: true })], alignment: AlignmentType.RIGHT })], margins: { top: 100, bottom: 100, left: 100, right: 100 } })
                ]
            }));

            let addedVisibleItem = false;

            itemsArray.forEach(item => {
                if (!item || item.isHiddenFromPrint) return;

                if (item.type === 'separator') {
                    rows.push(new TableRow({
                        children: [
                            new TableCell({ children: [emptyParagraph()], columnSpan: 4, borders: { top: { style: BorderStyle.DASHED, size: 1, color: "999999" }, bottom: { style: BorderStyle.NONE, size: 0, color: "auto" }, left: { style: BorderStyle.NONE, size: 0, color: "auto" }, right: { style: BorderStyle.NONE, size: 0, color: "auto" } } })
                        ]
                    }));
                    return;
                }

                addedVisibleItem = true;

                const itemCellChildren = [new Paragraph({ children: [new TextRun({ text: item.name || "", bold: true })] })];
                if (item.itemDescription) {
                    itemCellChildren.push(new Paragraph({ children: parseHtmlToRuns(item.itemDescription) }));
                }

                const priceStr = item.isPriceBakedIn ? "" : safeFormatPrice(item.targetPrice);

                rows.push(new TableRow({
                    children: [
                        new TableCell({ width: COL_WIDTHS[0], children: [new Paragraph({ children: [new TextRun({ text: String(item.itemNumber || "") })] })], margins: { top: 100, bottom: 100, left: 100, right: 100 }, verticalAlign: VerticalAlign.TOP }),
                        new TableCell({ width: COL_WIDTHS[1], children: itemCellChildren, margins: { top: 100, bottom: 100, left: 100, right: 100 }, verticalAlign: VerticalAlign.TOP }),
                        new TableCell({ width: COL_WIDTHS[2], children: [new Paragraph({ children: [new TextRun({ text: String(item.quantity || 0) })], alignment: AlignmentType.CENTER })], margins: { top: 100, bottom: 100, left: 100, right: 100 }, verticalAlign: VerticalAlign.TOP }),
                        new TableCell({ width: COL_WIDTHS[3], children: [new Paragraph({ children: [new TextRun({ text: priceStr })], alignment: AlignmentType.RIGHT })], margins: { top: 100, bottom: 100, left: 100, right: 100 }, verticalAlign: VerticalAlign.TOP })
                    ]
                }));

                (item.subItems || []).forEach(subItem => {
                    if (!subItem || subItem.isHiddenFromPrint) return;

                    const subCellChildren = [new Paragraph({ children: [new TextRun({ text: subItem.subItemName || "", bold: true, italics: true })], indent: { left: 360 } })];
                    if (subItem.subItemDescription) {
                        subCellChildren.push(new Paragraph({ children: parseHtmlToRuns(subItem.subItemDescription, true), indent: { left: 360 } }));
                    }

                    const subPriceStr = subItem.isPriceBakedIn ? "" : safeFormatPrice(subItem.subItemTargetPrice);

                    rows.push(new TableRow({
                        children: [
                            new TableCell({ width: COL_WIDTHS[0], children: [new Paragraph({ children: [new TextRun({ text: String(subItem.subItemNumber || "") })] })], margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: VerticalAlign.TOP }),
                            new TableCell({ width: COL_WIDTHS[1], children: subCellChildren, margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: VerticalAlign.TOP }),
                            new TableCell({ width: COL_WIDTHS[2], children: [new Paragraph({ children: [new TextRun({ text: String(subItem.subItemQuantity || 0) })], alignment: AlignmentType.CENTER })], margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: VerticalAlign.TOP }),
                            new TableCell({ width: COL_WIDTHS[3], children: [new Paragraph({ children: [new TextRun({ text: subPriceStr })], alignment: AlignmentType.RIGHT })], margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: VerticalAlign.TOP })
                        ]
                    }));
                });
            });

            if (!addedVisibleItem) return null;
            return new Table({ 
                width: { size: "100%", type: WidthType.PERCENTAGE }, 
                columnWidths: [1000, 5500, 1500, 2000], // Hardware DXA proportional backups
                borders: TABLE_BORDERS, 
                rows: rows 
            });
        };

        // --- Main Items & Totals ---
        const mainTable = buildItemTable(data.items, false);
        if (mainTable) docChildren.push(mainTable);

        if (data.quote?.removeTotal !== true) {
            let total = 0;
            (data.items || []).forEach(item => {
                if (item.isHiddenFromPrint || item.type === 'separator') return;
                total += parseFloat(item.targetPrice) || 0;
                (item.subItems || []).forEach(subItem => {
                    if (!subItem.isPriceBakedIn && !subItem.isHiddenFromPrint) {
                        total += parseFloat(subItem.subItemTargetPrice) || 0;
                    }
                });
            });

            const currency = data.quote?.useCustomCurrency ? (data.quote?.customCurrency || "SEK") : "SEK";
            docChildren.push(new Paragraph({
                children: [
                    new TextRun({ text: `${labels.totalLabel || t.totalLabel || "Total:"} `, bold: true, size: 28 }),
                    new TextRun({ text: `${safeFormatPrice(total)} ${currency}`, bold: true, size: 28 })
                ],
                alignment: AlignmentType.RIGHT,
                spacing: { before: 200, after: 400 }
            }));
        } else {
            docChildren.push(new Paragraph({ children: [new TextRun("")] , spacing: { after: 200 } })); 
        }

        // --- Optional Items ---
        if (visibility.optional && Array.isArray(data.optionalItems) && data.optionalItems.length > 0) {
            const optTable = buildItemTable(data.optionalItems, true);
            if (optTable) {
                docChildren.push(new Paragraph({
                    children: [new TextRun({ text: labels.optionalItemsHeading || t.optionalItemsHeading || "Optional", size: 28, bold: true })],
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 400, after: 100 }
                }));
                docChildren.push(optTable);
                docChildren.push(emptyParagraph());
            }
        }

        // --- Info & Images ---
        if (visibility.info && Array.isArray(data.infoImages) && data.infoImages.length > 0) {
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: labels.infoImagesHeading || t.infoImagesHeading || "Info/Images", size: 28, bold: true })],
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 100 }
            }));

            for (const info of data.infoImages) {
                if (!info) continue;
                const align = info.centering === 'left' ? AlignmentType.LEFT : AlignmentType.CENTER;

                if (info.type === 'text') {
                    docChildren.push(new Paragraph({
                        children: parseHtmlToRuns(info.content),
                        alignment: align,
                        spacing: { after: 200 }
                    }));
                } else if (info.type === 'table') {
                    const tableRows = [];
                    for (let r = 0; r < info.rows; r++) {
                        const cells = [];
                        for (let c = 0; c < info.cols; c++) {
                            const cellText = info.data && info.data[r] ? info.data[r][c] || "" : "";
                            cells.push(new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: cellText })] })],
                                margins: { top: 100, bottom: 100, left: 100, right: 100 }
                            }));
                        }
                        tableRows.push(new TableRow({ children: cells }));
                    }
                    docChildren.push(new Table({ width: { size: "100%", type: WidthType.PERCENTAGE }, borders: TABLE_BORDERS, rows: tableRows }));
                    docChildren.push(emptyParagraph());
                } else if (info.type === 'page-break') {
                    docChildren.push(new Paragraph({ children: [new TextRun("")], pageBreakBefore: true }));
                } else if (info.type === 'image' && info.src) {
                    try {
                        const dims = await getImageDimensions(info.src);
                        const buffer = base64ToArrayBuffer(info.src);
                        
                        if (buffer.byteLength > 0) {
                            let w = info.width || dims.width;
                            const maxA4Width = 600; 
                            if (w > maxA4Width) w = maxA4Width;
                            const h = Math.round((w / (dims.width || 1)) * (dims.height || w));

                            docChildren.push(new Paragraph({
                                children: [new ImageRun({ data: buffer, transformation: { width: Math.round(w), height: h } })],
                                alignment: align,
                                spacing: { after: 200 }
                            }));
                        }
                    } catch (e) { console.warn("Skipped malformed image block.", e); }
                }
            }
        }

        // --- Terms & Conditions ---
        if (visibility.terms && Array.isArray(data.terms) && data.terms.length > 0) {
            docChildren.push(new Paragraph({
                children: [new TextRun({ text: labels.termsHeading || t.termsHeading || "Terms", size: 28, bold: true })],
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 100 }
            }));

            data.terms.forEach(termHtml => {
                if (!termHtml) return;
                docChildren.push(new Paragraph({
                    children: parseHtmlToRuns(termHtml),
                    bullet: { level: 0 },
                    spacing: { after: 100 }
                }));
            });
        }

        // --- Signatures ---
        if (data.quote?.showSignature && data.signature) {
            docChildren.push(emptyParagraph());
            docChildren.push(emptyParagraph());
            
            const sig = data.signature;
            const dateStr = `${t.signatureDateLabel || "Datum"}: ${sig.date || ""}`;

            docChildren.push(new Paragraph({ children: [new TextRun({ text: dateStr })], spacing: { after: 400 } }));

            docChildren.push(new Table({
                width: { size: "100%", type: WidthType.PERCENTAGE },
                borders: INVISIBLE_BORDERS,
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                width: { size: "45%", type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({ children: [new TextRun({ text: sig.lessorLine || "______________________" })] }),
                                    new Paragraph({ children: [new TextRun({ text: sig.lessorText || "Uthyrare" })] })
                                ]
                            }),
                            new TableCell({ width: { size: "10%", type: WidthType.PERCENTAGE }, children: [emptyParagraph()] }),
                            new TableCell({
                                width: { size: "45%", type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({ children: [new TextRun({ text: sig.lesseeLine || "______________________" })] }),
                                    new Paragraph({ children: [new TextRun({ text: sig.lesseeText || "Hyrestagare" })] })
                                ]
                            })
                        ]
                    })
                ]
            }));
        }

        // --- Packaging and Download ---
        try {
            const doc = new Document({
                creator: "Quote Generator",
                sections: [{
                    properties: {
                        page: {
                            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } 
                        }
                    },
                    children: docChildren
                }]
            });

            const blob = await Packer.toBlob(doc);
            const refSafe = (data.quote?.reference || "").trim().replace(/[^\w\s\-]/gi, '');
            const numSafe = (data.quote?.quoteNumber || "").trim();
            const fileNameTitle = (labels.quoteTitle || t.quoteTitle || "Quote").replace(/[^\w\s\-]/gi, '');
            
            let exportName = `${fileNameTitle}`;
            if (refSafe) exportName += `_${refSafe}`;
            if (numSafe) exportName += `_${numSafe}`;
            exportName += ".docx";

            if (typeof saveAs === 'function') {
                saveAs(blob, exportName);
            } else {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = exportName;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }
        } catch (err) {
            console.error("Critical error during document packaging:", err);
            alert("Ett fel inträffade vid skapandet av Word-filen. Detaljer: " + err.message);
        }
    }

    // --- 4. Event Binding ---
    function attachExportButton() {
        const btn = document.getElementById('exportWordBtn');
        if (btn) {
            btn.removeEventListener('click', generateWordDocument); 
            btn.addEventListener('click', generateWordDocument);
        }
    }

    attachExportButton(); 
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', attachExportButton);
    }
    
})();
