/**
 * wordify.js - Full Robust Version
 * Fixes: Global Scope, Uint8Array Images, XML Entity Safety, and Binding.
 */

(function () {
    const TAG = "[Wordify]";
    console.log(`${TAG} 🚀 Script loaded and executing...`);

    // --- 1. Library Check ---
    if (typeof docx === 'undefined') {
        console.error(`${TAG} ❌ ERROR: docx library not found. Check CDN in index.html.`);
        return;
    }
    const { 
        Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
        ImageRun, WidthType, BorderStyle, AlignmentType, VerticalAlign, 
        HeadingLevel 
    } = docx;

    // --- 2. Robust Helpers ---

    // FIX: docx v8.x requires Uint8Array, not a raw ArrayBuffer
    function base64ToUint8Array(dataUrl) {
        try {
            const base64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
            const binaryString = window.atob(base64);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            return bytes;
        } catch (e) {
            console.error(`${TAG} ❌ Image decode failed.`, e);
            return null;
        }
    }

    function parseHtmlToRuns(htmlStr, defaultItalics = false, defaultBold = false) {
        if (!htmlStr) return [new TextRun({ text: "" })];
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

    async function getImageDimensions(dataUrl) {
        return new Promise(resolve => {
            const img = new Image();
            const timeout = setTimeout(() => resolve({ width: 200, height: 200 }), 3000); 
            img.onload = () => { clearTimeout(timeout); resolve({ width: img.width, height: img.height }); };
            img.onerror = () => { clearTimeout(timeout); resolve({ width: 200, height: 200 }); };
            img.src = dataUrl;
        });
    }

    function decryptCaesar(text, shift) {
        const mangleMap = { '€': 128, '‚': 130, 'ƒ': 131, '„': 132, '…': 133, '†': 134, '‡': 135, 'ˆ': 136, '‰': 137, 'Š': 138, '‹': 139, 'Œ': 140, 'Ž': 142, '‘': 145, '’': 146, '“': 147, '”': 148, '•': 149, '–': 150, '—': 151, '˜': 152, '™': 153, 'š': 154, '›': 155, 'œ': 156, 'ž': 158, 'Ÿ': 159 };
        return text.split('').map(char => {
            let code = char.charCodeAt(0);
            if (mangleMap[char] !== undefined) code = mangleMap[char];
            return String.fromCharCode(code - shift);
        }).join('');
    }

    const INVISIBLE_BORDERS = { top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 }, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 }, insideHorizontal: { style: BorderStyle.NONE, size: 0 }, insideVertical: { style: BorderStyle.NONE, size: 0 } };
    const LIGHT_BORDER = { style: BorderStyle.SINGLE, size: 1, color: "E0E0E0" };
    const TABLE_BORDERS = { top: LIGHT_BORDER, bottom: LIGHT_BORDER, insideHorizontal: LIGHT_BORDER, left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 }, insideVertical: { style: BorderStyle.NONE, size: 0 } };
    const COL_WIDTHS = [{ size: "10%", type: WidthType.PERCENTAGE }, { size: "55%", type: WidthType.PERCENTAGE }, { size: "15%", type: WidthType.PERCENTAGE }, { size: "20%", type: WidthType.PERCENTAGE }];

    // --- 3. Main Generator ---

    async function generateWordDocument() {
        console.log(`${TAG} 🎬 EXPORT STARTED`);

        // FIX: Check multiple scopes for data
        let data = window.jsonData || (typeof jsonData !== 'undefined' ? jsonData : null);
        
        if (!data || !data.quote) {
            console.warn(`${TAG} Global jsonData missing. Trying LocalStorage...`);
            try {
                const localData = localStorage.getItem('quoteData');
                if (localData) data = JSON.parse(decryptCaesar(localData, 17));
            } catch (e) { console.error(`${TAG} Rescue failed.`, e); }
        }

        if (!data || !data.quote) {
            alert("Export misslyckades: Ingen data hittades. Ladda upp en JSON-fil först.");
            return;
        }

        const lang = data.quote.language || 'sv';
        const t = (typeof translations !== 'undefined' && translations[lang]) || {};
        const labels = data.quote.labels || {};
        const visibility = data.quote.visibility || { optional: true, info: true, terms: true };
        const docChildren = [];
        const emptyParagraph = () => new Paragraph({ children: [new TextRun("")] });

        // Header & Logo
        const logoData = localStorage.getItem('companyLogo');
        const logoRuns = [];
        if (logoData) {
            const dims = await getImageDimensions(logoData);
            const targetWidth = data.quote.defaultLogoWidth || 200;
            const targetHeight = Math.round((targetWidth / (dims.width || 1)) * (dims.height || targetWidth));
            const uint8 = base64ToUint8Array(logoData);
            if (uint8) {
                logoRuns.push(new ImageRun({ data: uint8, transformation: { width: Math.round(targetWidth), height: targetHeight } }));
            }
        }

        docChildren.push(new Table({
            width: { size: "100%", type: WidthType.PERCENTAGE },
            borders: INVISIBLE_BORDERS,
            rows: [new TableRow({
                children: [
                    new TableCell({ width: { size: "50%", type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: logoRuns })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ width: { size: "50%", type: WidthType.PERCENTAGE }, children: [
                        new Paragraph({ children: [new TextRun({ text: labels.quoteTitle || t.quoteTitle || "Offert", size: 32, bold: true })], heading: HeadingLevel.HEADING_1, alignment: AlignmentType.RIGHT }),
                        new Paragraph({ children: [new TextRun({ text: `${t.quoteNumberLabel || "Nr:"} ${data.quote.quoteNumber || ""}` })], alignment: AlignmentType.RIGHT }),
                        new Paragraph({ children: [new TextRun({ text: `${t.dateLabel || "Datum:"} ${data.quote.date || ""}` })], alignment: AlignmentType.RIGHT })
                    ], verticalAlign: VerticalAlign.CENTER })
                ]
            })]
        }));

        docChildren.push(emptyParagraph());

        // Addresses
        const addressLines = (comp) => Object.values(comp || {}).filter(v => !!v).map(l => new Paragraph({ children: parseHtmlToRuns(l) }));
        docChildren.push(new Table({
            width: { size: "100%", type: WidthType.PERCENTAGE },
            borders: INVISIBLE_BORDERS,
            rows: [new TableRow({
                children: [
                    new TableCell({ width: { size: "50%", type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: t.toLabel || "Till:", bold: true })] }), ...addressLines(data.companyA)] }),
                    new TableCell({ width: { size: "50%", type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: t.fromLabel || "Från:", bold: true })] }), ...addressLines(data.companyB)] })
                ]
            })]
        }));

        docChildren.push(emptyParagraph());

        // Table logic
        const safeFormatPrice = (val) => (typeof formatPrice === 'function' ? formatPrice(val) : String(val || 0));
        const buildTable = (items, isOpt) => {
            if (!items || items.length === 0) return null;
            const rows = [new TableRow({
                tableHeader: true,
                children: [
                    new TableCell({ width: COL_WIDTHS[0], shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: "Nr", bold: true })] })] }),
                    new TableCell({ width: COL_WIDTHS[1], shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: "Namn", bold: true })] })] }),
                    new TableCell({ width: COL_WIDTHS[2], shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: "Antal", bold: true })], alignment: AlignmentType.CENTER })] }),
                    new TableCell({ width: COL_WIDTHS[3], shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: "Pris", bold: true })], alignment: AlignmentType.RIGHT })] })
                ]
            })];

            items.forEach(item => {
                if (item.isHiddenFromPrint) return;
                rows.push(new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ text: String(item.itemNumber || "") })] }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.name || "", bold: true })] }), ...(item.itemDescription ? [new Paragraph({ children: parseHtmlToRuns(item.itemDescription) })] : [])] }),
                        new TableCell({ children: [new Paragraph({ text: String(item.quantity || ""), alignment: AlignmentType.CENTER })] }),
                        new TableCell({ children: [new Paragraph({ text: item.isPriceBakedIn ? "" : safeFormatPrice(item.targetPrice), alignment: AlignmentType.RIGHT })] })
                    ]
                }));
            });
            return new Table({ width: { size: "100%", type: WidthType.PERCENTAGE }, borders: TABLE_BORDERS, rows });
        };

        const mainTbl = buildTable(data.items, false);
        if (mainTbl) docChildren.push(mainTbl);

        // Packaging
        try {
            const doc = new Document({ sections: [{ children: docChildren }] });
            const blob = await Packer.toBlob(doc);
            const fileName = `Offert_${data.quote.quoteNumber || "Draft"}.docx`;
            
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
            console.error(`${TAG} ❌ Packing Error:`, err);
            alert("Ett fel uppstod vid skapandet av filen.");
        }
    }

    // --- 4. Event Binding ---
    function attach() {
        const btn = document.getElementById('exportWordBtn');
        if (btn) {
            console.log(`${TAG} ✅ Button #exportWordBtn found.`);
            btn.removeEventListener('click', generateWordDocument);
            btn.addEventListener('click', generateWordDocument);
        } else {
            console.warn(`${TAG} ❌ Button #exportWordBtn NOT found in DOM.`);
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', attach);
    } else {
        attach();
    }
})();
