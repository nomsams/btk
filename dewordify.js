/**
 * dewordify.js - Production Ready Reverse-Parser for Wordify.js
 * Features: Zero-dependency setup (Auto-injects JSZip), Smart Table Recognition,
 * Dynamic Grid Mapping, Image Base64 Extraction, Hierarchy Reconstruction.
 * V3: Bottom-up Rewrite, Fault-tolerant Parsing, Dynamic Number Validation.
 */

(function () {
    const TAG = "[Dewordify]";
    console.log(`${TAG} 🚀 Script loaded. Waiting for UI...`);

    let zipObj = null;
    let relsMap = {};
    let mediaFiles = {};

    // --- Utility: Auto-Load JSZip ---
    async function ensureJSZip() {
        if (typeof window.JSZip !== 'undefined') return window.JSZip;
        console.log(`${TAG} 📦 Injecting JSZip...`);
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
            script.onload = () => resolve(window.JSZip);
            script.onerror = () => reject(new Error("Kunde inte ladda JSZip-biblioteket."));
            document.head.appendChild(script);
        });
    }

    // --- Utility: Setup UI Binding ---
    function attachUI() {
        let fileInput = document.getElementById('docxUploadInput');
        let btn = document.getElementById('uploadDocxBtn');

        if (!fileInput || !btn) {
            const footer = document.querySelector('.footer');
            if (!footer) {
                setTimeout(attachUI, 500);
                return;
            }
            
            fileInput = document.createElement('input');
            fileInput.type = 'file';
            fileInput.id = 'docxUploadInput';
            fileInput.accept = '.docx';
            fileInput.style.display = 'none';

            btn = document.createElement('button');
            btn.className = 'footer-button';
            btn.id = 'uploadDocxBtn';
            btn.style.cssText = 'background-color: #2b579a; color: white;';
            btn.textContent = 'Ladda Word-offert (.docx)';

            const pdfBtn = document.getElementById('uploadPdfBtn');
            if (pdfBtn) {
                footer.insertBefore(fileInput, pdfBtn);
                footer.insertBefore(btn, pdfBtn);
            } else {
                footer.appendChild(fileInput);
                footer.appendChild(btn);
            }
            console.log(`${TAG} 🎛️ UI Injected dynamically.`);
        } else {
            console.log(`${TAG} 🎛️ Bound to existing HTML elements.`);
        }

        btn.removeEventListener('click', triggerInput);
        btn.addEventListener('click', triggerInput);
        
        fileInput.removeEventListener('change', handleFileSelect);
        fileInput.addEventListener('change', handleFileSelect);

        function triggerInput() { fileInput.click(); }
    }

    // --- XML Parsing Helpers (Robust & Strict) ---
    function getDirect(parent, nodeName) {
        if (!parent) return [];
        return Array.from(parent.childNodes).filter(n => 
            n.localName === nodeName
        );
    }

    // Extracts text preserving paragraphs and line breaks perfectly
    function extractTextRobust(node) {
        if (!node) return "";
        let text = "";
        
        const paragraphs = node.getElementsByTagName("w:p");
        if (paragraphs.length === 0) {
            const runs = node.getElementsByTagName("w:t");
            for (let i = 0; i < runs.length; i++) text += runs[i].textContent;
            return text.trim();
        }

        for (let i = 0; i < paragraphs.length; i++) {
            const p = paragraphs[i];
            let pText = "";
            const runs = p.getElementsByTagName("w:r");
            for (let j = 0; j < runs.length; j++) {
                const children = runs[j].childNodes;
                for (let k = 0; k < children.length; k++) {
                    const child = children[k];
                    if (child.localName === "t") pText += child.textContent;
                    else if (child.localName === "br" || child.localName === "cr") pText += "\n";
                }
            }
            if (pText.trim() !== "") text += pText.trim() + "\n";
        }
        return text.trim();
    }

    // Parses diverse numbers: '218 288', '1,000.50', '1.000,50' -> 218288
    function parseRobustNumber(str) {
        if (!str) return 0;
        if (/ingår|included/i.test(str)) return 0; // Wordify specific baked items

        let s = str.replace(/[\s\xA0]/g, ''); // strip all variants of spaces
        s = s.replace(/[^\d,\.-]/g, ''); // keep only digits, comma, dot, minus
        if (!s) return 0;

        const lastComma = s.lastIndexOf(',');
        const lastDot = s.lastIndexOf('.');
        
        if (lastComma > -1 && lastDot > -1) {
            if (lastComma > lastDot) s = s.replace(/\./g, '').replace(',', '.'); // comma is decimal
            else s = s.replace(/,/g, ''); // dot is decimal
        } else if (lastComma > -1) {
            s = s.replace(',', '.'); // assume single comma is a decimal for Swedish format
        }
        
        return parseFloat(s) || 0;
    }

    // --- Extractor Logic ---
    async function extractImageBase64(rId) {
        try {
            const targetPath = relsMap[rId];
            if (!targetPath || !mediaFiles[targetPath]) return null;
            
            const fileData = mediaFiles[targetPath];
            const base64 = await fileData.async("base64");
            
            const ext = targetPath.split('.').pop().toLowerCase();
            let mime = "image/png";
            if (ext === "jpg" || ext === "jpeg") mime = "image/jpeg";
            if (ext === "gif") mime = "image/gif";

            return `data:${mime};base64,${base64}`;
        } catch (e) {
            console.warn(`${TAG} Image extraction failed for ID: ${rId}`, e);
            return null;
        }
    }

    async function extractImagesFromNode(node) {
        const blips = node.getElementsByTagName("a:blip");
        const results = [];
        for (let i = 0; i < blips.length; i++) {
            const rId = blips[i].getAttribute("r:embed");
            if (rId) {
                const b64 = await extractImageBase64(rId);
                if (b64) results.push(b64);
            }
        }
        return results;
    }

    // Analyzes a table dynamically without assuming strict columns
    function analyzeTable(tblNode) {
        const rows = getDirect(tblNode, "tr");
        const grid = rows.map(tr => getDirect(tr, "tc").map(tc => extractTextRobust(tc)));
        
        const flatTop = grid.slice(0, 2).map(r => r.join(" ").toLowerCase()).join(" ");
        
        let type = 'unknown';
        if (flatTop.includes("datum") || flatTop.includes("date") || flatTop.includes("offert nr")) type = 'header';
        else if (flatTop.includes("till:") || flatTop.includes("to:") || flatTop.includes("från:")) type = 'address';
        else if (flatTop.includes("antal") || flatTop.includes("quantity")) type = 'items';
        
        return { type, grid, tblNode };
    }

    // Processes item grids and maps columns via dynamic header discovery
    function processItemsGrid(grid, targetArray, extractedData) {
        let colMap = { nr: 0, name: 1, qty: 2, price: 3 };
        let headerRowIdx = -1;

        // 1. Map columns dynamically to avoid hidden/missing column bugs
        for (let i = 0; i < grid.length; i++) {
            const rowText = grid[i].join(" ").toLowerCase();
            if (rowText.includes("antal") || rowText.includes("quantity")) {
                headerRowIdx = i;
                const headerCells = grid[i].map(c => c.toLowerCase());
                headerCells.forEach((cText, idx) => {
                    if (cText.includes("nr") || cText.includes("no")) colMap.nr = idx;
                    if (cText.includes("artikel") || cText.includes("name")) colMap.name = idx;
                    if (cText.includes("antal") || cText.includes("quantity")) colMap.qty = idx;
                    if (cText.includes("pris") || cText.includes("price")) colMap.price = idx;
                });
                break;
            }
        }

        let lastMainItem = null;

        // 2. Scan rows
        for (let i = headerRowIdx + 1; i < grid.length; i++) {
            const rowTexts = grid[i];
            if (!rowTexts || rowTexts.length === 0) continue;

            const fullRowText = rowTexts.join(" ").toLowerCase();
            
            // Reached Total Footer? Break loop, extract custom currency if applicable
            if (fullRowText.includes("total:") || (fullRowText.includes("total") && rowTexts.length <= 2)) {
                const match = fullRowText.match(/([a-z]{3})$/i);
                if (match && match[1].toLowerCase() !== 'sek') {
                    extractedData.quote.useCustomCurrency = true;
                    extractedData.quote.customCurrency = match[1].toUpperCase();
                }
                break; 
            }

            const nrText = rowTexts[colMap.nr] || "";
            const nameDescRaw = rowTexts[colMap.name] || "";
            const qtyText = rowTexts[colMap.qty] || "";
            const priceText = rowTexts[colMap.price] || "";

            // Split Name and Desc perfectly by the first line break
            const lines = nameDescRaw.split('\n');
            const name = lines[0] ? lines[0].trim() : "";
            const desc = lines.length > 1 ? lines.slice(1).join('\n').trim() : "";

            const qty = parseRobustNumber(qtyText);
            const price = parseRobustNumber(priceText);

            if (!name && !desc && !nrText && price === 0) continue; // Skip genuinely empty rows

            const isPriceBakedIn = (price === 0 && (priceText.trim() === "" || priceText.trim() === "-" || priceText.toLowerCase().includes("ingår") || priceText.toLowerCase().includes("included")));

            // Smart SubItem Heuristic
            // Only considered a subitem if it has no number BUT a previous item exists, or starts with Parent Number (e.g., Parent 1, SubItem 1.1)
            const isSubItem = (lastMainItem && nrText && nrText.includes('.') && nrText.startsWith(lastMainItem.itemNumber.split('.')[0] + '.')) 
                           || (!nrText && lastMainItem && qty === 0 && price === 0);

            const itemObj = {
                type: 'item',
                itemNumber: nrText,
                name: name,
                itemDescription: desc,
                quantity: qty,
                targetPrice: price,
                originalPrice: qty ? price / qty : 0,
                isPriceBakedIn: isPriceBakedIn,
                isHiddenFromPrint: false,
                vendorNotes: "",
                subItems: []
            };

            if (isSubItem && lastMainItem) {
                lastMainItem.subItems.push({
                    subItemNumber: itemObj.itemNumber,
                    subItemName: itemObj.name,
                    subItemDescription: itemObj.itemDescription,
                    subItemQuantity: itemObj.quantity,
                    subItemTargetPrice: itemObj.targetPrice,
                    subItemOriginalPrice: itemObj.originalPrice,
                    isPriceBakedIn: itemObj.isPriceBakedIn,
                    isHiddenFromPrint: false,
                    vendorNotes: ""
                });
            } else {
                targetArray.push(itemObj);
                lastMainItem = itemObj;
            }
        }
    }

    // --- Main Processor Hook ---
    async function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        const msgEl = document.getElementById('statusMessage');
        if (msgEl) {
            msgEl.textContent = "Analyserar Word-dokument...";
            msgEl.style.color = "#2b579a";
        }

        try {
            await ensureJSZip();
            zipObj = await window.JSZip.loadAsync(file);
            
            // 1. Build Relationships Map for Media Extraction
            relsMap = {};
            mediaFiles = {};
            const relsFile = zipObj.file("word/_rels/document.xml.rels");
            if (relsFile) {
                const relsXml = await relsFile.async("text");
                const relsDoc = new DOMParser().parseFromString(relsXml, "application/xml");
                Array.from(relsDoc.getElementsByTagName("Relationship")).forEach(rel => {
                    relsMap[rel.getAttribute("Id")] = rel.getAttribute("Target");
                });
            }

            zipObj.folder("word/media").forEach((relativePath, fileObj) => {
                mediaFiles[`media/${relativePath}`] = fileObj;
            });

            // 2. Read Document Body
            const docFile = zipObj.file("word/document.xml");
            if (!docFile) throw new Error("Ogiltig Word-fil: word/document.xml saknas.");
            const docXml = await docFile.async("text");
            const doc = new DOMParser().parseFromString(docXml, "application/xml");
            const body = doc.getElementsByTagName("w:body")[0];

            // 3. Extracted Data Template
            const extractedData = {
                quote: { language: 'sv', visibility: { optional: false, info: false, terms: false } },
                companyA: {},
                companyB: {},
                items: [],
                optionalItems: [],
                infoImages: [],
                terms: []
            };

            let currentMode = 'pre-items'; // State flow: pre-items -> items -> optional -> info -> terms

            // 4. Sequential Iteration (Bottom-Up execution ensuring logical document flow)
            const nodes = Array.from(body.childNodes);
            for (let i = 0; i < nodes.length; i++) {
                const node = nodes[i];
                const nodeName = node.localName;

                if (nodeName === "tbl") {
                    const tblData = analyzeTable(node);
                    
                    if (tblData.type === 'header') {
                        // Attempt to grab Logo
                        const firstCell = getDirect(getDirect(node, "tr")[0], "tc")[0];
                        if (firstCell) {
                            const imgs = await extractImagesFromNode(firstCell);
                            if (imgs.length > 0) localStorage.setItem('companyLogo', imgs[0]);
                        }
                        
                        // Parse Quote No and Date
                        const text = tblData.grid.flat().join(" ");
                        const nrMatch = text.match(/(?:Nr|No):\s*(\S+)/i);
                        const dateMatch = text.match(/(?:Datum|Date):\s*([\d-]+)/i);
                        if (nrMatch) extractedData.quote.quoteNumber = nrMatch[1];
                        if (dateMatch) extractedData.quote.date = dateMatch[1];

                    } else if (tblData.type === 'address') {
                        const row0 = tblData.grid[0] || [];
                        extractedData.quote.language = row0.join(" ").toLowerCase().includes("to:") ? 'en' : 'sv';
                        
                        let colA = 0, colB = 1;
                        row0.forEach((cText, cIdx) => {
                            if (/till:|to:/i.test(cText)) colA = cIdx;
                            if (/från:|from:/i.test(cText)) colB = cIdx;
                        });

                        const parseAddress = (textBlock, targetObj) => {
                            textBlock.split('\n').map(l => l.replace(/Till:|To:|Från:|From:/gi, '').trim()).filter(Boolean)
                                     .forEach((l, idx) => targetObj[`line${idx + 1}`] = l);
                        };

                        if (row0[colA]) parseAddress(row0[colA], extractedData.companyA);
                        if (row0[colB]) parseAddress(row0[colB], extractedData.companyB);

                    } else if (tblData.type === 'items') {
                        if (currentMode === 'items' || currentMode === 'post-items') {
                            processItemsGrid(tblData.grid, extractedData.optionalItems, extractedData);
                            extractedData.quote.visibility.optional = true;
                            currentMode = 'info'; 
                        } else {
                            processItemsGrid(tblData.grid, extractedData.items, extractedData);
                            currentMode = 'post-items';
                        }
                    } else {
                        // Unrecognized tables go to Info section
                        if (currentMode === 'info' || currentMode === 'post-items') {
                            extractedData.infoImages.push({
                                type: 'table',
                                rows: tblData.grid.length,
                                cols: tblData.grid[0].length,
                                data: tblData.grid,
                                centering: 'center'
                            });
                            extractedData.quote.visibility.info = true;
                        }
                    }

                } else if (nodeName === "p") {
                    const text = extractTextRobust(node);
                    const lowerText = text.toLowerCase();
                    
                    // State Transitions based on Section Headers
                    if (/(alternativ|alternative)/i.test(lowerText) && (currentMode === 'items' || currentMode === 'post-items')) {
                        currentMode = 'optional';
                        extractedData.quote.visibility.optional = true;
                        continue;
                    } else if (/(info|bilder|images)/i.test(lowerText)) {
                        currentMode = 'info';
                        extractedData.quote.visibility.info = true;
                        continue;
                    } else if (/(villkor|terms)/i.test(lowerText)) {
                        currentMode = 'terms';
                        extractedData.quote.visibility.terms = true;
                        continue;
                    }
                    
                    // Always extract images inline
                    const images = await extractImagesFromNode(node);
                    images.forEach(b64 => {
                        extractedData.infoImages.push({ type: 'image', src: b64, width: 400, centering: 'center', compressionImmune: false });
                        extractedData.quote.visibility.info = true;
                        if (currentMode === 'pre-items' || currentMode === 'items' || currentMode === 'post-items') {
                            currentMode = 'info'; // Fallback jump if isolated images are found
                        }
                    });

                    // Consume free-floating text into appropriate bins
                    if (text && images.length === 0) {
                        if (currentMode === 'terms') {
                            extractedData.terms.push(text);
                        } else if (currentMode === 'info') {
                            extractedData.infoImages.push({ type: 'text', content: text.replace(/\n/g, '<br>'), centering: 'center' });
                        }
                    }
                }
            }

            // 5. Inject Data Back to App Ecosystem
            if (typeof window.processIncomingJson === 'function') {
                window.processIncomingJson(JSON.stringify(extractedData), file.name);
            } else {
                window.jsonData = extractedData;
                if (typeof window.loadQuoteData === 'function') window.loadQuoteData(extractedData, file.name);
                if (typeof window.reRenderAll === 'function') window.reRenderAll();
                if (typeof window.saveToLocal === 'function') window.saveToLocal();
            }

            if (msgEl) {
                msgEl.textContent = "Word-offert laddades framgångsrikt!";
                msgEl.style.color = "green";
                setTimeout(() => { if (msgEl.textContent.includes("Word-offert")) msgEl.textContent = ''; }, 3000);
            }

        } catch (err) {
            console.error(`${TAG} ❌ Kritiskt Fel:`, err);
            if (msgEl) {
                msgEl.textContent = "Kunde inte tolka Word-filen. Fel: " + err.message;
                msgEl.style.color = "red";
            }
            alert("Ett fel uppstod vid inläsning av Word-filen.");
        } finally {
            event.target.value = ''; // Reset input to allow re-upload
            zipObj = null;
            relsMap = {};
            mediaFiles = {};
        }
    }

    // --- Init ---
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', attachUI);
    } else {
        attachUI();
    }

})();
