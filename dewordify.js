/**
 * dewordify.js - Production Ready Reverse-Parser for Wordify.js
 * Features: Zero-dependency setup (Auto-injects JSZip), Smart Table Recognition,
 * Image Base64 Extraction, Hierarchy Reconstruction, Language Detection.
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
            script.onerror = () => reject(new Error("Kunde inte ladda JSZip-biblioteket. Krävs för att läsa Word-filer."));
            document.head.appendChild(script);
        });
    }

    // --- Utility: Setup UI Binding ---
    function attachUI() {
        let fileInput = document.getElementById('docxUploadInput');
        let btn = document.getElementById('uploadDocxBtn');

        // If not found in HTML, inject dynamically
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

        // Attach listeners safely
        btn.removeEventListener('click', triggerInput);
        btn.addEventListener('click', triggerInput);
        
        fileInput.removeEventListener('change', handleFileSelect);
        fileInput.addEventListener('change', handleFileSelect);

        function triggerInput() {
            fileInput.click();
        }
    }

    // --- XML Parsing Helpers ---
    function getText(node) {
        if (!node) return "";
        return Array.from(node.getElementsByTagName("w:t"))
            .map(t => t.textContent || "")
            .join("");
    }

    function getRuns(node) {
        if (!node) return [];
        const runs = [];
        Array.from(node.getElementsByTagName("w:r")).forEach(rNode => {
            const text = Array.from(rNode.getElementsByTagName("w:t")).map(t => t.textContent).join("");
            if (!text) return;
            const rPr = rNode.getElementsByTagName("w:rPr")[0];
            const isBold = rPr ? rPr.getElementsByTagName("w:b").length > 0 : false;
            const isItalic = rPr ? rPr.getElementsByTagName("w:i").length > 0 : false;
            runs.push({ text, isBold, isItalic });
        });
        return runs;
    }

    function getParagraphs(node) {
        if (!node) return [];
        return Array.from(node.children).filter(c => c.tagName === "w:p");
    }

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
            console.warn(`${TAG} Bildextraktion misslyckades för ID: ${rId}`, e);
            return null;
        }
    }

    function findImageRelId(pNode) {
        const blip = pNode.getElementsByTagName("a:blip")[0];
        if (blip) return blip.getAttribute("r:embed");
        return null;
    }

    function cleanNumber(str) {
        if (!str) return 0;
        // Strip spaces (thousand separators) and replace comma with dot
        const clean = str.replace(/\s+/g, '').replace(/[^\d.,-]/g, '').replace(',', '.');
        return parseFloat(clean) || 0;
    }

    // --- Main Processor ---
    async function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        const msgEl = document.getElementById('statusMessage');
        if (msgEl) {
            msgEl.textContent = "Laddar och tolkar Word-dokument...";
            msgEl.style.color = "#2b579a";
        }

        try {
            await ensureJSZip();
            zipObj = await window.JSZip.loadAsync(file);
            
            // 1. Build Relationships Map
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

            // Cache media files
            zipObj.folder("word/media").forEach((relativePath, fileObj) => {
                mediaFiles[`media/${relativePath}`] = fileObj;
            });

            // 2. Load Main Document XML
            const docFile = zipObj.file("word/document.xml");
            if (!docFile) throw new Error("Ogiltig Word-fil: word/document.xml saknas.");
            const docXml = await docFile.async("text");
            const doc = new DOMParser().parseFromString(docXml, "application/xml");
            const body = doc.getElementsByTagName("w:body")[0];

            // 3. Rebuild Data Object
            const extractedData = {
                quote: { language: 'sv', visibility: { optional: false, info: false, terms: false } },
                companyA: {},
                companyB: {},
                items: [],
                optionalItems: [],
                infoImages: [],
                terms: []
            };

            let currentSection = 'header'; // header -> address -> items -> optional -> info -> terms
            
            const nodes = Array.from(body.children);
            for (let i = 0; i < nodes.length; i++) {
                const node = nodes[i];
                const nodeType = node.tagName;

                if (nodeType === "w:tbl") {
                    const rows = Array.from(node.getElementsByTagName("w:tr"));
                    if (rows.length === 0) continue;
                    
                    const firstRowText = getText(rows[0]).toLowerCase();

                    // Detect: Header Table
                    if (currentSection === 'header' && (firstRowText.includes("datum") || firstRowText.includes("date"))) {
                        const cells = rows[0].getElementsByTagName("w:tc");
                        if (cells.length >= 2) {
                            const logoId = findImageRelId(cells[0]);
                            if (logoId) {
                                const logoBase64 = await extractImageBase64(logoId);
                                if (logoBase64) localStorage.setItem('companyLogo', logoBase64);
                            }
                            
                            const metaText = getText(cells[1]);
                            const nrMatch = metaText.match(/(?:Nr|No):\s*(\S+)/i);
                            const dateMatch = metaText.match(/(?:Datum|Date):\s*([\d-]+)/i);
                            
                            if (nrMatch) extractedData.quote.quoteNumber = nrMatch[1];
                            if (dateMatch) extractedData.quote.date = dateMatch[1];
                        }
                        currentSection = 'address';
                        continue;
                    }

                    // Detect: Address Table
                    if (currentSection === 'address' && (firstRowText.includes("till:") || firstRowText.includes("to:"))) {
                        extractedData.quote.language = firstRowText.includes("to:") ? 'en' : 'sv';
                        
                        const cells = rows[0].getElementsByTagName("w:tc");
                        if (cells.length >= 3) {
                            getParagraphs(cells[0]).forEach((p, idx) => {
                                const txt = getText(p).replace(/Till:|To:/gi, '').trim();
                                if (txt) extractedData.companyA[`line${idx + 1}`] = txt;
                            });

                            getParagraphs(cells[2]).forEach((p, idx) => {
                                const txt = getText(p).replace(/Från:|From:/gi, '').trim();
                                if (txt) extractedData.companyB[`line${idx + 1}`] = txt;
                            });
                        }
                        currentSection = 'items';
                        continue;
                    }

                    // Detect: Items or Optional Items
                    if (firstRowText.includes("antal") || firstRowText.includes("quantity")) {
                        const targetArray = currentSection === 'optional' ? extractedData.optionalItems : extractedData.items;
                        if (currentSection === 'optional') extractedData.quote.visibility.optional = true;
                        
                        let lastMainItem = null;

                        for (let r = 1; r < rows.length; r++) {
                            try {
                                const cells = Array.from(rows[r].getElementsByTagName("w:tc"));
                                if (cells.length < 4) continue;

                                const nrText = getText(cells[0]).trim();
                                const qty = cleanNumber(getText(cells[2]));
                                const priceText = getText(cells[3]);
                                const price = cleanNumber(priceText);

                                // Stop at totals row
                                if (getText(cells[2]).toLowerCase().includes("total") || getText(cells[1]).toLowerCase().includes("total")) {
                                    const currencyMatch = priceText.match(/[a-zA-Z]{3}$/);
                                    if (currencyMatch && currencyMatch[0].toUpperCase() !== "SEK") {
                                        extractedData.quote.useCustomCurrency = true;
                                        extractedData.quote.customCurrency = currencyMatch[0].toUpperCase();
                                    }
                                    break; 
                                }

                                const cell1Paras = getParagraphs(cells[1]);
                                let name = "";
                                let desc = [];
                                
                                cell1Paras.forEach((p, pIdx) => {
                                    const runs = getRuns(p);
                                    const pText = runs.map(r => r.text).join("").trim();
                                    if (!pText) return;

                                    const hasBold = runs.some(r => r.isBold);
                                    if (pIdx === 0 || (hasBold && !name)) {
                                        name += (name ? " " : "") + pText;
                                    } else {
                                        desc.push(pText);
                                    }
                                });

                                const isSubItem = nrText === "" || nrText.includes(".") || getRuns(cells[0]).some(r => r.isItalic);

                                if (isSubItem && lastMainItem) {
                                    lastMainItem.subItems.push({
                                        subItemNumber: nrText,
                                        subItemName: name,
                                        subItemDescription: desc.join("\n"),
                                        subItemQuantity: qty,
                                        subItemTargetPrice: price,
                                        subItemOriginalPrice: qty ? price / qty : 0,
                                        isPriceBakedIn: price === 0 && priceText.trim() === "",
                                        isHiddenFromPrint: false,
                                        vendorNotes: ""
                                    });
                                } else {
                                    const newItem = {
                                        type: 'item',
                                        itemNumber: nrText,
                                        name: name,
                                        itemDescription: desc.join("\n"),
                                        quantity: qty,
                                        targetPrice: price,
                                        originalPrice: qty ? price / qty : 0,
                                        isPriceBakedIn: price === 0 && priceText.trim() === "",
                                        isHiddenFromPrint: false,
                                        vendorNotes: "",
                                        subItems: []
                                    };
                                    targetArray.push(newItem);
                                    lastMainItem = newItem;
                                }
                            } catch (rowErr) {
                                console.warn(`${TAG} Fel vid parsning av artikelrad:`, rowErr);
                            }
                        }
                        continue;
                    }

                    // Detect: Info Table
                    if (currentSection === 'info') {
                        extractedData.quote.visibility.info = true;
                        const imageIds = Array.from(node.getElementsByTagName("a:blip")).map(b => b.getAttribute("r:embed"));
                        
                        if (imageIds.length > 0) {
                            for (let id of imageIds) {
                                const base64 = await extractImageBase64(id);
                                if (base64) {
                                    extractedData.infoImages.push({
                                        type: 'image',
                                        src: base64,
                                        width: 300,
                                        centering: 'center',
                                        compressionImmune: false
                                    });
                                }
                            }
                        } else {
                            const tblRows = [];
                            rows.forEach(r => {
                                const rowData = [];
                                Array.from(r.getElementsByTagName("w:tc")).forEach(c => rowData.push(getText(c).trim()));
                                if(rowData.some(v => v)) tblRows.push(rowData);
                            });
                            
                            if (tblRows.length > 0) {
                                extractedData.infoImages.push({
                                    type: 'table',
                                    rows: tblRows.length,
                                    cols: tblRows[0].length,
                                    data: tblRows,
                                    centering: 'center'
                                });
                            }
                        }
                    }
                } 
                else if (nodeType === "w:p") {
                    const text = getText(node).trim();
                    const imgId = findImageRelId(node);

                    if (!text && !imgId) continue;

                    const lowerText = text.toLowerCase();

                    // Broad transition checks
                    if (/(alternativ|alternative)/i.test(lowerText) && currentSection === 'items') {
                        currentSection = 'optional';
                        continue;
                    }
                    if (/(info|bilder|images)/i.test(lowerText) && (currentSection === 'items' || currentSection === 'optional')) {
                        currentSection = 'info';
                        continue;
                    }
                    if (/(villkor|terms)/i.test(lowerText)) {
                        currentSection = 'terms';
                        continue;
                    }

                    if (currentSection === 'info') {
                        extractedData.quote.visibility.info = true;
                        if (imgId) {
                            const base64 = await extractImageBase64(imgId);
                            if (base64) {
                                extractedData.infoImages.push({
                                    type: 'image',
                                    src: base64,
                                    width: 400,
                                    centering: 'center',
                                    compressionImmune: false
                                });
                            }
                        } else if (text) {
                            extractedData.infoImages.push({
                                type: 'text',
                                content: text.replace(/\n/g, '<br>'),
                                centering: 'center'
                            });
                        }
                    } 
                    else if (currentSection === 'terms' && text) {
                        extractedData.quote.visibility.terms = true;
                        extractedData.terms.push(text);
                    }
                }
            }

            // 4. Inject Data Back to App
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
                setTimeout(() => { if(msgEl.textContent.includes("Word-offert")) msgEl.textContent = ''; }, 3000);
            }

        } catch (err) {
            console.error(`${TAG} ❌ Kritiskt Fel:`, err);
            const msgEl = document.getElementById('statusMessage');
            if (msgEl) {
                msgEl.textContent = "Kunde inte tolka Word-filen. Fel: " + err.message;
                msgEl.style.color = "red";
            }
            alert("Ett fel uppstod vid inläsning av Word-filen.\nKontrollera att det är en offert skapad med detta system.");
        } finally {
            event.target.value = ''; // Ensure the input is reset so the same file can be loaded again
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
