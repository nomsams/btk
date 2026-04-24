/**
 * wordify.js - Complete Production-Ready DOCX Generator
 * Depends on: docx (library must be loaded via script tag beforehand)
 */

(function () {
    // Standard page content width safe zones in dxa (Twips)
    // Assumes A4 or US Letter with standard ~1 inch margins
    const PAGE_CONTENT_WIDTH = 9360; 
    
    // !!! CHANGE THIS TO YOUR ACTUAL LOGO PATH (can be relative or base64) !!!
    const logoUrl = 'path/to/your/logo.png'; 

    /**
     * Helper to get image dimensions async before adding to doc.
     * Prevents distortion by maintaining aspect ratio.
     */
    async function getImageDimensions(url) {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                resolve({ width: img.naturalWidth, height: img.naturalHeight });
            };
            img.onerror = () => {
                console.error(`Failed to load image: ${url}`);
                resolve(null); // Return null so script doesn't crash, just skips image
            };
            img.src = url;
        });
    }

    /**
     * Goal 3 & Smart Scaling Helper:
     * Returns ImageRun options with calculated dimensions based on target width.
     */
    async function createScaledImage(url, targetWidthDxa) {
        const dims = await getImageDimensions(url);
        if (!dims) return null;

        // DOCX handles images in pts internally (1pt = 20 dxa)
        const targetWidthPt = targetWidthDxa / 20; 
        
        // Calculate aspect ratio
        const aspectRatio = dims.width / dims.height;
        
        // Final pixel dimensions based on target width
        const finalWidth = targetWidthPt;
        const finalHeight = targetWidthPt / aspectRatio;

        // Fetch the raw data
        const response = await fetch(url);
        const data = await response.arrayBuffer();

        return new docx.ImageRun({
            data: data,
            transformation: {
                width: finalWidth,
                height: finalHeight,
            },
        });
    }

    /**
     * Goal 2: String parsing helper.
     * Converts raw text descriptions into docx Paragraphs/Runs, 
     * forcing new lines before "● " or standard newlines.
     */
    function parseDescriptionToRuns(text) {
        if (!text) return [new docx.TextRun("")];
        
        // 1. Force new lines before bullet points that might be squished
        // 2. Normalize existing standard newlines
        // 3. Handle HTML-like line breaks if they exist
        let normalizedText = text
            .replace(/<br\s*\/?>/gi, '\n') // Handle <br> just in case
            .replace(/●\s+/g, '\n● ')     // Force newline before bullet+space
            .trim();

        // Split by existing or generated newlines
        const lines = normalizedText.split('\n');
        const runs = [];

        lines.forEach((line, index) => {
            // Trim leading spaces from the start of a line (except maybe the bullet itself)
            let trimmedLine = line.trim();
            if (trimmedLine.startsWith('●')) {
                 trimmedLine = trimmedLine.replace('●', '●\u00A0'); // Ensure non-breaking space after bullet
            }

            runs.push(new docx.TextRun({
                text: trimmedLine,
                // Add a line break property to every line except the very last one
                break: index < lines.length - 1 ? 1 : 0 
            }));
        });

        return runs;
    }

    /**
     * Goal 5 Helper: Formatter for Price/Quantity cells.
     * If 0, returns space. Otherwise, normal formatted value.
     */
    function formatVal(val, formatterFn) {
        const numVal = parseFloat(val);
        if (isNaN(numVal) || numVal === 0) {
            return " "; // Empty cell if zero or nan
        }
        return formatterFn ? formatterFn(numVal) : String(numVal);
    }

    // Main currency formatter for consistent looks
    const currencyFormatter = new Intl.NumberFormat('sv-SE', {
        style: 'currency',
        currency: 'SEK',
        minimumFractionDigits: 0, 
    });

    /**
     * Main wordify function. Needs to be async to handle image dimension pre-loading.
     */
    async function wordify(jsonData) {
        console.log("Generating DOCX with improved layout rules...");

        if (!jsonData || !jsonData.quote) {
            alert("Feil: Ingen offert-data funnet.");
            return;
        }

        const quote = jsonData.quote;
        const children = []; // The list of elements for the final section

        // --- Styles Definition (for later reuse) ---
        const styles = {
            tableBorders: {
                top: { style: docx.BorderStyle.SINGLE, size: 1, color: "E0E0E0" },
                bottom: { style: docx.BorderStyle.SINGLE, size: 1, color: "E0E0E0" },
                insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1, color: "E0E0E0" },
                // Remove vertical borders for cleaner look like image_2.png
                left: { style: docx.BorderStyle.NONE }, 
                right: { style: docx.BorderStyle.NONE },
                insideVertical: { style: docx.BorderStyle.NONE },
            },
            noBorders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE },
                insideVertical: { style: docx.BorderStyle.NONE },
                insideHorizontal: { style: docx.BorderStyle.NONE },
            },
        };


        // --- Header Section ---
        // Top right details block (Title, Nr, Datum)
        children.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({ text: "Offert", bold: true, size: 32 }), // larger, bold
                ],
                alignment: docx.AlignmentType.RIGHT,
            })
        );
        children.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({ text: "Offert nr: ", bold: true }),
                    new docx.TextRun(quote.quoteNumber || "-"),
                ],
                alignment: docx.AlignmentType.RIGHT,
                spacing: { before: 120 } // slight gap below title
            })
        );
        children.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({ text: "Datum: ", bold: true }),
                    new docx.TextRun(quote.date || new Date().toISOString().split('T')[0]),
                ],
                alignment: docx.AlignmentType.RIGHT,
                spacing: { after: 400 } // gap below details block
            })
        );


        // --- Logo and Top Addresses Table (Goal 1 Fix) ---
        // Fetch and scale logo (target width: roughly 1/3 of page)
        const scaledLogoRun = await createScaledImage(logoUrl, 3200);

        // Define specific widths to ensure standard DOCX layout
        // Col 1 is Logo (roughly 40%)
        // Col 2 is From address (roughly 35%) -> pushed right
        // Col 3 is spacer / edge of "Till" block (roughly 25%)
        const topTableWidths = [
            (PAGE_CONTENT_WIDTH * 0.40),
            (PAGE_CONTENT_WIDTH * 0.35),
            (PAGE_CONTENT_WIDTH * 0.25),
        ];

        children.push(
            new docx.Table({
                width: { size: PAGE_CONTENT_WIDTH, type: docx.WidthType.DXA },
                borders: styles.noBorders,
                rows: [
                    new docx.TableRow({
                        children: [
                            // Column 1: Logo
                            new docx.TableCell({
                                width: { size: topTableWidths[0], type: docx.WidthType.DXA },
                                children: scaledLogoRun ? [new docx.Paragraph({ children: [scaledLogoRun] })] : [],
                            }),
                            // Column 2: empty spacer cell
                            new docx.TableCell({ width: { size: topTableWidths[1], type: docx.WidthType.DXA }, children: [] }),
                            // Column 3: "Från" block (pushed to right edge like image_2)
                            new docx.TableCell({
                                width: { size: topTableWidths[2], type: docx.WidthType.DXA },
                                children: [
                                    new docx.Paragraph({
                                        children: [new docx.TextRun({ text: "Från:", bold: true })],
                                        spacing: { after: 100 }
                                    }),
                                    // Map list items to paragraphs
                                    ...(quote.fromAddress || []).map(line => 
                                        new docx.Paragraph({ children: [new docx.TextRun(line)] })
                                    ),
                                    // Contact details
                                    new docx.Paragraph({ children: [new docx.TextRun({text: `Mob: ${quote.fromPhone || ""}`, break: 1})] }),
                                ],
                            }),
                        ],
                    }),
                ],
                spacing: { after: 300 } // gap below logo row
            })
        );


        // --- "Till" Address Block (Below logo row) ---
        children.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({ text: "Till:", bold: true }),
                    new docx.TextRun({ text: quote.toCompany || "", break: 1 }),
                    new docx.TextRun({ text: quote.toContactPerson || "", break: 1 }),
                    new docx.TextRun({ text: quote.toAddressLine1 || "", break: 1 }),
                    new docx.TextRun({ text: quote.toAddressLine2 || "", break: 1 }),
                    new docx.TextRun({ text: quote.toPostalCodeCity || "", break: 1 }),
                    quote.toOrgNumber ? new docx.TextRun({ text: `Org. nr: ${quote.toOrgNumber}`, break: 1 }) : null,
                ],
                spacing: { after: 600 } // large gap before item table
            })
        );


        // --- Main Items Table ---
        // CRITICAL: Define explicit column widths in DXA to prevent compressed columns (fixes image_0.png fail)
        // Values are examples based on standard layout ratios
        const colWidthsItems = {
            nr: 720,        // small width for item number
            name: 5800,     // wide for descriptions
            qty: 900,       // centered qty
            price: 1940,    // right aligned price
        };
        const tableHeaderShading = { fill: "F5F5F5", type: docx.ShadingType.CLEAR, color: "000000" };

        const itemTableRows = [
            // Table Header Row
            new docx.TableRow({
                tableHeader: true,
                children: [
                    new docx.TableCell({
                        width: { size: colWidthsItems.nr, type: docx.WidthType.DXA },
                        shading: tableHeaderShading,
                        children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "Nr", bold: true })] })],
                    }),
                    new docx.TableCell({
                        width: { size: colWidthsItems.name, type: docx.WidthType.DXA },
                        shading: tableHeaderShading,
                        children: [new docx.Paragraph({ children: [new TextRun({ text: "Artikel / Namn", bold: true })] })],
                    }),
                    new docx.TableCell({
                        width: { size: colWidthsItems.qty, type: docx.WidthType.DXA },
                        shading: tableHeaderShading,
                        children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "Antal", bold: true })], alignment: docx.AlignmentType.CENTER })],
                    }),
                    new docx.TableCell({
                        width: { size: colWidthsItems.price, type: docx.WidthType.DXA },
                        shading: tableHeaderShading,
                        children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "Pris", bold: true })], alignment: docx.AlignmentType.RIGHT })],
                    }),
                ],
            }),
        ];

        let totalPrice = 0;

        // Loop Main Items
        (jsonData.items || []).forEach(item => {
            if (item.isHidden) return; // Skip hidden items

            totalPrice += parseFloat(item.totalPrice || 0);

            // Goal 2 implemented here via parseDescriptionToRuns helper
            const descriptionRuns = parseDescriptionToRuns(item.description);

            // Create Main Item Row
            itemTableRows.push(
                new docx.TableRow({
                    children: [
                        // Nr
                        new docx.TableCell({ children: [new docx.Paragraph({ text: item.itemNumber || " " })] }),
                        // Name & Details (Goal 2 text representation)
                        new docx.TableCell({
                            children: [
                                new docx.Paragraph({ children: [new docx.TextRun({ text: item.name || " ", bold: true })] }), // Title Bold
                                new docx.Paragraph({ children: descriptionRuns, spacing: { before: 100 } }), // Details standard spacing
                            ],
                            verticalAlign: docx.VerticalAlign.TOP,
                            margins: { top: 120, bottom: 120 },
                        }),
                        // Antal (Goal 5: blank if 0)
                        new docx.TableCell({
                            children: [new docx.Paragraph({ 
                                children: [new docx.TextRun(formatVal(item.quantity, null))], 
                                alignment: docx.AlignmentType.CENTER 
                            })],
                            verticalAlign: docx.VerticalAlign.TOP,
                            margins: { top: 120, bottom: 120 },
                        }),
                        // Pris (Goal 5: blank if 0)
                        new docx.TableCell({
                            children: [new docx.Paragraph({ 
                                children: [new docx.TextRun(formatVal(item.totalPrice, val => currencyFormatter.format(val)))], 
                                alignment: docx.AlignmentType.RIGHT 
                            })],
                            verticalAlign: docx.VerticalAlign.TOP,
                            margins: { top: 120, bottom: 120 },
                        }),
                    ],
                })
            );

            // Loop Sub-items (Goal 4 Implemented Here)
            (item.subItems || []).forEach(subItem => {
                if (subItem.isHidden) return; // Skip hidden

                const subDescriptionRuns = parseDescriptionToRuns(subItem.description);
                const subItemNrFormatted = " " + (subItem.subNumber || ""); // leading space prefix

                itemTableRows.push(
                    new docx.TableRow({
                        children: [
                            // Nr Cell (Goal 4: Italic and with space prefix)
                            new docx.TableCell({
                                children: [new docx.Paragraph({
                                    children: [new docx.TextRun({ text: subItemNrFormatted, italics: true })] 
                                })],
                            }),
                            // Name & Details Cell (Goal 4: bold title, normal description)
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({ children: [new docx.TextRun({ text: subItem.name || " ", bold: true })] }), // bold title
                                    subDescriptionRuns.length > 0 && subItem.description // only add desc if exists
                                      ? new docx.Paragraph({ children: subDescriptionRuns, spacing: { before: 80 } }) // keep description normal
                                      : null,
                                ],
                                verticalAlign: docx.VerticalAlign.TOP,
                                margins: { top: 100, bottom: 100 },
                            }),
                            // Antal Cell (blank if 0)
                            new docx.TableCell({
                                children: [new docx.Paragraph({ 
                                    children: [new docx.TextRun(formatVal(subItem.quantity, null))], 
                                    alignment: docx.AlignmentType.CENTER 
                                })],
                                verticalAlign: docx.VerticalAlign.TOP,
                                margins: { top: 100, bottom: 100 },
                            }),
                            // Price Cell (Sub-items often don't have separate price visible, or it's 0 if bundled)
                            new docx.TableCell({
                                children: [new docx.Paragraph({ 
                                    // Using common bundled price scenario: show price if > 0, else blank if 0
                                    children: [new docx.TextRun(formatVal(subItem.totalPrice, val => currencyFormatter.format(val)))], 
                                    alignment: docx.AlignmentType.RIGHT 
                                })],
                                verticalAlign: docx.VerticalAlign.TOP,
                                margins: { top: 100, bottom: 100 },
                            }),
                        ],
                    })
                );
            });
        });

        // Add Total Row
        itemTableRows.push(
            new docx.TableRow({
                children: [
                    new TableCell({ children: [] }), // spacer nr
                    new TableCell({ // Total Label block
                        children: [new Paragraph({
                            children: [new TextRun({ text: "TOTALT ATT BETALA (exkl moms)", bold: true })],
                            alignment: AlignmentType.RIGHT,
                        })],
                        shading: { fill: "fcfcfc" } // light accent
                    }),
                    new TableCell({ children: [] }), // spacer qty
                    new docx.TableCell({ // Total Price Cell
                        children: [new docx.Paragraph({
                            children: [new docx.TextRun({ text: currencyFormatter.format(totalPrice), bold: true })],
                            alignment: docx.AlignmentType.RIGHT,
                        })],
                        verticalAlign: docx.VerticalAlign.CENTER,
                        margins: { top: 120, bottom: 120 },
                        shading: { fill: "fcfcfc" }
                    }),
                ],
            })
        );

        // Add Table to document
        children.push(
            new docx.Table({
                width: { size: PAGE_CONTENT_WIDTH, type: docx.WidthType.DXA },
                borders: styles.tableBorders, // borders horizontal only
                rows: itemTableRows,
                spacing: { after: 800 } // gap after table before terms
            })
        );


        // --- Villkor (Terms) Section (similar to image_1.png) ---
        if (quote.terms && quote.terms.length > 0) {
            children.push(
                new docx.Paragraph({
                    children: [new docx.TextRun({ text: "Villkor", bold: true, size: 28 })], // Larger section header
                    spacing: { after: 200 }
                })
            );

            quote.terms.forEach(term => {
                children.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({ text: term.label || "", bold: true }), // bold label e.g., "Giltighet:"
                            new docx.TextRun({ text: ` ${term.text || ""}` }), // normal text
                        ],
                        indent: { left: 400, hanging: 400 }, // clean hanging indent formatting
                        spacing: { after: 120 } // gap between terms
                    })
                );
            });
        }


        // --- Packaging and Returning Document Object ---
        return new docx.Document({
            sections: [{
                properties: {},
                children: children,
            }],
        });
    }

    // --- Integration with standard UI download pattern ---
    async function handleDocxDownload(jsonData) {
        try {
            const doc = await wordify(jsonData);
            const quoteNumber = jsonData.quote.quoteNumber || "draft";
            const fileName = `Offert_${quoteNumber}.docx`;

            docx.Packer.toBlob(doc).then(blob => {
                saveAs(blob, fileName);
                console.log("DOCX generated successfully!");
            });
        } catch (error) {
            console.error("Error generating DOCX:", error);
            alert("Ett fel uppstod vid generering av DOCX-filen. Kontrollera webbläsarens konsol för detaljer.");
        }
    }

    // Bind function to window for UI usage
    window.wordifyData = handleDocxDownload;

})();
