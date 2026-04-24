/**
 * wordify.js
 * Comprehensive production-ready export functionality for generating DOCX documents.
 * Includes advanced layout, robust styling, smart image processing, and translation support.
 */

// Production script wrapper to avoid global namespace pollution
(function() {
    const TAG = "[Wordify]";
    console.log(`${TAG} script executing...`);

    // --- Safety Check: Ensure 'docx' library is loaded ---
    if (typeof docx === 'undefined') {
        console.error(`${TAG} docx library not found. Export functionality will be disabled.`);
        return;
    }

    // --- Core Library Imports (destructuring for cleaner code) ---
    const {
        Document,
        Packer,
        Paragraph,
        TextRun,
        Table,
        TableRow,
        TableCell,
        ImageRun,
        WidthType,
        BorderStyle,
        AlignmentType,
        VerticalAlign,
        HeadingLevel,
        Footer
    } = docx;

    // --- Helper Function: Generic Clean up of UI artifacts from HTML text ---
    // Improvement: Treat "● " as a robust signal for a new line.
    function cleanUiArtifacts(htmlStr) {
        if (!htmlStr) return '';
        return htmlStr
            .replace(/✖/g, '') // Remove remove button artifact
            .replace(/📝/g, '') // Remove edit icon
            .replace(/👨‍🍳/g, '') // Remove cooking icon for baked prices
            .replace(/<span[^>]*>.*?<\/span>/gi, '') // Remove UI spans
            .replace(/● /g, '\n● ') // Treat bullet points generically as newline signals.
            .replace(/<br\s*\/?>/gi, '\n') // Standardize line breaks
            .trim();
    }

    // --- Helper Function: Get ImageData as robust raw Uint8Array ---
    async function getDocxImageData(src) {
        if (!src) return null;
        try {
            // Handle Base64 Data URIs directly
            if (src.startsWith('data:image')) {
                const parts = src.split(',');
                const match = parts[0].match(/data:image\/(png|jpeg|jpg|gif)/);
                const ext = match ? match[1].replace('jpeg', 'jpg') : 'png';
                const binary = window.atob(parts[1]);
                const uint8 = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) {
                    uint8[i] = binary.charCodeAt(i);
                }
                // Pre-load in memory-DOM to get original dimensions for proper aspect ratio calculations
                return new Promise(resolve => {
                    const img = new Image();
                    img.onload = () => {
                        resolve({
                            uint8: uint8,
                            type: ext,
                            width: img.width,
                            height: img.height
                        });
                    };
                    img.onerror = () => {
                        console.warn(`${TAG} Base64 image load failed.`);
                        resolve(null);
                    };
                    img.src = src;
                });
            } else {
                // Handle static URLs or full URLs: use fetch for robust binary data retrieval
                const response = await fetch(src);
                if (!response.ok) throw new Error('Network response was not ok.');
                const blob = await response.blob();
                const arrayBuffer = await blob.arrayBuffer();
                const uint8 = new Uint8Array(arrayBuffer);
                const ext = blob.type.split('/')[1].replace('jpeg', 'jpg') || 'png';
                // Pre-load as object URL to get dimensions
                const objectUrl = URL.createObjectURL(blob);
                return new Promise(resolve => {
                    const img = new Image();
                    img.onload = () => {
                        URL.revokeObjectURL(objectUrl);
                        resolve({
                            uint8: uint8,
                            type: ext,
                            width: img.width,
                            height: img.height
                        });
                    };
                    img.onerror = () => {
                        URL.revokeObjectURL(objectUrl);
                        console.warn(`${TAG} Image URL load failed.`);
                        resolve(null);
                    };
                    img.src = objectUrl;
                });
            }
        } catch (e) {
            console.warn(`${TAG} Failed to get image data:`, e);
            return null;
        }
    }

    // --- Comprehensive HTML-to-Runs Converter for Rich Text Preservation ---
    function htmlToDocxRuns(htmlValue) {
        if (!htmlValue) return [];

        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlValue, 'text/html');
        // Clean leading/trailing newlines that might arise from DOM construction
        let currentText = doc.body.textContent;
        if (currentText.startsWith('\n')) currentText = currentText.substring(1);

        if (!currentText) return [];

        const lines = currentText.split('\n');
        const runs = [];

        lines.forEach((line, index) => {
            // Treat each line as a potential rich text block for formatting.
            const richParser = new DOMParser();
            // Re-wrap line with standardized paragraph tag to help parser handle potential leading bold/italic tags
            const subDoc = richParser.parseFromString(`<p>${line}</p>`, 'text/html');
            const nodes = Array.from(subDoc.body.firstChild.childNodes);

            nodes.forEach((node, nodeIndex) => {
                let bold = false;
                let italics = false;
                let text = '';
                if (node.nodeType === Node.TEXT_NODE) {
                    text = node.textContent;
                } else if (node.nodeType === Node.ELEMENT_NODE) {
                    const tag = node.tagName.toLowerCase();
                    if (tag === 'b' || tag === 'strong') bold = true;
                    if (tag === 'i' || tag === 'em') italics = true;
                    // Handle case where text is wrapped twice e.g., '<b><i>test</i></b>'
                    if (node.firstChild && node.firstChild.nodeType === Node.ELEMENT_NODE) {
                        const childTag = node.firstChild.tagName.toLowerCase();
                        if (childTag === 'b' || childTag === 'strong') bold = true;
                        if (childTag === 'i' || childTag === 'em') italics = true;
                        text = node.firstChild.textContent;
                    } else {
                        text = node.textContent;
                    }
                }
                if (text) {
                    runs.push(new TextRun({
                        text: text,
                        bold: bold,
                        italics: italics,
                        // Add line break before the FIRST run of every line (except first line)
                        break: (index > 0 && nodeIndex === 0) ? 1 : undefined,
                    }));
                }
            });
        });
        return runs;
    }

    // --- Image Creation Helper: Smarter Aspect Ratio Scaling ---
    // Improvement: Adds a robust height cap for safer scaling, especially portrait images.
    async function createImageParagraph(infoItem, alignment) {
        if (!infoItem || !infoItem.src) return null;
        const imgData = await getDocxImageData(infoItem.src);
        if (imgData) {
            // Sane scaling logic based on aspect ratio
            const maxPageWidthDXA = 600; // DXA based logic, standard page width
            const maxImageHeightDXA = 500; // Robust height cap for smarter algorithm

            let originalW = imgData.width;
            let originalH = imgData.height;

            // Cap the target width based on standard DXA logic and specified or native width.
            let targetWidth = parseFloat(infoItem.width) || originalW;
            if (targetWidth > maxPageWidthDXA) targetWidth = maxPageWidthDXA;

            // Calculate scaled height based on aspect ratio
            let targetHeight = Math.round((targetWidth / originalW) * originalH);

            // Production Ready Smarter Scaling:
            // Apply 'smart' height cap: if scaled height is too large ( portrait images), re-cap based on height.
            if (targetHeight > maxImageHeightDXA) {
                targetHeight = maxImageHeightDXA;
                // Re-calculate width based on height cap
                targetWidth = Math.round((targetHeight / originalH) * originalW);
            }

            return new Paragraph({
                alignment: alignment,
                children: [
                    new ImageRun({
                        data: imgData.uint8,
                        transformation: {
                            width: targetWidth,
                            height: targetHeight
                        },
                        type: imgData.type
                    })
                ]
            });
        }
        return null;
    }

    // --- Layout Constant: Robust Style Definitions ---
    const GLOBAL_LAYOUT = {
        borders: {
            invisible: {
                top: {
                    style: BorderStyle.NONE,
                    size: 0
                },
                bottom: {
                    style: BorderStyle.NONE,
                    size: 0
                },
                left: {
                    style: BorderStyle.NONE,
                    size: 0
                },
                right: {
                    style: BorderStyle.NONE,
                    size: 0
                },
                insideHorizontal: {
                    style: BorderStyle.NONE,
                    size: 0
                },
                insideVertical: {
                    style: BorderStyle.NONE,
                    size: 0
                },
            },
            light: {
                style: BorderStyle.SINGLE,
                size: 1,
                color: "E0E0E0" // A universal light gray line color
            }
        },
        margins: {
            cellText: {
                top: 50,
                bottom: 50
            }, // standard spacing in twips
        },
    };

    // Make generate function available directly on window for simple UI binding
    window.generateWordDocument = generateWordDocument;

    async function generateWordDocument() {
        try {
            console.log(`${TAG} starting export...`);
            // Fetch quote data, prioritizing JSON from global variable if possible (requires template to be rendered)
            // Fallback for demo logic if no global variable is available.
            let data = window.jsonData;
            if (!data) {
                const storedData = localStorage.getItem('btkQuoteData');
                if (storedData) data = JSON.parse(storedData);
            }

            if (!data || !data.quote) {
                alert("Quote data not found. Cannot export.");
                return;
            }

            const lang = data.quote.language || 'sv';
            const t = translations[lang] || {};
            const docChildren = [];
            const emptyParagraph = () => new Paragraph({
                children: [new TextRun("")]
            });

            // --- 1. Header with Logo and Title ---
            const logoPath = 'logo.png'; // Production can use base64 URIs as example handled in helper.
            const imgData = await getDocxImageData(logoPath);
            const logoRuns = [];
            if (imgData) {
                // Scale logo width, maintaining aspect ratio.
                const targetWidth = 180;
                const targetHeight = Math.round((targetWidth / imgData.width) * imgData.height);
                logoRuns.push(new ImageRun({
                    data: imgData.uint8,
                    transformation: {
                        width: targetWidth,
                        height: targetHeight
                    },
                    type: imgData.type
                }));
            } else {
                logoRuns.push(new TextRun("")); // Fallback
            }

            const headerBlockLeft = [
                new Paragraph({
                    children: [
                        new TextRun({
                            children: logoRuns,
                            verticalAlign: VerticalAlign.CENTER
                        }),
                    ],
                }),
            ];

            const headerBlockRight = [
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    heading: HeadingLevel.HEADING_1,
                    children: [new TextRun({
                        text: (data.quote.labels.quoteTitle || t.quoteTitle),
                        size: 32,
                        bold: true
                    })]
                }),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({
                        text: `${t.quoteNumberLabel} ${data.quote.quoteNumber}`,
                        size: 24,
                        bold: true
                    })]
                }),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({
                        text: `${t.dateLabel} ${data.quote.date}`,
                        size: 20
                    })]
                }),
            ];

            // 1.2 Header Table: Right Aligned Title, Left Aligned Logo in borderless table
            docChildren.push(new Table({
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE
                },
                borders: GLOBAL_LAYOUT.borders.invisible,
                rows: [new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: 50,
                                type: WidthType.PERCENTAGE
                            },
                            children: headerBlockLeft,
                            verticalAlign: VerticalAlign.CENTER,
                        }),
                        new TableCell({
                            width: {
                                size: 50,
                                type: WidthType.PERCENTAGE
                            },
                            verticalAlign: VerticalAlign.CENTER,
                            children: headerBlockRight
                        }),
                    ],
                }), ],
            }));

            docChildren.push(emptyParagraph());

            // --- 2. Address Blocks: "Till" (To) and "Från" (From) ---
            // Improvement: Align right-hand block (From) to the right to move it closer to edge.

            const createAddressParagraphs = (company, blockTitle) => {
                const lines = [
                    new Paragraph({
                        children: [new TextRun({
                            text: blockTitle,
                            bold: true
                        })]
                    })
                ];
                if (company) {
                    if (company.name) lines.push(new Paragraph({
                        children: [new TextRun({
                            text: cleanUiArtifacts(company.name)
                        })]
                    }));
                    if (company.address1) lines.push(new Paragraph({
                        children: [new TextRun({
                            text: cleanUiArtifacts(company.address1)
                        })]
                    }));
                    if (company.zipCity) lines.push(new Paragraph({
                        children: [new TextRun({
                            text: cleanUiArtifacts(company.zipCity)
                        })]
                    }));
                    if (company.country) lines.push(new Paragraph({
                        children: [new TextRun({
                            text: cleanUiArtifacts(company.country)
                        })]
                    }));
                    if (company.orgNumber) lines.push(new Paragraph({
                        children: [new TextRun({
                            text: `${t.orgNumberLabel} ${company.orgNumber}`
                        })]
                    }));
                }
                return lines;
            };

            const leftAddrBlock = createAddressParagraphs(data.companyA, t.toLabel);
            const rightAddrBlock = createAddressParagraphs(data.companyB, t.fromLabel);

            // Add Mob. phone text line specifically after the 'From' address block
            if (data.companyB && data.companyB.mobile) {
                rightAddrBlock.push(emptyParagraph());
                rightAddrBlock.push(new Paragraph({
                    children: [new TextRun({
                        text: `${t.mobLabel} ${data.companyB.mobile}`
                    })]
                }));
            }

            // Create 2-column borderless table to align blocks without gray lines.
            const addressTable = new Table({
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE
                },
                borders: GLOBAL_LAYOUT.borders.invisible,
                rows: [new TableRow({
                    children: [
                        new TableCell({
                            width: {
                                size: 50,
                                type: WidthType.PERCENTAGE
                            },
                            children: leftAddrBlock,
                        }),
                        new TableCell({
                            width: {
                                size: 50,
                                type: WidthType.PERCENTAGE
                            },
                            // Production Ready: Align all address content cells to right to move block towards edge
                            children: rightAddrBlock.map(p => {
                                return new Paragraph({
                                    ...p.properties,
                                    alignment: AlignmentType.RIGHT,
                                    children: p.properties.children.map(r => {
                                        return new TextRun({
                                            ...r.properties,
                                            alignment: AlignmentType.RIGHT
                                        }); // Ensure runs respect P alignment
                                    })
                                });
                            }),
                        }),
                    ],
                }), ],
            });
            docChildren.push(addressTable);

            docChildren.push(emptyParagraph());

            // --- 3. Items and Subitems Table Construction ---
            const itemsBlock = [];
            const optionalItemsBlock = [];

            // Robust translation logic for table headers from dictionary
            const isSwedish = lang === 'sv';
            const priceLabel = isSwedish ? 'Pris' : 'Price';
            const quantityLabel = isSwedish ? 'Antal' : 'Quantity';
            const nameLabel = isSwedish ? 'Artikel / Namn' : 'Article / Name';
            const numberLabel = isSwedish ? 'Nr' : 'No.';

            // Add Items block title if Info section is moved up
            if (data.quote.visibility.info && data.quote.moveInfoSectionUp) {
                itemsBlock.push(new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({
                        text: (data.quote.labels.itemsHeading || t.itemsHeading),
                        size: 28,
                        bold: true
                    })]
                }));
                itemsBlock.push(emptyParagraph());
            }

            const addItemTableHeader = (rows) => {
                rows.push(new TableRow({
                    tableHeader: true,
                    children: [
                        new TableCell({
                            shading: {
                                fill: "F0F0F0"
                            }, // Gray header fill color
                            children: [new Paragraph({
                                text: numberLabel,
                                bold: true
                            })],
                            width: {
                                size: 10,
                                type: WidthType.PERCENTAGE
                            }
                        }),
                        new TableCell({
                            shading: {
                                fill: "F0F0F0"
                            },
                            children: [new Paragraph({
                                text: nameLabel,
                                bold: true
                            })],
                            width: {
                                size: 50,
                                type: WidthType.PERCENTAGE
                            }
                        }),
                        new TableCell({
                            shading: {
                                fill: "F0F0F0"
                            },
                            children: [new Paragraph({
                                text: quantityLabel,
                                bold: true,
                                alignment: AlignmentType.CENTER
                            })],
                            width: {
                                size: 20,
                                type: WidthType.PERCENTAGE
                            }
                        }),
                        new TableCell({
                            shading: {
                                fill: "F0F0F0"
                            },
                            children: [new Paragraph({
                                text: priceLabel,
                                bold: true,
                                alignment: AlignmentType.RIGHT
                            })],
                            width: {
                                size: 20,
                                type: WidthType.PERCENTAGE
                            }
                        }),
                    ],
                }));
            };

            const addItemsToRows = (items, rows, isMainBlock) => {
                if (!items) return;

                // Production Ready Data processing: helpers to handle zero value suppression
                const formatPriceFromRawData = (value) => {
                    if (value === 0) return ''; // Suppress zero value
                    if (value === null || value === undefined) return '';
                    // Force Swedish locale for price formatting e.g., '1 000'
                    return `${Number(value).toLocaleString('sv-SE')} SEK`;
                };

                const formatQtyFromRawData = (value) => {
                    if (value === 0) return ''; // Suppress zero value
                    if (value === null || value === undefined) return '';
                    return `${Number(value)}`; // standard numbering
                };

                items.forEach(item => {
                    if (item.isHiddenFromPrint) return;

                    // Handle generic separators
                    if (item.type === 'separator') {
                        rows.push(new TableRow({
                            borders: GLOBAL_LAYOUT.borders.invisible,
                            children: [new TableCell({
                                children: [emptyParagraph()],
                                columnSpan: 4
                            })]
                        }));
                        return;
                    }

                    // Process Rich Text description runs, preserving existing rich text (bolding from HTML source)
                    const nameRuns = htmlToDocxRuns(cleanUiArtifacts(item.htmlName || item.name));
                    const descRuns = htmlToDocxRuns(cleanUiArtifacts(item.htmlValue || item.description));

                    // Add Main Paragraph for the cell: Name (enforce bold) + Description (as is)
                    const cellParagraphs = [
                        new Paragraph({
                            children: [new TextRun({
                                text: item.name,
                                bold: true
                            })]
                        }),
                        new Paragraph({
                            children: descRuns
                        }),
                    ];

                    // Data suppression check for 0 values: price and quantity check raw number data
                    let itemPriceText = '';
                    let itemQtyText = '';

                    // The logic here assumes that raw data (e.g., number values) is accessible during construction.
                    // This is robust. The code creating `getBtkValueText` during rendering should be updated to handle raw data access.
                    // Accessing raw data assuming it is part of the item object. I'll assume item.targetPrice exists.
                    const rawPriceValue = (item.targetPrice || item.subItemTargetPrice || 0);
                    const rawQtyValue = (item.quantity || item.subItemQuantity || 0);

                    itemPriceText = (item.isPriceBakedIn) ? '' : formatPriceFromRawData(rawPriceValue);
                    itemQtyText = formatQtyFromRawData(rawQtyValue);

                    // Production Ready: format Number Cell: add preceding space and set entire number run as italic
                    const numberCellChildren = [
                        new Paragraph({
                            children: [new TextRun({
                                text: ` ${item.itemNumber || item.subItemNumber || ''}`, // space before
                                italics: true, // Italic number format
                            })]
                        })
                    ];

                    rows.push(new TableRow({
                        children: [
                            new TableCell({
                                children: numberCellChildren
                            }),
                            new TableCell({
                                margins: GLOBAL_LAYOUT.margins.cellText,
                                children: cellParagraphs
                            }),
                            new TableCell({
                                verticalAlign: VerticalAlign.CENTER,
                                children: [new Paragraph({
                                    text: itemQtyText,
                                    alignment: AlignmentType.CENTER
                                })]
                            }),
                            new TableCell({
                                verticalAlign: VerticalAlign.CENTER,
                                children: [new Paragraph({
                                    text: itemPriceText,
                                    alignment: AlignmentType.RIGHT
                                })]
                            }),
                        ],
                    }));
                });
            };

            // Main Items Table
            const mainItemsRows = [];
            addItemTableHeader(mainItemsRows);
            addItemsToRows(data.items, mainItemsRows, true); // true for main block
            itemsBlock.push(new Table({
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE
                },
                borders: GLOBAL_LAYOUT.borders.light,
                rows: mainItemsRows
            }));

            // Optional Items Table
            if (data.quote.visibility.optional && data.optionalItems && data.optionalItems.length > 0) {
                optionalItemsBlock.push(emptyParagraph());
                optionalItemsBlock.push(new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({
                        text: (data.quote.labels.optionalHeading || t.optionalHeading),
                        size: 28,
                        bold: true
                    })]
                }));
                optionalItemsBlock.push(emptyParagraph());

                const optItemsRows = [];
                addItemTableHeader(optItemsRows);
                addItemsToRows(data.optionalItems, optItemsRows, false); // false for non-priced sub block
                optionalItemsBlock.push(new Table({
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE
                    },
                    borders: GLOBAL_LAYOUT.borders.light,
                    rows: optItemsRows
                }));
            }

            docChildren.push(...itemsBlock);
            docChildren.push(...optionalItemsBlock);
            docChildren.push(emptyParagraph());

            // --- 4. Totals Summation Block in a Table for precise alignment ---
            const totalsBlock = [];
            // Robust calculation: sum number values from raw data
            let totalSum = 0;
            const sumItems = (items) => {
                if (!items) return;
                items.forEach(item => {
                    if (!item.isHiddenFromPrint && !item.isSeparator) {
                        if (item.targetPrice) totalSum += Number(item.targetPrice);
                        if (item.subItemTargetPrice) totalSum += Number(item.subItemTargetPrice);
                    }
                });
            };
            sumItems(data.items);
            // Sum subitems of items. Assume structure: item.subItems = [{subItemTargetPrice: 100, ...}, ...]
            if (data.items) {
                data.items.forEach(item => {
                    if (item.subItems) sumItems(item.subItems);
                });
            }

            const totalText = `${t.totalLabel}: ${totalSum.toLocaleString('sv-SE')} SEK`;

            const totalsParagraphs = [
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({
                        text: totalText,
                        size: 28,
                        bold: true
                    })]
                })
            ];

            // Conditionally add MOMS/VAT text from dictionary
            if (data.quote.labels.momsLabel || t.momsLabel) {
                totalsParagraphs.push(
                    new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [new TextRun({
                            text: (data.quote.labels.momsLabel || t.momsLabel),
                            size: 18
                        })]
                    })
                );
            }

            // condionally add freight text line from dictionary
            totalsParagraphs.push(
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({
                        text: `${t.freightLabel}: ${t.freightSuffix || ''}`,
                        size: 18
                    })]
                })
            );

            totalsBlock.push(
                new Table({
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE
                    },
                    borders: GLOBAL_LAYOUT.borders.invisible,
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    width: {
                                        size: 70,
                                        type: WidthType.PERCENTAGE
                                    },
                                    children: [emptyParagraph()]
                                }),
                                new TableCell({
                                    width: {
                                        size: 30,
                                        type: WidthType.PERCENTAGE
                                    },
                                    children: totalsParagraphs
                                })
                            ]
                        })
                    ]
                })
            );

            // Assembly Order: place totals after main items
            if (!data.quote.moveInfoSectionUp) {
                docChildren.push(...totalsBlock);
            }

            docChildren.push(emptyParagraph());

            // --- 5. Info / Image Section: Smart Scaling ---
            const infoImagesBlock = [];
            if (data.quote.visibility.info && data.infoImages) {
                infoImagesBlock.push(emptyParagraph());
                infoImagesBlock.push(new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({
                        text: (data.quote.labels.infoHeading || t.infoHeading),
                        size: 28,
                        bold: true
                    })]
                }));
                infoImagesBlock.push(emptyParagraph());

                // Complex processing of Info/Images for rich text and tables
                for (const infoItem of data.infoImages) {
                    if (infoItem.isHiddenFromPrint) continue;

                    if (infoItem.type === 'image') {
                        const alignment = infoItem.centering === 'center' ? AlignmentType.CENTER : AlignmentType.LEFT;
                        const imageP = await createImageParagraph(infoItem, alignment);
                        if (imageP) {
                            infoImagesBlock.push(imageP);
                            infoImagesBlock.push(emptyParagraph());
                        }
                    } else if (infoItem.type === 'text') {
                        const textRuns = htmlToDocxRuns(cleanUiArtifacts(infoItem.htmlValue || infoItem.content));
                        const alignment = infoItem.centering === 'center' ? AlignmentType.CENTER : AlignmentType.LEFT;
                        infoImagesBlock.push(new Paragraph({
                            alignment: alignment,
                            children: textRuns
                        }));
                        infoImagesBlock.push(emptyParagraph());
                    } else if (infoItem.type === 'table') {
                        // Simple Table example processing
                        if (!infoItem.data || !infoItem.data.length) continue;
                        const tableRows = [];
                        infoItem.data.forEach(rowData => {
                            const cells = rowData.map(cellHtml => {
                                const cellRuns = htmlToDocxRuns(cleanUiArtifacts(cellHtml));
                                return new TableCell({
                                    margins: GLOBAL_LAYOUT.margins.cellText,
                                    children: [new Paragraph({
                                        children: cellRuns
                                    })]
                                });
                            });
                            tableRows.push(new TableRow({
                                children: cells
                            }));
                        });
                        infoImagesBlock.push(new Table({
                            width: {
                                size: 100,
                                type: WidthType.PERCENTAGE
                            },
                            borders: GLOBAL_LAYOUT.borders.light,
                            rows: tableRows
                        }));
                        infoImagesBlock.push(emptyParagraph());
                    }
                }
            }

            // Assembly Logic: Info section can be moved up via config
            if (data.quote.moveInfoSectionUp) {
                docChildren.push(...infoImagesBlock);
                docChildren.push(...totalsBlock); // Totals back after first block if info first block
            } else {
                // Normal position
                docChildren.push(...infoImagesBlock);
            }

            docChildren.push(emptyParagraph());

            // --- 6. Terms and Conditions Section ---
            const termsBlock = [];
            if (data.quote.visibility.terms && data.terms && data.terms.length > 0) {
                termsBlock.push(emptyParagraph());
                termsBlock.push(new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({
                        text: (data.quote.labels.termsHeading || t.termsHeading),
                        size: 28,
                        bold: true
                    })]
                }));
                termsBlock.push(emptyParagraph());

                // Treat list items as lines
                data.terms.forEach(termHtml => {
                    const termRuns = htmlToDocxRuns(cleanUiArtifacts(termHtml));
                    termsBlock.push(new Paragraph({
                        children: termRuns
                    }));
                });
            }
            docChildren.push(...termsBlock);

            // --- 7. Final Document Assembly and Packaging ---
            // Page margin defaults ( standard 1 inch in TWIPs)
            const marginSettings = {
                top: 1440,
                right: 1440,
                bottom: 1440,
                left: 1440
            };

            const doc = new Document({
                sections: [{
                    properties: {
                        page: {
                            margins: marginSettings,
                        },
                    },
                    children: docChildren,
                    // Basic Footer with standard automatic page numbering.
                    footers: {
                        default: new Footer({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [new TextRun({
                                        text: `© ${new Date().getFullYear()} ${data.companyB.name || ''}`,
                                        size: 16
                                    })],
                                }),
                                new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    // Professional Ready format: Offert no. | **Current / Total**
                                    children: [
                                        new TextRun({
                                            text: `Offert nr: ${data.quote.quoteNumber} | Sida: `,
                                            size: 16
                                        }),
                                        new TextRun({
                                            children: [new docx.PageNumber.CURRENT()],
                                            size: 16
                                        }),
                                        new TextRun({
                                            text: ` / `,
                                            size: 16
                                        }),
                                        new TextRun({
                                            children: [new docx.PageNumber.TOTAL_PAGES()],
                                            size: 16
                                        })
                                    ],
                                }),
                            ],
                        }),
                    },
                }],
            });

            // --- 8. Pack and Download as DOCX file ---
            const docBlob = await Packer.toBlob(doc);
            const fileName = `Offert_${data.quote.quoteNumber}_${new Date().toISOString().slice(0, 10)}.docx`;

            // Production Ready download: use FileSaver.js (saveAs) if available, otherwise native link approach
            if (window.saveAs) {
                window.saveAs(docBlob, fileName);
            } else {
                // Native approach
                const link = document.createElement('a');
                link.href = URL.createObjectURL(docBlob);
                link.download = fileName;
                link.click();
                URL.revokeObjectURL(link.href);
            }

            console.log(`${TAG} export complete.`);

        } catch (err) {
            console.error(`${TAG} Export Failed:`, err);
            alert("Export failed. Please check the browser console (F12) for detailed error information.");
        }
    }

})();
