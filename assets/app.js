// assets/app.js

// Co-Author Management
let coAuthorCount = 0;

document.getElementById('addCoAuthor').addEventListener('click', function() {
  coAuthorCount++;
  const coAuthorsList = document.getElementById('coAuthorsList');

  const coAuthorDiv = document.createElement('div');
  coAuthorDiv.className = 'coauthor-entry';
  coAuthorDiv.id = `coauthor-${coAuthorCount}`;
  coAuthorDiv.innerHTML = `
    <input type="text" placeholder="Last Name" class="coauthor-lastname" required>
    <input type="text" placeholder="First Name" class="coauthor-firstname" required>
    <input type="text" placeholder="Phone (e.g., 44-7398344190)" class="coauthor-phone" required>
    <button type="button" class="remove-coauthor" onclick="removeCoAuthor(${coAuthorCount})">Remove</button>
  `;

  coAuthorsList.appendChild(coAuthorDiv);
});

function removeCoAuthor(id) {
  const coAuthorDiv = document.getElementById(`coauthor-${id}`);
  if (coAuthorDiv) coAuthorDiv.remove();
}

// ====== Equity Research section toggle (NEW) ======
const noteTypeEl = document.getElementById('noteType');
const equitySectionEl = document.getElementById('equitySection');

function toggleEquitySection() {
  const isEquity = noteTypeEl.value === 'Equity Research';
  equitySectionEl.style.display = isEquity ? 'block' : 'none';
}
noteTypeEl.addEventListener('change', toggleEquitySection);
toggleEquitySection();

// Format date and time
function formatDateTime(date) {
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const month = months[date.getMonth()];
  const day = date.getDate();
  const year = date.getFullYear();

  let hours = date.getHours();
  const minutes = date.getMinutes().toString().padStart(2, '0');
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12 || 12;

  return `${month} ${day}, ${year} ${hours}:${minutes} ${ampm}`;
}

// Process and add images
async function addImages(files) {
  const imageParagraphs = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    try {
      console.log(`Processing image ${i + 1}: ${file.name}`);
      const arrayBuffer = await file.arrayBuffer();
      const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, "");

      imageParagraphs.push(
        new docx.Paragraph({
          children: [
            new docx.ImageRun({
              data: arrayBuffer,
              transformation: { width: 600, height: 450 }
            })
          ],
          spacing: { before: 200, after: 100 },
          alignment: docx.AlignmentType.CENTER
        }),
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: `Figure ${i + 1}: ${fileNameWithoutExt}`,
              italics: true,
              size: 18,
              font: "Book Antiqua"
            })
          ],
          spacing: { after: 300 },
          alignment: docx.AlignmentType.CENTER
        })
      );
    } catch (error) {
      console.error(`Error processing image ${file.name}:`, error);
    }
  }
  return imageParagraphs;
}

// Helpers: split into paragraphs preserving blank lines
function linesToParagraphs(text, spacingAfter = 150) {
  const lines = (text || "").split('\n');
  return lines.map(line => {
    if (line.trim() === '') {
      return new docx.Paragraph({ text: "", spacing: { after: spacingAfter } });
    }
    return new docx.Paragraph({ text: line, spacing: { after: spacingAfter } });
  });
}

// Create Word Document
async function createDocument(data) {
  const {
    noteType,
    title,
    topic,
    authorLastName,
    authorFirstName,
    authorPhone,
    coAuthors,
    analysis,
    keyTakeaways,
    content,
    cordobaView,
    imageFiles,
    dateTimeString,
    // equity fields (NEW)
    ticker,
    modelType,
    valuationSummary,
    keyAssumptions,
    scenarioNotes,
    modelFiles
  } = data;

  // Key takeaways bullets (preserve empty lines)
  const takeawayLines = (keyTakeaways || "").split('\n');
  const takeawayBullets = takeawayLines.map(line => {
    if (line.trim() === '') {
      return new docx.Paragraph({ text: "", spacing: { after: 100 } });
    }
    const cleanLine = line.replace(/^[-*•]\s*/, '').trim();
    return new docx.Paragraph({
      text: cleanLine,
      bullet: { level: 0 },
      spacing: { after: 100 }
    });
  });

  const analysisParagraphs = linesToParagraphs(analysis, 150);
  const contentParagraphs = linesToParagraphs(content, 150);
  const cordobaViewParagraphs = linesToParagraphs(cordobaView, 150);

  // Images
  console.log('About to call addImages...');
  const imageParagraphs = await addImages(imageFiles);
  console.log('addImages completed, paragraphs:', imageParagraphs.length);

  // Info table for title/topic/authors
  const infoTable = new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: docx.BorderStyle.NONE },
      bottom: { style: docx.BorderStyle.NONE },
      left: { style: docx.BorderStyle.NONE },
      right: { style: docx.BorderStyle.NONE },
      insideHorizontal: { style: docx.BorderStyle.NONE },
      insideVertical: { style: docx.BorderStyle.NONE }
    },
    rows: [
      new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [
              new docx.Paragraph({
                text: title,
                bold: true,
                size: 28,
                spacing: { after: 100 }
              })
            ],
            width: { size: 60, type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.TOP
          }),
          new docx.TableCell({
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    text: `${authorLastName.toUpperCase()}, ${authorFirstName.toUpperCase()} (${authorPhone})`,
                    bold: true,
                    size: 28
                  })
                ],
                alignment: docx.AlignmentType.RIGHT,
                spacing: { after: 100 }
              })
            ],
            width: { size: 40, type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.TOP
          })
        ]
      }),
      new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({ text: "TOPIC: ", bold: true, size: 28 }),
                  new docx.TextRun({ text: topic, bold: true, size: 28 })
                ],
                spacing: { after: 200 }
              })
            ],
            width: { size: 60, type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.TOP
          }),
          new docx.TableCell({
            children: coAuthors.length > 0 ? coAuthors.map(coAuthor =>
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    text: `${coAuthor.lastName.toUpperCase()}, ${coAuthor.firstName.toUpperCase()} (${coAuthor.phone})`,
                    bold: true,
                    size: 28
                  })
                ],
                alignment: docx.AlignmentType.RIGHT,
                spacing: { after: 100 }
              })
            ) : [new docx.Paragraph({ text: "" })],
            width: { size: 40, type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.TOP
          })
        ]
      })
    ]
  });

  // Build doc children
  const documentChildren = [
    infoTable,

    // Horizontal line
    new docx.Paragraph({
      border: {
        bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 }
      },
      spacing: { after: 300 }
    })
  ];

  // ====== Equity Research section in DOC (NEW) ======
  if (noteType === "Equity Research") {
    // Create a clean “Model Pack” block, plus list of attached model files by name
    const attachedModelNames = (modelFiles && modelFiles.length)
      ? Array.from(modelFiles).map(f => f.name)
      : [];

    documentChildren.push(
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: "Equity Research — Financial Model Pack", bold: true, size: 24, font: "Book Antiqua" })
        ],
        spacing: { after: 200 }
      }),

      ...(ticker && ticker.trim()
        ? [new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Ticker / Company: ", bold: true }),
              new docx.TextRun({ text: ticker.trim() })
            ],
            spacing: { after: 120 }
          })]
        : []),

      ...(modelType && modelType.trim()
        ? [new docx.Paragraph({
            children: [
              new docx.TextRun({ text: "Model Type: ", bold: true }),
              new docx.TextRun({ text: modelType.trim() })
            ],
            spacing: { after: 150 }
          })]
        : []),

      ...(attachedModelNames.length
        ? [
            new docx.Paragraph({
              children: [new docx.TextRun({ text: "Attached model files:", bold: true })],
              spacing: { after: 120 }
            }),
            ...attachedModelNames.map((name) =>
              new docx.Paragraph({
                text: name,
                bullet: { level: 0 },
                spacing: { after: 80 }
              })
            ),
            new docx.Paragraph({ spacing: { after: 150 } })
          ]
        : [
            new docx.Paragraph({
              children: [new docx.TextRun({ text: "Attached model files: ", bold: true }), new docx.TextRun({ text: "None uploaded" })],
              spacing: { after: 150 }
            })
          ]),

      ...(valuationSummary && valuationSummary.trim()
        ? [
            new docx.Paragraph({
              children: [new docx.TextRun({ text: "Valuation Summary", bold: true })],
              spacing: { after: 120 }
            }),
            ...linesToParagraphs(valuationSummary, 120),
            new docx.Paragraph({ spacing: { after: 150 } })
          ]
        : []),

      ...(keyAssumptions && keyAssumptions.trim()
        ? [
            new docx.Paragraph({
              children: [new docx.TextRun({ text: "Key Assumptions", bold: true })],
              spacing: { after: 120 }
            }),
            ...keyAssumptions.split('\n').map(line => {
              if (line.trim() === '') return new docx.Paragraph({ text: "", spacing: { after: 80 } });
              return new docx.Paragraph({
                text: line.replace(/^[-*•]\s*/, '').trim(),
                bullet: { level: 0 },
                spacing: { after: 80 }
              });
            }),
            new docx.Paragraph({ spacing: { after: 150 } })
          ]
        : []),

      ...(scenarioNotes && scenarioNotes.trim()
        ? [
            new docx.Paragraph({
              children: [new docx.TextRun({ text: "Scenario / Sensitivity Notes", bold: true })],
              spacing: { after: 120 }
            }),
            ...linesToParagraphs(scenarioNotes, 120),
            new docx.Paragraph({ spacing: { after: 250 } })
          ]
        : [new docx.Paragraph({ spacing: { after: 200 } })])
    );
  }

  // KEY TAKEAWAYS
  documentChildren.push(
    new docx.Paragraph({
      children: [new docx.TextRun({ text: "Key Takeaways", bold: true, size: 24, font: "Book Antiqua" })],
      spacing: { after: 200 }
    }),
    ...takeawayBullets,
    new docx.Paragraph({ spacing: { after: 300 } }),

    // ANALYSIS AND COMMENTARY
    new docx.Paragraph({
      children: [new docx.TextRun({ text: "Analysis and Commentary", bold: true, size: 24, font: "Book Antiqua" })],
      spacing: { after: 200 }
    }),
    ...analysisParagraphs,
    ...contentParagraphs
  );

  // Cordoba View section
  if ((cordobaView || "").trim()) {
    documentChildren.push(
      new docx.Paragraph({ spacing: { after: 300 } }),
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "The Cordoba View", bold: true, size: 24, font: "Book Antiqua" })],
        spacing: { after: 200 }
      }),
      ...cordobaViewParagraphs
    );
  }

  // Figures section
  if (imageParagraphs.length > 0) {
    documentChildren.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "Figures and Charts", bold: true, size: 24, font: "Book Antiqua" })],
        spacing: { before: 400, after: 200 }
      }),
      ...imageParagraphs
    );
  }

  // Document
  const doc = new docx.Document({
    styles: {
      default: {
        document: {
          run: { font: "Book Antiqua", size: 20, color: "000000" },
          paragraph: { spacing: { after: 150 } }
        }
      }
    },
    sections: [{
      properties: {
        page: {
          margin: { top: 720, right: 720, bottom: 720, left: 720 },
          pageSize: {
            orientation: docx.PageOrientation.LANDSCAPE,
            width: 15840,
            height: 12240
          }
        }
      },
      headers: {
        default: new docx.Header({
          children: [
            new docx.Paragraph({
              children: [
                new docx.TextRun({
                  text: `Cordoba Research Group | ${noteType} | ${dateTimeString}`,
                  size: 16,
                  font: "Book Antiqua"
                })
              ],
              alignment: docx.AlignmentType.RIGHT,
              spacing: { after: 100 },
              border: {
                bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 }
              }
            })
          ]
        })
      },
      footers: {
        default: new docx.Footer({
          children: [
            new docx.Paragraph({
              border: {
                top: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 }
              },
              spacing: { after: 0 }
            }),
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: "\t" }),
                new docx.TextRun({
                  text: "Cordoba Research Group Internal Information",
                  size: 16,
                  font: "Book Antiqua",
                  italics: true
                }),
                new docx.TextRun({ text: "\t" }),
                new docx.TextRun({
                  children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
                  size: 16,
                  font: "Book Antiqua",
                  italics: true
                })
              ],
              spacing: { before: 0, after: 0 },
              tabStops: [
                { type: docx.TabStopType.CENTER, position: 5000 },
                { type: docx.TabStopType.RIGHT, position: 10000 }
              ]
            })
          ]
        })
      },
      children: documentChildren
    }]
  });

  return doc;
}

// Main Form Submission
document.getElementById('researchForm').addEventListener('submit', async function(e) {
  e.preventDefault();

  const button = e.target.querySelector('button[type="submit"]');
  const messageDiv = document.getElementById('message');

  // Loading state
  button.disabled = true;
  button.classList.add('loading');
  button.textContent = 'Generating Document...';
  messageDiv.className = 'message';
  messageDiv.textContent = '';

  try {
    if (typeof docx === 'undefined') throw new Error('docx library not loaded. Please refresh the page.');
    if (typeof saveAs === 'undefined') throw new Error('FileSaver library not loaded. Please refresh the page.');

    // Base fields
    const noteType = document.getElementById('noteType').value;
    const title = document.getElementById('title').value;
    const topic = document.getElementById('topic').value;
    const authorLastName = document.getElementById('authorLastName').value;
    const authorFirstName = document.getElementById('authorFirstName').value;
    const authorPhone = document.getElementById('authorPhone').value;
    const analysis = document.getElementById('analysis').value;
    const keyTakeaways = document.getElementById('keyTakeaways').value;
    const content = document.getElementById('content').value;
    const cordobaView = document.getElementById('cordobaView').value;
    const imageFiles = document.getElementById('imageUpload').files;

    // Equity fields (NEW)
    const ticker = document.getElementById('ticker') ? document.getElementById('ticker').value : "";
    const modelType = document.getElementById('modelType') ? document.getElementById('modelType').value : "";
    const valuationSummary = document.getElementById('valuationSummary') ? document.getElementById('valuationSummary').value : "";
    const keyAssumptions = document.getElementById('keyAssumptions') ? document.getElementById('keyAssumptions').value : "";
    const scenarioNotes = document.getElementById('scenarioNotes') ? document.getElementById('scenarioNotes').value : "";
    const modelFiles = document.getElementById('modelFiles') ? document.getElementById('modelFiles').files : null;

    // Date/time
    const now = new Date();
    const dateTimeString = formatDateTime(now);

    // Co-authors
    const coAuthors = [];
    const coAuthorEntries = document.querySelectorAll('.coauthor-entry');
    coAuthorEntries.forEach(entry => {
      const lastName = entry.querySelector('.coauthor-lastname').value;
      const firstName = entry.querySelector('.coauthor-firstname').value;
      const phone = entry.querySelector('.coauthor-phone').value;
      if (lastName && firstName && phone) coAuthors.push({ lastName, firstName, phone });
    });

    console.log('Generating document with:', { noteType, title, topic, coAuthors: coAuthors.length });

    const doc = await createDocument({
      noteType,
      title,
      topic,
      authorLastName,
      authorFirstName,
      authorPhone,
      coAuthors,
      analysis,
      keyTakeaways,
      content,
      cordobaView,
      imageFiles,
      dateTimeString,
      // equity payload
      ticker,
      modelType,
      valuationSummary,
      keyAssumptions,
      scenarioNotes,
      modelFiles
    });

    console.log('Creating blob...');
    const blob = await docx.Packer.toBlob(doc);
    console.log('Blob created, size:', blob.size);

    const fileName = `${title.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_${noteType.replace(/\s+/g, '_').toLowerCase()}.docx`;
    saveAs(blob, fileName);

    messageDiv.className = 'message success';
    messageDiv.textContent = `✓ Document "${fileName}" generated successfully!`;

    setTimeout(() => {
      if (confirm('Document generated! Would you like to create another document?')) {
        document.getElementById('researchForm').reset();
        document.getElementById('coAuthorsList').innerHTML = '';
        coAuthorCount = 0;
        toggleEquitySection();
        messageDiv.className = 'message';
        messageDiv.textContent = '';
      }
    }, 1500);

  } catch (error) {
    console.error('Error generating document:', error);
    messageDiv.className = 'message error';
    messageDiv.textContent = `✗ Error: ${error.message}`;
  } finally {
    button.disabled = false;
    button.classList.remove('loading');
    button.textContent = 'Generate Word Document';
  }
});

console.log('app.js loaded successfully');
