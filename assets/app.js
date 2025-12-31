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
    if (coAuthorDiv) {
        coAuthorDiv.remove();
    }
}

// Main Form Submission
document.getElementById('researchForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    const button = e.target.querySelector('button[type="submit"]');
    const messageDiv = document.getElementById('message');
    
    // Show loading state
    button.disabled = true;
    button.classList.add('loading');
    button.textContent = 'Generating Document...';
    messageDiv.className = 'message';
    messageDiv.textContent = '';
    
    try {
        // Check if libraries are loaded
        if (typeof docx === 'undefined') {
            throw new Error('docx library not loaded. Please refresh the page.');
        }
        if (typeof saveAs === 'undefined') {
            throw new Error('FileSaver library not loaded. Please refresh the page.');
        }
        
        // Get form data
        const noteType = document.getElementById('noteType').value;
        const title = document.getElementById('title').value;
        const topic = document.getElementById('topic').value;
        const authorLastName = document.getElementById('authorLastName').value;
        const authorFirstName = document.getElementById('authorFirstName').value;
        const authorPhone = document.getElementById('authorPhone').value;
        const analysis = document.getElementById('analysis').value;
        const keyTakeaways = document.getElementById('keyTakeaways').value;
        const content = document.getElementById('content').value;
        const imageFiles = document.getElementById('imageUpload').files;
        
        // Get current date and time
        const now = new Date();
        const dateTimeString = formatDateTime(now);
        
        // Collect co-authors
        const coAuthors = [];
        const coAuthorEntries = document.querySelectorAll('.coauthor-entry');
        coAuthorEntries.forEach(entry => {
            const lastName = entry.querySelector('.coauthor-lastname').value;
            const firstName = entry.querySelector('.coauthor-firstname').value;
            const phone = entry.querySelector('.coauthor-phone').value;
            if (lastName && firstName && phone) {
                coAuthors.push({ lastName, firstName, phone });
            }
        });
        
        console.log('Generating document with:', { noteType, title, topic, coAuthors: coAuthors.length });
        
        // Create the document
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
            imageFiles,
            dateTimeString
        });
        
        // Generate and download
        console.log('Creating blob...');
        const blob = await docx.Packer.toBlob(doc);
        console.log('Blob created, size:', blob.size);
        
        const fileName = `${title.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_${noteType.replace(/\s+/g, '_').toLowerCase()}.docx`;
        saveAs(blob, fileName);
        
        // Show success message
        messageDiv.className = 'message success';
        messageDiv.textContent = `✓ Document "${fileName}" generated successfully!`;
        
        // Reset form after delay
        setTimeout(() => {
            if (confirm('Document generated! Would you like to create another document?')) {
                document.getElementById('researchForm').reset();
                document.getElementById('coAuthorsList').innerHTML = '';
                coAuthorCount = 0;
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

// Format date and time
function formatDateTime(date) {
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                   'July', 'August', 'September', 'October', 'November', 'December'];
    
    const month = months[date.getMonth()];
    const day = date.getDate();
    const year = date.getFullYear();
    
    let hours = date.getHours();
    const minutes = date.getMinutes().toString().padStart(2, '0');
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12 || 12;
    
    return `${month} ${day}, ${year} ${hours}:${minutes} ${ampm}`;
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
        imageFiles,
        dateTimeString
    } = data;
    
    // Process key takeaways into bullet points
    const takeawayLines = keyTakeaways.split('\n').filter(line => line.trim());
    const takeawayBullets = takeawayLines.map(line => {
        const cleanLine = line.replace(/^[-*•]\s*/, '').trim();
        return new docx.Paragraph({
            text: cleanLine,
            bullet: {
                level: 0
            },
            spacing: { after: 100 }
        });
    });
    
    // Process content into paragraphs
    const contentParagraphs = content.split('\n').filter(line => line.trim()).map(line => 
        new docx.Paragraph({
            text: line,
            spacing: { after: 150 }
        })
    );
    
    // Process images
    const imageParagraphs = await addImages(imageFiles);
    
    // Build author section (right-aligned)
    const authorSection = [
        new docx.Paragraph({
            text: `${authorLastName}, ${authorFirstName}`,
            alignment: docx.AlignmentType.RIGHT,
            spacing: { after: 50 }
        }),
        new docx.Paragraph({
            text: `(${authorPhone})`,
            alignment: docx.AlignmentType.RIGHT,
            spacing: { after: 100 }
        })
    ];
    
    // Add co-authors
    coAuthors.forEach(coAuthor => {
        authorSection.push(
            new docx.Paragraph({
                text: `${coAuthor.lastName}, ${coAuthor.firstName}`,
                alignment: docx.AlignmentType.RIGHT,
                spacing: { after: 50 }
            }),
            new docx.Paragraph({
                text: `(${coAuthor.phone})`,
                alignment: docx.AlignmentType.RIGHT,
                spacing: { after: 100 }
            })
        );
    });
    
    // Create document
    const doc = new docx.Document({
        styles: {
            default: {
                document: {
                    run: {
                        font: "Book Antiqua",
                        size: 20 // 10pt = 20 half-points
                    }
                }
            }
        },
        sections: [{
            properties: {
                page: {
                    margin: {
                        top: 1440,
                        right: 1440,
                        bottom: 1440,
                        left: 1440
                    }
                }
            },
            headers: {
                default: new docx.Header({
                    children: [
                        new docx.Paragraph({
                            text: `Cordoba Research Group | ${noteType} | ${dateTimeString}`,
                            alignment: docx.AlignmentType.LEFT,
                            spacing: { after: 200 },
                            border: {
                                bottom: {
                                    color: "000000",
                                    space: 1,
                                    style: docx.BorderStyle.SINGLE,
                                    size: 6
                                }
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
                                top: {
                                    color: "000000",
                                    space: 1,
                                    style: docx.BorderStyle.SINGLE,
                                    size: 6
                                }
                            },
                            spacing: { before: 100 }
                        }),
                        new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                                new docx.TextRun({
                                    text: "Cordoba Research Group Internal Information",
                                    size: 20
                                }),
                                new docx.TextRun({
                                    text: "\t\t",
                                }),
                                new docx.TextRun({
                                    children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
                                    size: 20
                                })
                            ]
                        })
                    ]
                })
            },
            children: [
                // Title (Size 14, Bold, Left-aligned)
                new docx.Paragraph({
                    text: title,
                    bold: true,
                    size: 28, // 14pt = 28 half-points
                    spacing: { after: 100 }
                }),
                
                // Horizontal line
                new docx.Paragraph({
                    border: {
                        bottom: {
                            color: "000000",
                            space: 1,
                            style: docx.BorderStyle.SINGLE,
                            size: 6
                        }
                    },
                    spacing: { after: 200 }
                }),
                
                // Topic label and value (Size 14)
                new docx.Paragraph({
                    text: "Topic:",
                    bold: true,
                    size: 28,
                    spacing: { after: 100 }
                }),
                new docx.Paragraph({
                    text: topic,
                    size: 20,
                    spacing: { after: 200 }
                }),
                
                // Author section (right side)
                ...authorSection,
                
                // Horizontal line
                new docx.Paragraph({
                    border: {
                        bottom: {
                            color: "000000",
                            space: 1,
                            style: docx.BorderStyle.SINGLE,
                            size: 6
                        }
                    },
                    spacing: { after: 300, before: 100 }
                }),
                
                // Analysis and Commentary
                new docx.Paragraph({
                    text: "Analysis and Commentary",
                    heading: docx.HeadingLevel.HEADING_2,
                    bold: true,
                    size: 24, // 12pt
                    spacing: { after: 200 }
                }),
                new docx.Paragraph({
                    text: analysis,
                    spacing: { after: 300 }
                }),
                
                // Key Takeaways
                new docx.Paragraph({
                    text: "Key Takeaways",
                    heading: docx.HeadingLevel.HEADING_2,
                    bold: true,
                    size: 24,
                    spacing: { after: 200 }
                }),
                ...takeawayBullets,
                
                new docx.Paragraph({
                    spacing: { after: 300 }
                }),
                
                // Content
                new docx.Paragraph({
                    text: "Content",
                    heading: docx.HeadingLevel.HEADING_2,
                    bold: true,
                    size: 24,
                    spacing: { after: 200 }
                }),
                ...contentParagraphs,
                
                // Images
                ...(imageParagraphs.length > 0 ? [
                    new docx.Paragraph({
                        text: "Figures and Charts",
                        heading: docx.HeadingLevel.HEADING_2,
                        bold: true,
                        size: 24,
                        spacing: { before: 400, after: 200 }
                    }),
                    ...imageParagraphs
                ] : [])
            ]
        }]
    });
    
    return doc;
}

// Process and add images
async function addImages(files) {
    const imageParagraphs = [];
    
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        
        try {
            console.log(`Processing image ${i + 1}: ${file.name}`);
            const arrayBuffer = await file.arrayBuffer();
            
            imageParagraphs.push(
                new docx.Paragraph({
                    children: [
                        new docx.ImageRun({
                            data: arrayBuffer,
                            transformation: {
                                width: 500,
                                height: 375
                            }
                        })
                    ],
                    spacing: { before: 200, after: 100 },
                    alignment: docx.AlignmentType.CENTER
                }),
                new docx.Paragraph({
                    text: `Figure ${i + 1}: ${file.name}`,
                    italics: true,
                    spacing: { after: 300 },
                    alignment: docx.AlignmentType.CENTER,
                    size: 18 // 9pt
                })
            );
        } catch (error) {
            console.error(`Error processing image ${file.name}:`, error);
        }
    }
    
    return imageParagraphs;
}
