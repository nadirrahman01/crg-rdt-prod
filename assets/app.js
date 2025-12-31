// assets/app.js

document.getElementById('researchForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    const button = e.target.querySelector('button[type="submit"]');
    const messageDiv = document.getElementById('message');
    
    // Show loading state
    button.disabled = true;
    button.classList.add('loading');
    button.textContent = 'Generating...';
    messageDiv.className = 'message';
    messageDiv.textContent = '';
    
    try {
        // Check if libraries are loaded
        if (typeof docx === 'undefined') {
            throw new Error('docx library not loaded');
        }
        if (typeof saveAs === 'undefined') {
            throw new Error('FileSaver library not loaded');
        }
        
        const title = document.getElementById('title').value;
        const content = document.getElementById('content').value;
        const imageFiles = document.getElementById('imageUpload').files;
        
        console.log('Title:', title);
        console.log('Content length:', content.length);
        console.log('Number of images:', imageFiles.length);
        
        // Create paragraphs from content (split by line breaks)
        const contentParagraphs = content.split('\n').filter(line => line.trim()).map(line => 
            new docx.Paragraph({
                text: line,
                spacing: { after: 200 }
            })
        );
        
        // Process images
        const imageParagraphs = await addImages(imageFiles);
        
        console.log('Creating document...');
        
        // Create Word document with simpler structure
        const doc = new docx.Document({
            sections: [{
                children: [
                    // Title
                    new docx.Paragraph({
                        text: title,
                        heading: docx.HeadingLevel.HEADING_1,
                        spacing: { after: 400 }
                    }),
                    
                    // Content
                    ...contentParagraphs,
                    
                    // Images
                    ...imageParagraphs
                ]
            }]
        });
        
        console.log('Generating blob...');
        
        // Generate and download
        const blob = await docx.Packer.toBlob(doc);
        
        console.log('Blob created, size:', blob.size);
        
        const fileName = `${title.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_research.docx`;
        saveAs(blob, fileName);
        
        console.log('Download initiated');
        
        // Show success message
        messageDiv.className = 'message success';
        messageDiv.textContent = `✓ Document "${fileName}" generated successfully!`;
        
        // Reset form
        setTimeout(() => {
            document.getElementById('researchForm').reset();
        }, 1000);
        
    } catch (error) {
        console.error('Detailed error:', error);
        console.error('Error stack:', error.stack);
        messageDiv.className = 'message error';
        messageDiv.textContent = `✗ Error: ${error.message}. Check console for details.`;
    } finally {
        // Reset button
        button.disabled = false;
        button.classList.remove('loading');
        button.textContent = 'Generate Word Document';
    }
});

// Function to process and add images to document
async function addImages(files) {
    const imageParagraphs = [];
    
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        
        try {
            console.log(`Processing image ${i + 1}: ${file.name}`);
            
            const arrayBuffer = await file.arrayBuffer();
            
            console.log(`Image ${i + 1} loaded, size:`, arrayBuffer.byteLength);
            
            // Add image with caption
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
                    spacing: { before: 200, after: 100 }
                }),
                new docx.Paragraph({
                    text: `Figure ${i + 1}: ${file.name}`,
                    italics: true,
                    spacing: { after: 300 }
                })
            );
        } catch (error) {
            console.error(`Error processing image ${file.name}:`, error);
            // Continue with other images even if one fails
        }
    }
    
    return imageParagraphs;
}
