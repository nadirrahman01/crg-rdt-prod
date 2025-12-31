document.getElementById('researchForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    const title = document.getElementById('title').value;
    const content = document.getElementById('content').value;
    const imageFiles = document.getElementById('imageUpload').files;
    
    // Create Word document
    const doc = new docx.Document({
        sections: [{
            properties: {},
            children: [
                new docx.Paragraph({
                    text: title,
                    heading: docx.HeadingLevel.HEADING_1,
                    spacing: { after: 200 }
                }),
                new docx.Paragraph({
                    text: content,
                    spacing: { after: 200 }
                }),
                // Add images
                ...await addImages(imageFiles)
            ]
        }]
    });
    
    // Generate and download
    docx.Packer.toBlob(doc).then(blob => {
        saveAs(blob, `${title.replace(/\s+/g, '_')}_research.docx`);
    });
});

async function addImages(files) {
    const imageParagraphs = [];
    
    for (let file of files) {
        const arrayBuffer = await file.arrayBuffer();
        imageParagraphs.push(
            new docx.Paragraph({
                children: [
                    new docx.ImageRun({
                        data: arrayBuffer,
                        transformation: {
                            width: 400,
                            height: 300
                        }
                    })
                ]
            })
        );
    }
    
    return imageParagraphs;
}
