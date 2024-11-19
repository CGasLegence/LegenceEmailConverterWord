async function saveDocument() {
    const format = document.getElementById("format").value; // Get the selected format from the dropdown

    await Word.run(async (context) => {
        const body = context.document.body;
        body.load("paragraphs, tables, inlinePictures");
        await context.sync();

        // Start building the HTML content
        let htmlContent = `<html><head><style>body { font-family: Arial, sans-serif; }</style></head><body>`;

        if (format === "full") {
            // Process paragraphs with formatting
            const paragraphs = body.paragraphs.items;
            for (let para of paragraphs) {
                para.load("font, text");
                await context.sync();

                const styles = getParagraphStyles(para.font);
                htmlContent += `<p style="${styles}">${para.text.replace(/\n/g, "<br>")}</p>`;
            }

            // Convert inline images to Base64
            const images = body.inlinePictures.items;
            if (images.length > 0) {
                for (let i = 0; i < images.length; i++) {
                    images[i].load("base64");
                }
                await context.sync();

                images.forEach((img, index) => {
                    htmlContent += `<img src="data:image/png;base64,${img.base64}" alt="Image ${index}" /><br>`;
                });
            }

            // Convert tables to HTML
            const tables = body.tables.items;
            for (let table of tables) {
                table.load("values");
                await context.sync();
                htmlContent += convertTableToHTML(table);
            }
        } else if (format === "text") {
            // Convert only the plain text to HTML
            const paragraphs = body.paragraphs.items;
            for (let para of paragraphs) {
                htmlContent += `<p>${para.text.replace(/\n/g, "<br>")}</p>`;
            }
        }

        htmlContent += "</body></html>";

        // Save the generated HTML content to a local file
        saveAsFile(htmlContent);
    });
}

// Helper function to extract paragraph styles
function getParagraphStyles(font) {
    let styles = "";
    if (font.color) styles += `color: ${font.color};`;
    if (font.size) styles += `font-size: ${font.size}px;`;
    if (font.bold) styles += "font-weight: bold;";
    if (font.italic) styles += "font-style: italic;";
    if (font.underline !== "None") styles += "text-decoration: underline;";
    if (font.name) styles += `font-family: ${font.name};`;
    return styles;
}

// Helper function to convert a Word table to HTML
function convertTableToHTML(table) {
    let html = "<table border='1' cellspacing='0' cellpadding='5'>";
    for (const row of table.values) {
        html += "<tr>";
        for (const cell of row) {
            html += `<td>${cell}</td>`;
        }
        html += "</tr>";
    }
    html += "</table><br>";
    return html;
}

// Function to save the HTML content as a local file
function saveAsFile(content) {
    const blob = new Blob([content], { type: "text/html" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "ConvertedDocument.html";
    link.click();
    URL.revokeObjectURL(link.href); // Clean up the URL object
}
