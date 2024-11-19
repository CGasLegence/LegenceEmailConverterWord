async function saveDocument() {
    await Word.run(async (context) => {
        const body = context.document.body;
        body.load("paragraphs, tables, inlinePictures");
        await context.sync();

        // Start building the HTML content
        let htmlContent = `<html><head><style>${getDefaultStyles()}</style></head><body>`;

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
            for (let img of images) {
                img.load("base64"); // Explicitly load the Base64 content
            }
            await context.sync();

            // Process images after loading
            images.forEach((img, index) => {
                if (img.base64) {
                    htmlContent += `<img src="data:image/png;base64,${img.base64}" alt="Image ${index}" style="max-width: 100%;"/><br>`;
                } else {
                    console.warn(`Image ${index} has no Base64 data.`);
                }
            });
        }

        // Convert tables with borders, shading, and styles
        const tables = body.tables.items;
        for (let table of tables) {
            table.load("values");
            await context.sync();
            htmlContent += convertTableToHTML(table);
        }

        htmlContent += "</body></html>";

        // Save the generated HTML content to a local file
        saveAsFile(htmlContent);
    });
}

// Default styles for the HTML document
function getDefaultStyles() {
    return `
        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 20px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        td, th { border: 1px solid black; padding: 8px; text-align: left; }
    `;
}

// Extract paragraph styles
function getParagraphStyles(font) {
    let styles = "";
    if (font.color) styles += `color: ${font.color};`;
    if (font.size) styles += `font-size: ${font.size}px;`;
    if (font.bold) styles += "font-weight: bold;";
    if (font.italic) styles += "font-style: italic;";
    if (font.underline !== "None") styles += "text-decoration: underline;";
    if (font.highlightColor) styles += `background-color: ${font.highlightColor};`;
    if (font.name) styles += `font-family: ${font.name};`;
    return styles;
}

// Convert a Word table to HTML with borders, shading, and styles
function convertTableToHTML(table) {
    let html = "<table style='border-collapse: collapse; width: 100%;'>";

    for (const row of table.values) {
        html += "<tr>";
        for (const cell of row) {
            const styles = getCellStyles(cell);
            html += `<td style="${styles}">${cell}</td>`;
        }
        html += "</tr>";
    }

    html += "</table><br>";
    return html;
}

// Extract table cell styles, including shading and border color
function getCellStyles(cell) {
    let styles = "border: 1px solid black;"; // Default border style

    // Extract specific border and shading styles
    if (cell.shading) styles += `background-color: ${cell.shading};`;
    if (cell.borderTop) styles += `border-top: ${cell.borderTop};`;
    if (cell.borderBottom) styles += `border-bottom: ${cell.borderBottom};`;
    if (cell.borderLeft) styles += `border-left: ${cell.borderLeft};`;
    if (cell.borderRight) styles += `border-right: ${cell.borderRight};`;

    return styles + " padding: 5px;";
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
