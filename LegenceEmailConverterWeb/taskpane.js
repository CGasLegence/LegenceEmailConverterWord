async function extractOpenXmlAndConvert() {
    await Word.run(async (context) => {
        const body = context.document.body;
        const openXml = body.getOoxml(); // Get the document content as Open XML
        await context.sync();
        convertOpenXmlToHtml(openXml.value); // Convert Open XML to HTML
    });
}

// Function to parse Open XML and convert it to HTML
function convertOpenXmlToHtml(openXml) {
    const xmlDoc = parseOpenXml(openXml); // Parse Open XML using DOMParser
    const bodyNode = xmlDoc.getElementsByTagName("w:body")[0]; // Get the document body
    const relationships = extractRelationships(xmlDoc); // Extract image relationships

    // Initialize HTML content with styles
    let htmlContent = "<html><head><style>";
    htmlContent += "body { font-family: Arial, sans-serif; } table { border-collapse: collapse; }";
    htmlContent += "table, th, td { border: 1px solid black; padding: 5px; }";
    htmlContent += "</style></head><body>";

    // Process the body node (paragraphs, tables, etc.)
    htmlContent += processBody(bodyNode, relationships);

    htmlContent += "</body></html>";

    // Save the HTML file locally
    saveAsHtml(htmlContent);
}

// Function to parse Open XML into a DOM-like structure
function parseOpenXml(openXml) {
    const parser = new DOMParser();
    return parser.parseFromString(openXml, "application/xml");
}

// Extract relationships (e.g., images) from the Open XML
function extractRelationships(xmlDoc) {
    const relationshipsNode = xmlDoc.getElementsByTagName("Relationships")[0];
    const relationships = {};
    if (relationshipsNode) {
        relationshipsNode.childNodes.forEach((rel) => {
            if (rel.nodeName === "Relationship" && rel.getAttribute("Type").includes("image")) {
                const id = rel.getAttribute("Id");
                const target = rel.getAttribute("Target");
                relationships[id] = target; // Map relationship ID to image path
            }
        });
    }
    return relationships;
}

// Process the document body (<w:body>) and extract content
function processBody(bodyNode, relationships) {
    let html = "";

    bodyNode.childNodes.forEach((child) => {
        if (child.nodeName === "w:p") {
            html += processParagraph(child); // Handle paragraphs
        } else if (child.nodeName === "w:tbl") {
            html += processTable(child); // Handle tables
        } else if (child.nodeName === "w:drawing") {
            html += processImage(child, relationships); // Handle images
        }
    });

    return html;
}

// Process paragraphs (<w:p>) and text runs (<w:r>)
function processParagraph(paragraphNode) {
    let html = "<p>";

    paragraphNode.childNodes.forEach((runNode) => {
        if (runNode.nodeName === "w:r") {
            html += processRun(runNode); // Handle individual text runs
        }
    });

    html += "</p>";
    return html;
}

// Process individual text runs (<w:r>) and apply styles
function processRun(runNode) {
    let html = "";
    let styles = "";

    // Check for formatting (bold, italic, underline, etc.)
    if (runNode.getElementsByTagName("w:b").length > 0) styles += "font-weight: bold;";
    if (runNode.getElementsByTagName("w:i").length > 0) styles += "font-style: italic;";
    if (runNode.getElementsByTagName("w:u").length > 0) styles += "text-decoration: underline;";

    // Font size and color
    const sizeNode = runNode.getElementsByTagName("w:sz")[0];
    if (sizeNode) styles += `font-size: ${sizeNode.getAttribute("w:val") / 2}pt;`;

    const colorNode = runNode.getElementsByTagName("w:color")[0];
    if (colorNode) styles += `color: #${colorNode.getAttribute("w:val")};`;

    // Extract text content
    const textNode = runNode.getElementsByTagName("w:t")[0];
    if (textNode) html += `<span style="${styles}">${textNode.textContent}</span>`;

    return html;
}

// Process tables (<w:tbl>) and convert to HTML
function processTable(tableNode) {
    let html = "<table>";
    tableNode.childNodes.forEach((rowNode) => {
        if (rowNode.nodeName === "w:tr") {
            html += "<tr>";
            rowNode.childNodes.forEach((cellNode) => {
                if (cellNode.nodeName === "w:tc") {
                    html += "<td>";
                    html += processBody(cellNode); // Process cell content
                    html += "</td>";
                }
            });
            html += "</tr>";
        }
    });
    html += "</table><br>";
    return html;
}

// Process images (<w:drawing>) and embed them as Base64
function processImage(imageNode, relationships) {
    const blip = imageNode.getElementsByTagName("a:blip")[0];
    if (!blip) return "";

    const relId = blip.getAttribute("r:embed"); // Get the relationship ID
    const imagePath = relationships[relId]; // Get the image path from relationships

    if (!imagePath) return "";

    // Simulate fetching the image as Base64 (replace this with actual logic for fetching binary data)
    const base64Image = simulateFetchBase64(imagePath);

    return `<img src="data:image/png;base64,${base64Image}" alt="Embedded Image" />`;
}

// Simulate fetching Base64 image data (replace with actual file read logic)
function simulateFetchBase64(imagePath) {
    // Replace this with logic to read image data as Base64
    return "BASE64_ENCODED_IMAGE_DATA"; // Placeholder
}

// Save the generated HTML as a local file
function saveAsHtml(content) {
    const blob = new Blob([content], { type: "text/html" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "ConvertedDocument.html";
    link.click();
    URL.revokeObjectURL(link.href); // Clean up the URL object
}

// Attach the main function to a button or event
Office.onReady(() => {
    document.getElementById("convertButton").onclick = extractOpenXmlAndConvert;
});
