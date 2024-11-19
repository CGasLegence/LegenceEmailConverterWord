async function convertDocumentToHtml() {
    await Word.run(async (context) => {
        // Extract the document as OOXML
        const body = context.document.body;
        const openXml = body.getOoxml();
        await context.sync();

        // Convert Open XML to HTML using OfficeToHtml.js
        const htmlContent = OfficeToHtml.convert(openXml.value);

        // Save the generated HTML locally
        saveAsHtml(htmlContent);
    });
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
    document.getElementById("convertButton").onclick = convertDocumentToHtml;
});
