async function convertLocalDocxToHtml() {
    try {
        // Path to the existing DOCX file
        const filePath = "Templates/CMTA.docx";

        // Fetch the file as a Blob
        const response = await fetch(filePath);
        if (!response.ok) throw new Error(`Failed to fetch ${filePath}`);
        const docxBlob = await response.blob();

        // Read the Blob as ArrayBuffer (required by OfficeToHtml.js)
        const arrayBuffer = await docxBlob.arrayBuffer();

        // Convert the DOCX file to HTML using OfficeToHtml.js
        const htmlContent = OfficeToHtml.convert(arrayBuffer);

        // Save the generated HTML locally
        saveAsHtml(htmlContent);
    } catch (error) {
        console.error("Error converting DOCX to HTML:", error);
        alert("Failed to convert the document. Check console for details.");
    }
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
    document.getElementById("convertButton").onclick = convertLocalDocxToHtml;
});
