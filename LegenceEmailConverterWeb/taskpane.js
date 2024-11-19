async function convertLocalDocxToHtml() {
    try {
        // Path to the DOCX file relative to your web server
        const filePath = "Templates/CMTA.docx";

        // Fetch the file as a Blob
        const response = await fetch(filePath);
        if (!response.ok) {
            throw new Error(`Failed to fetch the file from ${filePath}. Status: ${response.status}`);
        }
        const docxBlob = await response.blob();

        // Read the Blob as an ArrayBuffer (required by OfficeToHtml.js)
        const arrayBuffer = await docxBlob.arrayBuffer();

        // Convert the DOCX content to HTML using OfficeToHtml.js
        const htmlContent = await OfficeToHtml.convert(arrayBuffer);

        // Log and save the HTML
        console.log("Generated HTML:", htmlContent);
        saveAsHtml(htmlContent);
    } catch (error) {
        console.error("Error converting DOCX to HTML:", error);
        alert(`Error: ${error.message}`);
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
