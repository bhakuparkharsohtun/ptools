// Ensure Office.js is ready before running any code
Office.onReady(() => {
    console.log("PowerPoint Add-in is ready!");
});

async function comparePPT() {
    console.log("Compare button clicked!"); // Debugging log

    let originalPPTFile = document.getElementById("originalFile").files[0];
    let workingPPTFile = document.getElementById("workingFile").files[0];

    if (!originalPPTFile || !workingPPTFile) {
        showDialog("Please select both files.");
        console.log("File selection missing!"); // Debugging log
        return;
    }

    console.log("Files selected:", originalPPTFile.name, workingPPTFile.name); // Debugging log

    try {
        const originalText = await extractText(originalPPTFile);
        const workingText = await extractText(workingPPTFile);

        console.log("Extracted text from original file:", originalText);
        console.log("Extracted text from working file:", workingText);

        let differences = [];
        const maxSlides = Math.max(originalText.length, workingText.length);

        for (let i = 0; i < maxSlides; i++) {
            if (originalText[i] !== workingText[i]) {
                differences.push(`Difference in Slide ${i + 1}`);
            }
        }

        if (differences.length > 0) {
            showDialog(differences.join("\n"));
        } else {
            showDialog("No differences found.");
        }
    } catch (error) {
        console.error("Error comparing presentations:", error);
        showDialog("An error occurred while processing the files.");
    }
}



async function extractText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsArrayBuffer(file);

        reader.onload = async function () {
            console.log("File loaded successfully:", file.name);
            const arrayBuffer = reader.result;

            try {
                let extractedText = await parsePPTText(arrayBuffer);
                console.log("Extracted text:", extractedText);
                resolve(extractedText);
            } catch (error) {
                console.error("Error extracting text:", error);
                reject(error);
            }
        };

        reader.onerror = function (error) {
            console.error("File read error:", error);
            reject(error);
        };
    });
}


async function parsePPTText(arrayBuffer) {
    let textArray = [];

    try {
        const zip = new JSZip();
        const pptx = await zip.loadAsync(arrayBuffer);

        for (let i = 1; i <= 50; i++) { // Assuming max 50 slides
            let slidePath = `ppt/slides/slide${i}.xml`;
            if (pptx.files[slidePath]) {
                let slideXml = await pptx.files[slidePath].async("string");
                let parser = new DOMParser();
                let xmlDoc = parser.parseFromString(slideXml, "text/xml");

                let texts = xmlDoc.getElementsByTagName("a:t");
                let slideText = [];

                for (let textNode of texts) {
                    slideText.push(textNode.textContent);
                }

                textArray.push(slideText.join(" "));
            }
        }
    } catch (error) {
        console.error("Error parsing PowerPoint:", error);
    }

    return textArray;
}

function showDialog(message) {
    let url = "https://bhakuparkharsohtun.github.io/ptools/dialog.html"; // Use GitHub Pages URL

    Office.context.ui.displayDialogAsync(url, { height: 30, width: 40 }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            let dialog = asyncResult.value;
            
            // Listen for messages from dialog
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                console.log("Dialog message:", arg.message);
            });

            // Allow some time before sending message
            setTimeout(() => {
                dialog.messageChild(message);
            }, 1000);
        } else {
            console.error("Failed to open dialog:", asyncResult.error.message);
        }
    });
}
