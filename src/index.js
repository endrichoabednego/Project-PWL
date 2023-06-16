import "regenerator-runtime";
import "./style/style.css";
import main from "./scripts/main.js";

const fileInput = document.getElementById('fileInput'); // Assuming you have an input field with id 'fileInput'

fileInput.addEventListener('change', handleFileUpload);

function handleFileUpload(event) {
    const file = event.target.files[0];

    if (file) {
        uploadFile(file);
    }
}

async function uploadFile(file) {
    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        if (response.ok) {
            // File uploaded successfully, handle the response
        } else {
            // Handle error response
        }
    } catch (error) {
        // Handle network or other errors
    }
}

main();