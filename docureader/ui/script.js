const fixedHeadings = [
    "Introduction",
    "Technical Implementation",
    "Architecture",
    "Results"
];

async function uploadPDF() {
    let result = await window.pywebview.api.upload_pdf();
    if (!result) return;

    // Show preview
    let preview = document.getElementById("preview");
    preview.innerHTML = "";

    result.images.forEach(img => {
        let image = document.createElement("img");
        image.src = img;
        preview.appendChild(image);
    });

    // Populate table
    populateTable(result.headings);
}

async function populateTable(autoHeadings) {
    let table = document.getElementById("statusTable");

    // Clear old rows except header
    table.innerHTML = `
        <tr>
            <th>Heading</th>
            <th>Status</th>
        </tr>
    `;

    // 1. FIXED HEADINGS FIRST
    for (let heading of fixedHeadings) {
        let row = table.insertRow();

        let cell1 = row.insertCell(0);
        let cell2 = row.insertCell(1);

        cell1.innerText = heading;

        let result = await window.pywebview.api.check_heading(heading);
        cell2.innerText = result;

        cell2.className = result === "Present"
            ? "status-present"
            : "status-missing";
    }

    // 2. AUTO-DETECTED HEADINGS
    autoHeadings.forEach(h => {
        // Avoid duplicates with fixed ones
        if (fixedHeadings.includes(h)) return;

        let row = table.insertRow();
        let cell1 = row.insertCell(0);
        let cell2 = row.insertCell(1);

        cell1.innerText = h;
        cell2.innerText = "Detected";
        cell2.className = "status-present";
    });
}

function addEntry() {
    let table = document.getElementById("statusTable");

    let row = table.insertRow();
    let cell1 = row.insertCell(0);
    let cell2 = row.insertCell(1);

    let input = document.createElement("input");
    input.placeholder = "Enter heading";

    let button = document.createElement("button");
    button.innerText = "Check";

    button.onclick = async () => {
    let heading = input.value;
    let result = await window.pywebview.api.check_heading(heading);

    cell2.innerText = result;

    if (result === "Present") {
        cell2.className = "status-present";
    } else {
        cell2.className = "status-missing";
    }
};

    cell1.appendChild(input);
    cell2.appendChild(button);
}

async function downloadDocx() {
    let res = await window.pywebview.api.export_docx();
    alert(res);
}

async function downloadPDF() {
    let res = await window.pywebview.api.export_pdf();
    alert(res);
}

function previewTemplate() {
    alert("Template preview coming soon");
}
