document.getElementById('startBtn').addEventListener('click', function () {
    let resultDisplay = document.getElementById('result');
    let spinner = document.querySelector('.spinner');

    spinner.style.display = 'block';
    resultDisplay.textContent = "";

    fetch("AIT_details.xlsx")
        .then(response => response.arrayBuffer())
        .then(data => {
            let workbook = XLSX.read(data, { type: "array" });
            let sheets = workbook.SheetNames;

            let facultyList = [];

            sheets.forEach(sheetName => {
                let sheet = workbook.Sheets[sheetName];
                let sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                sheetData.forEach(row => {
                    if (row.length > 0) facultyList.push(row[0]); 
                });
            });

            if (facultyList.length === 0) {
                alert("No faculty data found!");
                spinner.style.display = "none";
                return;
            }

            setTimeout(() => {
                let randomFaculty = facultyList[Math.floor(Math.random() * facultyList.length)];
                resultDisplay.textContent = `Selected: ${randomFaculty}`;
                spinner.style.display = "none";
            }, 3000);
        })
        .catch(error => {
            console.error("Error loading file:", error);
            alert("Failed to load AIT_details.xlsx");
            spinner.style.display = "none";
        });
});

