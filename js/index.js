function convertWordToExcel() {
            const fileInput = document.getElementById("wordFile");
            if (fileInput.files.length === 0) {
                alert("Please upload a Word file.");
                return;
            }
            const reader = new FileReader();
            reader.onload = function(event) {
                const arrayBuffer = reader.result;
                mammoth.extractRawText({ arrayBuffer: arrayBuffer })
                    .then(result => createExcel(result.value))
                    .catch(err => console.error("Error processing file", err));
            };
            reader.readAsArrayBuffer(fileInput.files[0]);
        }
 
        function convertTextToExcel() {
            const text = document.getElementById("textInput").value;
            if (!text.trim()) {
                alert("Please enter some text.");
                return;
            }
            createExcel(text);
        }
 
        function createExcel(content) {
            const ws = XLSX.utils.aoa_to_sheet([["Content"], [content]]);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
            XLSX.writeFile(wb, "converted.xlsx");
        }
