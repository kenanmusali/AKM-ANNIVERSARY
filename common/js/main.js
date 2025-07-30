flatpickr("#datePicker", {
    dateFormat: "d/m",
    enableTime: false,
    noCalendar: false
});

document.getElementById('processBtn').addEventListener('click', processExcel);

function processExcel() {
    const file = document.getElementById('excelFile').files[0];
    const selectedDate = document.getElementById('datePicker').value;

    if (!file || !selectedDate) {
        alert("Please upload file and select date");
        return;
    }

    const [day, month] = selectedDate.split('/');
    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);

        const currentYear = new Date().getFullYear();
        const resultContainer = document.getElementById('resultContainer');
        resultContainer.innerHTML = '';

        jsonData.forEach(row => {
            // Check both possible field names (StartDate and StartDate2)
            const dateField = row.StartDate || row.StartDate2;
            if (!dateField) return;

            // Handle both formats (with time and without time)
            let excelDay, excelMonth, excelYear;
            
            if (dateField.includes(' ')) {
                // Old format with time (20.04.2018 0:00:00)
                [excelDay, excelMonth, excelYear] = dateField.split(' ')[0].split('.');
            } else {
                // New format without time (20.04.2018)
                [excelDay, excelMonth, excelYear] = dateField.split('.');
            }

            if (excelDay === day && excelMonth === month) {
                const yearsOfService = currentYear - parseInt(excelYear);
                const imageNumber = Math.min(Math.max(yearsOfService, 1), 8);
                const imgSrc = `./assets/img/${imageNumber}.png`;
                const employeeName = row.ссылка || row['ПИН'];
                let job = row.Department;
                const started = dateField.split(' ')[0]; // This will work for both formats

                const card = document.createElement('div');
                card.className = 'image-card';
                card.innerHTML = `<div class="employee-info">
                    <p><strong>Employee:</strong> ${employeeName}</p>
                    <p><strong>Job:</strong> ${job}</p>
                    <p><strong>Years:</strong> ${yearsOfService}</p>
                    <p><strong>Started:</strong> ${started}</p>
                    <button onclick="downloadImage(this)">Download</button>
                </div>`;

                resultContainer.appendChild(card);

                fetch(imgSrc)
                    .then(res => res.blob())
                    .then(blob => {
                        const reader = new FileReader();
                        reader.onloadend = () => {
                            const img = new Image();
                            img.onload = function () {
                                const canvas = document.createElement('canvas');
                                const ctx = canvas.getContext('2d');

                                canvas.width = img.width;
                                canvas.height = img.height;

                                ctx.drawImage(img, 0, 0);

                                const rootStyles = getComputedStyle(document.documentElement);

                                const marginLeft = parseInt(rootStyles.getPropertyValue('--margin-left')) || 100;
                                const marginTop = parseInt(rootStyles.getPropertyValue('--margin-top')) || 570;
                                const marginTop1 = parseInt(rootStyles.getPropertyValue('--margin-top1')) || 630;

                                const fontColor = rootStyles.getPropertyValue('--main-text-color') || '#0346B8';
                                const fontColor1 = rootStyles.getPropertyValue('--main1-text-color') || '#156fd8';
                                const fontFamily = rootStyles.getPropertyValue('--font-family') || 'Inter, sans-serif';

                                const jobOriginal = job.trim();
                                const jobReplaceList = [
                                    "BBGİ ofisi",
                                    "Bəyannamə bölməsi",
                                    "Gömrük təmsilçiliyi bölməsi",
                                    "HNBGİ ofisi",
                                    "Koordinasiya şöbəsi",
                                    "Koordinasiya şöbəsi / Sertifikatlaşdırma bölməsi",
                                    "Mühasibatlıq şöbəsi",
                                    "MB Broker",
                                    "Mühasibatlıq şöbəsi / Kassa - hesablaşmalar bölməsi",
                                    "Satış şöbəsi",
                                ];

                                let jobTransformed = jobOriginal;
                                let jobLines;

                                if (jobOriginal === "TIR Park və minik avtomobillərinin dayanacağı şöbəsi") {
                                    jobLines = [
                                        "TIR Park və minik avtomobillərinin",
                                        "dayanacağı şöbəsi"
                                    ];
                                } else if (jobOriginal.startsWith("DDD / Daxili daşımalar departamenti")) {
                                    jobTransformed = "Daxili daşımalar departamenti";
                                    jobLines = [jobTransformed];
                                } else if (jobReplaceList.some(prefix => jobOriginal.startsWith(prefix))) {
                                    jobLines = ["MB BROKER"];
                                } else {
                                    jobLines = [jobTransformed];
                                }

                                document.fonts.ready.then(() => {
                                    ctx.textAlign = 'left';

                                    ctx.fillStyle = fontColor.trim();
                                    ctx.font = `bold 51px ${fontFamily}`;
                                    ctx.fillText(`${employeeName}`, marginLeft, marginTop);

                                    ctx.fillStyle = fontColor1.trim();
                                    ctx.font = `normal 51px ${fontFamily}`;
                                    const lineHeight = 60;
                                    jobLines.forEach((line, i) => {
                                        ctx.fillText(line, marginLeft, marginTop1 + i * lineHeight);
                                    });

                                    const preview = document.createElement('img');
                                    preview.className = 'preview-image';
                                    preview.src = canvas.toDataURL('image/png');
                                    preview.style = 'display:block; margin-top:10px; max-width:300px; border:1px solid #ccc;';
                                    card.insertBefore(preview, card.querySelector('.employee-info'));
                                });
                            };
                            img.src = reader.result;
                        };
                        reader.readAsDataURL(blob);
                    });
            }
        });

        if (resultContainer.innerHTML === '') {
            resultContainer.innerHTML = '<p>No employees found for this anniversary date</p>';
        }
    };

    reader.readAsArrayBuffer(file);
}

function downloadImage(button) {
    const card = button.closest('.image-card');
    const preview = card.querySelector('.preview-image');
    const employeeName = card.querySelector('.employee-info p:nth-child(1)').innerText.replace('Employee: ', '');
    const years = card.querySelector('.employee-info p:nth-child(3)').innerText.replace('Years: ', '').trim();

    const link = document.createElement('a');
    link.download = `${employeeName} - ${years} il.png`;
    link.href = preview.src;
    link.click();
}