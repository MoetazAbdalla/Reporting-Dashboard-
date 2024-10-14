window.dash_clientside = Object.assign({}, window.dash_clientside, {
    clientside: {
        generatePDF: function(n_clicks) {
            if (n_clicks > 0) {
                const element = document.body;  // Capture the entire page
                html2canvas(element, {
                    useCORS: true,  // Allow cross-origin content
                    scrollX: 0,     // Ensures capturing full width
                    scrollY: 0,     // Ensures capturing full height
                    windowWidth: document.documentElement.scrollWidth,
                    windowHeight: document.documentElement.scrollHeight
                }).then(canvas => {
                    const imgData = canvas.toDataURL('image/png');
                    const pdf = new jsPDF('p', 'pt', 'a4');
                    const imgWidth = 595.28; // A4 width in points
                    const pageHeight = 841.89; // A4 height in points
                    const imgHeight = canvas.height * imgWidth / canvas.width;
                    let heightLeft = imgHeight;
                    let position = 0;

                    pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
                    heightLeft -= pageHeight;

                    while (heightLeft >= 0) {
                        position = heightLeft - imgHeight;
                        pdf.addPage();
                        pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
                        heightLeft -= pageHeight;
                    }
                    pdf.save('dashboard.pdf');  // Save the PDF
                });
            }
            return window.dash_clientside.no_update;
        }
    }
});

// Listen for the click event from the Dash button
document.addEventListener('DOMContentLoaded', function() {
    const downloadBtn = document.getElementById('download-pdf');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', function() {
            window.dash_clientside.clientside.generatePDF(downloadBtn.n_clicks);
        });
    }
});
