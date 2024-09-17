
document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('auditForm');
    const language = "{{ language }}";

    form.addEventListener('submit', function(event) {
        event.preventDefault(); // Prevents the default form submission behavior

        const formData = new FormData(form);
        const formObject = {};

        formData.forEach((value, key) => {
            formObject[key] = value;
        });
        const overlay = document.getElementById('overlay');
        const status = document.getElementById('status');
        
        overlay.style.display = 'block';

        console.log(formObject);

        fetch('/generate_report', {method: 'POST',
        body: formData
    })
    .then(response => response.blob())

            .then(blob => {
               
                // Create a link element
                const link = document.createElement('a');
                link.href = window.URL.createObjectURL(blob);
                if (language == 'arabic') {
                    const reportNumber = document.querySelector('input[name="رقم_التقرير"]').value;
                    if (reportNumber == '') {
                        const filename = `Manzili_Energy_Audit_Report_${reportNumber}.docx`;
                    link.download = filename;
                    }  
                }
                else{
                    const reportNumber = document.querySelector('input[name="report_number"]').value;
                    if(reportNumber == ''){
                        const filename = `Manzili_Energy_Audit_Report_${reportNumber}.docx`;
                        link.download = filename;
                    }
              }

                
             overlay.style.display = 'none';

                // Append link to the body and trigger a click
                document.body.appendChild(link);
                link.click();
                
                // Clean up
                document.body.removeChild(link);
            })
            .catch(error => {
                overlay.style.display = 'none';
                status.textContent = 'Error generating report: ' + error;
});

    });

    
});