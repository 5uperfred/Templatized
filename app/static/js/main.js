document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('uploadForm');
    const result = document.getElementById('result');
    const downloads = document.getElementById('downloads');
    const progress = document.querySelector('.progress');

    form.addEventListener('submit', async function(e) {
        e.preventDefault();

        // Reset UI
        result.style.display = 'none';
        downloads.innerHTML = '';
        progress.style.width = '0%';

        const formData = new FormData(form);

        try {
            // Show progress
            progress.style.width = '50%';

            const response = await fetch('/process', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (response.ok) {
                progress.style.width = '100%';

                // Show download links
                downloads.innerHTML = `
                    <div class="success">Document processed successfully!</div>
                    <p><a href="/download/${data.template_path}" class="download-link">
                        Download Template Document
                    </a></p>
                    <p><a href="/download/${data.variables_path}" class="download-link">
                        Download Variables JSON
                    </a></p>
                `;
                result.style.display = 'block';
            } else {
                throw new Error(data.error || 'Processing failed');
            }
        } catch (error) {
            progress.style.width = '100%';
            progress.style.backgroundColor = '#dc3545';

            downloads.innerHTML = `
                <div class="error">
                    Error: ${error.message}
                </div>
            `;
            result.style.display = 'block';
        }
    });
});
