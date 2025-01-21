function startProgress() {
    document.getElementById('progress-container').style.display = 'block';
    var progressBar = document.getElementById('progress-bar');
    var progressText = document.getElementById('progress-text');
    var successScreen = document.getElementById('success-screen');

    function updateProgress() {
        fetch('/progress')
            .then(response => response.json())
            .then(data => {
                var progress = data.progress;
                progressBar.style.width = progress + '%';
                progressText.textContent = progress + '% completed';
                if (progress < 100) {
                    setTimeout(updateProgress, 1000);
                } else {
                    successScreen.style.display = 'block';
                    document.getElementById('progress-container').style.display = 'none';
                }
            });
    }

    updateProgress();
}