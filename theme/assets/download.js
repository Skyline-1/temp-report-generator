function downloadBlob(url, filename) {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename; // The name you want to give the downloaded file
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}