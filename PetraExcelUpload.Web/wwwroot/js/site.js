// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

var uploadInput = document.getElementById("upload_btn");
var fileInput = document.getElementById("excel-upload");
let form = document.getElementById('upload-form');

if (uploadInput) {
    uploadInput.addEventListener('click', displayLoadingButton);
}
if (fileInput) {
    fileInput.addEventListener('change', updateInputValue);
}


function displayLoadingButton() {
    

    //! Create span element for spinner and add classes/attr
    let spinner = document.createElement('span');
    spinner.classList.add('spinner-border', 'spinner-border-sm');
    spinner.setAttribute('role', 'status');
    spinner.setAttribute('aria-hidden', 'true');

    this.classList.remove('btn-secondary');
    this.classList.add('btn-primary');
    this.innerText = " Loading...";
    this.prepend(spinner);
    this.disabled = true;

    form.submit();
}

function updateInputValue(element) {
    let fileLabel = document.getElementsByClassName("custom-file-label")[0];
    let fullPath = fileInput.value;

    let startIndex = (fullPath.indexOf('\\') >= 0
                        ? fullPath.lastIndexOf('\\')
                        : fullPath.lastIndexOf('/'));

    let filename = fullPath.substring(startIndex);
    if (filename.indexOf('\\') === 0 || filename.indexOf('/') === 0) {
        filename = filename.substring(1);
    }
    fileLabel.textContent = filename;
}