const codeInput = document.getElementById('code');
const submitButton = document.getElementById('submit');

codeInput.addEventListener('input', () => {
    if (codeInput.value.length === 6 && /^\d{6}$/.test(codeInput.value)) { // Copilot did the regex thank you so much
        submitButton.disabled = false;
    } else {
        submitButton.disabled = true;
    }
});