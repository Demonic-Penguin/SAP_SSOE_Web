// JavaScript for SAP Service Order Automation Web Interface

// Wait for the DOM to be fully loaded
document.addEventListener('DOMContentLoaded', function() {
    // Add a small animation to the page title
    const pageTitle = document.querySelector('header h1');
    if (pageTitle) {
        pageTitle.style.opacity = '0';
        setTimeout(() => {
            pageTitle.style.transition = 'opacity 0.5s ease-in-out';
            pageTitle.style.opacity = '1';
        }, 300);
    }
    
    // Auto-focus the service order input field on the index page
    const serviceOrderInput = document.getElementById('service_order');
    if (serviceOrderInput) {
        serviceOrderInput.focus();
    }
    
    // Add confirmation for "No" responses in the wizard
    const noButtons = document.querySelectorAll('button[name="response"][value="no"]');
    noButtons.forEach(button => {
        button.addEventListener('click', function(e) {
            if (!confirm('Are you sure you want to select "No"? This may terminate the SSOE process.')) {
                e.preventDefault();
            }
        });
    });
    
    // Enhance input field validation for part numbers and serial numbers
    const manualInput = document.getElementById('manual_input');
    if (manualInput) {
        manualInput.addEventListener('input', function() {
            // Convert to uppercase for consistency
            this.value = this.value.toUpperCase();
        });
    }
});

// Function to provide visual feedback for completed steps
function markStepComplete(stepNumber) {
    const progressBar = document.querySelector('.progress-bar');
    if (progressBar) {
        progressBar.classList.add('bg-success');
    }
}

// Function to show a warning modal before terminating the process
function confirmTermination(reason) {
    return confirm(`Warning: ${reason}\n\nDo you want to terminate the SSOE process?`);
}