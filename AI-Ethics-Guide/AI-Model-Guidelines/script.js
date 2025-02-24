// You can add interactive features here, such as toggling visibility of detailed content
document.addEventListener('DOMContentLoaded', function() {
    const sections = document.querySelectorAll('section');
    sections.forEach(section => {
        section.addEventListener('click', function() {
            const details = section.querySelector('p');
            details.style.display = details.style.display === 'none' ? 'block' : 'none';
        });
    });
});
