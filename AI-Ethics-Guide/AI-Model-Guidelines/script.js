// You can add interactive features here, such as toggling visibility of detailed content
document.addEventListener('DOMContentLoaded', function() {
    
    fetch('https://script.google.com/macros/s/AKfycbyjO968g4tuex86kWcpmF7ivzQU0C4uCCwLqDSd9lAcELYHj-obFXAhCPcKYMmpzna3Xg/exec')
    .then(response => response.json())
    .then(data => {
        console.log(data);
        const commentsList = document.getElementById('comments-list');
        data.forEach(comment => {
            const listItem = document.createElement('li');
            listItem.textContent = `"${comment.Feedback}" - ${comment.Name} ${comment.Email}`;
            commentsList.appendChild(listItem);
        });
    })
    .catch(error => console.error('Error fetching comments:', error));
    
    const sections = document.querySelectorAll('section');
    sections.forEach(section => {
        section.addEventListener('click', function() {
            const details = section.querySelector('p');
            details.style.display = details.style.display === 'none' ? 'block' : 'none';
        });
    });
});
