document.addEventListener('DOMContentLoaded', function() {
    const exploreBtn = document.getElementById('exploreBtn');
    centerexploreBtn.addEventListener('click', function() {
        alert('Explore nossos produtos e descubra mais sobre a moda surf!');
        
    });

    const contactForm = document.getElementById('contactForm');
    contactForm.addEventListener('submit', function(event) {
        event.preventDefault();
        alert('Obrigado por entrar em contato! Responderemos em breve.');
        contactForm.reset();
    });
});