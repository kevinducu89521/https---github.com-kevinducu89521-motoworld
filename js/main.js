document.addEventListener('DOMContentLoaded', () => {
    const cards = document.querySelectorAll('.card');

    cards.forEach(card => {
        card.addEventListener('click', (event) => {
            if (event.target.closest('a')) return;
            card.classList.toggle('is-flipped');
        });
        card.addEventListener('keydown', (event) => {
            if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                card.classList.toggle('is-flipped');
            }
        });
    });

    const contactForm = document.getElementById('contactForm');
    const formStatus = document.getElementById('formStatus');

    if (contactForm) {
        contactForm.addEventListener('submit', (e) => {
            e.preventDefault();

            const nombre = contactForm.nombre.value.trim();
            const email = contactForm.email.value.trim();
            const asunto = contactForm.asunto.value;
            const mensaje = contactForm.mensaje.value.trim();

            if (!nombre || !email || !asunto || !mensaje) {
                formStatus.textContent = 'Por favor completa todos los campos.';
                formStatus.className = 'form-status error';
                return;
            }

            const submitBtn = contactForm.querySelector('.btn-submit');
            submitBtn.disabled = true;
            submitBtn.querySelector('span').textContent = 'Enviando...';

            setTimeout(() => {
                formStatus.textContent = '¡Mensaje enviado con éxito! Te responderemos pronto.';
                formStatus.className = 'form-status success';
                contactForm.reset();
                submitBtn.disabled = false;
                submitBtn.querySelector('span').textContent = 'Enviar Mensaje';
            }, 1500);
        });
    }
});