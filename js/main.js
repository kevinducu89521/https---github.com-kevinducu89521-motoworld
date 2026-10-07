// ==========================================================================
// MotoWorld - Interactividad y Animaciones
// ==========================================================================

document.addEventListener('DOMContentLoaded', () => {
    // Soporte para voltear tarjetas al hacer clic o en pantallas táctiles / móviles
    const cards = document.querySelectorAll('.card');

    cards.forEach(card => {
        card.addEventListener('click', (event) => {
            // Si el usuario hace clic en un enlace o botón dentro de la tarjeta, no interferir
            if (event.target.closest('a')) return;

            // Alternar la clase que mantiene la tarjeta volteada
            card.classList.toggle('is-flipped');
        });

        // Soporte de accesibilidad por teclado (Enter o Espacio)
        card.addEventListener('keydown', (event) => {
            if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                card.classList.toggle('is-flipped');
            }
        });
    });
});
