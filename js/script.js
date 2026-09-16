/**
 * MotoWorld - Interactive Scripts
 * Horizontal Timeline Navigation & Controls
 */

document.addEventListener('DOMContentLoaded', () => {
    // Horizontal Timeline Controller
    const timeline = document.getElementById('timelineTrack');
    const prevBtn = document.getElementById('timelinePrev');
    const nextBtn = document.getElementById('timelineNext');

    if (timeline && prevBtn && nextBtn) {
        const scrollAmount = 350;

        // Button Click Scrolling
        prevBtn.addEventListener('click', () => {
            timeline.scrollBy({ left: -scrollAmount, behavior: 'smooth' });
        });

        nextBtn.addEventListener('click', () => {
            timeline.scrollBy({ left: scrollAmount, behavior: 'smooth' });
        });

        // Update button states on scroll
        const updateButtonState = () => {
            const maxScrollLeft = timeline.scrollWidth - timeline.clientWidth;
            prevBtn.style.opacity = timeline.scrollLeft <= 10 ? '0.35' : '1';
            prevBtn.style.pointerEvents = timeline.scrollLeft <= 10 ? 'none' : 'auto';
            nextBtn.style.opacity = timeline.scrollLeft >= maxScrollLeft - 10 ? '0.35' : '1';
            nextBtn.style.pointerEvents = timeline.scrollLeft >= maxScrollLeft - 10 ? 'none' : 'auto';
        };

        timeline.addEventListener('scroll', updateButtonState);
        window.addEventListener('resize', updateButtonState);
        updateButtonState();

        // Mouse Drag to Scroll (Desktop)
        let isDown = false;
        let startX;
        let scrollLeftPos;

        timeline.addEventListener('mousedown', (e) => {
            isDown = true;
            timeline.classList.add('active-dragging');
            startX = e.pageX - timeline.offsetLeft;
            scrollLeftPos = timeline.scrollLeft;
        });

        timeline.addEventListener('mouseleave', () => {
            isDown = false;
            timeline.classList.remove('active-dragging');
        });

        timeline.addEventListener('mouseup', () => {
            isDown = false;
            timeline.classList.remove('active-dragging');
        });

        timeline.addEventListener('mousemove', (e) => {
            if (!isDown) return;
            e.preventDefault();
            const x = e.pageX - timeline.offsetLeft;
            const walk = (x - startX) * 1.5; // Drag scroll sensitivity
            timeline.scrollLeft = scrollLeftPos - walk;
        });
    }
});
