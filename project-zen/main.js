document.addEventListener('DOMContentLoaded', () => {
    // Elements
    const splashBg = document.getElementById('splash-bg');
    const splashOverlay = document.getElementById('splash-overlay');
    const introContainer = document.getElementById('intro-container');
    const animatedLogoWrapper = document.getElementById('animated-logo-wrapper');
    const animatedText = document.getElementById('animated-text');
    const mainContent = document.getElementById('main-content');
    const agentCards = document.querySelectorAll('.agent-card');

    // Check if the splash screen has already been played in this session
    if (sessionStorage.getItem('splashPlayed') === 'true') {
        // Skip animation and set everything to its final state immediately
        animatedText.innerHTML = "Project<span class='zen-text'>Zen</span>";
        animatedText.classList.add('visible', 'top-left');
        animatedLogoWrapper.classList.add('top-left', 'settled');
        
        splashBg.classList.add('background-mode');
        splashOverlay.classList.add('hidden');
        document.body.style.overflow = 'auto';
        
        mainContent.classList.remove('hidden');
        mainContent.classList.add('visible');
        agentCards.forEach(card => card.classList.add('visible'));
        
        return; // Exit the function to prevent the animation sequence
    }

    // Mark as played so it doesn't run again if user navigates back
    sessionStorage.setItem('splashPlayed', 'true');

    // 1. Text appears first in the center (agentic typing effect)
    setTimeout(() => {
        const textToType = "ProjectZen";
        animatedText.classList.add('visible');
        animatedText.classList.add('agentic-cursor');
        animatedText.innerHTML = ''; // Ensure it's empty
        
        let i = 0;
        const typeInterval = setInterval(() => {
            if (i < textToType.length) {
                const span = document.createElement('span');
                span.textContent = textToType.charAt(i);
                span.className = 'letter-in';
                
                // Add specific styling to the "Zen" part (starts at index 7)
                if (i >= 7) {
                    span.classList.add('zen-text');
                }
                
                animatedText.appendChild(span);
                i++;
            } else {
                clearInterval(typeInterval);
                
                // Remove cursor shortly after typing finishes
                setTimeout(() => {
                    animatedText.classList.remove('agentic-cursor');
                }, 400);
                
                // 2. Start moving text and logo to top left simultaneously
                setTimeout(() => {
                    animatedText.classList.add('top-left');
                    animatedLogoWrapper.classList.add('top-left');
                    
                    // Dim background
                    splashBg.classList.add('background-mode');
                    splashOverlay.classList.add('hidden');
                    
                    document.body.style.overflow = 'auto';

                    // 3. Once it reaches the position (after 1.5s flight), fade opacity from 20% to 100%
                    setTimeout(() => {
                        animatedLogoWrapper.classList.add('settled');
                    }, 1500);

                    // 4. Fade in blocks
                    setTimeout(() => {
                        mainContent.classList.remove('hidden');
                        setTimeout(() => {
                            mainContent.classList.add('visible');
                            agentCards.forEach((card, index) => {
                                setTimeout(() => {
                                    card.classList.add('visible');
                                }, index * 200);
                            });
                        }, 50);
                    }, 1000);
                }, 1200); // Hold completed text for 1.2s before moving
            }
        }, 250); // Slower typing speed (250ms per letter)
    }, 500); // Initial delay
});
