import { AppCore } from './core.js';

window.onload = () => {
    AppCore.init('QA');
};

window.toggleMenu = function(event, id) {
    event.stopPropagation();
    document.querySelectorAll(".dropdown").forEach(d => d.style.display = "none");
    const menu = document.getElementById(`menu-${id}`);
    if (menu) menu.style.display = "block";
};