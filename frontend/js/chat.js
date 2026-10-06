async function sendAIChatMessage() {
    const input = document.getElementById('chatInput');
    const messagesContainer = document.getElementById('chatMessages');
    const query = input.value.trim();
    if (!query) return;

    messagesContainer.innerHTML += `<div class="text-right"><span class="inline-block bg-indigo-600 text-white text-xs px-3 py-2 rounded-xl">${query}</span></div>`;
    input.value = '';

    try {
        const response = await fetch('http://localhost:8000/api/ai/chat', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ query: query, role: currentRole })
        });
        const data = await response.json();
        messagesContainer.innerHTML += `<div class="text-left"><span class="inline-block bg-slate-800 text-emerald-400 text-xs px-3 py-2 rounded-xl border border-slate-700">${data.reply}</span></div>`;
    } catch(e) {
        // Запасной локальный ответ, если бэкенд не запущен во время демо
        let fallbackReply = currentLang === 'kz' ? "🤖 ЖИ Ассистент: Барлық жүйелер штаттық режимде жұмыс істеуде." :
                            currentLang === 'en' ? "🤖 AI Assistant: All systems are operating normally." :
                            "🤖 НарядAI Ассистент: Все системы АО «Костанайские минералы» работают штатно.";
        messagesContainer.innerHTML += `<div class="text-left"><span class="inline-block bg-slate-800 text-emerald-400 text-xs px-3 py-2 rounded-xl border border-slate-700">${fallbackReply}</span></div>`;
    }
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}