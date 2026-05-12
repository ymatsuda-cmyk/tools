document.addEventListener('DOMContentLoaded', () => {
    const state = {
        currentDate: new Date().toISOString().split('T')[0],
        gasUrl: localStorage.getItem('gas_url') || '',
        newsData: {}
    };

    // Elements
    const displayDate = document.getElementById('display-date');
    const dateTabs = document.getElementById('date-tabs');
    const summarySection = document.getElementById('daily-summary');
    const trendsList = document.getElementById('trends-list');
    const articlesList = document.getElementById('articles-list');
    const settingsBtn = document.getElementById('settings-btn');
    const settingsModal = document.getElementById('settings-modal');
    const closeSettings = document.getElementById('close-settings');
    const saveSettings = document.getElementById('save-settings');
    const gasInput = document.getElementById('gas-url');
    const importBtn = document.getElementById('import-local');
    const toast = document.getElementById('toast');

    // Init
    gasInput.value = state.gasUrl;
    setupDateTabs();
    loadNews(state.currentDate);

    // Event Listeners
    settingsBtn.onclick = () => settingsModal.classList.add('active');
    closeSettings.onclick = () => settingsModal.classList.remove('active');
    
    saveSettings.onclick = () => {
        state.gasUrl = gasInput.value;
        localStorage.setItem('gas_url', state.gasUrl);
        settingsModal.classList.remove('active');
        showToast('Settings Saved');
        loadNews(state.currentDate);
    };

    importBtn.onclick = async () => {
        try {
            // In a real app, this would be a file input or direct fetch
            // Since we know the file exists on the user's machine, 
            // for this demo we'll try to fetch it locally if running on a server,
            // or ask the user to paste the content.
            const response = await fetch('./json/ai-news.json');
            const data = await response.json();
            
            if (data.date) {
                state.newsData[data.date] = data;
                saveToLocal(data.date, data);
                
                // Update state and UI tabs
                state.currentDate = data.date;
                setupDateTabs(); 
                
                // If GAS URL exists, also sync to sheets
                if (state.gasUrl) {
                    syncToSheets(data);
                }
                
                loadNews(data.date);
                showToast('Imported Successfully');
            }
        } catch (e) {
            console.error('Import failed:', e);
            alert('Could not find local ai-news.json or file is invalid. Path: ./json/ai-news.json');
        }
    };

    function setupDateTabs() {
        dateTabs.innerHTML = '';
        for (let i = 0; i < 7; i++) {
            const d = new Date();
            d.setDate(d.getDate() - i);
            const dateStr = d.toISOString().split('T')[0];
            const displayStr = i === 0 ? 'Today' : i === 1 ? 'Yesterday' : `${d.getMonth() + 1}/${d.getDate()}`;
            
            const tab = document.createElement('div');
            tab.className = `date-tab ${dateStr === state.currentDate ? 'active' : ''}`;
            tab.textContent = displayStr;
            tab.onclick = () => {
                document.querySelectorAll('.date-tab').forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                state.currentDate = dateStr;
                loadNews(dateStr);
            };
            dateTabs.appendChild(tab);
        }
    }

    async function loadNews(date) {
        renderLoader();
        
        // 1. Try Memory
        if (state.newsData[date]) {
            renderNews(state.newsData[date]);
            return;
        }

        // 2. Try Local Storage
        const cached = localStorage.getItem(`news_${date}`);
        if (cached) {
            const data = JSON.parse(cached);
            state.newsData[date] = data;
            renderNews(data);
            return;
        }

        // 3. Try GAS
        if (state.gasUrl) {
            try {
                const res = await fetch(`${state.gasUrl}?date=${date}`);
                if (!res.ok) {
                    throw new Error(`HTTP error! status: ${res.status}`);
                }
                const data = await res.json();
                if (data && !data.error) {
                    state.newsData[date] = data;
                    saveToLocal(date, data);
                    renderNews(data);
                    return;
                } else if (data && data.error) {
                    console.warn('GAS returned error:', data.error);
                }
            } catch (e) {
                console.error('GAS Fetch failed:', e.message || e);
            }
        }

        renderEmpty();
    }

    function renderNews(data) {
        if (!data) {
            renderEmpty();
            return;
        }

        displayDate.textContent = data.date ? formatDate(data.date) : 'Unknown Date';
        
        // Summary
        summarySection.innerHTML = `<p>${data.summary || 'No summary available.'}</p>`;
        
        // Trends
        const trends = Array.isArray(data.trends) ? data.trends : [];
        trendsList.innerHTML = trends.map(t => `<li>${t}</li>`).join('');
        
        // Articles
        const articles = Array.isArray(data.articles) ? data.articles : [];
        articlesList.innerHTML = articles.map(a => {
            if (!a) return '';
            return `
            <div class="article-card">
                <div class="article-header">
                    <span class="category-tag">${a.category || 'General'}</span>
                    <span class="impact-badge impact-${a.impact === '高' ? 'high' : a.impact === '中' ? 'med' : 'low'}">
                        Impact: ${a.impact || 'Unknown'}
                    </span>
                </div>
                <div class="article-body">
                    <h3>${a.title || 'Untitled'}</h3>
                    <p>${a.summary || ''}</p>
                    <div class="expandable-section">
                        <span class="section-label">Beginner Note</span>
                        <div class="section-content">${a.beginnerNote || ''}</div>
                    </div>
                    <div class="expandable-section">
                        <span class="section-label">Impact Note</span>
                        <div class="section-content">${a.impactNote || ''}</div>
                    </div>
                    <div style="font-size: 11px; color: var(--text-secondary); margin-top: 15px;">Source: ${a.source || 'Unknown'}</div>
                </div>
            </div>
        `}).join('');

        if (articles.length === 0 && trends.length === 0 && !data.summary) {
            renderEmpty();
        }
    }

    function renderLoader() {
        summarySection.innerHTML = '<div class="loader"></div>';
        trendsList.innerHTML = '';
        articlesList.innerHTML = '';
    }

    function renderEmpty() {
        summarySection.innerHTML = '<p style="text-align:center; color:var(--text-secondary);">No news found for this date.<br>Try importing local JSON in settings.</p>';
        trendsList.innerHTML = '';
        articlesList.innerHTML = '';
    }

    function saveToLocal(date, data) {
        localStorage.setItem(`news_${date}`, JSON.stringify(data));
    }

    async function syncToSheets(data) {
        if (!state.gasUrl) return;
        try {
            await fetch(state.gasUrl, {
                method: 'POST',
                mode: 'no-cors', // GAS web apps often require no-cors for simple POSTs
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            });
        } catch (e) {
            console.error('Sync failed', e);
        }
    }

    function formatDate(dateStr) {
        const d = new Date(dateStr);
        return d.toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' });
    }

    function showToast(msg) {
        toast.textContent = msg;
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 3000);
    }
});
