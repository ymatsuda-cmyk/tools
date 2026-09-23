(function(root, factory){
  const api = factory();
  if(typeof module === 'object' && module.exports) module.exports = api;
  root.DashboardConfig = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function(){
  const DEFAULT_TAB_IDS = new Set(['links', 'monitor', 'info']);

  function customTabs(value){
    if(!Array.isArray(value)) return [];
    const used = new Set();
    return value.reduce((tabs, tab) => {
      const id = String(tab && tab.id || '').trim();
      const label = String(tab && (tab.label || tab.title) || '').trim();
      if(!/^[A-Za-z0-9_-]+$/.test(id) || !label || DEFAULT_TAB_IDS.has(id) || used.has(id)) return tabs;
      used.add(id);
      tabs.push({ id, label });
      return tabs;
    }, []);
  }

  function cardTab(card, fallback, tabs){
    const requested = String(card && card.tab || '').trim();
    return tabs.some(tab => tab.id === requested) ? requested : fallback;
  }

  function tabOptions(defaultTab, tabs){
    const labels = { links:'リンク', monitor:'稼働状況', info:'情報' };
    return [{ id:defaultTab, label:labels[defaultTab] || defaultTab }].concat(tabs);
  }

  function isVisible(card){
    return !card || card.hidden !== true;
  }

  return { customTabs, cardTab, tabOptions, isVisible };
});