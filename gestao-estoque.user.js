// ==UserScript==
// @name         Gestão de Estoque
// @namespace    http://tampermonkey.net/
// @version      1.0.5
// @description  Controle de encomendas, notificações e análise contínua de trade sobreposta.
// @author       ivanils0n
// @match        *://*/*
// @updateURL    https://raw.githubusercontent.com/ivanils0n/script-gestao-de-estoque/refs/heads/main/gestao-estoque.user.js
// @downloadURL  https://raw.githubusercontent.com/ivanils0n/script-gestao-de-estoque/refs/heads/main/gestao-estoque.user.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/dist/umd/supabase.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// ==/UserScript==

(function() {
    'use strict';

    // CSS Otimizado
    GM_addStyle(`
        #tm-content-wrapper { transition: margin-right 0.28s cubic-bezier(0.4, 0, 0.2, 1); margin-right: 0; }
        #tm-content-wrapper.tm-pushed { margin-right: 380px; }

        #tm-manager-btn { position: fixed; top: 50%; right: 16px; z-index: 999999999; width: 48px; height: 48px; border-radius: 50%; background-color: #ffffff; color: #0f172a; border: none; box-shadow: 0 4px 14px rgba(0,0,0,0.18); cursor: grab; display: flex; align-items: center; justify-content: center; font-size: 22px; opacity: 0.8; transition: opacity 0.25s ease, right 0.28s cubic-bezier(0.4, 0, 0.2, 1); user-select: none; touch-action: none; }
        #tm-manager-btn:hover { opacity: 1; }
        #tm-manager-btn.dragging { opacity: 1; cursor: grabbing; transition: none; box-shadow: 0 6px 20px rgba(0,0,0,0.28); }
        #tm-manager-btn.closed-state { right: 16px; }
        #tm-manager-btn.open-state { right: 396px; }

        #tm-manager-panel { position: fixed; top: 0; right: -400px; width: 380px; height: 100vh; background-color: #ffffff; color: #334155; z-index: 999999998; box-shadow: -5px 0 25px rgba(0,0,0,0.2); transition: right 0.28s cubic-bezier(0.4, 0, 0.2, 1); font-family: system-ui, -apple-system, sans-serif; display: flex; flex-direction: column; box-sizing: border-box; }
        #tm-manager-panel.open { right: 0; }

        #tm-history-panel { position: fixed; top: 0; right: 380px; width: 0; height: 100vh; background-color: #ffffff; color: #334155; z-index: 999999997; overflow: hidden; transition: width 0.2s cubic-bezier(0.075, 0.82, 0.165, 1); font-family: system-ui, -apple-system, sans-serif; box-sizing: border-box; }
        #tm-history-panel.open { width: 400px; box-shadow: -5px 0 20px rgba(0,0,0,0.12); border-left: 1px solid #e2e8f0; }
        #tm-history-panel-inner { height: 100%; width: 400px; box-sizing: border-box; padding: 24px 28px; display: flex; flex-direction: column; overflow: hidden; }
        #tm-history-panel-inner h3 { margin: 0 0 14px 0; color: #0f172a; font-size: 1.1rem; flex-shrink: 0; }
        #tm-history-list-slot { flex: 1; display: flex; flex-direction: column; min-height: 0; overflow: hidden; }
        #tm-history-panel .tm-list-container { flex: 1; min-height: 0; }
        #tm-history-panel .tm-list { max-height: none; flex: 1; min-height: 0; }

        .tm-header { padding: 15px 20px; background-color: #0f172a; color: #ffffff; display: flex; justify-content: space-between; align-items: center; }
        .tm-header h3 { margin: 0; font-size: 1.1rem; font-weight: 600; }
        .tm-header-actions { display: flex; gap: 12px; align-items: center; }
        .tm-icon-btn { background: none; border: none; color: #94a3b8; cursor: pointer; font-size: 18px; padding: 0; transition: color 0.2s; }
        .tm-icon-btn:hover { color: #ffffff; }
        .tm-close-btn { margin-left: 5px; font-size: 20px; }
        .tm-close-btn:hover { color: #ef4444; }

        .tm-tabs { display: flex; background-color: #1e293b; }
        .tm-tab { flex: 1; padding: 12px; text-align: center; cursor: pointer; background: none; border: none; color: #94a3b8; font-weight: 500; transition: all 0.2s; font-size: 13px; }
        .tm-tab.active { color: #0f172a; background-color: #ffffff; font-weight: bold; }
        .tm-content { flex: 1; padding: 20px; display: flex; flex-direction: column; overflow: hidden; }

        .tm-form { background-color: #f8fafc; padding: 15px; border-radius: 8px; margin-bottom: 15px; display: flex; flex-direction: column; gap: 10px; border: 1px solid #e2e8f0; flex-shrink: 0; }
        .tm-form input, .tm-form textarea, .tm-form select { background-color: #ffffff; border: 1px solid #cbd5e1; color: #0f172a; padding: 8px 12px; border-radius: 6px; font-size: 14px; outline: none; }
        .tm-form input:focus, .tm-form textarea:focus, .tm-form select:focus { border-color: #3b82f6; }
        .tm-form button { background-color: #0f172a; color: white; border: none; padding: 10px; border-radius: 6px; cursor: pointer; font-weight: bold; transition: background 0.2s; }
        .tm-form button:hover { background-color: #1e293b; }

        .tm-search-container { margin-bottom: 10px; display: flex; gap: 8px; flex-shrink: 0; }
        .tm-search-input { flex: 1; background-color: #f1f5f9; border: 1px solid #cbd5e1; padding: 8px 12px; border-radius: 6px; font-size: 13px; outline: none; transition: all 0.2s; }
        .tm-search-input:focus { border-color: #3b82f6; background-color: #ffffff; box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.15); }

        .tm-list-container { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
        .tm-list-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; padding: 0 5px; font-size: 13px; flex-shrink: 0; }
        .tm-list-header label { display: flex; align-items: center; gap: 8px; cursor: pointer; color: #475569; font-weight: 500; }
        .tm-list { display: flex; flex-direction: column; gap: 12px; overflow-y: auto; padding-right: 5px; padding-bottom: 20px; max-height: 380px; }

        .tm-list::-webkit-scrollbar { width: 6px; }
        .tm-list::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
        .tm-list::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        .tm-list::-webkit-scrollbar-thumb:hover { background: #94a3b8; }

        .tm-card { background-color: #ffffff; border: 1px solid #e2e8f0; border-left: 4px solid #3b82f6; padding: 12px; border-radius: 4px 8px 8px 4px; position: relative; box-shadow: 0 1px 3px rgba(0,0,0,0.05); display: flex; gap: 10px; flex-shrink: 0; }
        .tm-card.obs-card { border-left-color: #f59e0b; }
        .tm-card-content { flex: 1; }
        .tm-card h4 { margin: 0 0 5px 0; font-size: 14px; color: #0f172a; }
        .tm-card p { margin: 3px 0; font-size: 13px; color: #475569; }
        .tm-card-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 10px; }
        .tm-btn-small { padding: 4px 8px; font-size: 11px; border-radius: 4px; cursor: pointer; border: none; color: white; font-weight: 500; }
        .tm-btn-edit { background-color: #3b82f6; } .tm-btn-edit:hover { background-color: #2563eb; }
        .tm-btn-del { background-color: #ef4444; } .tm-btn-del:hover { background-color: #dc2626; }
        .tm-checkbox { cursor: pointer; width: 16px; height: 16px; margin-top: 2px; }

        .tm-notify-check { position: absolute; top: 12px; right: 12px; cursor: pointer; width: 18px; height: 18px; accent-color: #f59e0b; }

        #tm-notification-wrapper { position: fixed; bottom: 20px; left: 20px; z-index: 999999999; display: flex; flex-direction: column; gap: 10px; font-family: system-ui, sans-serif; }
        .tm-alert-card { background: #ffffff; border-left: 5px solid #f59e0b; box-shadow: 0 10px 25px rgba(0,0,0,0.2); padding: 15px 20px; border-radius: 8px; width: 320px; position: relative; animation: slideInNotify 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275); }
        .tm-alert-card.tm-alert-out { animation: slideOutNotify 0.35s cubic-bezier(0.4, 0, 0.2, 1) forwards; }
        .tm-alert-close { position: absolute; top: 10px; right: 10px; background: none; border: none; font-size: 16px; cursor: pointer; color: #94a3b8; }
        .tm-alert-close:hover { color: #ef4444; }
        .tm-toast-body { transition: background 0.2s; padding: 5px; border-radius: 6px; margin-top: -5px; }
        .tm-toast-body:hover.clickable { background: #f8fafc; }
        .tm-toast-body h4 { margin: 0 0 5px 0; color: #0f172a; font-size: 15px; }
        .tm-toast-body p { margin: 0; font-size: 13px; color: #475569; }

        .tm-trade-sync { line-height: 24px !important; padding-top: 4px !important; padding-bottom: 4px !important; background-image: repeating-linear-gradient(transparent, transparent 23px, #e2e8f0 23px, #e2e8f0 24px) !important; background-attachment: local !important; }

        #tm-confirm-modal, #tm-config-modal, #tm-sql-modal { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(15, 23, 42, 0.6); z-index: 999999999; display: flex; align-items: center; justify-content: center; font-family: system-ui, sans-serif; }
        .tm-modal-content { background: #ffffff; padding: 25px; border-radius: 12px; width: 360px; text-align: center; box-shadow: 0 20px 40px rgba(0,0,0,0.3); border: 1px solid #e2e8f0; }
        .tm-modal-content h3 { margin: 0 0 12px 0; color: #0f172a; font-size: 18px; }
        .tm-modal-content p { color: #475569; font-size: 14px; margin-bottom: 24px; line-height: 1.5; }
        .tm-modal-actions { display: flex; gap: 12px; justify-content: center; }
        .tm-modal-btn { padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 14px; }
        .tm-btn-confirm { background: #0f172a; color: white; transition: background 0.2s; }
        .tm-btn-confirm:hover { background: #1e293b; }
        .tm-btn-cancel { background: #e2e8f0; color: #334155; transition: background 0.2s; }
        .tm-btn-cancel:hover { background: #cbd5e1; }

        .tm-cfg-label { display: block; font-size: 11px; color: #94a3b8; margin-bottom: 6px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em; }
        .tm-cfg-input, .tm-cfg-select { width: 100%; padding: 10px 12px; border-radius: 10px; border: 1px solid #e2e8f0; margin-bottom: 14px; box-sizing: border-box; font-size: 13px; background: #f8fafc; transition: border-color 0.2s, box-shadow 0.2s, background 0.2s; }
        .tm-cfg-input:focus, .tm-cfg-select:focus { outline: none; border-color: #94a3b8; background: #ffffff; box-shadow: 0 0 0 3px rgba(15, 23, 42, 0.06); }
        .tm-cfg-input:disabled { background: #f1f5f9; color: #94a3b8; cursor: not-allowed; border-style: dashed; }
        .tm-sql-box { width: 100%; height: 220px; font-family: monospace; font-size: 11px; padding: 8px; border-radius: 6px; border: 1px solid #cbd5e1; box-sizing: border-box; resize: none; }

        #tm-config-modal .tm-modal-content { width: 380px; text-align: left; padding: 0; overflow: hidden; border-radius: 16px; }
        .tm-cfg-header { background: #ffffff; color: #0f172a; padding: 20px 24px 4px; display: flex; align-items: center; gap: 10px; }
        .tm-cfg-header span { font-size: 16px; opacity: 0.7; }
        .tm-cfg-header h3 { margin: 0; font-size: 15px; font-weight: 700; letter-spacing: -0.01em; }
        .tm-cfg-body { padding: 16px 24px 6px; }
        .tm-cfg-section { margin-bottom: 14px; }
        .tm-cfg-status { display: flex; align-items: center; gap: 8px; padding: 9px 12px; border-radius: 10px; font-size: 12px; font-weight: 600; margin-bottom: 16px; }
        .tm-cfg-status.on { background: #f0fdf4; color: #16a34a; }
        .tm-cfg-status.off { background: #f8fafc; color: #94a3b8; }
        .tm-cfg-divider { border: none; border-top: 1px solid #f1f5f9; margin: 4px 0 16px; }
        .tm-cfg-actions-secondary { display: flex; flex-direction: column; gap: 8px; margin-bottom: 6px; }
        .tm-cfg-actions-secondary button { width: 100%; border: 1px solid transparent; padding: 10px; border-radius: 10px; cursor: pointer; font-weight: 600; font-size: 12.5px; transition: filter 0.15s, background 0.15s; }
        .tm-cfg-actions-secondary button:hover { filter: brightness(0.96); }
        #tm-config-modal .tm-modal-actions { padding: 14px 24px 22px; margin: 0; }
        .tm-modal-btn:disabled { opacity: 0.5; cursor: not-allowed; }

        @keyframes slideInNotify { from { transform: translateX(-120%); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
        @keyframes slideOutNotify { from { transform: translateX(0); opacity: 1; } to { transform: translateX(-120%); opacity: 0; } }
    `);

    // --- Core Functions e Banco de Dados ---
    let db = GM_getValue('tm_store_db', { encomendas: [], observacoes: [] });
    let realtimeChannel = null;

    if (!Array.isArray(db.encomendas)) db.encomendas = [];
    if (!Array.isArray(db.observacoes)) db.observacoes = [];

    const saveDB = () => GM_setValue('tm_store_db', db);
    const generateUniqueId = () => Date.now().toString(36) + Math.random().toString(36).substring(2, 9);

    const todayISO = () => {
        const d = new Date();
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };

    const isoDateOffset = (days) => {
        const d = new Date();
        d.setDate(d.getDate() + days);
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };

    // --- Configuração de Armazenamento ---
    const getConfig = () => GM_getValue('tm_config', { mode: 'local', supabaseUrl: '', supabaseKey: '' });
    const saveConfig = (cfg) => GM_setValue('tm_config', cfg);

    const SQL_SCHEMA = `create table if not exists public.encomendas (
  id text primary key,
  loja text,
  produto text,
  ean text,
  qtd text,
  data date,
  created_at timestamptz default now()
);

create table if not exists public.observacoes (
  id text primary key,
  titulo text,
  data date,
  "desc" text,
  notify boolean default false,
  notify_weekday text,
  created_at timestamptz default now()
);

alter table public.observacoes add column if not exists notify_weekday text;

alter table public.encomendas enable row level security;
alter table public.observacoes enable row level security;

create policy "encomendas_all" on public.encomendas for all using (true) with check (true);
create policy "observacoes_all" on public.observacoes for all using (true) with check (true);

grant all on public.encomendas to anon, authenticated;
grant all on public.observacoes to anon, authenticated;

alter table public.encomendas replica identity full;
alter table public.observacoes replica identity full;

do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'encomendas'
  ) then
    alter publication supabase_realtime add table public.encomendas;
  end if;
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'observacoes'
  ) then
    alter publication supabase_realtime add table public.observacoes;
  end if;
end $$;`;

    // Limite de dias para evitar filtros de data muito amplos (consomem egress do banco)
    const MAX_FILTER_RANGE_DAYS = 60;

    function resyncFilteredRange() {
        const cfg = getConfig();
        if (cfg.mode === 'supabase' && cfg.supabaseUrl && cfg.supabaseKey) supaSyncAll(false, true);
    }

    function applyDateFilterChange(newFrom, newTo, revertFn) {
        if (newFrom && newTo) {
            const span = Math.abs((new Date(newTo) - new Date(newFrom)) / 86400000);
            if (span > MAX_FILTER_RANGE_DAYS) {
                showConfirmModal(
                    "⚠️ Intervalo muito grande",
                    "Filtrar um período maior que 2 meses consome bastante egress do banco de dados. Deseja continuar mesmo assim?",
                    () => { dateFrom = newFrom; dateTo = newTo; applyFilters(); resyncFilteredRange(); },
                    "Aplicar Mesmo Assim",
                    true,
                    revertFn
                );
                return;
            }
        }
        dateFrom = newFrom; dateTo = newTo; applyFilters();
        resyncFilteredRange();
    }

    function normalizeSupabaseUrl(url) {
        return (url || '').trim().replace(/\/+$/, '').replace(/\/rest\/v1$/i, '');
    }

    function supaHeaders(cfg, extraPrefer) {
        return {
            'apikey': cfg.supabaseKey,
            'Authorization': `Bearer ${cfg.supabaseKey}`,
            'Content-Type': 'application/json',
            'Prefer': extraPrefer || 'return=minimal,resolution=merge-duplicates'
        };
    }

    async function supaFetchTable(cfg, table, from, to) {
        const base = normalizeSupabaseUrl(cfg.supabaseUrl);
        const columns = table === 'encomendas'
            ? 'id,loja,produto,ean,qtd,data'
            : 'id,titulo,data,desc,notify,notify_weekday';
        let url = `${base}/rest/v1/${table}?select=${columns}&order=created_at.desc`;
        if (from && to) url += `&or=(data.is.null,and(data.gte.${from},data.lte.${to}))`;
        const res = await fetch(url, { headers: { 'apikey': cfg.supabaseKey, 'Authorization': `Bearer ${cfg.supabaseKey}` } });
        if (!res.ok) throw new Error(`Erro ao buscar ${table}: ${res.status}`);
        return res.json();
    }

    async function supaFetchNotifyCandidates(cfg) {
        const base = normalizeSupabaseUrl(cfg.supabaseUrl);
        const columns = 'id,titulo,data,desc,notify,notify_weekday';
        const url = `${base}/rest/v1/observacoes?select=${columns}&notify=eq.true`;
        const res = await fetch(url, { headers: { 'apikey': cfg.supabaseKey, 'Authorization': `Bearer ${cfg.supabaseKey}` } });
        if (!res.ok) throw new Error(`Erro ao buscar notificações: ${res.status}`);
        return res.json();
    }

    function mergeSyncedItems(table, fetchedItems, from, to) {
        const fetchedMap = new Map(fetchedItems.map(it => [String(it.id), it]));
        const inWindow = (it) => !it.data || (from && to ? (it.data >= from && it.data <= to) : true);

        const kept = db[table].filter(it => !inWindow(it) || fetchedMap.has(String(it.id)));
        fetchedItems.forEach(it => {
            const idx = kept.findIndex(x => String(x.id) === String(it.id));
            if (idx > -1) kept[idx] = it; else kept.push(it);
        });
        db[table] = kept;
    }

    // --- Realtime ---
    function handleRealtimeChange(table) {
        return (payload) => {
            const list = db[table];
            if (payload.eventType === 'INSERT') {
                if (!list.find(x => String(x.id) === String(payload.new.id))) list.unshift(payload.new);
            } else if (payload.eventType === 'UPDATE') {
                const idx = list.findIndex(x => String(x.id) === String(payload.new.id));
                if (idx > -1) list[idx] = payload.new; else list.unshift(payload.new);
            } else if (payload.eventType === 'DELETE' && payload.old && payload.old.id !== undefined) {
                db[table] = list.filter(x => String(x.id) !== String(payload.old.id));
            }
            GM_setValue('tm_store_db', db);
            render();
        };
    }

    function setupRealtime(cfg) {
        if (realtimeChannel) {
            try { realtimeChannel.unsubscribe(); } catch (e) {}
            realtimeChannel = null;
        }
        if (cfg.mode !== 'supabase' || !cfg.supabaseUrl || !cfg.supabaseKey) return;
        if (typeof supabase === 'undefined' || !supabase.createClient) return;

        const client = supabase.createClient(normalizeSupabaseUrl(cfg.supabaseUrl), cfg.supabaseKey);
        realtimeChannel = client
            .channel('tm-store-changes')
            .on('postgres_changes', { event: '*', schema: 'public', table: 'encomendas' }, handleRealtimeChange('encomendas'))
            .on('postgres_changes', { event: '*', schema: 'public', table: 'observacoes' }, handleRealtimeChange('observacoes'))
            .subscribe();
    }

    async function supaSyncAll(notify = false, force = false) {
        const cfg = getConfig();
        if (cfg.mode !== 'supabase' || !cfg.supabaseUrl || !cfg.supabaseKey) return;

        const THROTTLE_MS = 5 * 60 * 1000;
        const lastSync = GM_getValue('tm_last_sync_at', 0);
        if (!force && (Date.now() - lastSync) < THROTTLE_MS) return;

        try {
            const [enc, obs] = await Promise.all([
                supaFetchTable(cfg, 'encomendas', dateFrom, dateTo),
                supaFetchTable(cfg, 'observacoes', dateFrom, dateTo)
            ]);
            mergeSyncedItems('encomendas', enc || [], dateFrom, dateTo);
            mergeSyncedItems('observacoes', obs || [], dateFrom, dateTo);
            GM_setValue('tm_store_db', db);
            GM_setValue('tm_last_sync_at', Date.now());
            render();
            if (notify) showToast("☁️ Sincronizado", "Dados carregados do banco de dados.", "#10b981");
        } catch (err) {
            showToast("⚠️ Erro Banco de dados", "Não foi possível carregar os dados. Verifique URL/chave.", "#ef4444");
        }
    }

    async function supaUpsert(table, itemsOrItem) {
        const cfg = getConfig();
        if (cfg.mode !== 'supabase' || !cfg.supabaseUrl || !cfg.supabaseKey) return;
        const base = normalizeSupabaseUrl(cfg.supabaseUrl);
        const body = Array.isArray(itemsOrItem) ? itemsOrItem : [itemsOrItem];
        if (body.length === 0) return;
        try {
            const res = await fetch(`${base}/rest/v1/${table}?on_conflict=id`, {
                method: 'POST', headers: supaHeaders(cfg), body: JSON.stringify(body)
            });
            if (!res.ok) throw new Error(res.status);
        } catch (err) { }
    }

    async function supaDelete(table, ids) {
        const cfg = getConfig();
        if (cfg.mode !== 'supabase' || !cfg.supabaseUrl || !cfg.supabaseKey) return;
        if (!ids || ids.length === 0) return;
        const base = normalizeSupabaseUrl(cfg.supabaseUrl);
        try {
            const filter = ids.length === 1 ? `id=eq.${encodeURIComponent(ids[0])}` : `id=in.(${ids.map(encodeURIComponent).join(',')})`;
            await fetch(`${base}/rest/v1/${table}?${filter}`, { method: 'DELETE', headers: supaHeaders(cfg) });
        } catch (err) { }
    }

    const body = document.body;

    // --- Wrapper de conteúdo (efeito "push", estilo Sider) ---
    const contentWrapper = document.createElement('div');
    contentWrapper.id = 'tm-content-wrapper';
    while (body.firstChild) contentWrapper.appendChild(body.firstChild);
    body.appendChild(contentWrapper);

    // --- Inputs e Componentes Globais ---
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx, .xls, .csv';
    fileInput.style.display = 'none';
    body.appendChild(fileInput);

    const toggleBtn = document.createElement('button');
    toggleBtn.id = 'tm-manager-btn';
    toggleBtn.className = 'closed-state';
    toggleBtn.innerHTML = '📦';
    toggleBtn.title = 'Gestão de Estoque';
    body.appendChild(toggleBtn);

    // Posição vertical salva (arraste pela lateral)
    const savedBtnTop = GM_getValue('tm_btn_top', null);
    toggleBtn.style.top = (savedBtnTop !== null ? savedBtnTop : (window.innerHeight / 2 - 24)) + 'px';

    // --- Arrastar o botão flutuante verticalmente pela lateral ---
    (() => {
        let dragging = false, moved = false, startY = 0, startTop = 0;
        const MARGIN = 8;

        const onMove = (clientY) => {
            const delta = clientY - startY;
            if (Math.abs(delta) > 3) moved = true;
            let newTop = startTop + delta;
            newTop = Math.max(MARGIN, Math.min(window.innerHeight - 48 - MARGIN, newTop));
            toggleBtn.style.top = newTop + 'px';
        };

        const onDown = (clientY) => {
            dragging = true;
            moved = false;
            startY = clientY;
            startTop = parseFloat(toggleBtn.style.top) || 0;
            toggleBtn.classList.add('dragging');
        };

        const onUp = () => {
            if (!dragging) return;
            dragging = false;
            toggleBtn.classList.remove('dragging');
            if (moved) GM_setValue('tm_btn_top', parseFloat(toggleBtn.style.top) || 0);
        };

        toggleBtn.addEventListener('mousedown', (e) => { onDown(e.clientY); e.preventDefault(); });
        document.addEventListener('mousemove', (e) => { if (dragging) onMove(e.clientY); });
        document.addEventListener('mouseup', onUp);

        toggleBtn.addEventListener('touchstart', (e) => { onDown(e.touches[0].clientY); }, { passive: true });
        document.addEventListener('touchmove', (e) => { if (dragging) onMove(e.touches[0].clientY); }, { passive: true });
        document.addEventListener('touchend', onUp);

        // Bloqueia o clique de abrir/fechar caso tenha havido arraste
        toggleBtn.addEventListener('click', (e) => { if (moved) { e.stopPropagation(); e.preventDefault(); } }, true);
    })();

    // Botão fixo na borda direita: clique abre/fecha a sidebar e empurra o conteúdo da página
    toggleBtn.addEventListener('click', () => {
        const isOpen = panel.classList.contains('open');

        if (isOpen) {
            closeWidget();
        } else {
            // Abrir a Sidebar e empurrar o conteúdo da página
            panel.classList.add('open');
            toggleBtn.className = 'open-state';
            contentWrapper.classList.add('tm-pushed');
            // Em seguida o histórico, caso devesse estar aberto e não estarmos no Trade
            if (isExpanded && activeTab !== 'trade') {
                setTimeout(() => { historyPanel.classList.add('open'); }, PANEL_ANIM_MS);
            }
            const cfg = getConfig();
            if (cfg.mode === 'supabase' && cfg.supabaseUrl && cfg.supabaseKey) supaSyncAll();
        }
    });

    // --- Painel Estrutura Base ---
    const panel = document.createElement('div');
    panel.id = 'tm-manager-panel';
    panel.innerHTML = `
        <div class="tm-header">
            <h3>Gestão de Estoque</h3>
            <div class="tm-header-actions">
                <button class="tm-icon-btn" id="tm-import-btn" title="Importar Planilha">📥</button>
                <button class="tm-icon-btn" id="tm-export-btn" title="Exportar Planilha">📤</button>
                <button class="tm-icon-btn" id="tm-config-btn" title="Configurações">⚙️</button>
                <button class="tm-icon-btn" id="tm-expand-btn" title="Expandir">⤢</button>
                <button class="tm-close-btn" id="tm-close-panel" title="Fechar">✕</button>
            </div>
        </div>
        <div class="tm-tabs">
            <button class="tm-tab active" id="tab-encomendas">Encomendas</button>
            <button class="tm-tab" id="tab-observacoes">Observações</button>
            <button class="tm-tab" id="tab-trade">Trade</button>
        </div>
        <div class="tm-content" id="tm-panel-content"></div>
    `;
    body.appendChild(panel);

    const historyPanel = document.createElement('div');
    historyPanel.id = 'tm-history-panel';
    historyPanel.innerHTML = `<div id="tm-history-panel-inner"><h3 id="tm-history-title">Histórico</h3><div id="tm-history-list-slot"></div></div>`;
    body.appendChild(historyPanel);
    const historyListSlot = document.getElementById('tm-history-list-slot');

    function closeWidget() {
        contentWrapper.classList.remove('tm-pushed');
        if (isExpanded) {
            historyPanel.classList.remove('open');
            setTimeout(() => {
                panel.classList.remove('open');
                toggleBtn.className = 'closed-state';
            }, HISTORY_ANIM_MS); // Aguarda fechar o histórico primeiro
        } else {
            panel.classList.remove('open');
            toggleBtn.className = 'closed-state';
        }
    }

    document.getElementById('tm-close-panel').addEventListener('click', closeWidget);
    document.getElementById('tm-import-btn').addEventListener('click', () => fileInput.click());
    document.getElementById('tm-export-btn').addEventListener('click', () => exportCurrentTab());
    document.getElementById('tm-config-btn').addEventListener('click', () => showConfigModal());

    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        fileInput.value = ''; // permite reimportar o mesmo arquivo depois
        if (!file) return;
        if (activeTab === 'trade') {
            showToast("⚠️ Aba Inválida", "Selecione a aba Encomendas ou Observações antes de importar.", "#ef4444");
            return;
        }
        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const wb = XLSX.read(evt.target.result, { type: 'binary' });
                const sheet = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                importRows(rows);
            } catch (err) {
                showToast("⚠️ Erro na Importação", "Não foi possível ler a planilha selecionada.", "#ef4444");
            }
        };
        reader.readAsBinaryString(file);
    });

    let isExpanded = false;
    const expandBtn = document.getElementById('tm-expand-btn');
    expandBtn.addEventListener('click', () => {
        isExpanded = !isExpanded;
        historyPanel.classList.toggle('open', isExpanded);
        expandBtn.innerHTML = isExpanded ? '⤡' : '⤢';
        expandBtn.title = isExpanded ? 'Recolher Histórico' : 'Expandir Histórico';
        render();
    });

    let activeTab = 'encomendas', editingId = null, selectedIds = new Set(), searchQuery = "";
    let hideUndated = false, dateFrom = isoDateOffset(-5), dateTo = isoDateOffset(5), tradeInterval = null;

    // Duração (ms) sincronizada com as transições CSS de #tm-manager-panel e #tm-history-panel
    const PANEL_ANIM_MS = 250;
    const HISTORY_ANIM_MS = 200;

    const tabEnc = document.getElementById('tab-encomendas');
    const tabObs = document.getElementById('tab-observacoes');
    const tabTrade = document.getElementById('tab-trade');

    tabEnc.addEventListener('click', () => switchTab('encomendas'));
    tabObs.addEventListener('click', () => switchTab('observacoes'));
    tabTrade.addEventListener('click', () => switchTab('trade'));

    function switchTab(tabName) {
        if (tradeInterval) {
            clearInterval(tradeInterval);
            tradeInterval = null;
            clearTradeHighlights();
        }
        activeTab = tabName; editingId = null; selectedIds.clear(); searchQuery = "";
        dateFrom = isoDateOffset(-5); dateTo = isoDateOffset(5);

        tabEnc.classList.toggle('active', tabName === 'encomendas');
        tabObs.classList.toggle('active', tabName === 'observacoes');
        tabTrade.classList.toggle('active', tabName === 'trade');

        // Desativa expandir histórico na aba trade para evitar bugs
        if (tabName === 'trade') {
            expandBtn.style.opacity = '0.3';
            expandBtn.style.pointerEvents = 'none';
            if (isExpanded) {
                isExpanded = false;
                historyPanel.classList.remove('open');
                expandBtn.innerHTML = '⤢';
            }
        } else {
            expandBtn.style.opacity = '1';
            expandBtn.style.pointerEvents = 'auto';
            if (isExpanded && panel.classList.contains('open')) {
                historyPanel.classList.add('open');
            }
        }

        render();
    }

    // --- Sistema de Toasts e Notificações (Snooze + Abertura Integrada) ---
    function showToast(title, message, borderColor = '#3b82f6', actionData = null) {
        let wrapper = document.getElementById('tm-notification-wrapper');
        if (!wrapper) {
            wrapper = document.createElement('div');
            wrapper.id = 'tm-notification-wrapper';
            body.appendChild(wrapper);
        }
        const alertCard = document.createElement('div');
        alertCard.className = 'tm-alert-card';
        alertCard.style.borderLeftColor = borderColor;

        let actionHTML = '';
        if (actionData) {
            actionHTML = `
                <div style="margin-top: 10px; border-top: 1px solid #f1f5f9; padding-top: 8px;">
                    <label style="font-size: 12px; cursor: pointer; color: #475569; display: flex; align-items: center; gap: 6px; font-weight: 500;">
                        <input type="checkbox" class="tm-snooze-chk" data-id="${actionData.id}"> Lembrar mais tarde (+4h)
                    </label>
                </div>`;
        }

        alertCard.innerHTML = `
            <button class="tm-alert-close">✕</button>
            <div class="tm-toast-body ${actionData ? 'clickable' : ''}" style="${actionData ? 'cursor:pointer;' : ''}">
                <h4>${title}</h4>
                ${message ? `<p>${message}</p>` : ''}
            </div>
            ${actionHTML}
        `;
        wrapper.appendChild(alertCard);

        const closeBtn = alertCard.querySelector('.tm-alert-close');
        closeBtn.addEventListener('click', () => removeCard(alertCard));

        if (actionData) {
            // Abrir e focar o registro específico na lateral
            alertCard.querySelector('.tm-toast-body').addEventListener('click', () => {
                openRecordInSidebar(actionData.tab, actionData.id);
                removeCard(alertCard);
            });
            // Opção de adiar 4h
            const snoozeChk = alertCard.querySelector('.tm-snooze-chk');
            if(snoozeChk) {
                snoozeChk.addEventListener('change', (e) => {
                    if (e.target.checked) {
                        let snoozed = GM_getValue('tm_snoozed_notifications', {});
                        snoozed[actionData.id] = Date.now() + (4 * 60 * 60 * 1000); // 4 horas
                        GM_setValue('tm_snoozed_notifications', snoozed);
                        removeCard(alertCard);
                        showToast("⏳ Adiado", "Você será lembrado novamente em 4 horas.", "#10b981");
                    }
                });
            }
        }
        setTimeout(() => removeCard(alertCard), 3000);
    }

    function removeCard(card) {
        if (!card.parentNode || card.classList.contains('tm-alert-out')) return;
        card.classList.add('tm-alert-out');
        setTimeout(() => card.remove(), 350); // Aguarda a animação de saída terminar
    }

    function openRecordInSidebar(tab, id) {
        const isOpen = panel.classList.contains('open');
        if (!isOpen) {
            panel.classList.add('open');
            toggleBtn.className = 'open-state';
        }
        switchTab(tab);

        // Espera a animação da sidebar/histórico e foca
        setTimeout(() => {
            if (isExpanded && tab !== 'trade') historyPanel.classList.add('open');

            setTimeout(() => {
                const card = document.querySelector(`.tm-card[data-id="${id}"]`);
                if (card) {
                    card.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    const originalBg = card.style.backgroundColor;
                    card.style.transition = 'all 0.5s ease';
                    card.style.backgroundColor = '#fef08a'; // Destaca amarelo claro
                    card.style.transform = 'scale(1.02)';

                    setTimeout(() => {
                        card.style.backgroundColor = originalBg;
                        card.style.transform = 'scale(1)';
                    }, 2500); // Tira o destaque depois de 2,5s
                }
            }, isExpanded ? HISTORY_ANIM_MS : 100);
        }, isOpen ? 0 : PANEL_ANIM_MS);
    }

    function showConfirmModal(title, message, onConfirm, confirmBtnText = "Confirmar", isDanger = false, onCancel = null) {
        const overlay = document.createElement('div');
        overlay.id = 'tm-confirm-modal';
        overlay.innerHTML = `
            <div class="tm-modal-content">
                <h3>${title}</h3><p>${message}</p>
                <div class="tm-modal-actions">
                    <button class="tm-modal-btn tm-btn-cancel" id="tm-modal-no">Cancelar</button>
                    <button class="tm-modal-btn tm-btn-confirm" id="tm-modal-yes" style="${isDanger ? 'background:#ef4444;' : 'background:#0f172a;'}">${confirmBtnText}</button>
                </div>
            </div>`;
        body.appendChild(overlay);

        document.getElementById('tm-modal-no').addEventListener('click', () => { overlay.remove(); if (onCancel) onCancel(); });
        document.getElementById('tm-modal-yes').addEventListener('click', () => { overlay.remove(); onConfirm(); });
    }

    // --- Configurações (Modo Local/Supabase) ---
    function showConfigModal() {
        const cfg = getConfig();
        const isConnected = cfg.mode === 'supabase' && cfg.supabaseUrl && cfg.supabaseKey;
        const overlay = document.createElement('div');
        overlay.id = 'tm-config-modal';
        overlay.innerHTML = `
            <div class="tm-modal-content">
                <div class="tm-cfg-header"><span>⚙️</span><h3>Configurações</h3></div>
                <div class="tm-cfg-body">
                    <div class="tm-cfg-status ${isConnected ? 'on' : 'off'}">
                        ${isConnected ? '🟢 Conectado ao Supabase (nuvem)' : '⚪ Modo Local (neste navegador)'}
                    </div>
                    <div class="tm-cfg-section">
                        <label class="tm-cfg-label">Modo de Armazenamento</label>
                        <select id="cfg-mode" class="tm-cfg-select" ${isConnected ? 'disabled' : ''}>
                            <option value="local" ${cfg.mode === 'local' ? 'selected' : ''}>Local (neste navegador)</option>
                            <option value="supabase" ${cfg.mode === 'supabase' ? 'selected' : ''}>Supabase (nuvem)</option>
                        </select>
                    </div>
                    <div id="cfg-supabase-fields" style="display:${cfg.mode === 'supabase' ? 'block' : 'none'};">
                        <div class="tm-cfg-section">
                            <label class="tm-cfg-label">URL do Projeto Supabase</label>
                            <input type="password" autocomplete="new-password" id="cfg-url" class="tm-cfg-input" placeholder="https://xxxx.supabase.co" value="${cfg.supabaseUrl || ''}" ${isConnected ? 'disabled' : ''}>
                            <label class="tm-cfg-label">Chave Anon/Public</label>
                            <input type="password" autocomplete="new-password" id="cfg-key" class="tm-cfg-input" placeholder="eyJhbGciOi..." value="${cfg.supabaseKey || ''}" ${isConnected ? 'disabled' : ''}>
                        </div>
                        <hr class="tm-cfg-divider">
                        <div class="tm-cfg-actions-secondary">
                            <button type="button" id="cfg-view-sql" style="background:#64748b; color:#fff;">📋 Ver Script SQL das Tabelas</button>
                            ${isConnected ? '<button type="button" id="cfg-disconnect" style="background:#ef4444; color:#fff;">🔌 Desconectar Banco de Dados</button>' : ''}
                        </div>
                    </div>
                </div>
                <div class="tm-modal-actions">
                    <button class="tm-modal-btn tm-btn-cancel" id="cfg-cancel">Cancelar</button>
                    <button class="tm-modal-btn tm-btn-confirm" id="cfg-save" ${isConnected ? 'disabled' : ''}>Salvar</button>
                </div>
            </div>`;
        body.appendChild(overlay);

        const modeSelect = document.getElementById('cfg-mode');
        const supaFields = document.getElementById('cfg-supabase-fields');
        modeSelect.addEventListener('change', () => {
            supaFields.style.display = modeSelect.value === 'supabase' ? 'block' : 'none';
        });

        document.getElementById('cfg-view-sql').addEventListener('click', showSqlModal);
        document.getElementById('cfg-cancel').addEventListener('click', () => overlay.remove());

        const disconnectBtn = document.getElementById('cfg-disconnect');
        if (disconnectBtn) disconnectBtn.addEventListener('click', () => {
            showConfirmModal("🔌 Desconectar Banco de Dados?", "O modo voltará para Local e as credenciais salvas serão apagadas.", () => {
                const newCfg = { mode: 'local', supabaseUrl: '', supabaseKey: '' };
                saveConfig(newCfg);
                overlay.remove();
                showToast("🔌 Banco Desconectado", "Configuração alterada para modo Local.", "#ef4444");
                setupRealtime(newCfg);
            }, "Desconectar", true);
        });

        if (!isConnected) {
            document.getElementById('cfg-save').addEventListener('click', () => {
                const newCfg = {
                    mode: modeSelect.value,
                    supabaseUrl: document.getElementById('cfg-url').value.trim(),
                    supabaseKey: document.getElementById('cfg-key').value.trim()
                };
                saveConfig(newCfg);
                overlay.remove();
                showToast("✅ Configurações Salvas", "As alterações foram aplicadas.", "#10b981");
                setupRealtime(newCfg);
                if (newCfg.mode === 'supabase' && newCfg.supabaseUrl && newCfg.supabaseKey) supaSyncAll(true, true);
            });
        }
    }

    function showSqlModal() {
        const overlay = document.createElement('div');
        overlay.id = 'tm-sql-modal';
        overlay.innerHTML = `
            <div class="tm-modal-content" style="width: 420px;">
                <h3>📋 Script SQL (Supabase)</h3>
                <p style="margin-bottom: 10px;">Execute este script no SQL Editor do seu projeto Supabase para criar as tabelas necessárias:</p>
                <textarea class="tm-sql-box" readonly></textarea>
                <div class="tm-modal-actions" style="margin-top: 14px;">
                    <button class="tm-modal-btn tm-btn-cancel" id="sql-close">Fechar</button>
                    <button class="tm-modal-btn tm-btn-confirm" id="sql-copy">Copiar</button>
                </div>
            </div>`;
        body.appendChild(overlay);
        overlay.querySelector('.tm-sql-box').value = SQL_SCHEMA;

        document.getElementById('sql-close').addEventListener('click', () => overlay.remove());
        document.getElementById('sql-copy').addEventListener('click', async () => {
            try {
                await navigator.clipboard.writeText(SQL_SCHEMA);
                showToast("📋 Copiado", "Script SQL copiado para a área de transferência.", "#10b981");
            } catch (err) {
                const box = overlay.querySelector('.tm-sql-box');
                box.select();
                document.execCommand('copy');
                showToast("📋 Copiado", "Script SQL copiado para a área de transferência.", "#10b981");
            }
        });
    }

    // --- Importação / Exportação de Planilhas ---
    function getFieldValue(row, ...keys) {
        for (const key of keys) {
            for (const rowKey in row) {
                if (rowKey.trim().toLowerCase() === key) return String(row[rowKey]).trim();
            }
        }
        return '';
    }

    function importRows(rows) {
        if (!rows || rows.length === 0) {
            showToast("⚠️ Planilha Vazia", "Nenhum registro encontrado no arquivo.", "#ef4444");
            return;
        }
        showConfirmModal("📥 Importar Planilha?", `Foram encontrados ${rows.length} registros para importar na aba atual.`, () => {
            const imported = [];
            let ignorados = 0; // Contador para saber quantos foram cortados

            rows.forEach(row => {
                let record;
                // Tenta pegar o ID da planilha. Se não tiver, gera um novo.
                const rowId = getFieldValue(row, 'id');
                const finalId = rowId ? rowId : generateUniqueId();

                if (activeTab === 'encomendas') {
                    // Verifica se o ID já existe no banco de encomendas
                    if (db.encomendas.find(e => String(e.id) === String(finalId))) {
                        ignorados++;
                        return; // Corta (pula) este registro
                    }

                    const loja = getFieldValue(row, 'loja');
                    const produto = getFieldValue(row, 'produto');
                    if (!loja && !produto) return;

                    record = {
                        id: finalId,
                        loja, produto,
                        ean: getFieldValue(row, 'ean'),
                        qtd: getFieldValue(row, 'qtd', 'quantidade'),
                        data: getFieldValue(row, 'data') || todayISO()
                    };
                    db.encomendas.unshift(record);
                } else {
                    // Verifica se o ID já existe no banco de observações
                    if (db.observacoes.find(o => String(o.id) === String(finalId))) {
                        ignorados++;
                        return; // Corta (pula) este registro
                    }

                    const titulo = getFieldValue(row, 'titulo', 'título', 'assunto');
                    if (!titulo) return;

                    record = {
                        id: finalId,
                        titulo,
                        data: getFieldValue(row, 'data') || null,
                        notify_weekday: getFieldValue(row, 'notify_weekday', 'repeticao', 'repetição') || null,
                        desc: getFieldValue(row, 'desc', 'descricao', 'descrição'),
                        notify: false
                    };
                    db.observacoes.unshift(record);
                }
                imported.push(record);
            });

            saveDB();
            render();
            if (imported.length > 0) supaUpsert(activeTab, imported);

            // Mostra o resultado final com a quantidade de ignorados
            if (ignorados > 0) {
                showToast("✅ Importação Concluída", `${imported.length} importados. ${ignorados} ignorados (já existiam).`, "#10b981");
            } else {
                showToast("✅ Importação Concluída", `${imported.length} registros importados com sucesso.`, "#10b981");
            }
        }, "Importar", false);
    }

    function exportCurrentTab() {
        if (activeTab === 'trade') {
            showToast("⚠️ Aba Inválida", "Selecione a aba Encomendas ou Observações antes de exportar.", "#ef4444");
            return;
        }
        const items = db[activeTab];
        if (!items || items.length === 0) {
            showToast("⚠️ Nada para Exportar", "Não há registros nesta aba.", "#ef4444");
            return;
        }

        // Agora o 'ID' é a primeira coluna a ser exportada
        const rows = items.map(item => activeTab === 'encomendas'
            ? { ID: item.id, Loja: item.loja, Produto: item.produto, EAN: item.ean, Qtd: item.qtd, Data: item.data }
            : { ID: item.id, Titulo: item.titulo, Data: item.data, Repeticao: item.notify_weekday, Descricao: item.desc, Notificar: item.notify ? 'Sim' : 'Não' }
        );

        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, activeTab === 'encomendas' ? 'Encomendas' : 'Observacoes');
        XLSX.writeFile(wb, `${activeTab}_${todayISO()}.xlsx`);
        showToast("📤 Exportado", "A planilha foi baixada com sucesso.", "#10b981");
    }

    // --- Notificações Diárias / Semanais e Auto-Desmarque ---
    const WEEKDAYS = ['domingo', 'segunda', 'terca', 'quarta', 'quinta', 'sexta', 'sabado'];
    const WEEKDAY_LABELS = { domingo: 'Domingo', segunda: 'Segunda', terca: 'Terça', quarta: 'Quarta', quinta: 'Quinta', sexta: 'Sexta', sabado: 'Sábado' };
    const todayWeekdayKey = () => WEEKDAYS[new Date().getDay()];

    function checkNotifications() {
        const todayStr = todayISO();
        const todayWd = todayWeekdayKey();
        const now = Date.now();
        let snoozed = GM_getValue('tm_snoozed_notifications', {});
        let hasChanges = false;

        const toNotify = db.observacoes.filter(o => {
            if (!o.notify) return false;

            // Regra: Se a notificação for de uma data passada, desmarca sozinho.
            if (!o.notify_weekday && o.data && o.data < todayStr) {
                o.notify = false;
                hasChanges = true;
                supaUpsert('observacoes', o);
                return false;
            }

            if (snoozed[o.id] && now < snoozed[o.id]) return false; // Verifica se foi adiada
            if (o.notify_weekday) return o.notify_weekday === todayWd;
            return !o.data || o.data === todayStr;
        });

        if (hasChanges) {
            saveDB();
            render(); // Atualiza visual caso um checkbox tenha desmarcado
        }

        toNotify.forEach((item, idx) => {
            setTimeout(() => showToast(`🔔 Lembrete: ${item.titulo}`, item.desc || '', '#f59e0b', { id: item.id, tab: 'observacoes' }), idx * 400);
        });
    }

    // --- Sistema Multifuncional de Análise de Trade (Local) ---
    const STORE_CODE_REGEX = /^([A-ZÀ-Ú]+)\s*0*(\d+)$/;
    function normalizeStoreCode(rawText) {
        if (!rawText) return '';
        const cleaned = String(rawText).trim().toUpperCase().replace(/\s+/g, ' ');
        if (!cleaned) return '';
        const match = cleaned.match(STORE_CODE_REGEX);
        return match ? (match[1] + parseInt(match[2], 10)) : cleaned;
    }

    function clearTradeHighlights() {
        const marked = document.querySelectorAll('tr[data-tm-trade]');
        for (let i = 0; i < marked.length; i++) {
            const row = marked[i];
            row.style.removeProperty('background-color');
            row.style.removeProperty('color');
            row.removeAttribute('data-tm-trade');
        }
    }

    function highlightTradeRow(row, hasEnoughStock) {
        if (hasEnoughStock) {
            row.style.setProperty('background-color', '#d1fae5', 'important');
            row.style.setProperty('color', '#065f46', 'important');
        } else {
            row.style.setProperty('background-color', '#fee2e2', 'important');
            row.style.setProperty('color', '#991b1b', 'important');
        }
        row.setAttribute('data-tm-trade', '1');
    }

    // Descobre o índice da coluna "Estoque" a partir do cabeçalho real da tabela,
    // evitando o antigo bug de pegar o primeiro número qualquer da linha (que podia
    // ser a coluna errada, ex: código de filial numérico).
    function getStockColIndex(table, cache) {
        if (!table) return -1;
        if (cache.has(table)) return cache.get(table);
        let idx = -1;
        const headerRow = (table.tHead && table.tHead.rows[0]) || table.rows[0];
        if (headerRow) {
            const headerCells = headerRow.cells;
            for (let i = 0; i < headerCells.length; i++) {
                if (headerCells[i].textContent.trim().toUpperCase() === 'ESTOQUE') { idx = i; break; }
            }
        }
        cache.set(table, idx);
        return idx;
    }

    function runTradeAnalysis(isSilent = false) {
        const lojasText = document.getElementById('trade-lojas')?.value || '';
        const qtdsText = document.getElementById('trade-qtds')?.value || '';
        const stockColCache = new Map();

        const lojasArr = lojasText.split('\n');
        const qtdsArr = qtdsText.split('\n');

        const tradeMap = new Map();
        for (let i = 0; i < lojasArr.length; i++) {
            const norm = normalizeStoreCode(lojasArr[i]);
            if (norm) {
                // A quantidade buscará o exato índice (linha) da Loja correspondente
                const qtdVal = parseInt((qtdsArr[i] || '').trim(), 10);
                tradeMap.set(norm, isNaN(qtdVal) ? 0 : qtdVal);
            }
        }

        if (tradeMap.size === 0) {
            if (!isSilent) showToast("⚠️ Lista vazia", "Por favor, preencha os campos de lojas e quantidades.", "#ef4444");
            return false;
        }

        clearTradeHighlights();
        let matchedCount = 0;
        const allRows = document.querySelectorAll('tr');

        // Varredura célula-a-célula: não depende de índice fixo de coluna, então continua
        // funcionando mesmo se a tabela tiver colunas/th ocultos deslocando a posição real
        // do cabeçalho em relação às linhas de dados (causa da contagem incorreta).
        for (let i = 0; i < allRows.length; i++) {
            const row = allRows[i];
            const cells = row.cells; // HTMLCollection nativa: mais rápida que querySelectorAll
            if (!cells || cells.length === 0) continue;

            const table = row.closest('table');
            const stockIdx = getStockColIndex(table, stockColCache);
            if (stockIdx === -1) continue; // Sem coluna "Estoque" reconhecida: não é a tabela alvo, pula

            for (let ci = 0; ci < cells.length; ci++) {
                const cellTextRaw = cells[ci].textContent.trim();
                if (!cellTextRaw || cellTextRaw.toUpperCase() === 'FILIAL') continue;

                const cellNorm = normalizeStoreCode(cellTextRaw);
                if (cellNorm && tradeMap.has(cellNorm)) {
                    matchedCount++;
                    const targetQtd = tradeMap.get(cellNorm);
                    const n = cells[stockIdx] ? parseInt(cells[stockIdx].textContent.replace(/[^\d-]/g, ''), 10) : NaN;
                    const stockNum = isNaN(n) ? 0 : n;
                    highlightTradeRow(row, stockNum >= targetQtd);
                    break; // uma correspondência por linha é suficiente
                }
            }
        }

        if (!isSilent) showToast("📈 Análise em Execução", `Monitoramento ativo! ${matchedCount} lojas destacadas.`, "#10b981");
        return true;
    }

    // --- Renderização Dinâmica e Otimizada ---
    const contentArea = document.getElementById('tm-panel-content');
    const historyTitle = document.getElementById('tm-history-title');

    function render() {
        const isEnc = activeTab === 'encomendas';
        const isObs = activeTab === 'observacoes';
        const isTrade = activeTab === 'trade';
        let formHTML = '', listHTML = '';

        if (isEnc) {
            formHTML = buildEncFormHTML(); listHTML = buildEncListHTML(); historyTitle.textContent = 'Histórico de Encomendas';
        } else if (isObs) {
            formHTML = buildObsFormHTML(); listHTML = buildObsListHTML(); historyTitle.textContent = 'Histórico de Observações';
        } else if (isTrade) {
            formHTML = buildTradeFormHTML(); listHTML = ''; historyTitle.textContent = 'Análise de Trade';
        }

        if (isExpanded) {
            contentArea.innerHTML = formHTML; historyListSlot.innerHTML = listHTML;
        } else {
            contentArea.innerHTML = formHTML + listHTML; historyListSlot.innerHTML = '';
        }

        if (!isTrade) setupFilters();

        if (isTrade) {
            const tl = document.getElementById('trade-lojas');
            const tq = document.getElementById('trade-qtds');
            const countLojas = document.getElementById('tm-count-lojas');
            const countQtds = document.getElementById('tm-count-qtds');
            const countNonEmptyLines = (val) => val.split('\n').filter(l => l.trim() !== '').length;
            const updateCounts = () => {
                if (countLojas) countLojas.textContent = `(${countNonEmptyLines(tl.value)})`;
                if (countQtds) countQtds.textContent = `(${countNonEmptyLines(tq.value)})`;
            };
            if (tl && tq) {
                // Sincroniza Scroll do Trade
                tl.addEventListener('scroll', () => { tq.scrollTop = tl.scrollTop; });
                tq.addEventListener('scroll', () => { tl.scrollTop = tq.scrollTop; });
                // Contagem de linhas preenchidas em cada coluna
                tl.addEventListener('input', updateCounts);
                tq.addEventListener('input', updateCounts);
                updateCounts();
            }
        }
    }

    // Delegação Central de Eventos
    body.addEventListener('click', (e) => {
        const btnEdit = e.target.closest('.tm-btn-edit');
        if (btnEdit) { editingId = btnEdit.dataset.id; return render(); }
        if (e.target.id === 'enc-cancel' || e.target.id === 'obs-cancel') { editingId = null; return render(); }

        const btnDel = e.target.closest('.single-delete');
        if (btnDel) {
            const id = btnDel.dataset.id;
            showConfirmModal("🚨 Excluir Registro?", "Esta ação não pode ser desfeita.", () => {
                db[activeTab] = db[activeTab].filter(item => String(item.id) !== String(id));
                selectedIds.delete(String(id)); saveDB(); render();
                supaDelete(activeTab, [id]);
                showToast("🗑️ Apagado", "Registro deletado com sucesso.", "#ef4444");
            }, "Deletar", true);
            return;
        }

        // Exclusão Múltipla
        if (e.target.id === 'tm-delete-selected') {
            const ids = Array.from(selectedIds);
            showConfirmModal("🚨 Excluir Selecionados?", `Apagar ${selectedIds.size} registros definitivamente?`, () => {
                db[activeTab] = db[activeTab].filter(item => !selectedIds.has(String(item.id)));
                selectedIds.clear(); saveDB(); render();
                supaDelete(activeTab, ids);
                showToast("🗑️ Limpeza Concluída", "Os registros foram apagados.", "#ef4444");
            }, "Excluir Todos", true);
            return;
        }

        if (e.target.id === 'enc-submit') {
            const loja = document.getElementById('enc-loja').value.trim();
            const produto = document.getElementById('enc-produto').value.trim();
            const ean = document.getElementById('enc-ean').value.trim();
            const qtd = document.getElementById('enc-qtd').value.trim();
            let record;
            if (editingId) {
                const idx = db.encomendas.findIndex(x => String(x.id) === String(editingId));
                const dataOriginal = idx > -1 ? (db.encomendas[idx].data || todayISO()) : todayISO();
                record = { id: editingId, loja, produto, ean, qtd, data: dataOriginal };
                if (idx > -1) db.encomendas[idx] = record;
                editingId = null;
            } else {
                record = { id: generateUniqueId(), loja, produto, ean, qtd, data: todayISO() };
                db.encomendas.unshift(record);
            }
            saveDB(); render(); supaUpsert('encomendas', record);
        }

        if (e.target.id === 'obs-submit') {
            const titulo = document.getElementById('obs-titulo').value.trim();
            const data = document.getElementById('obs-data').value || null;
            const notify_weekday = document.getElementById('obs-weekday').value || null;
            const desc = document.getElementById('obs-desc').value.trim();
            const notify = document.getElementById('obs-notify-now').checked;
            let record;
            if (editingId) {
                record = { id: editingId, titulo, data, notify_weekday, desc, notify };
                const idx = db.observacoes.findIndex(x => String(x.id) === String(editingId));
                if (idx > -1) db.observacoes[idx] = record;
                editingId = null;
            } else {
                record = { id: generateUniqueId(), titulo, data, notify_weekday, desc, notify };
                db.observacoes.unshift(record);
            }
            saveDB(); render(); supaUpsert('observacoes', record);
        }

        if (e.target.id === 'trade-start') {
            if (tradeInterval) {
                clearInterval(tradeInterval); tradeInterval = null; clearTradeHighlights();
                e.target.textContent = 'Iniciar Análise'; e.target.style.backgroundColor = '#10b981';
            } else {
                if (runTradeAnalysis(false)) {
                    e.target.textContent = '⏸️ Parar Análise'; e.target.style.backgroundColor = '#ef4444';
                    tradeInterval = setInterval(() => { runTradeAnalysis(true); }, 100);
                }
            }
        }
    });

    body.addEventListener('change', (e) => {
        if (e.target.classList.contains('tm-checkbox') && e.target.id !== 'tm-select-all') {
            const id = String(e.target.dataset.id);
            if (e.target.checked) selectedIds.add(id); else selectedIds.delete(id);
            updateDeleteBtnState();
        }
        if (e.target.id === 'tm-select-all') {
            const isChecked = e.target.checked;
            document.querySelectorAll('#tm-items-list .tm-card').forEach(card => {
                if (card.style.display === 'none') return;
                const chk = card.querySelector('.tm-checkbox');
                if (!chk) return;
                chk.checked = isChecked;
                const id = String(chk.dataset.id);
                if (isChecked) selectedIds.add(id); else selectedIds.delete(id);
            });
            updateDeleteBtnState();
        }
        if (e.target.id === 'tm-toggle-undated') { hideUndated = e.target.checked; return render(); }
        if (e.target.classList.contains('tm-notify-check')) {
            const id = String(e.target.dataset.id);
            const item = db.observacoes.find(o => String(o.id) === id);
            if (item) { item.notify = e.target.checked; saveDB(); supaUpsert('observacoes', item); }
        }
    });

    function updateDeleteBtnState() {
        const btn = document.getElementById('tm-delete-selected');
        const chkAll = document.getElementById('tm-select-all');
        const cards = document.querySelectorAll('#tm-items-list .tm-card');
        const totalVisible = Array.from(cards).filter(c => c.style.display !== 'none').length;
        if (btn) btn.style.display = selectedIds.size > 0 ? 'block' : 'none';
        if (chkAll) chkAll.checked = (selectedIds.size === totalVisible && totalVisible > 0);
    }

    function buildFilterBarHTML(searchPlaceholder) {
        return `
            <div class="tm-search-container"><input type="text" id="tm-search-input" class="tm-search-input" placeholder="${searchPlaceholder}" value="${searchQuery}"></div>
            <div class="tm-search-container" style="gap:6px;">
                <input type="date" id="tm-filter-from" class="tm-search-input" value="${dateFrom}" title="De">
                <input type="date" id="tm-filter-to" class="tm-search-input" value="${dateTo}" title="Até">
                <button id="tm-filter-clear" class="tm-btn-small" style="background:#64748b; color:#fff;" title="Resetar filtro de data (5 dias)">↺</button>
            </div>`;
    }

    function buildEncFormHTML() {
        const itemEdit = editingId ? db.encomendas.find(e => String(e.id) === String(editingId)) : null;
        return `
            <div class="tm-form">
                <strong style="color:#3b82f6;">${itemEdit ? 'Editar Encomenda' : 'Nova Encomenda'}</strong>
                <input type="text" id="enc-loja" placeholder="Nome da Loja" value="${itemEdit?.loja || ''}">
                <input type="text" id="enc-produto" placeholder="Produto" value="${itemEdit?.produto || ''}">
                <input type="text" id="enc-ean" placeholder="EAN" value="${itemEdit?.ean || ''}">
                <input type="number" id="enc-qtd" placeholder="Quantidade" value="${itemEdit?.qtd || ''}">
                <button id="enc-submit">${itemEdit ? 'Salvar Alterações' : 'Registrar Encomenda'}</button>
                ${itemEdit ? '<button id="enc-cancel" style="background:#64748b;">Cancelar</button>' : ''}
            </div>` + (isExpanded ? '' : buildFilterBarHTML("Buscar por loja, produto ou EAN..."));
    }

    function buildEncListHTML() {
        let listHTML = (isExpanded ? buildFilterBarHTML("Buscar por loja, produto ou EAN...") : '') +
            `<div class="tm-list-container"><div class="tm-list-header"><label><input type="checkbox" id="tm-select-all"> Marcar Todos</label><button id="tm-delete-selected" class="tm-btn-small tm-btn-del" style="display:none;">Excluir Selecionados</button></div><div class="tm-list" id="tm-items-list">`;
        if (db.encomendas.length === 0) { listHTML += '<p id="tm-empty-msg" style="text-align:center; color:#94a3b8; font-size:13px; margin-top:20px;">Nenhuma encomenda.</p>'; }
        else {
            db.encomendas.forEach(item => {
                const isChecked = selectedIds.has(String(item.id)) ? 'checked' : '';
                let dataFormatada = item.data;
                try { const [a, m, d] = String(item.data || '').split('-'); if (d) dataFormatada = `${d}/${m}/${a}`; } catch(e) {}
                listHTML += `
                    <div class="tm-card" data-id="${item.id}">
                        <input type="checkbox" class="tm-checkbox" data-id="${item.id}" ${isChecked}>
                        <div class="tm-card-content">
                            <h4>${item.loja || 'Sem loja'}</h4>
                            <p><strong>Produto:</strong> ${item.produto || '-'} (x${item.qtd || '0'})</p>
                            ${item.ean ? `<p><strong>EAN:</strong> ${item.ean}</p>` : ''}
                            ${item.data ? `<p><strong>Data:</strong> ${dataFormatada}</p>` : ''}
                            <div class="tm-card-actions"><button class="tm-btn-small tm-btn-edit" data-id="${item.id}">Editar</button><button class="tm-btn-small tm-btn-del single-delete" data-id="${item.id}">Excluir</button></div>
                        </div>
                    </div>`;
            });
        }
        return listHTML + '</div></div>';
    }

    function buildObsFormHTML() {
        const itemEdit = editingId ? db.observacoes.find(e => String(e.id) === String(editingId)) : null;
        return `
            <div class="tm-form">
                <strong style="color:#f59e0b;">${itemEdit ? 'Editar Observação' : 'Nova Observação'}</strong>
                <input type="text" id="obs-titulo" placeholder="Título / Assunto" value="${itemEdit?.titulo || ''}">
                <div id="obs-data-wrap" style="display:${itemEdit?.notify_weekday ? 'none' : 'block'};">
                    <input type="date" id="obs-data" value="${itemEdit?.data || ''}" style="width:100%; box-sizing:border-box;">
                </div>
                <select id="obs-weekday" class="tm-cfg-select" style="margin-bottom:0;">
                    <option value="">Sem repetição semanal</option>
                    ${WEEKDAYS.map(wd => `<option value="${wd}" ${itemEdit?.notify_weekday === wd ? 'selected' : ''}>Repetir toda ${WEEKDAY_LABELS[wd]}</option>`).join('')}
                </select>
                <textarea id="obs-desc" placeholder="Detalhes..." rows="2">${itemEdit?.desc || ''}</textarea>
                <label style="display:flex; gap:8px; font-size:13px; color:#475569; cursor:pointer;"><input type="checkbox" id="obs-notify-now" ${itemEdit?.notify ? 'checked' : ''}>Ativar notificação / Lembrete</label>
                <button id="obs-submit" style="background:#f59e0b; color:#fff;">${itemEdit ? 'Salvar Alterações' : 'Registrar Observação'}</button>
                ${itemEdit ? '<button id="obs-cancel" style="background:#64748b;">Cancelar</button>' : ''}
            </div>` + (isExpanded ? '' : buildFilterBarHTML("Buscar por título ou descrição..."));
    }

    function buildObsListHTML() {
        let listHTML = (isExpanded ? buildFilterBarHTML("Buscar por título ou descrição...") : '') +
            `<div class="tm-list-container"><div class="tm-list-header"><label><input type="checkbox" id="tm-select-all"> Marcar Todos</label><div style="display:flex; gap:12px; align-items:center;"><label style="display:flex; align-items:center; gap:6px; cursor:pointer; font-weight:500; color:#475569;" title="Itens sem data e sem repetição semanal ignoram o filtro de data"><input type="checkbox" id="tm-toggle-undated" ${hideUndated ? 'checked' : ''}> ${hideUndated ? 'Reexibir' : 'Ocultar'}</label><button id="tm-delete-selected" class="tm-btn-small tm-btn-del" style="display:none;">Excluir Selecionados</button></div></div><div class="tm-list" id="tm-items-list">`;
        if (db.observacoes.length === 0) { listHTML += '<p id="tm-empty-msg" style="text-align:center; color:#94a3b8; font-size:13px; margin-top:20px;">Nenhuma observação.</p>'; }
        else {
            db.observacoes.forEach(item => {
                let dataFormatada = item.data; try { const [a, m, d] = String(item.data || '').split('-'); if(d) dataFormatada = `${d}/${m}/${a}`; } catch(e){}
                const isChecked = selectedIds.has(String(item.id)) ? 'checked' : '';
                listHTML += `
                    <div class="tm-card obs-card" data-id="${item.id}">
                        <input type="checkbox" class="tm-notify-check" data-id="${item.id}" title="Marcar para notificar" ${item.notify ? 'checked' : ''}>
                        <input type="checkbox" class="tm-checkbox" data-id="${item.id}" ${isChecked}>
                        <div class="tm-card-content">
                            <h4 style="padding-right:20px;">${item.titulo || 'Sem Assunto'}</h4>
                            ${item.notify_weekday ? `<p><strong>Repetição:</strong> Toda ${WEEKDAY_LABELS[item.notify_weekday]}</p>` : `<p><strong>Data:</strong> ${dataFormatada || 'Sempre ativo'}</p>`}
                            ${item.desc ? `<p><strong>Obs:</strong> ${item.desc}</p>` : ''}
                            <div class="tm-card-actions"><button class="tm-btn-small tm-btn-edit" data-id="${item.id}">Editar</button><button class="tm-btn-small tm-btn-del single-delete" data-id="${item.id}">Excluir</button></div>
                        </div>
                    </div>`;
            });
        }
        return listHTML + '</div></div>';
    }

    function buildTradeFormHTML() {
        const isRunning = !!tradeInterval;
        return `
            <div class="tm-form" style="gap: 8px;">
                <strong style="color:#10b981;">Trade Analytics (Local)</strong>
                <p style="font-size: 12px; color: #64748b; margin: 0 0 4px 0;">As linhas de lojas e quantidades rolam juntas. Pule as linhas (Enter) para bater a QTD exatamente com a linha da sua Loja.</p>
                <div style="display: flex; gap: 10px;">
                    <div style="flex: 1; display: flex; flex-direction: column;">
                        <label style="font-size: 11px; font-weight: bold; margin-bottom: 4px; color: #475569;">Coluna Lojas <span id="tm-count-lojas" style="font-weight:normal; color:#64748b;">(0)</span></label>
                        <textarea id="trade-lojas" class="tm-trade-sync" placeholder="PVH 1\\n\\nJAR 2" rows="12" style="background-color: #ffffff; border: 1px solid #cbd5e1; color: #0f172a; border-radius: 6px; font-size: 12px; font-family: monospace; resize: none; box-sizing: border-box; width: 100%;"></textarea>
                    </div>
                    <div style="flex: 1; display: flex; flex-direction: column;">
                        <label style="font-size: 11px; font-weight: bold; margin-bottom: 4px; color: #475569;">Coluna Qtds <span id="tm-count-qtds" style="font-weight:normal; color:#64748b;">(0)</span></label>
                        <textarea id="trade-qtds" class="tm-trade-sync" placeholder="15\\n\\n5" rows="12" style="background-color: #ffffff; border: 1px solid #cbd5e1; color: #0f172a; border-radius: 6px; font-size: 12px; font-family: monospace; resize: none; box-sizing: border-box; width: 100%;"></textarea>
                    </div>
                </div>
                <button id="trade-start" style="background-color: ${isRunning ? '#ef4444' : '#10b981'}; color: white; margin-top: 6px;">${isRunning ? 'Parar Análise' : 'Iniciar Análise'}</button>
            </div>`;
    }

    function setupFilters() {
        const searchInput = document.getElementById('tm-search-input');
        const fromInput = document.getElementById('tm-filter-from');
        const toInput = document.getElementById('tm-filter-to');
        const clearBtn = document.getElementById('tm-filter-clear');

        if (searchInput) searchInput.addEventListener('input', (e) => { searchQuery = e.target.value.toLowerCase().trim(); applyFilters(); });
        if (fromInput) fromInput.addEventListener('change', (e) => { applyDateFilterChange(e.target.value, dateTo, () => { fromInput.value = dateFrom; }); });
        if (toInput) toInput.addEventListener('change', (e) => { applyDateFilterChange(dateFrom, e.target.value, () => { toInput.value = dateTo; }); });
        if (clearBtn) clearBtn.addEventListener('click', () => {
            dateFrom = isoDateOffset(-5); dateTo = isoDateOffset(5);
            if (fromInput) fromInput.value = dateFrom;
            if (toInput) toInput.value = dateTo;
            applyFilters();
            resyncFilteredRange();
        });

        const weekdaySelect = document.getElementById('obs-weekday');
        const dataWrap = document.getElementById('obs-data-wrap');
        if (weekdaySelect && dataWrap) weekdaySelect.addEventListener('change', (e) => {
            const isWeekly = !!e.target.value;
            dataWrap.style.display = isWeekly ? 'none' : 'block';
            if (isWeekly) document.getElementById('obs-data').value = '';
        });

        applyFilters();
    }

    function applyFilters() {
        let visibleCount = 0;
        document.querySelectorAll('#tm-items-list .tm-card').forEach(card => {
            const id = card.dataset.id;
            const item = db[activeTab].find(x => String(x.id) === String(id));
            let visible = true;

            if (searchQuery) visible = card.innerText.toLowerCase().includes(searchQuery);

            const hasDate = !!(item && item.data);
            if (!hasDate) {
                if (hideUndated) visible = false;
            } else if (visible && item && (dateFrom || dateTo)) {
                const d = item.data || '';
                if (!d) { visible = false; }
                else {
                    if (dateFrom && d < dateFrom) visible = false;
                    if (dateTo && d > dateTo) visible = false;
                }
            }

            card.style.display = visible ? 'flex' : 'none';
            if (visible) visibleCount++;
        });
        const emptyMsg = document.getElementById('tm-empty-msg');
        if (emptyMsg) emptyMsg.style.display = visibleCount === 0 ? 'block' : 'none';
        updateDeleteBtnState();
    }

    // Inicialização
    render();
    (async () => {
        const cfg = getConfig();
        if (cfg.mode === 'supabase' && cfg.supabaseUrl && cfg.supabaseKey) {
            try {
                const candidates = await supaFetchNotifyCandidates(cfg);
                (candidates || []).forEach(it => {
                    const idx = db.observacoes.findIndex(x => String(x.id) === String(it.id));
                    if (idx > -1) db.observacoes[idx] = it; else db.observacoes.unshift(it);
                });
                GM_setValue('tm_store_db', db);
            } catch (err) {}
        }
        checkNotifications();
    })();
    setupRealtime(getConfig());

})();