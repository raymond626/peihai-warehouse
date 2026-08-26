window.PEIHAI_SUPABASE_CONFIG = {
    url: "https://gverdzrvvalkmsntqwof.supabase.co",
    anonKey: "sb_publishable_t2dK2_2SjsPqMFScKTewFA_Iw8zLJEq",
    allowSignUp: false
};

document.addEventListener('DOMContentLoaded', function(){
    var pageSize = 1000;

    if(typeof MAT !== 'undefined' && MAT.L && MAT.L.specs){
        MAT.L.specs[10] = '壓紋 Embossed / 荔枝紋 Pebble Grain';
    }

    if(typeof autoClassifyImport === 'function'){
        var originalAutoClassifyImport = autoClassifyImport;
        autoClassifyImport = function(matType){
            var mt = String(matType || '').toLowerCase();
            if(mt.includes('荔枝紋') || mt.includes('壓紋') || mt.includes('emboss') || mt.includes('pebble')){
                return {cat:'L', type:'C', spec:'10'};
            }
            return originalAutoClassifyImport(matType);
        };
    }

    async function fetchAllMaterialsCloud(){
        var rows = [];
        for(var from = 0;; from += pageSize){
            var to = from + pageSize - 1;
            var result = await supabaseClient
                .from('materials')
                .select('code,data,updated_at')
                .order('updated_at', { ascending: false })
                .order('code', { ascending: true })
                .range(from, to);
            if(result.error) return { data: null, error: result.error };
            var page = result.data || [];
            rows.push.apply(rows, page);
            if(page.length < pageSize) break;
        }
        return { data: rows, error: null };
    }

    window.loadCloud = async function(showAlert){
        if(!supabaseClient || !currentUser) return;
        if(cloudLoadTimer){ clearTimeout(cloudLoadTimer); cloudLoadTimer = null; }
        setCloud('雲端同步中', 'cloud-ok');
        var result = await fetchAllMaterialsCloud();
        if(result.error){
            setCloud('雲端讀取失敗', 'cloud-err');
            if(showAlert) alert('雲端讀取失敗');
            return;
        }
        inventory = (result.data || []).map(function(r){ return r.data; }).filter(Boolean);
        inventory.sort(function(a, b){ return (b.timestamp || 0) - (a.timestamp || 0); });
        if(db) replaceInDB(inventory, function(){ renderAll(); }, false);
        else renderAll();
        setCloud('雲端同步完成', 'cloud-ok');
    };
});
