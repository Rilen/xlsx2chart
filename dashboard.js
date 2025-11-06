// dashboard.js (refatorado – foco em correção de gráficos e leitura de múltiplas sheets)

(() => {
    let chart = null, data = null;
    const palette = ['#0d6efd','#dc3545','#198754','#ffc107','#6c757d','#6f42c1','#20c997','#0dcaf0','#641a96','#00bcd4','#ff9800','#8bc34a'];
    const monthOrder = {JANEIRO:'01',FEVEREIRO:'02',MARÇO:'03',ABRIL:'04',MAIO:'05',JUNHO:'06',JULHO:'07',AGOSTO:'08',SETEMBRO:'09',OUTUBRO:'10',NOVEMBRO:'11',DEZEMBRO:'12'};

    const monthKey = f => (f.match(/(JANEIRO|FEVEREIRO|MARÇO|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)\s*-\s*(\d{4})/i)||[])[0]?.toUpperCase()||null;
    const sortKey = m => m.split('/').reverse().join('-') + '-' + (monthOrder[m.split('/')[0]]||'00');

    const normalize = (rows, key) => {
        if (!rows?.length) return [];
        const [cat, sub] = [rows[0], rows[1]];
        const map = [], cur = {cat:''};
        for (let i=1;i<sub.length;i++){
            if (cat[i] && cat[i].trim()) cur.cat = cat[i].trim();
            if (sub[i] && sub[i].trim() && sub[i].trim()!=='Nº SEMANA')
                map.push({idx:i, name: cur.cat ? `${cur.cat} - ${sub[i].trim()}` : sub[i].trim()});
        }
        const out = [];
        rows.slice(2).forEach(r => {
            if (!r[0] || /total/i.test(r[0])) return;
            map.forEach(c => { const q = +r[c.idx]; if (q>0) out.push({m:key, s:c.name, q}); });
        });
        return out;
    };

    const aggregate = arr => {
        const byMonth = {}, total = {}, months = new Set(), subs = new Set(), monthTot = {};
        arr.forEach(({m,s,q}) => {
            months.add(m); subs.add(s);
            (byMonth[m]??={})[s] = (byMonth[m][s]||0) + q;
            total[s] = (total[s]||0) + q;
            monthTot[m] = (monthTot[m]||0) + q;
        });
        const mList = [...months].sort((a,b)=>sortKey(a).localeCompare(sortKey(b)));
        const sList = [...subs].sort();
        const colors = sList.reduce((o,s,i)=>(o[s]=palette[i%palette.length],o),{});
        const mTotals = mList.map(m=>({m, total:monthTot[m].toLocaleString('pt-BR')}));
        return {months:mList, subs:sList, byMonth, total, mTotals, colors};
    };

    const render = (d, type) => {
        if (chart) chart.destroy();
        const canvas = document.getElementById('myChart');
        const msg = document.getElementById('initial-message');
        const area = document.getElementById('chart-area');
        canvas.style.display='block'; msg.style.display='none';
        area.style.height = (type==='pie'||type==='doughnut') ? '0' : '400px';
        area.style.paddingBottom = (type==='pie'||type==='doughnut') ? '100%' : '0';

        if (type==='pie'||type==='doughnut'){
            chart = new Chart(canvas, {type, data:{labels:Object.keys(d.total), datasets:[{data:Object.values(d.total), backgroundColor:Object.values(d.colors)}]}, options:{responsive:true, maintainAspectRatio:true}});
        } else {
            const ds = d.subs.map(s=>({
                label:s, data:d.months.map(m=>d.byMonth[m][s]||0),
                                      backgroundColor: type==='bar'?d.colors[s]:null, borderColor:d.colors[s],
                                      stack: type==='bar'?'s':undefined, fill: type==='line'?false:true
            }));
            chart = new Chart(canvas, {type:type==='bar'?'bar':'line', data:{labels:d.months, datasets:ds}, options:{responsive:true, maintainAspectRatio:false, scales:{x:{stacked:type==='bar'}, y:{stacked:type==='bar', beginAtZero:true}}}});
        }
    };

    const renderTable = d => {
        const tbody = document.querySelector('#data-table tbody');
        tbody.innerHTML = d.mTotals.length ? d.mTotals.map((r,i)=>`<tr><td>${i+1}</td><td>${r.m}</td><td>${r.total}</td></tr>`).join('') : '<tr><td colspan="3">Sem dados</td></tr>';
    };

    window.onload = () => {
        feather.replace();
        const input = document.getElementById('excel-upload');
        const sel = document.getElementById('chart-type-select');
        const status = document.getElementById('upload-status');
        const err = document.getElementById('error-message');
        let all = [];

        const proc = f => new Promise(res=>{
            const r = new FileReader();
            r.onload = e=>{
                try{
                    const wb = XLSX.read(e.target.result, {type:'array'});
                    const norm = wb.SheetNames.flatMap(sn=>{
                        const key = monthKey(sn) || monthKey(f.name);
                        if (!key) throw new Error('Mês/Ano não encontrado');
                        return normalize(XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1}), key);
                    });
                    res(norm);
                }catch(ex){ err.textContent=`Erro em ${f.name}: ${ex.message}`; res([]); }
            };
            r.readAsArrayBuffer(f);
        });

        input.onchange = async ev=>{
            const files = [...ev.target.files];
            status.textContent=`Processando ${files.length}…`; err.textContent=''; all=[];
            const res = await Promise.all(files.map(proc));
            all = res.flat().filter(x=>x.q>0);
            if (all.length){ data=aggregate(all); render(data, sel.value); renderTable(data); status.textContent='Concluído'; }
            else { status.textContent='Sem dados'; }
            input.value='';
        };

        sel.onchange = () => data && render(data, sel.value);
        window.onresize = () => chart && chart.resize();
    };
})();
