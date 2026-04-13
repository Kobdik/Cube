import { actionCreators, TYPES } from './sutils.min.js';
import kobdikFactory from './kobdik1.2.min.js'

const { 
    askBinaryOrdsAsync, askBinaryDataAsync, 
    askConfOrdsAsync, askConfDataAsync, 
    askJsonOrdsAsync, askJsonDataAsync,
} = actionCreators;

const PAD_ROOT = 25, VER_HEIGHT = 25, VER_WIDTH = 210, HOR_HEIGHT = 175, HOR_WIDTH = 100;

window.addEventListener('load', function() {
    // const root = document.querySelector(":root");
    // root.style.setProperty("--HOR_WIDTH", HOR_WIDTH);
    const engine = kobdikFactory();
    const svgProp = { 
        height: 1000, width: 1000, posX: 0, posY: 0, overX: 0, overY: 0,
        touchMove: false, touchStartX: 0, touchStartY: 0, touchMoveX: 0, touchMoveY: 0 
    };
    window.cobdy = engine;
    const svg = d3.select('svg.multi');
    // const corner = svg.select('g.corner');
    const horOrds = svg.select('g.hor-ords');
    const verOrds = svg.select('g.ver-ords');
    const matArea = svg.select('g.matrix');
    const horClip = svg.select('#hor-clip rect');
    const verClip = svg.select('#ver-clip rect');
    const matClip = svg.select('#mat-clip rect');

    horOrds.attr('transform', `translate(${VER_WIDTH}, 0)`);
    verOrds.attr('transform', `translate(0, ${HOR_HEIGHT})`);
    matArea.attr('transform', `translate(${VER_WIDTH}, ${HOR_HEIGHT})`);

    window.addEventListener("resize", resetView);
    
    svg.on('touchstart', function (event) {
        // console.log(event);
        svgProp.touchStartX = event.touches[0].clientX;
        svgProp.touchStartY = event.touches[0].clientY;
    });

    svg.on('touchmove', function (event) {
        // console.log('touchmove 2', svgProp.touchMoveY, event.touches[0].clientY);
        svgProp.touchMoveX = event.touches[0].clientX;
        svgProp.touchMoveY = event.touches[0].clientY;
        svgProp.touchMove = true;
    });

    svg.on('touchend', function (event) {
        const deltaX = 2*(svgProp.touchStartX - svgProp.touchMoveX);
        const deltaY = 2*(svgProp.touchStartY - svgProp.touchMoveY);
        if (svgProp.touchMove) {
            event.preventDefault();
            touchMove(deltaX, deltaY);
            svgProp.touchMove = false;
        }
    });

    svg.on('wheel', function (event) {
        event.preventDefault();
        // console.log('wheel', event);
        if (event.shiftKey || event.deltaX != 0) {
            horMove((event.deltaX != 0) ? event.deltaX : event.deltaY);
        } else verMove(event.deltaY);
    });

    function touchMove(deltaX, deltaY) {
        const valY = svgProp.posY + PAD_ROOT * Math.round(deltaY / PAD_ROOT);
        const posY = (valY < 0) ? 
            0 : (valY < svgProp.height - 5 * VER_HEIGHT - HOR_HEIGHT) ?
                valY : svgProp.height - 5 * VER_HEIGHT - HOR_HEIGHT;
        svgProp.posY = posY;
        const valX = svgProp.posX + PAD_ROOT * Math.round(deltaX / PAD_ROOT);
        const posX = (valX < 0) ? 
            0 : (valX < svgProp.width - 5 * HOR_WIDTH - VER_WIDTH) ?
                valX : svgProp.width - 5 * HOR_WIDTH - VER_WIDTH;
        svgProp.posX = posX;
        verClip.attr('y', posY);
        verOrds.attr('transform', `translate(0, ${HOR_HEIGHT - posY})`);
        horClip.attr('x', posX);
        horOrds.attr('transform', `translate(${VER_WIDTH - posX}, 0)`);
        matClip.attr('x', posX);
        matClip.attr('y', posY);
        matArea.attr('transform', `translate(${VER_WIDTH - posX}, ${HOR_HEIGHT - posY})`);
    }

    function verMove(delta) {
        const val = svgProp.posY + PAD_ROOT * Math.round(delta / PAD_ROOT);
        const pos = (val < 0) ? 
            0 : val < svgProp.height - 5 * VER_HEIGHT - HOR_HEIGHT ?
                val : svgProp.height - 5 * VER_HEIGHT - HOR_HEIGHT;
        svgProp.posY = pos;
        verClip.attr('y', pos);
        verOrds.attr('transform', `translate(0, ${HOR_HEIGHT - pos})`);
        matClip.attr('y', pos);
        matArea.attr('transform', `translate(${VER_WIDTH - svgProp.posX}, ${HOR_HEIGHT - pos})`);
    }
    
    function horMove(delta) {
        const val = svgProp.posX + PAD_ROOT * Math.round(delta / PAD_ROOT);
        const pos = (val < 0) ? 
            0 : val < svgProp.width - 5 * HOR_WIDTH - VER_WIDTH ?
                val : svgProp.width - 5 * HOR_WIDTH - VER_WIDTH;
        svgProp.posX = pos;
        horClip.attr('x', pos);
        horOrds.attr('transform', `translate(${VER_WIDTH - pos}, 0)`);
        matClip.attr('x', pos);
        matArea.attr('transform', `translate(${VER_WIDTH - pos}, ${HOR_HEIGHT - svgProp.posY})`);
    }

    function dispatch(action) {
        switch (action.type) {
            case TYPES.ON_ERROR: console.error(action); break;
            case TYPES.INIT_ORDS: setTimeout(initOrds, 0, action.ords); break;
            case TYPES.INIT_DATA: setTimeout(initData, 0, action.rows, action.add); break;
            default: console.log('action.type', action.type);
        }
    }

    function resetView() {
        const clientWidth = document.body.clientWidth;
        // const clientHeight = document.body.clientHeight;
        const scale = (clientWidth <= 720) ?
            0.4 : (clientWidth <= 1080) ? 
            0.6 : 0.8;
        const hor_pad = engine.calcHorPad(false);
        const ver_pad = engine.calcVerPad(false);
        const width  = hor_pad + VER_WIDTH;
        const height = ver_pad + HOR_HEIGHT;
        const scaled_width = Math.round(width * scale);
        const scaled_height = Math.round(height * scale);
        horClip.attr('width', Math.round(width / scale)).attr('height', HOR_HEIGHT);
        verClip.attr('width', Math.round(VER_WIDTH)).attr('height', height);
        matClip.attr('width', Math.round(width / scale)).attr('height', height);
        svg.attr('viewBox', `0, 0, ${width}, ${height}`);
        svg.attr('width', scaled_width).attr('height', scaled_height);
        // console.log(clientWidth, clientHeight, width, scaled_width);
        // corner.select('text').text(`${clientWidth} ${clientHeight}`);
        svgProp.height = height;
        svgProp.width = width;
    }

    function initOrds(ords) {
        engine.loadOrds(ords);
        engine.setLays({ PAD_ROOT, VER_HEIGHT, VER_WIDTH, HOR_HEIGHT, HOR_WIDTH });
        engine.setDimHor(0, true);
        engine.setDimHor(3, true);
        // const t0 = engine.getDim(0).nodes.get('2024r');
        // engine.selNodes(t0, true);
        setTimeout(askJsonDataAsync('contest-rows.json', false), 0, dispatch);
        // setTimeout(askJsonDataAsync('exam-adds.json', true), 2000, dispatch);
        drawOrds(true);
    }

    function drawOrds(hide=true) {
        engine.createMatrix();
        resetView();
        drawVerOrds(hide);
        drawHorOrds(hide);
    }

    askJsonOrdsAsync('contest-ords.json')(dispatch);

    function cmdClick(event, top) {
        // console.log(top, event)
        if (top.act === 1) {
            setTimeout(selClick, 0, top);
            return;
        }
        if (top.act === 3) {
            setTimeout(subClick, 0, top, event.target.closest('g'));
            return;
        } 
    }

    function subClick(top, elem) {
        const dim = engine.getDim(top.ord);
        if (!dim) return;
        if (dim.subn && (dim.subn != top)) {
            console.log('active sub node differ')
            return;
        };
        const sub = (dim.mode != 'sub'); 
        dim.mode = (sub) ? 'sub' : 'sum';
        dim.subn = (sub) ? top : null;
        const sa_nodes = engine.selNodes(top.sa, sub);
        const sb_nodes = engine.selNodes(top.sb, sub);
        engine.posNodes(top.sa, sub);
        engine.negNodes(top.sb, sub);
        const grids = engine.getGrids();
        for (const grid of grids) {
            engine.zeroGrid(grid);
            engine.flowGrid(grid);
            engine.fillGridSub(grid);
        }
        d3.select(elem).classed('sub', sub);
        const sel_group = (dim.hor) ? horOrds.selectAll(`g.hor-node[ord="${top.ord}"]`) : verOrds.selectAll(`g.ver-node[ord="${top.ord}"]`);
        sel_group.data(sa_nodes, d => d.tid).classed('sel', d => d.sel).classed('pos', d => d.pos);
        sel_group.data(sb_nodes, d => d.tid).classed('sel', d => d.sel).classed('neg', d => d.neg);
        // sel_group.data(sa_nodes.concat(sb_nodes), d => d.tid).classed('sel', d => d.sel);
        setTimeout(drawMatrix, 90, true);
    }

    function selClick(top) {
        const dim = engine.getDim(top.ord);
        if (!dim) return;
        if (dim.mode === 'sub' && !top.parent.sel) {
            console.log('parent not selected');
            return;
        } 
        const sel_nodes = engine.selNodes(top, !top.sel);
        const grids = engine.getFiltGrids(dim.ord);
        // console.log(dim.ord, grids);
        for (const grid of grids) {
            // console.time(`grid${grid.hOrd.ord}-${grid.vOrd.ord}`)
            const start = performance.now();
            engine.zeroGrid(grid);
            engine.flowGrid(grid);
            if (dim.mode != 'sub') engine.fillGrid(grid);
            else engine.fillGridSub(grid);
            performance.measure(
                `grid_${grid.hOrd.ord}-${grid.vOrd.ord}`, 
                { start, end: performance.now() }
            );
            // console.timeEnd(`grid${grid.hOrd.ord}-${grid.vOrd.ord}`)
        }
        // const start = performance.now();
        // performance.measure(`draw`, { start, end: performance.now() });
        for (const entry of performance.getEntries()) {
            if (entry.entryType == 'measure') {
                console.log(`${entry.name}. Duration ${entry.duration}ms`);
            }
        }
        performance.clearMeasures();
        const sel_group = (dim.hor) ? horOrds.selectAll(`g.hor-node[ord="${top.ord}"]`) : verOrds.selectAll(`g.ver-node[ord="${top.ord}"]`);
        sel_group.data(sel_nodes, d => d.tid).classed('sel', d => d.sel);
        setTimeout(drawMatrix, 0, true);
    }

    function txtClick(event, top) {
        // console.log(top, event)
        if (!top.fst_child) {
            console.log('txt', top);
            return;
        }
        top.exp = !top.exp;
        const dim = engine.getDim(top.ord);
        if (!dim) return;
        if (dim.hor) {
            drawHorOrds(true);
            setTimeout(drawMatrix, 0, true);
        } else {
            drawVerOrds(true);
            setTimeout(drawMatrix, 0, true);
        }
    }

    function exitData(exit, hide) {
        if (hide) exit.classed('hide', true);
        else exit.remove()
    }

    function insTreeVer(enter) {
        const g = enter.append('svg:g')
            .attr('id', d => d.tid)
            .attr('ord', d => d.ord)
            .classed('ver-node', true);
        // append content
        g.append('svg:rect')
            .classed('cmd-ver', true)
            .attr('x', 0)
            .attr('y', 0)
            .on('click', cmdClick)
            .append('svg:title').text(d => d.title);
        g.append('svg:text')
            .attr('transform', (d) => `translate(${d.lev*15}, ${VER_HEIGHT/2+5})`)
            .text(d => d.caption)
            .on('click', txtClick);
        g.filter(d => d.lev > 1).append('svg:path').classed('curve', true);
        g.call(updTreeVer);
    }

    function updTreeVer(update) {
        update
            .classed('hide', false)
            .classed('sel', d => d.sel)
            .classed('exp', d => d.fst_child != null)
            .classed('cls', d => (d.fst_child) ? !d.exp : false)
            .attr('transform', d => `translate(0, ${d.pad})`)
            .select('path.curve')
            .attr('d', d => `M${d.lev*15-5},12L${(d.lev-1)*15-5},12,${(d.lev-1)*15-5},${d.parent.pad-d.pad+12}`);
    }

    function drawVerOrds(hide=false) {
        engine.setVerRows(hide);
        const rows = engine.getVerRows(0);
        verOrds.selectAll('g.ver-node').data(rows, d => d.tid).join(
            enter => insTreeVer(enter),
            update => updTreeVer(update),
            exit => exitData(exit, hide)
        );
    }

    function insTreeHor(enter) {
        const g = enter.append('svg:g')
            .attr('id', d => d.tid)
            .attr('ord', d => d.ord)
            .classed('hor-node', true);
        // append content
        g.append('svg:rect')
            .classed('cmd-hor', true)
            .attr('x', 0)
            .attr('y', 0)
            .on('click', cmdClick)
            .append('svg:title').text(d => d.title);
        g.append('svg:text')
            .attr('transform', (d) => `translate(${HOR_WIDTH/2-20}, ${HOR_HEIGHT-d.lev*15+10}) rotate (-60)`)
            .text(d => d.caption)
            .on('click', txtClick);
        g.filter(d => d.lev > 1).append('svg:path').classed('curve', true);
        g.call(updTreeHor);
    }

    function updTreeHor(update) {
        update
            .classed('hide', false)
            .classed('sel', d => d.sel)
            .classed('exp', d => d.fst_child != null)
            .classed('cls', d => (d.fst_child) ? !d.exp : false)
            .attr('transform', d => `translate(${d.pad}, 0)`)
            .select('path.curve')
            .attr('d', d => `M${HOR_WIDTH/2+10},${HOR_HEIGHT-d.lev*15+5}L${HOR_WIDTH/2+10},${HOR_HEIGHT-(d.lev-1)*15+5}
                ${d.parent.pad-d.pad+HOR_WIDTH/2+10},${HOR_HEIGHT-(d.lev-1)*15+5}`);
    }

    function drawHorOrds(hide=false) {
        engine.setHorRows(hide);
        const rows = engine.getHorRows(0);
        horOrds.selectAll('g.hor-node').data(rows, d => d.tid).join(
            enter => insTreeHor(enter),
            update => updTreeHor(update),
            exit => exitData(exit, hide)
        );
    }

    function initData(rows, add=false) {
        // window.rows = JSON.stringify(rows);
        engine.initFlow(rows, add);
        // initial selection
        if (!add) {
            const click = new MouseEvent('click');
            // verOrds.select('#cem').select('rect').node().dispatchEvent(click);
            // verOrds.select('#vnd1').select('rect').node().dispatchEvent(click);
            // verOrds.select('#vnd2').select('rect').node().dispatchEvent(click);
            // verOrds.select('#st1').select('rect').node().dispatchEvent(click);
            // verOrds.select('#st2').select('rect').node().dispatchEvent(click);
            // horOrds.select('#c2025-09').select('rect').node().dispatchEvent(click);
            // horOrds.select('#c2025-09').select('text').node().dispatchEvent(click);
            // horOrds.select('#p2025').select('rect').node().dispatchEvent(click);
        }
        setTimeout(drawMatrix, 90, true);
        console.log('rows loaded:', rows.length);
    }

    function insTape(enter, hide, cols) {
        enter.append('svg:g')
            .classed('tape', true)
            .attr('key', d => d.tid)
            .attr('row', d => d.row)
            .attr('ord', d => d.ord)
            .call(updTape, hide, cols);
    }

    function insCell(enter) {
        const g = enter.append('svg:g')
            .classed('cell', true)
            .attr('key', d => d.tid)
            .attr('row', d => d.row)
            .attr('col', d => d.col)
            .attr('ord', d => d.ord);
        g.append('svg:rect')
            .attr('x', 0)
            .attr('y', 0);
        g.append('svg:text')
            .attr('transform', `translate(${HOR_WIDTH-5}, ${VER_HEIGHT/2+5})`);
        g.call(updCell);
    }

    function updCell(update) {
        update
            .classed('hide', false)
            .classed('sel', c => c.sel)
            .attr('transform', c => `translate(${c.pad}, 0)`)
            .select('text').text(c => engine.getFixVal(c.row, c.col, 0));
            // .select('text').text(c => `${c.row}-${c.col}`);
    }

    function updTape(update, hide, cols) {
        update
            .classed('hide', false)
            .classed('sel', d => d.sel)
            .attr('transform', d => `translate(0, ${d.pad})`)
            .selectAll('g.cell').data(d => cols.map(c => ({ ...c, row: d.row })), c => c.tid)
            .join(
                enter => insCell(enter),
                update => updCell(update),
                exit => exitData(exit, hide)
            );
    }

    function drawMatrix(hide=false) {
        // console.time('draw_matrix');
        const rows = engine.getVerRows(0);
        const cols = engine.getHorRows(0);
        matArea.selectAll('g.tape').data(rows, r => r.tid).join(
            enter => insTape(enter, hide, cols),
            update => updTape(update, hide, cols),
            exit => exitData(exit, hide)
        );
        // console.timeEnd('draw_matrix');
    }

});
