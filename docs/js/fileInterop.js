window.fileInterop = window.fileInterop || {};

window.fileInterop.downloadFile = function (filename, content) {
    const blob = new Blob([content], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
};

window.fileInterop.clickElement = function (elementId) {
    const el = document.getElementById(elementId);
    if (el) el.click();
};

window.fileInterop.getImageDimensions = function (dataUrl) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve({ width: img.width, height: img.height });
        img.onerror = reject;
        img.src = dataUrl;
    });
};

window.fileInterop.downloadSvg = function (svgElementId, filename, viewBoxStr) {
    const svg = document.getElementById(svgElementId);
    if (!svg) return;
    
    // Clone to avoid modifying the active editor
    const clone = svg.cloneNode(true);
    
    // Remove the grid and background rect for clean export
    const grid = clone.querySelector('pattern#grid');
    if (grid) grid.parentElement.removeChild(grid);
    const bgRect = clone.querySelector('rect[fill*="url(#grid)"]');
    if (bgRect) bgRect.parentElement.removeChild(bgRect);

    // Remove the group transform because the viewBox handles the translation/zoom
    const mainG = clone.querySelector('g[transform*="translate"]');
    if (mainG) mainG.removeAttribute('transform');

    // Apply calculated viewBox if provided
    if (viewBoxStr) {
        clone.setAttribute('viewBox', viewBoxStr);
        clone.removeAttribute('width');
        clone.removeAttribute('height');
    }

    const serializer = new XMLSerializer();
    let source = serializer.serializeToString(clone);
    if (!source.match(/^<svg[^>]+xmlns="http\:\/\/www\.w3\.org\/2000\/svg"/)) {
        source = source.replace(/^<svg/, '<svg xmlns="http://www.w3.org/2000/svg"');
    }
    const preamble = '<?xml version="1.0" standalone="no"?>\r\n';
    const blob = new Blob([preamble + source], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
};

window.fileInterop.downloadBinary = function (filename, base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'model/gltf-binary' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
};

window.fileInterop.renderSvgToPng = function (svgElementId, viewBoxStr, maxDim) {
    const svg = document.getElementById(svgElementId);
    if (!svg) return Promise.resolve("");

    const clone = svg.cloneNode(true);

    if (viewBoxStr) {
        clone.setAttribute('viewBox', viewBoxStr);
        const parts = viewBoxStr.split(/\s+/).map(v => parseFloat(v));
        if (parts.length === 4 && parts.every(v => !isNaN(v))) {
            clone.setAttribute('width', String(parts[2]));
            clone.setAttribute('height', String(parts[3]));
        } else {
            clone.removeAttribute('width');
            clone.removeAttribute('height');
        }
    }

    const serializer = new XMLSerializer();
    let source = serializer.serializeToString(clone);
    if (!source.match(/^<svg[^>]+xmlns="http\:\/\/www\.w3\.org\/2000\/svg"/)) {
        source = source.replace(/^<svg/, '<svg xmlns="http://www.w3.org/2000/svg"');
    }

    const svgBlob = new Blob([source], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(svgBlob);

    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            let width = img.width || 1024;
            let height = img.height || 1024;
            if (viewBoxStr) {
                const parts = viewBoxStr.split(/\s+/).map(v => parseFloat(v));
                if (parts.length === 4 && parts.every(v => !isNaN(v))) {
                    width = parts[2];
                    height = parts[3];
                }
            }

            const targetMax = maxDim || 4096;
            const scale = targetMax / Math.max(width, height);
            const canvas = document.createElement('canvas');
            canvas.width = Math.max(1, Math.round(width * scale));
            canvas.height = Math.max(1, Math.round(height * scale));
            const ctx = canvas.getContext('2d');
            const finish = () => {
                ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
                URL.revokeObjectURL(url);
                const pngDataUrl = canvas.toDataURL('image/png');
                const base64 = pngDataUrl.split(',')[1];
                resolve(base64);
            };

            if (window.createImageBitmap) {
                createImageBitmap(img)
                    .then(bitmap => {
                        ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
                        URL.revokeObjectURL(url);
                        const pngDataUrl = canvas.toDataURL('image/png');
                        const base64 = pngDataUrl.split(',')[1];
                        resolve(base64);
                    })
                    .catch(() => finish());
            } else {
                finish();
            }
        };
        img.onerror = (e) => {
            URL.revokeObjectURL(url);
            reject(e);
        };
        img.src = url;
    });
};

window.fileInterop.getSvgMarkup = function (svgElementId, viewBoxStr) {
    const svg = document.getElementById(svgElementId);
    if (!svg) return "";
    const clone = svg.cloneNode(true);
    if (viewBoxStr) {
        clone.setAttribute('viewBox', viewBoxStr);
        clone.removeAttribute('width');
        clone.removeAttribute('height');
    }
    const serializer = new XMLSerializer();
    let source = serializer.serializeToString(clone);
    if (!source.match(/^<svg[^>]+xmlns="http\:\/\/www\.w3\.org\/2000\/svg"/)) {
        source = source.replace(/^<svg/, '<svg xmlns="http://www.w3.org/2000/svg"');
    }
    return source;
};


window.fileInterop.getBoundingClientRect = function (element) {
    if (!element) return null;
    const rect = element.getBoundingClientRect();
    return {
        width: rect.width,
        height: rect.height,
        top: rect.top,
        left: rect.left,
        bottom: rect.bottom,
        right: rect.right
    };
};

window.fileInterop.scrollToBottom = function (element) {
    if (!element) return;
    
    const doScroll = () => {
        element.scrollTop = element.scrollHeight;
        if (element.lastElementChild) {
            element.lastElementChild.scrollIntoView({ behavior: 'auto', block: 'end' });
        }
    };

    // Attempt immediately
    doScroll();
    
    // And after short delays to catch late layout updates
    setTimeout(doScroll, 50);
    setTimeout(doScroll, 150);
    setTimeout(doScroll, 300);
};

window.fileInterop.scrollToElementId = function (elementId) {
    const element = document.getElementById(elementId);
    if (element) {
        element.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "nearest" });
    } else {
        console.warn("Could not find element to scroll to: " + elementId);
    }
};

window.fileInterop.makeNonPassive = function (element, dotNetHelper) {
    if (!element) return;
    element.addEventListener('wheel', function (e) {
        e.preventDefault();
        dotNetHelper.invokeMethodAsync('OnWheelJs', {
            deltaY: e.deltaY,
            ctrlKey: e.ctrlKey,
            shiftKey: e.shiftKey,
            altKey: e.altKey
        });
    }, { passive: false });
};

window.fileInterop.enableSmartHorizontalScroll = function (element) {
    if (!element) return;
    element.addEventListener('wheel', function (e) {
        const canScrollX = element.scrollWidth > element.clientWidth + 1;
        const canScrollY = element.scrollHeight > element.clientHeight + 1;

        if (canScrollX) {
            const isAtTop = element.scrollTop <= 0;
            const isAtBottom = Math.abs(element.scrollHeight - element.clientHeight - element.scrollTop) < 2;

            if (e.shiftKey || !canScrollY || (isAtTop && e.deltaY < 0) || (isAtBottom && e.deltaY > 0)) {
                if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    let delta = e.deltaY;
                    if (e.deltaMode === 1) delta *= 33;
                    else if (e.deltaMode === 2) delta *= element.clientWidth;
                    
                    element.scrollLeft += delta;
                }
            }
        }
    }, { passive: false });
};

window.fileInterop.enableSignalGroupScrollSync = function (root, container, scrollbar, contentSelector) {
    if (!container || !scrollbar) return;
    const selector = contentSelector || ".signal-groups-list";
    const existing = scrollbar._signalGroupScrollSync;
    if (existing && existing.container === container) {
        existing.refresh();
        return;
    }

    const scrollbarContent = scrollbar.querySelector(".signal-groups-scrollbar-content");
    let syncing = false;

    const getContents = () => Array.from(container.querySelectorAll(selector));

    let pendingFrame = 0;
    const syncSizes = () => {
        const contents = getContents();
        let maxWidth = 0;
        for (const el of contents) {
            maxWidth = Math.max(maxWidth, el.scrollWidth);
        }
        const width = Math.max(maxWidth, scrollbar.clientWidth);
        if (scrollbarContent) {
            scrollbarContent.style.width = width + "px";
        }
        const shouldShow = maxWidth > scrollbar.clientWidth + 1;
        if (scrollbar.classList) {
            scrollbar.classList.toggle("is-hidden", !shouldShow);
        }
        if (pendingFrame) {
            cancelAnimationFrame(pendingFrame);
            pendingFrame = 0;
        }
    };

    const syncFromScrollbar = () => {
        if (syncing) return;
        syncing = true;
        const left = scrollbar.scrollLeft;
        for (const el of getContents()) {
            el.scrollLeft = left;
        }
        syncing = false;
    };

    const syncFromContent = (e) => {
        if (syncing) return;
        syncing = true;
        scrollbar.scrollLeft = e.target.scrollLeft;
        syncing = false;
    };

    const attachContentListeners = () => {
        for (const el of getContents()) {
            el.removeEventListener("scroll", syncFromContent);
            el.addEventListener("scroll", syncFromContent, { passive: true });
        }
    };

    const refresh = () => {
        attachContentListeners();
        syncSizes();
    };

    scrollbar.addEventListener("scroll", syncFromScrollbar, { passive: true });

    const resizeObserver = new ResizeObserver(() => refresh());
    resizeObserver.observe(container);
    resizeObserver.observe(scrollbar);

    refresh();

    scrollbar._signalGroupScrollSync = {
        container,
        refresh,
        cleanup: () => {
            scrollbar.removeEventListener("scroll", syncFromScrollbar);
            resizeObserver.disconnect();
        }
    };
};

window.fileInterop.portalDropdown = function (dropdown, trigger, options) {
    if (!dropdown || !trigger) return;

    const opts = options || {};
    const zIndex = opts.zIndex || 4000;
    const gutter = opts.gutter || 6;

    const getScrollParents = (el) => {
        const parents = [];
        let node = el && el.parentElement;
        while (node) {
            const style = window.getComputedStyle(node);
            const overflow = (style.overflow || "") + (style.overflowX || "") + (style.overflowY || "");
            if (/(auto|scroll)/.test(overflow)) {
                parents.push(node);
            }
            node = node.parentElement;
        }
        return parents;
    };

    const state = dropdown._portalState || {};
    if (!dropdown._portalState) {
        dropdown._portalState = state;
        state.scrollParents = getScrollParents(trigger);
        state.onScroll = () => state.position && state.position();
        state.onResize = () => state.position && state.position();
        for (const p of state.scrollParents) {
            p.addEventListener("scroll", state.onScroll, { passive: true });
        }
        window.addEventListener("scroll", state.onScroll, { passive: true });
        window.addEventListener("resize", state.onResize, { passive: true });

        state.observer = new MutationObserver(() => {
            if (!document.body.contains(dropdown)) {
                cleanup();
            }
        });
        state.observer.observe(document.body, { childList: true, subtree: true });
    }

    const cleanup = () => {
        const st = dropdown._portalState;
        if (!st) return;
        for (const p of st.scrollParents || []) {
            p.removeEventListener("scroll", st.onScroll);
        }
        window.removeEventListener("scroll", st.onScroll);
        window.removeEventListener("resize", st.onResize);
        if (st.observer) st.observer.disconnect();
        dropdown._portalState = null;
    };

    const position = () => {
        const rect = trigger.getBoundingClientRect();
        dropdown.style.position = "fixed";
        dropdown.style.zIndex = String(zIndex);
        dropdown.style.right = "auto";
        dropdown.style.bottom = "auto";
        dropdown.style.maxHeight = "";
        dropdown.style.overflowY = "";

        // Measure dropdown size after it is in the document.
        const dropRect = dropdown.getBoundingClientRect();
        const width = Math.max(dropRect.width || 0, rect.width);
        dropdown.style.minWidth = width ? width + "px" : "";

        const spaceBelow = window.innerHeight - rect.bottom;
        const spaceAbove = rect.top;
        let top;
        let maxHeight;

        if (spaceBelow >= dropRect.height || spaceBelow >= spaceAbove) {
            top = rect.bottom + gutter;
            maxHeight = Math.max(80, spaceBelow - gutter - 8);
        } else {
            top = Math.max(8, rect.top - dropRect.height - gutter);
            maxHeight = Math.max(80, spaceAbove - gutter - 8);
        }

        let left = rect.left;
        const maxLeft = Math.max(8, window.innerWidth - (dropRect.width || width) - 8);
        if (left > maxLeft) left = maxLeft;
        if (left < 8) left = 8;

        dropdown.style.left = left + "px";
        dropdown.style.top = top + "px";
        dropdown.style.maxHeight = maxHeight + "px";
        dropdown.style.overflowY = "auto";
    };

    state.position = position;

    if (dropdown.parentElement !== document.body) {
        document.body.appendChild(dropdown);
    }

    position();
};

window.fileInterop.portalizeOpenDropdowns = function (root, dropdownSelector, triggerSelector) {
    const scope = root || document;
    const dropdowns = Array.from(scope.querySelectorAll(dropdownSelector || ".custom-dropdown-list"));
    for (const dropdown of dropdowns) {
        if (!dropdown._portalOrigin) {
            dropdown._portalOrigin = dropdown.parentElement;
        }
        const origin = dropdown._portalOrigin || dropdown.parentElement;
        let trigger = null;
        if (origin) {
            trigger = origin.querySelector(triggerSelector || ".inline-dropdown-trigger.open");
        }
        if (!trigger) {
            const prev = dropdown.previousElementSibling;
            if (prev && prev.matches && prev.matches(triggerSelector || ".inline-dropdown-trigger")) {
                trigger = prev;
            }
        }
        if (!trigger) continue;
        window.fileInterop.portalDropdown(dropdown, trigger, {});
    }
};

window.fileInterop.enableMiddleMousePan = function (element, dotNetHelper) {
    if (!element || !dotNetHelper) return;
    if (element._middlePan) return;

    const state = { active: false };

    const onDown = (e) => {
        if (e.button !== 1) return;
        e.preventDefault();
        state.active = true;
        dotNetHelper.invokeMethodAsync("BeginMiddlePan", e.clientX, e.clientY);
    };

    const onUp = (e) => {
        if (!state.active) return;
        if (e.button !== 1) return;
        e.preventDefault();
        state.active = false;
        dotNetHelper.invokeMethodAsync("EndMiddlePan");
    };

    element.addEventListener("mousedown", onDown, true);
    window.addEventListener("mouseup", onUp, { passive: false });

    element._middlePan = {
        cleanup: () => {
            element.removeEventListener("mousedown", onDown, true);
            window.removeEventListener("mouseup", onUp);
        }
    };
};

window.fileInterop.enablePanButtons = function (element, dotNetHelper, buttons) {
    if (!element || !dotNetHelper) return;
    if (element._panButtons) return;

    const activeButtons = Array.isArray(buttons) ? buttons : [1, 2];
    const state = { active: false };

    const onDown = (e) => {
        if (!activeButtons.includes(e.button)) return;
        e.preventDefault();
        state.active = true;
        dotNetHelper.invokeMethodAsync("BeginPan", e.clientX, e.clientY);
    };

    const onUp = (e) => {
        if (!state.active) return;
        if (!activeButtons.includes(e.button)) return;
        e.preventDefault();
        state.active = false;
        dotNetHelper.invokeMethodAsync("EndPan");
    };

    element.addEventListener("mousedown", onDown, true);
    window.addEventListener("mouseup", onUp, { passive: false });
    window.addEventListener("contextmenu", onUp, { passive: false });

    element._panButtons = {
        cleanup: () => {
            element.removeEventListener("mousedown", onDown, true);
            window.removeEventListener("mouseup", onUp);
            window.removeEventListener("contextmenu", onUp);
        }
    };
};

window.fileInterop.stopWheelPropagation = function (element) {
    if (!element) return;
    element.addEventListener('wheel', function (e) {
        e.stopPropagation();
    }, { passive: true });
};

window.fileInterop.focusAndSelect = function (elementOrId) {
    let element = elementOrId;
    if (typeof elementOrId === 'string') {
        element = document.getElementById(elementOrId);
    }
    
    if (element && typeof element.focus === 'function') {
        element.focus();
        if (typeof element.select === 'function') {
            element.select();
        }
    } else {
        console.warn("focusAndSelect: Element not found or not focusable", elementOrId);
    }
};

window.fileInterop.clientToSvgPoint = function (svgElement, clientX, clientY) {
    if (!svgElement) {
        return { x: 0, y: 0 };
    }

    const pt = svgElement.createSVGPoint();
    pt.x = clientX;
    pt.y = clientY;
    const ctm = svgElement.getScreenCTM();
    if (!ctm) {
        return { x: 0, y: 0 };
    }

    const local = pt.matrixTransform(ctm.inverse());
    return { x: local.x, y: local.y };
};

window.fileInterop.setPointerCapture = function (element, pointerId) {
    if (!element || typeof element.setPointerCapture !== "function") return;
    try {
        element.setPointerCapture(pointerId);
    } catch (_) {
    }
};

window.fileInterop.releasePointerCapture = function (element, pointerId) {
    if (!element || typeof element.releasePointerCapture !== "function") return;
    try {
        if (typeof element.hasPointerCapture === "function" && element.hasPointerCapture(pointerId)) {
            element.releasePointerCapture(pointerId);
        }
    } catch (_) {
    }
};

window.fileInterop.triangulateSvg = async function (svgText) {
    if (!svgText || typeof svgText !== "string") return null;

    if (!window.fileInterop._svgTriangulator) {
        const [THREE, svgLoaderModule] = await Promise.all([
            import("https://unpkg.com/three@0.160.0/build/three.module.js"),
            import("/js/vendor/three-svg-loader.js")
        ]);
        window.fileInterop._svgTriangulator = {
            THREE,
            SVGLoader: svgLoaderModule.SVGLoader,
            loader: new svgLoaderModule.SVGLoader()
        };
    }

    const tri = window.fileInterop._svgTriangulator;
    const parsed = tri.loader.parse(svgText);

    const isNoneColor = (value) => {
        const v = (value || "").toString().trim().toLowerCase();
        return v === "" || v === "none" || v === "transparent";
    };

    const normalizeColor = (value, fallbackHex) => {
        try {
            const c = new tri.THREE.Color();
            c.setStyle((value || "").toString().trim());
            return "#" + c.getHexString();
        } catch (_) {
            return fallbackHex;
        }
    };

    let vbX = 0, vbY = 0, vbW = 600, vbH = 600;
    try {
        const doc = new DOMParser().parseFromString(svgText, "image/svg+xml");
        const root = doc?.documentElement;
        const vb = root?.getAttribute("viewBox");
        if (vb) {
            const parts = vb.trim().split(/[\s,]+/).map(Number);
            if (parts.length >= 4 && parts.every(n => Number.isFinite(n))) {
                vbX = parts[0]; vbY = parts[1]; vbW = parts[2]; vbH = parts[3];
            }
        } else {
            const w = parseFloat((root?.getAttribute("width") || "600").replace(/[^\d.\-eE]/g, ""));
            const h = parseFloat((root?.getAttribute("height") || "600").replace(/[^\d.\-eE]/g, ""));
            if (Number.isFinite(w) && w > 0) vbW = w;
            if (Number.isFinite(h) && h > 0) vbH = h;
        }
    } catch (_) {
        // keep defaults
    }

    const meshes = [];
    const pushGeometry = (geom, fill) => {
        if (!geom) return;
        const posAttr = geom.getAttribute("position");
        if (!posAttr || posAttr.count < 3) {
            geom.dispose?.();
            return;
        }

        const vertices = [];
        const pos = posAttr.array;
        for (let i = 0; i < pos.length; i += 3) {
            vertices.push(pos[i], pos[i + 1]);
        }

        let indices = [];
        if (geom.index && geom.index.array) {
            indices = Array.from(geom.index.array, v => Number(v));
        } else {
            const vcount = posAttr.count;
            for (let i = 0; i < vcount; i++) indices.push(i);
        }

        meshes.push({ fill, vertices, indices });
        geom.dispose?.();
    };

    for (const path of parsed.paths) {
        const style = path.userData?.style || {};
        const rawFill = (style.fill || "black").toString();
        if (!isNoneColor(rawFill)) {
            let shapes = tri.SVGLoader.createShapes(path);

            // For even-odd fills, prefer ShapePath extraction if it preserves more holes.
            const fillRule = (style.fillRule || "").toString().toLowerCase();
            if (fillRule === "evenodd" && typeof path.toShapes === "function") {
                try {
                    const altShapes = path.toShapes(true, false) || [];
                    const holeCount = (arr) => arr.reduce((sum, s) => sum + ((s.holes && s.holes.length) || 0), 0);
                    if (holeCount(altShapes) > holeCount(shapes)) {
                        shapes = altShapes;
                    }
                } catch (_) {
                    // Keep SVGLoader.createShapes result.
                }
            }

            const fill = normalizeColor(rawFill, "#000000");
            for (const shape of shapes) {
                const geom = new tri.THREE.ShapeGeometry(shape, 24);
                pushGeometry(geom, fill);
            }
        }

        const rawStroke = (style.stroke || "").toString();
        const strokeWidth = Number(style.strokeWidth || 0);
        if (!isNoneColor(rawStroke) && strokeWidth > 0) {
            const stroke = normalizeColor(rawStroke, "#000000");
            for (const subPath of path.subPaths || []) {
                try {
                    const points = subPath.getPoints();
                    const strokeGeom = tri.SVGLoader.pointsToStroke(points, style);
                    pushGeometry(strokeGeom, stroke);
                } catch (_) {
                    // keep best-effort behavior
                }
            }
        }
    }

    return { vbX, vbY, vbW, vbH, meshes };
};

window.fileInterop.beginPointDrag = function (svgElement, dotNetHelper, pointerId, startClientX, startClientY) {
    if (!svgElement || !dotNetHelper) return;

    const state = window.fileInterop._pointDragState || {};

    if (state.cleanup) {
        state.cleanup();
    }

    const toLocal = (ev) => {
        const rect = svgElement.getBoundingClientRect();
        const vb = svgElement.viewBox && svgElement.viewBox.baseVal ? svgElement.viewBox.baseVal : { x: 0, y: 0, width: rect.width, height: rect.height };
        if (!rect || rect.width <= 0 || rect.height <= 0) return null;
        const x = vb.x + (ev.clientX - rect.left) * (vb.width / rect.width);
        const y = vb.y + (ev.clientY - rect.top) * (vb.height / rect.height);
        return { x, y };
    };

    let latest = null;
    let rafScheduled = false;
    const pump = () => {
        rafScheduled = false;
        if (!latest) return;
        const local = latest;
        latest = null;
        dotNetHelper.invokeMethodAsync("OnDragMoveJs", local.x, local.y).catch(() => {});
    };

    const onMove = (ev) => {
        const local = toLocal(ev);
        if (!local) return;
        latest = local;
        if (!rafScheduled) {
            rafScheduled = true;
            requestAnimationFrame(pump);
        }
    };

    const onUp = (ev) => {
        dotNetHelper.invokeMethodAsync("OnDragEndJs")
            .catch(() => {});
        cleanup();
    };

    const cleanup = () => {
        window.removeEventListener("pointermove", onMove, true);
        window.removeEventListener("mousemove", onMove, true);
        window.removeEventListener("pointerup", onUp, true);
        window.removeEventListener("mouseup", onUp, true);
        window.removeEventListener("pointercancel", onUp, true);
        latest = null;
        rafScheduled = false;
        window.fileInterop._pointDragState = null;
    };

    window.addEventListener("pointermove", onMove, true);
    window.addEventListener("mousemove", onMove, true);
    window.addEventListener("pointerup", onUp, true);
    window.addEventListener("mouseup", onUp, true);
    window.addEventListener("pointercancel", onUp, true);

    if (typeof startClientX === "number" && typeof startClientY === "number") {
        const fakeEvent = { clientX: startClientX, clientY: startClientY, pointerId };
        onMove(fakeEvent);
    }

window.fileInterop._pointDragState = { cleanup };
};

window.fileInterop.clearStageDefinitionHover = function (container) {
    if (!container) return;
    const cells = container.querySelectorAll("td.state-cell.hover-target, td.state-cell.hover-conflict, td.state-cell.hover-nonconflict");
    cells.forEach((cell) => {
        cell.classList.remove("hover-target", "hover-conflict", "hover-nonconflict");
    });
};

window.fileInterop.highlightStageDefinitionColumn = function (container, stageNumber, hoveredSg, conflictSgs, nonConflictSgs) {
    if (!container) return;
    window.fileInterop.clearStageDefinitionHover(container);

    const conflictSet = new Set(conflictSgs || []);
    const targetStage = String(stageNumber);
    const targetSg = String(hoveredSg);

    const cells = container.querySelectorAll("td.state-cell[data-stage]");
    cells.forEach((cell) => {
        if (cell.getAttribute("data-stage") !== targetStage) return;
        const sg = cell.getAttribute("data-sg");
        if (sg === targetSg) {
            cell.classList.add("hover-target");
            return;
        }
        const sgNum = parseInt(sg || "", 10);
        if (Number.isFinite(sgNum) && conflictSet.has(sgNum)) {
            cell.classList.add("hover-conflict");
        }
    });
};
