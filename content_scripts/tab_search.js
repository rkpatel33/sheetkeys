/**
 * TabSearch: a keyboard-driven overlay for jumping between sheet tabs by typing.
 *
 * Opens when the user presses the configured shortcut (Shift+P by default).
 * Renders an input box and a filtered list of the current spreadsheet's tabs.
 * Typing filters the list (case-insensitive substring). Up/Down navigate,
 * Enter activates the selected tab, Esc (or click-outside) closes.
 */
class TabSearch {
    // Visual constants kept here to avoid magic numbers scattered through render code.
    static OVERLAY_ID = "sheetkeys-tab-search-overlay";
    static LIST_MAX_HEIGHT_PX = 320;
    static PANEL_WIDTH_PX = 480;

    constructor() {
        this.isVisible = false;
        this.overlay = null;
        this.input = null;
        this.listEl = null;
        /** @type {{name: string, el: Element, isActive: boolean}[]} */
        this.tabs = [];
        /** @type {{name: string, el: Element, isActive: boolean}[]} */
        this.filteredTabs = [];
        this.selectedIndex = 0;
    }

    /** Opens the overlay, collects tabs, and focuses the input. */
    show() {
        if (this.isVisible) return;
        this.tabs = this.collectTabs();
        if (this.tabs.length === 0) {
            console.log("TabSearch: no tabs found");
            return;
        }
        this.buildOverlay();
        document.body.appendChild(this.overlay);
        this.isVisible = true;
        this.applyFilter("");
        // Defer focus until the element is actually in the DOM and painted.
        requestAnimationFrame(() => this.input && this.input.focus());
    }

    /** Closes the overlay and returns focus to the sheet so normal-mode resumes. */
    hide() {
        if (!this.isVisible) return;
        this.isVisible = false;
        if (this.overlay) this.overlay.remove();
        this.overlay = null;
        this.input = null;
        this.listEl = null;
        // Refocus the sheet's cell editor so SheetActions.mode flips back to "normal".
        const editor = document.getElementById("waffle-rich-text-editor");
        if (editor) editor.focus();
    }

    /**
     * Reads tab DOM elements from the sheet and extracts display names.
     * @returns {{name: string, el: Element, isActive: boolean}[]}
     */
    collectTabs() {
        const tabEls = SheetActions.getTabEls();
        const tabs = [];
        for (const el of tabEls) {
            const name = TabSearch.extractTabName(el);
            if (!name) continue;
            tabs.push({
                name,
                el,
                isActive: el.classList.contains("docs-sheet-active-tab"),
            });
        }
        return tabs;
    }

    /**
     * Extracts the display name of a single tab element.
     *
     * Sheets puts the name in a purpose-built `.docs-sheet-tab-name` span
     * alongside several sibling indicators (comment count, lock, form, etc.)
     * — the comment-count `<text>` node in particular contains the literal
     * digit "0" even when hidden, which is why a naive textContent read on
     * the whole tab leaks that "0" as a prefix.
     *
     * Preference order:
     *   1. `.docs-sheet-tab-name` child (current Sheets).
     *   2. `aria-label` on the tab element.
     *   3. textContent of the tab itself (last-resort fallback).
     */
    static extractTabName(tabEl) {
        const nameEl = tabEl.querySelector(".docs-sheet-tab-name");
        if (nameEl) {
            const t = (nameEl.textContent || "").trim();
            if (t) return t;
        }

        const ariaLabel = tabEl.getAttribute("aria-label");
        if (ariaLabel && ariaLabel.trim()) return ariaLabel.trim();

        return (tabEl.textContent || "").trim();
    }

    buildOverlay() {
        this.overlay = document.createElement("div");
        this.overlay.id = TabSearch.OVERLAY_ID;
        this.overlay.style.cssText = `
            position: fixed;
            top: 0; left: 0;
            width: 100vw; height: 100vh;
            background: rgba(0, 0, 0, 0.35);
            z-index: 10001;
            display: flex;
            align-items: flex-start;
            justify-content: center;
            padding-top: 15vh;
            font-family: 'Google Sans', Roboto, Arial, sans-serif;
        `;
        this.overlay.addEventListener("click", (e) => {
            if (e.target === this.overlay) this.hide();
        });

        const panel = document.createElement("div");
        panel.style.cssText = `
            width: ${TabSearch.PANEL_WIDTH_PX}px;
            max-width: 90vw;
            background: #ffffff;
            border-radius: 8px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
            overflow: hidden;
            display: flex;
            flex-direction: column;
        `;

        this.input = document.createElement("input");
        this.input.type = "text";
        this.input.placeholder = "Search sheets…";
        this.input.setAttribute("aria-label", "Search sheets");
        this.input.style.cssText = `
            padding: 14px 16px;
            border: none;
            outline: none;
            font-size: 15px;
            border-bottom: 1px solid #e0e0e0;
            width: 100%;
            box-sizing: border-box;
            color: #202124;
        `;
        this.input.addEventListener("input", () =>
            this.applyFilter(this.input.value)
        );
        this.input.addEventListener("keydown", (e) => this.onInputKeyDown(e));

        this.listEl = document.createElement("div");
        this.listEl.style.cssText = `
            max-height: ${TabSearch.LIST_MAX_HEIGHT_PX}px;
            overflow-y: auto;
        `;

        panel.appendChild(this.input);
        panel.appendChild(this.listEl);
        this.overlay.appendChild(panel);
    }

    /**
     * Handles navigation/activation keys inside the input.
     * Normal character keys fall through so the input updates naturally.
     */
    onInputKeyDown(e) {
        switch (e.key) {
            case "Escape":
                e.preventDefault();
                this.hide();
                return;
            case "Enter":
                e.preventDefault();
                this.activateSelected();
                return;
            case "ArrowDown":
                e.preventDefault();
                this.moveSelection(1);
                return;
            case "ArrowUp":
                e.preventDefault();
                this.moveSelection(-1);
                return;
        }
    }

    /** Filters tabs by case-insensitive substring and re-renders. */
    applyFilter(query) {
        const q = query.trim().toLowerCase();
        this.filteredTabs = q
            ? this.tabs.filter((t) => t.name.toLowerCase().includes(q))
            : this.tabs.slice();
        this.selectedIndex = 0;
        this.renderList();
    }

    renderList() {
        this.listEl.innerHTML = "";
        if (this.filteredTabs.length === 0) {
            const empty = document.createElement("div");
            empty.textContent = "No matching sheets";
            empty.style.cssText =
                "padding: 12px 16px; color: #888; font-size: 14px;";
            this.listEl.appendChild(empty);
            return;
        }
        this.filteredTabs.forEach((tab, i) => {
            const row = document.createElement("div");
            row.dataset.index = String(i);
            const isSelected = i === this.selectedIndex;
            row.style.cssText = `
                padding: 10px 16px;
                font-size: 14px;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 8px;
                ${isSelected ? "background: #e8f0fe; color: #1a73e8;" : "color: #202124;"}
            `;

            const nameSpan = document.createElement("span");
            nameSpan.textContent = tab.name;
            nameSpan.style.cssText =
                "overflow: hidden; text-overflow: ellipsis; white-space: nowrap;";
            row.appendChild(nameSpan);

            if (tab.isActive) {
                const badge = document.createElement("span");
                badge.textContent = "current";
                badge.style.cssText = `
                    font-size: 11px;
                    color: #5f6368;
                    background: #f1f3f4;
                    padding: 2px 6px;
                    border-radius: 10px;
                    flex-shrink: 0;
                `;
                row.appendChild(badge);
            }

            row.addEventListener("mouseenter", () => {
                if (this.selectedIndex !== i) {
                    this.selectedIndex = i;
                    this.renderList();
                }
            });
            row.addEventListener("click", () => this.activateSelected());
            this.listEl.appendChild(row);
        });
    }

    /** Moves the selection with wrap-around and scrolls it into view. */
    moveSelection(delta) {
        const count = this.filteredTabs.length;
        if (count === 0) return;
        this.selectedIndex = (this.selectedIndex + delta + count) % count;
        this.renderList();
        const selected = this.listEl.children[this.selectedIndex];
        if (selected && selected.scrollIntoView) {
            selected.scrollIntoView({ block: "nearest" });
        }
    }

    activateSelected() {
        if (this.filteredTabs.length === 0) return;
        const tab = this.filteredTabs[this.selectedIndex];
        this.hide();
        KeyboardUtils.simulateClick(tab.el);
    }
}

window.TabSearch = TabSearch;
