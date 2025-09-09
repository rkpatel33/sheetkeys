// Utilities
window.invertObjectMap = (o) => {
    const o2 = {};
    for (let k of Object.keys(o)) {
        const v = o[k];
        o2[v] = k;
    }
    return o2;
};

// Add an event listener which removes itself once the event is fired once.
const addOneTimeListener = function (dispatcher, eventType, listenerFn) {
    const handlerFn = function (e) {
        dispatcher.removeEventListener(eventType, handlerFn, true);
        return listenerFn(e);
    };
    return dispatcher.addEventListener(eventType, handlerFn, true);
};

const UI = {
    // An arbitrary limit that should instead be equal to the longest key sequence that's actually bound.
    maxKeyMappingLength: 6,
    // Keys which were typed recently
    keyQueue: [],
    // A map of mode -> comma-separated keys -> bool. The keys are prefixes to the user's bound key mappings.
    keyMappingsPrefixes: null,
    richTextEditorId: "waffle-rich-text-editor",
    modeToKeyToCommand: null,

    init() {
        this.injectPageScript();
        SheetActions.typeKeyFn = this.typeKey;
        window.addEventListener("focus", (e) => this.onFocus(e), true);
        // When we first focus the spreadsheet, if we're in fullscreen mode, dismiss Sheet's "info" message.
        addOneTimeListener(window, "focus", () => {
            // We have to wait 1 second because the DOM is not yet ready to be clicked on.
            return setTimeout(
                () => SheetActions.dismissFullScreenNotificationMessage(),
                1000
            );
        });

        // Auto-zoom to 90% when opening in a new tab
        this.setupAutoZoom();

        // Key event handlers fire on window before they do on document. Prefer window for key events so the page
        // can't set handlers to grab keys before this extension does.
        window.addEventListener("keydown", (e) => this.onKeydown(e), true);

        this.loadKeyMappings();

        // If a key mapping setting is changed from another tab, update this tab's key mappings.
        chrome.runtime.onMessage.addListener((message) => {
            if (message == "keyMappingChange") this.loadKeyMappings();
        });
    },

    /**
     * Sets up auto-zoom functionality for when a Google Sheet opens.
     * Waits for the sheet to fully load before applying zoom.
     */
    setupAutoZoom() {
        // Always wait for the sheet to load and apply zoom
        // This works for new tabs, reloads, and navigation
        this.waitForSheetLoad().then(() => {
            // Apply 90% zoom after a short delay to ensure UI is ready
            setTimeout(() => {
                console.log("SheetKeys: Auto-zooming to 90%");
                SheetActions.setZoom90();
                this.showToast("Zoomed to 90%", 2000);
            }, 500);
        });
    },

    /**
     * Waits for the Google Sheet to be fully loaded and ready.
     * Returns a promise that resolves when the sheet is ready.
     */
    waitForSheetLoad() {
        return new Promise((resolve) => {
            const checkReady = () => {
                // Check for key Google Sheets elements that indicate the sheet is loaded
                const gridContainer = document.getElementById("waffle-grid-container");
                const menuBar = document.querySelector('[role="menubar"]');
                const cellInput = document.getElementById("t-formula-bar-input-container");
                
                if (gridContainer && menuBar && cellInput) {
                    // Additional check: ensure the grid has content
                    const cells = gridContainer.querySelector('.waffle-grid-container');
                    if (cells) {
                        resolve();
                        return;
                    }
                }
                
                // If not ready, check again in 100ms
                setTimeout(checkReady, 100);
            };
            
            // Start checking immediately
            checkReady();
            
            // Fallback: resolve after 5 seconds even if elements aren't found
            setTimeout(resolve, 5000);
        });
    },

    /**
     * Shows a non-blocking toast notification.
     * @param {string} message - The message to display
     * @param {number} duration - How long to show the toast in milliseconds (default: 3000)
     */
    showToast(message, duration = 3000) {
        // Create toast container if it doesn't exist
        let toastContainer = document.getElementById('sheetkeys-toast-container');
        if (!toastContainer) {
            toastContainer = document.createElement('div');
            toastContainer.id = 'sheetkeys-toast-container';
            toastContainer.style.cssText = `
                position: fixed;
                bottom: 20px;
                left: 20px;
                z-index: 10000;
                pointer-events: none;
            `;
            document.body.appendChild(toastContainer);
        }

        // Create toast element
        const toast = document.createElement('div');
        toast.style.cssText = `
            background: rgba(0, 0, 0, 0.8);
            color: white;
            padding: 12px 20px;
            border-radius: 6px;
            margin-top: 10px;
            font-family: 'Google Sans', Roboto, Arial, sans-serif;
            font-size: 14px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.2);
            opacity: 0;
            transform: translateX(-100%);
            transition: all 0.3s ease;
            pointer-events: auto;
            display: flex;
            align-items: center;
            gap: 8px;
        `;
        
        // Add icon and message
        toast.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <circle cx="12" cy="12" r="10"></circle>
                <polyline points="16 12 12 8 8 12"></polyline>
                <line x1="12" y1="16" x2="12" y2="8"></line>
            </svg>
            <span>${message}</span>
        `;
        
        toastContainer.appendChild(toast);

        // Animate in
        requestAnimationFrame(() => {
            toast.style.opacity = '1';
            toast.style.transform = 'translateX(0)';
        });

        // Remove after duration
        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateX(-100%)';
            setTimeout(() => {
                toast.remove();
                // Remove container if no more toasts
                if (toastContainer.children.length === 0) {
                    toastContainer.remove();
                }
            }, 300);
        }, duration);
    },

    async loadKeyMappings() {
        const mappings = await Settings.loadUserKeyMappings();
        this.modeToKeyToCommand = {};
        for (const mode of Object.keys(mappings)) {
            const m = mappings[mode];
            this.modeToKeyToCommand[mode] = invertObjectMap(m);
        }

        // Since we don't expose in the UI the concept of mappings for insert mode commands, for command sthat
        // exist in both modes, use the mappings defined for normal mode.
        for (const [commandName, insertKey] of Object.entries(
            mappings["insert"]
        )) {
            const normalKey = mappings.normal[commandName];
            if (normalKey) {
                delete this.modeToKeyToCommand["insert"][insertKey];
                this.modeToKeyToCommand["insert"][normalKey] = commandName;
            }
        }

        this.keyMappingsPrefixes = this.buildKeyMappingsPrefixes(mappings);
    },

    // We inject the page_script into the page so that we can simulate keypress events, which must be done by a
    // page script, and not a content script.
    // See here for docs on how to inject page scripts: http://stackoverflow.com/a/9517879/46237
    injectPageScript() {
        const script = document.createElement("script");
        script.src = chrome.runtime.getURL("page_scripts/page_script.js");
        return document.documentElement.appendChild(script);
    },

    isEditable(el) {
        // Note that the window object doesn't have a tagname.
        const tagName = el.tagName ? el.tagName.toLowerCase() : null;
        return (
            el.isContentEditable ||
            tagName === "input" ||
            tagName === "textarea"
        );
    },

    onFocus(e) {
        if (!this.editor) {
            this.setupEditor();
        }
        const el = event.target;
        if (el.id === this.richTextEditorId) {
            if (SheetActions.mode === "disabled") {
                SheetActions.setMode("normal");
            }
        } else if (this.isEditable(el)) {
            SheetActions.setMode("disabled");
        }
    },

    setupEditor() {
        if (!this.editor) {
            this.editor = document.getElementById(this.richTextEditorId);
            if (this.editor) {
                // Listen for when the editor's style attribute changes. This indicates that a cell is now being
                // edited, perhaps due to double clicking into a cell.
                const observer = new MutationObserver((mutations) => {
                    if (SheetActions.mode === "disabled") return;
                    this.isEditorEditing()
                        ? SheetActions.setMode("insert")
                        : SheetActions.setMode("normal");
                });
                observer.observe(this.editor.parentNode, {
                    attributes: true,
                    attributeFilter: ["style"],
                });
            }
        }

        return this.editor;
    },

    isEditorEditing() {
        if (!this.editor) {
            return false;
        }
        // There's no obvious way to determine directly that the cell editor is currently editing a cell.
        // However, when this happens, the parent node of the editor gets a big long style attribute to portray
        // the cell editor input box.
        var style = this.editor.parentNode.getAttribute("style");
        return style != null && style != "";
    },

    // Returns a map of (partial keyString) => is_bound?
    // Note that the keys only include partial keystrings for mappings. So the mapping "d•a•p" will add "d" and
    // "d•a" keys to this map, but not "d•a•p".
    buildKeyMappingsPrefixes(keyMappings) {
        const prefixes = {};
        for (let mode in keyMappings) {
            prefixes[mode] = {};
            const modeKeyMappings = keyMappings[mode];
            for (let command of Object.keys(modeKeyMappings)) {
                const keyString = modeKeyMappings[command];
                // If the bound action is null, then treat this key as unbound.
                if (!keyString) {
                    continue;
                }
                const keys = keyString.split(Commands.KEY_SEPARATOR);
                for (let i = 0; i < keys.length - 1; i++) {
                    let prefix = keys
                        .slice(0, i + 1)
                        .join(Commands.KEY_SEPARATOR);
                    prefixes[mode][prefix] = true;
                }
            }
        }
        return prefixes;
    },

    cancelEvent(e) {
        e.preventDefault();
        e.stopPropagation();
    },

    onKeydown(e) {
        const keyString = KeyboardUtils.getKeyString(e);
        // console.log "keydown event. keyString:", keyString, e.keyCode, e.keyIdentifier, e
        if (this.ignoreKeys || SheetActions.mode == "disabled") {
            return;
        }

        if (!keyString) {
            return;
        } // Ignore key presses which are just modifiers.

        // In replace mode, we're waiting for one character to be typed, and we will replace the cell's contents
        // with that character and then return to normal mode.
        if (SheetActions.mode === "replace") {
            if (keyString === "esc") {
                this.cancelEvent(e);
                SheetActions.setMode("normal");
            } else {
                SheetActions.changeCell();
                setTimeout(() => SheetActions.commitCellChanges(), 0);
            }
            return;
        }

        this.keyQueue.push(keyString);
        // There are keymaps for two different modes: insert and normal. When we're in one of the visual modes,
        // use the normal keymap. The commands themselves may implement mode-specific behavior.
        const modeToUse = SheetActions.mode == "insert" ? "insert" : "normal";
        if (this.keyQueue.length > this.maxKeyMappingLength) {
            this.keyQueue.shift();
        }
        const modeMappings = this.modeToKeyToCommand[modeToUse] || [];
        const modePrefixes = this.keyMappingsPrefixes[modeToUse] || [];
        // See if a bound command matches the typed key sequence. If so, execute it.
        // Prioritize longer mappings over shorter mappings.
        for (
            let i = Math.min(this.maxKeyMappingLength, this.keyQueue.length);
            i >= 1;
            i--
        ) {
            var fn;
            const keySequence = this.keyQueue
                .slice(this.keyQueue.length - i, this.keyQueue.length)
                .join(Commands.KEY_SEPARATOR);
            // If this key could be part of one of the bound key mapping, don't pass it through to the page.
            // Also, if some longer mapping partially matches this key sequence, then wait for more keys, and
            // don't immediately apply a shorter mapping which also matches this key sequence.
            if (modePrefixes[keySequence]) {
                this.cancelEvent(e);
                return;
            }

            // Rishi: Debugging statement
            console.log("Key sequence pressed", keySequence);

            if ((commandName = modeMappings[keySequence])) {
                this.keyQueue = [];
                this.cancelEvent(e);
                Commands.commands[commandName].fn();
                console.log("Command done:", commandName);
            }
        }
    },

    // modifiers: Optional; an object with these boolean properties: meta, alt, shift, control.
    typeKey(keyCode, modifiers) {
        // Rishi: Print log
        // var modifiers_str = JSON.stringify(modifiers) || "no modifiers";
        // console.log(`Typing: ${keyCode} + ${modifiers_str}`);

        if (keyCode == null) {
            throw "The keyCode provided to typeKey() is null.";
        }
        this.ignoreKeys = true;
        if (!modifiers) {
            modifiers = {};
        }
        document.getElementById("sheetkeys-json-message").innerText =
            JSON.stringify({ keyCode, mods: modifiers });
        window.dispatchEvent(
            new CustomEvent("sheetkeys-simulate-key-event", {})
        );
        this.ignoreKeys = false;
    },

    reflowGrid() {
        // When you hide a DOM element, Google's Waffle grid doesn't know to reflow and take up the full viewport.
        // You can trigger a reflow by resizing the browser or by clicking on the Explore button in the lower-left
        // corner.
        const exploreButton = document.querySelector(
            ".waffle-assistant-entry [role=button]"
        );
        KeyboardUtils.simulateClick(exploreButton);
        KeyboardUtils.simulateClick(exploreButton); // Click twice to show and then hide.
    },
};

// Don't initialize this Sheets UI if this code is being loaded from our extension's options page.
if (window.document && !document.location.pathname.endsWith("harness.html"))
    UI.init();

window.UI = UI;
