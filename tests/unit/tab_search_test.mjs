/**
 * Scripted test for TabSearch.extractTabName.
 *
 * Runs under plain Node using jsdom to simulate the DOM. The production
 * content script (tab_search.js) is evaluated in a jsdom-backed vm
 * context so we can call TabSearch.extractTabName against HTML fixtures.
 *
 * Run:   npm test
 * Or:    node tests/unit/tab_search_test.mjs
 */
import { JSDOM } from "jsdom";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import vm from "node:vm";

const here = dirname(fileURLToPath(import.meta.url));
const src = readFileSync(
    resolve(here, "../../content_scripts/tab_search.js"),
    "utf8"
);

const dom = new JSDOM("<!doctype html><html><body></body></html>");
const { window } = dom;

// Evaluate tab_search.js in a context that exposes the jsdom globals it
// expects (document, window, NodeFilter) plus a stub for SheetActions
// which is referenced inside TabSearch.collectTabs (not under test here).
const context = vm.createContext({
    window,
    document: window.document,
    NodeFilter: window.NodeFilter,
    Element: window.Element,
    Node: window.Node,
    SheetActions: { getTabEls: () => [] },
    KeyboardUtils: { simulateClick() {} },
    console,
});
vm.runInContext(src, context);
const TabSearch = context.window.TabSearch;

// --- tiny test runner ---------------------------------------------------
let passed = 0;
let failed = 0;
const failures = [];

function test(name, fn) {
    try {
        fn();
        console.log(`  ✓ ${name}`);
        passed++;
    } catch (e) {
        console.log(`  ✗ ${name}`);
        console.log(`      ${e.message.replace(/\n/g, "\n      ")}`);
        failed++;
        failures.push(name);
    }
}

function eq(actual, expected, label = "") {
    if (actual !== expected) {
        throw new Error(
            `${label ? label + ": " : ""}expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`
        );
    }
}

function buildTab(html) {
    const container = window.document.createElement("div");
    container.innerHTML = html.trim();
    return container.firstElementChild;
}

// --- fixtures -----------------------------------------------------------
console.log("TabSearch.extractTabName");

// Captured verbatim from Rishi's live Google Sheets on 2026-04-23.
// This is the authoritative fixture for the current Sheets DOM.
const REAL_SHEETS_TAB_MARKETS = `
<div class="goog-inline-block docs-sheet-tab docs-material" role="button" aria-expanded="false" tabindex="0" aria-haspopup="true" id=":y">
  <div class="goog-inline-block docs-sheet-tab-outer-box">
    <div class="goog-inline-block docs-sheet-tab-inner-box">
      <div class="goog-inline-block docs-sheet-tab-caption">
        <div class="docs-icon goog-inline-block docs-sheet-form-icon-container" aria-hidden="true" data-tooltip="Form responses" aria-label="Form responses" style="display: none;">
          <div class="docs-icon-img-container docs-icon-img docs-icon-form"></div>
        </div>
        <g pointer-events="all" class="docs-sheet-comment-indicator-container docos-comments-pe" style="display: none;">
          <svg class="docs-sheet-comment-indicator" xmlns="http://www.w3.org/2000/svg" height="20" width="20" viewBox="0 0 20 20">
            <path d="M20 10C20 4.47715 15.5228 0 10 0C4.47715 0 0 4.47715 0 10V19.2308C0 19.6556 0.344397 20 0.76923 20H10C15.5228 20 20 15.5228 20 10Z"></path>
          </svg>
          <text class="docs-sheet-comment-indicator-text">0</text>
        </g>
        <div class="docs-icon goog-inline-block docs-sheet-lock-container" aria-hidden="true" style="display: none;">
          <div class="docs-icon-img-container docs-icon-img docs-icon-locked" title="Protected"></div>
        </div>
        <div class="docs-icon goog-inline-block docs-sheet-database-icon-container" aria-hidden="true" style="display: none;">
          <div class="docs-icon-img-container docs-icon-img docs-icon-database"></div>
        </div>
        <div class="docs-icon goog-inline-block docs-sheet-timeline-icon-container" aria-hidden="true" style="display: none;">
          <div class="docs-icon-img-container docs-icon-img docs-sheet-timeline-icon"></div>
        </div>
        <span dir="ltr" class="docs-sheet-tab-name" spellcheck="false">Markets</span>
        <div class="docs-sheet-tab-color" style="background: transparent;"></div>
      </div>
      <div class="docs-icon goog-inline-block docs-sheet-tab-dropdown" aria-hidden="true">
        <div class="docs-icon-img-container docs-icon-img goog-inline-block docs-icon-arrow-dropdown">&nbsp;</div>
      </div>
    </div>
  </div>
</div>`;

test("real Sheets DOM: returns sheet name without comment-count prefix", () => {
    const tab = buildTab(REAL_SHEETS_TAB_MARKETS);
    eq(TabSearch.extractTabName(tab), "Markets");
});

test("prefers .docs-sheet-tab-name over aria-label", () => {
    const tab = buildTab(`
        <div class="docs-sheet-tab" aria-label="Stale label">
            <span class="docs-sheet-tab-name">Markets</span>
        </div>
    `);
    eq(TabSearch.extractTabName(tab), "Markets");
});

test("trims whitespace around the name", () => {
    const tab = buildTab(`
        <div class="docs-sheet-tab">
            <span class="docs-sheet-tab-name">  Lifetime income  </span>
        </div>
    `);
    eq(TabSearch.extractTabName(tab), "Lifetime income");
});

test("falls back to aria-label when name span is missing", () => {
    const tab = buildTab(
        `<div class="docs-sheet-tab" aria-label="Contacts"></div>`
    );
    eq(TabSearch.extractTabName(tab), "Contacts");
});

test("falls back to aria-label when name span is empty", () => {
    const tab = buildTab(`
        <div class="docs-sheet-tab" aria-label="Reena">
            <span class="docs-sheet-tab-name"></span>
        </div>
    `);
    eq(TabSearch.extractTabName(tab), "Reena");
});

test("falls back to textContent when neither span nor aria-label exists", () => {
    const tab = buildTab(
        `<div class="docs-sheet-tab">  Macro history  </div>`
    );
    eq(TabSearch.extractTabName(tab), "Macro history");
});

test("ignores hidden comment-indicator text inside the tab caption", () => {
    // Stripped-down version of the real fixture: name span wins even when
    // a sibling <text> node holds a stray digit.
    const tab = buildTab(`
        <div class="docs-sheet-tab">
            <g class="docs-sheet-comment-indicator-container">
                <text class="docs-sheet-comment-indicator-text">0</text>
            </g>
            <span class="docs-sheet-tab-name">Assets</span>
        </div>
    `);
    eq(TabSearch.extractTabName(tab), "Assets");
});

// Regression guard for the original bug Rishi reported: every tab name
// came back prefixed with the digit "0" because textContent on the tab
// leaked the hidden comment-count indicator.
test("regression: does not leak comment-indicator '0' prefix", () => {
    const names = ["Markets", "Reena", "Contacts", "CSV scratch", "Assets"];
    for (const name of names) {
        const tab = buildTab(`
            <div class="docs-sheet-tab">
                <g class="docs-sheet-comment-indicator-container" style="display:none;">
                    <text class="docs-sheet-comment-indicator-text">0</text>
                </g>
                <span class="docs-sheet-tab-name">${name}</span>
            </div>
        `);
        eq(TabSearch.extractTabName(tab), name, `tab "${name}"`);
    }
});

// --- summary ------------------------------------------------------------
console.log(`\n${passed} passed, ${failed} failed`);
if (failed > 0) {
    console.log("failed: " + failures.join(", "));
    process.exit(1);
}
