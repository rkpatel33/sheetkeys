/**
 * QuickHelp - A visual keyboard shortcut reference modal for SheetKeys
 * Displays all keyboard mappings organized by category with icons and visual aids
 */

class QuickHelp {
  constructor() {
    this.el = null;
    this.isVisible = false;
    this.mappings = null;

    // Color mappings for visual display
    this.colorMap = {
      colorCellLightYellow: { bg: '#fff2cc', label: 'Yellow' },
      colorCellWhite: { bg: '#ffffff', label: 'White', border: true },
      colorCellLightCornflowerBlue3: { bg: '#cfe2f3', label: 'Blue' },
      colorCellLightBlue3: { bg: '#cfe2f3', label: 'Blue 3' },
      colorCellLightPurple: { bg: '#d9d2e9', label: 'Purple' },
      colorCellLightRed3: { bg: '#f4cccc', label: 'Red' },
      colorCellLightGray2: { bg: '#efefef', label: 'Gray' },
      colorCellDarkGray1: { bg: '#cccccc', label: 'Dark Gray' },
      colorCellYellow: { bg: '#fff2cc', label: 'Yellow' },

      colorCellFontColorBlue: { color: '#0000ff', label: 'Blue' },
      colorCellFontColorBlack: { color: '#000000', label: 'Black' },
      colorCellFontColorRed: { color: '#ff0000', label: 'Red' },
      colorCellFontColorDarkRed: { color: '#990000', label: 'Dark Red' },
    };

    // Category configurations with display metadata
    this.categories = [
      {
        id: 'movement',
        title: 'Movement',
        icon: 'move',
        layout: 'directional-grid',
        description: 'Navigate through cells and scroll the sheet'
      },
      {
        id: 'selection',
        title: 'Selection',
        icon: 'square',
        layout: 'compact-list',
        description: 'Select cells, rows, and columns'
      },
      {
        id: 'editing',
        title: 'Editing',
        icon: 'edit-3',
        layout: 'two-column',
        description: 'Edit cell contents, insert/delete rows, copy/paste'
      },
      {
        id: 'formatting',
        title: 'Formatting',
        icon: 'type',
        layout: 'compact-list',
        description: 'Text alignment, wrapping, and font sizes'
      },
      {
        id: 'cell color',
        title: 'Cell Colors',
        icon: 'palette',
        layout: 'color-swatches',
        colorType: 'background'
      },
      {
        id: 'font color',
        title: 'Font Colors',
        icon: 'droplet',
        layout: 'color-swatches',
        colorType: 'text'
      },
      {
        id: 'number format',
        title: 'Number Formats',
        icon: 'hash',
        layout: 'visual-examples'
      },
      {
        id: 'border',
        title: 'Borders',
        icon: 'square',
        layout: 'visual-examples'
      },
      {
        id: 'zoom',
        title: 'Zoom',
        icon: 'zoom-in',
        layout: 'compact-list'
      },
      {
        id: 'tabs',
        title: 'Tabs',
        icon: 'columns-3',
        layout: 'compact-list'
      },
      {
        id: 'filter',
        title: 'Filtering',
        icon: 'filter',
        layout: 'compact-list'
      },
      {
        id: 'ui',
        title: 'UI Controls',
        icon: 'grid-3x3',
        layout: 'compact-list'
      },
      {
        id: 'other',
        title: 'Other',
        icon: 'circle-help',
        layout: 'compact-list'
      },
    ];
  }

  /**
   * Formats a key string for display (replaces • with · for better readability)
   */
  formatKeyString(keyString) {
    if (!keyString) return '';
    return keyString
      .replace(/•/g, '·')
      .replace(/<M-/g, '⌘')
      .replace(/<C-/g, '⌃')
      .replace(/<A-/g, '⌥')
      .replace(/<S-/g, '⇧')
      .replace(/>/g, '')
      .replace(/slash/g, '/');
  }

  /**
   * Creates a key pill element
   */
  createKeyPill(keyString, options = {}) {
    const pill = document.createElement('span');
    pill.className = 'key-pill';
    pill.textContent = this.formatKeyString(keyString);

    if (options.backgroundColor) {
      pill.style.backgroundColor = options.backgroundColor;
      if (options.border) {
        pill.style.border = '1px solid #ddd';
      }
    }
    if (options.color) {
      pill.style.color = options.color;
    }

    return pill;
  }

  /**
   * Loads an SVG icon by name
   */
  async loadIcon(iconName) {
    try {
      const iconPath = chrome.runtime.getURL(`icons/lucide/${iconName}.svg`);
      const response = await fetch(iconPath);
      const svgText = await response.text();
      return svgText;
    } catch (error) {
      console.error(`Failed to load icon: ${iconName}`, error);
      return '<svg></svg>';
    }
  }

  /**
   * Creates a command item element
   */
  createCommandItem(commandName, keyString, options = {}) {
    const command = Commands.commands[commandName];
    if (!command) return null;

    const item = document.createElement('div');
    item.className = 'command-item';

    // Add icon if specified
    if (options.icon) {
      const iconContainer = document.createElement('span');
      iconContainer.className = 'command-icon';
      iconContainer.innerHTML = options.icon;
      item.appendChild(iconContainer);
    }

    // Add key pill
    const pillOptions = {};
    if (options.backgroundColor) pillOptions.backgroundColor = options.backgroundColor;
    if (options.color) pillOptions.color = options.color;
    if (options.border) pillOptions.border = options.border;

    item.appendChild(this.createKeyPill(keyString, pillOptions));

    // Add label
    const label = document.createElement('span');
    label.className = 'command-label';
    label.textContent = command.name || commandName;
    item.appendChild(label);

    // Add preview if specified
    if (options.preview) {
      const preview = document.createElement('span');
      preview.className = 'example-preview';
      preview.textContent = options.preview;
      item.appendChild(preview);
    }

    return item;
  }

  /**
   * Gets commands for a specific category
   */
  getCommandsForCategory(categoryId) {
    const commands = [];
    for (const [commandName, command] of Object.entries(Commands.commands)) {
      if (command.group === categoryId && !command.hiddenFromHelp) {
        const keyString = this.mappings[commandName];
        if (keyString) {
          commands.push({ name: commandName, key: keyString, command });
        }
      }
    }
    return commands;
  }

  /**
   * Renders a category section
   */
  async renderCategory(category) {
    const section = document.createElement('div');
    section.className = 'help-category';

    // Header
    const header = document.createElement('h2');
    header.className = 'help-category-header';

    const icon = await this.loadIcon(category.icon);
    const iconSpan = document.createElement('span');
    iconSpan.innerHTML = icon;
    header.appendChild(iconSpan);

    const title = document.createElement('span');
    title.textContent = category.title;
    header.appendChild(title);

    section.appendChild(header);

    // Content
    const content = document.createElement('div');
    content.className = `help-category-content layout-${category.layout}`;

    const commands = this.getCommandsForCategory(category.id);

    if (commands.length === 0) {
      return null; // Skip empty categories
    }

    // Render based on layout type
    switch (category.layout) {
      case 'color-swatches':
        await this.renderColorSwatches(content, commands, category);
        break;
      case 'visual-examples':
        await this.renderVisualExamples(content, commands, category);
        break;
      default:
        await this.renderStandardLayout(content, commands, category);
        break;
    }

    section.appendChild(content);
    return section;
  }

  /**
   * Renders color swatches layout
   */
  async renderColorSwatches(container, commands, category) {
    for (const { name, key } of commands) {
      const colorInfo = this.colorMap[name];
      if (!colorInfo) continue;

      const item = document.createElement('div');
      item.className = 'color-item';

      const pillOptions = {};
      if (category.colorType === 'background') {
        pillOptions.backgroundColor = colorInfo.bg;
        if (colorInfo.border) pillOptions.border = true;
      } else if (category.colorType === 'text') {
        pillOptions.color = colorInfo.color;
      }

      item.appendChild(this.createKeyPill(key, pillOptions));

      const label = document.createElement('div');
      label.className = 'command-label';
      label.textContent = colorInfo.label;
      item.appendChild(label);

      container.appendChild(item);
    }
  }

  /**
   * Renders visual examples layout (for borders, number formats)
   */
  async renderVisualExamples(container, commands, category) {
    const exampleMap = {
      // Number formats
      numberFormatNumber2: '1,234.56',
      numberFormatDollar2: '$1,234.56',
      numberFormatPercent2: '12.34%',
      decimalIncrease: '1.2 → 1.23',
      decimalDecrease: '1.23 → 1.2',

      // Borders
      borderTop: '▔▔ Top',
      borderBottom: '▁▁ Bottom',
      borderLeft: '▕ Left',
      borderRight: '▏ Right',
      borderClear: '⊠ Clear',
    };

    for (const { name, key, command } of commands) {
      const item = document.createElement('div');
      item.className = 'example-item';

      item.appendChild(this.createKeyPill(key));

      const label = document.createElement('span');
      label.className = 'command-label';
      label.textContent = command.name;
      item.appendChild(label);

      if (exampleMap[name]) {
        const preview = document.createElement('span');
        preview.className = 'example-preview';
        preview.textContent = exampleMap[name];
        item.appendChild(preview);
      }

      container.appendChild(item);
    }
  }

  /**
   * Renders standard layout (directional-grid, two-column, compact-list)
   */
  async renderStandardLayout(container, commands, category) {
    // Load common icons
    const iconMap = {
      moveUp: await this.loadIcon('arrow-up'),
      moveDown: await this.loadIcon('arrow-down'),
      moveLeft: await this.loadIcon('arrow-left'),
      moveRight: await this.loadIcon('arrow-right'),
      moveEndUp: await this.loadIcon('arrow-up-to-line'),
      moveEndUp2: await this.loadIcon('arrow-up-to-line'),
      moveEndDown: await this.loadIcon('arrow-down-to-line'),
      moveEndDown2: await this.loadIcon('arrow-down-to-line'),
      moveEndLeft: await this.loadIcon('arrow-left-to-line'),
      moveEndLeft2: await this.loadIcon('arrow-left-to-line'),
      moveEndRight: await this.loadIcon('arrow-right-to-line'),
      moveEndRight2: await this.loadIcon('arrow-right-to-line'),
      editCell: await this.loadIcon('edit-3'),
      editCellAppend: await this.loadIcon('edit-3'),
      undo: await this.loadIcon('undo'),
      redo: await this.loadIcon('redo'),
      clear: await this.loadIcon('trash-2'),
      deleteRowsOrColumns: await this.loadIcon('trash-2'),
      deleteColumns: await this.loadIcon('trash-2'),
      copy: await this.loadIcon('copy'),
      copyRowOrSelection: await this.loadIcon('copy'),
      paste: await this.loadIcon('clipboard-paste'),
      pasteValuesOnly: await this.loadIcon('clipboard-paste'),
      pasteFormatOnly: await this.loadIcon('clipboard-paste'),
      pasteFormulaOnly: await this.loadIcon('clipboard-paste'),
    };

    for (const { name, key } of commands) {
      const item = this.createCommandItem(name, key, {
        icon: iconMap[name]
      });
      if (item) {
        container.appendChild(item);
      }
    }
  }

  /**
   * Creates the modal HTML structure
   */
  async createModal() {
    if (this.el) return;

    // Load mappings
    this.mappings = (await Settings.loadUserKeyMappings()).normal;

    // Create overlay
    const overlay = document.createElement('div');
    overlay.className = 'sheetkeys-quick-help-overlay';
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) {
        this.hide();
      }
    });

    // Create modal
    const modal = document.createElement('div');
    modal.className = 'sheetkeys-quick-help-modal';
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-modal', 'true');
    modal.setAttribute('aria-labelledby', 'quick-help-title');

    // Header
    const header = document.createElement('div');
    header.className = 'sheetkeys-quick-help-header';

    const title = document.createElement('h1');
    title.className = 'sheetkeys-quick-help-title';
    title.id = 'quick-help-title';
    const keyboardIcon = await this.loadIcon('keyboard');
    title.innerHTML = `${keyboardIcon} SheetKeys Quick Reference`;
    header.appendChild(title);

    const closeBtn = document.createElement('button');
    closeBtn.className = 'sheetkeys-quick-help-close';
    closeBtn.setAttribute('aria-label', 'Close');
    const closeIcon = await this.loadIcon('x');
    closeBtn.innerHTML = closeIcon;
    closeBtn.addEventListener('click', () => this.hide());
    header.appendChild(closeBtn);

    modal.appendChild(header);

    // Body
    const body = document.createElement('div');
    body.className = 'sheetkeys-quick-help-body';
    body.setAttribute('tabindex', '0');

    // Render all categories
    for (const category of this.categories) {
      const section = await this.renderCategory(category);
      if (section) {
        body.appendChild(section);
      }
    }

    modal.appendChild(body);

    // Footer
    const footer = document.createElement('div');
    footer.className = 'sheetkeys-quick-help-footer';
    footer.textContent = 'Press ;·q or ESC to close';
    modal.appendChild(footer);

    overlay.appendChild(modal);

    // Add keyboard handler with capture phase to intercept before UI.js
    overlay.addEventListener('keydown', (e) => this.onKeydown(e), true);

    this.el = overlay;
    document.body.appendChild(this.el);
  }

  /**
   * Handles keydown events when modal is open
   */
  onKeydown(e) {
    const keyString = KeyboardUtils.getKeyString(e);

    if (keyString === 'esc') {
      e.preventDefault();
      e.stopPropagation();
      this.hide();
    }
  }

  /**
   * Shows the quick help modal
   */
  async show() {
    if (this.isVisible) return;

    await this.createModal();

    this.el.style.display = 'flex';
    this.isVisible = true;

    // Focus the body for keyboard scrolling
    setTimeout(() => {
      const body = this.el.querySelector('.sheetkeys-quick-help-body');
      if (body) body.focus();
    }, 100);
  }

  /**
   * Hides the quick help modal
   */
  hide() {
    if (!this.isVisible || !this.el) return;

    this.el.classList.add('closing');
    setTimeout(() => {
      if (this.el) {
        this.el.style.display = 'none';
        this.el.classList.remove('closing');
      }
      this.isVisible = false;
    }, 200);
  }

  /**
   * Toggles the quick help modal
   */
  async toggle() {
    if (this.isVisible) {
      this.hide();
    } else {
      await this.show();
    }
  }
}

// Export for use in ui.js
window.QuickHelp = QuickHelp;
