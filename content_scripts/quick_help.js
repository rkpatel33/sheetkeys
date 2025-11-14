/**
 * QuickHelp - A visual keyboard shortcut reference modal for SheetKeys
 * Displays all keyboard mappings organized by category with icons and visual aids
 */

class QuickHelp {
  constructor() {
    this.el = null;
    this.isVisible = false;
    this.mappings = null;

    // Inline SVG icons (subset of Lucide icons)
    this.icons = {
      'keyboard': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="16" x="2" y="4" rx="2"/><path d="M6 8h.001"/><path d="M10 8h.001"/><path d="M14 8h.001"/><path d="M18 8h.001"/><path d="M8 12h.001"/><path d="M12 12h.001"/><path d="M16 12h.001"/><path d="M7 16h10"/></svg>',
      'x': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18"/><path d="m6 6 12 12"/></svg>',
      'move': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="5 9 2 12 5 15"/><polyline points="9 5 12 2 15 5"/><polyline points="15 19 12 22 9 19"/><polyline points="19 9 22 12 19 15"/><line x1="2" x2="22" y1="12" y2="12"/><line x1="12" x2="12" y1="2" y2="22"/></svg>',
      'square': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="3" rx="2"/></svg>',
      'edit-3': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>',
      'type': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="4 7 4 4 20 4 20 7"/><line x1="9" x2="15" y1="20" y2="20"/><line x1="12" x2="12" y1="4" y2="20"/></svg>',
      'palette': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="13.5" cy="6.5" r=".5"/><circle cx="17.5" cy="10.5" r=".5"/><circle cx="8.5" cy="7.5" r=".5"/><circle cx="6.5" cy="12.5" r=".5"/><path d="M12 2C6.5 2 2 6.5 2 12s4.5 10 10 10c.926 0 1.648-.746 1.648-1.688 0-.437-.18-.835-.437-1.125-.29-.289-.438-.652-.438-1.125a1.64 1.64 0 0 1 1.668-1.668h1.996c3.051 0 5.555-2.503 5.555-5.554C21.965 6.012 17.461 2 12 2z"/></svg>',
      'droplet': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z"/></svg>',
      'hash': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="4" x2="20" y1="9" y2="9"/><line x1="4" x2="20" y1="15" y2="15"/><line x1="10" x2="8" y1="3" y2="21"/><line x1="16" x2="14" y1="3" y2="21"/></svg>',
      'zoom-in': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/><path d="M11 8v6"/><path d="M8 11h6"/></svg>',
      'columns-3': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="3" rx="2"/><path d="M9 3v18"/><path d="M15 3v18"/></svg>',
      'filter': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>',
      'grid-3x3': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="3" rx="2"/><path d="M3 9h18"/><path d="M3 15h18"/><path d="M9 3v18"/><path d="M15 3v18"/></svg>',
      'circle-help': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><path d="M12 17h.01"/></svg>',
      'arrow-up': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m5 12 7-7 7 7"/><path d="M12 19V5"/></svg>',
      'arrow-down': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14"/><path d="m19 12-7 7-7-7"/></svg>',
      'arrow-left': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m12 19-7-7 7-7"/><path d="M19 12H5"/></svg>',
      'arrow-right': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>',
      'arrow-up-to-line': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 3h14"/><path d="m18 13-6-6-6 6"/><path d="M12 7v14"/></svg>',
      'arrow-down-to-line': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 17V3"/><path d="m6 11 6 6 6-6"/><path d="M19 21H5"/></svg>',
      'arrow-left-to-line': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 19V5"/><path d="m13 6-6 6 6 6"/><path d="M7 12h14"/></svg>',
      'arrow-right-to-line': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 12H3"/><path d="m11 18 6-6-6-6"/><path d="M21 5v14"/></svg>',
      'undo': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7v6h6"/><path d="M21 17a9 9 0 0 0-9-9 9 9 0 0 0-6 2.3L3 13"/></svg>',
      'redo': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 7v6h-6"/><path d="M3 17a9 9 0 0 1 9-9 9 9 0 0 1 6 2.3l3 2.7"/></svg>',
      'trash-2': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/><line x1="10" x2="10" y1="11" y2="17"/><line x1="14" x2="14" y1="11" y2="17"/></svg>',
      'copy': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg>',
      'clipboard-paste': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 2H9a1 1 0 0 0-1 1v2c0 .6.4 1 1 1h6c.6 0 1-.4 1-1V3c0-.6-.4-1-1-1Z"/><path d="M8 4H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2M16 4h2a2 2 0 0 1 2 2v2M11 14h10"/><path d="m17 10 4 4-4 4"/></svg>',
      'align-start-horizontal': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="6" height="16" x="4" y="6" rx="2"/><rect width="6" height="9" x="14" y="6" rx="2"/><path d="M22 2H2"/></svg>',
      'align-center-horizontal': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12h20"/><rect width="6" height="16" x="9" y="4" rx="2"/></svg>',
      'align-end-horizontal': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="6" height="16" x="4" y="2" rx="2"/><rect width="6" height="9" x="14" y="9" rx="2"/><path d="M22 22H2"/></svg>',
      'align-start-vertical': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="9" height="6" x="6" y="14" rx="2"/><rect width="16" height="6" x="6" y="4" rx="2"/><path d="M2 2v20"/></svg>',
      'align-end-vertical': '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="16" height="6" x="2" y="4" rx="2"/><rect width="9" height="6" x="9" y="14" rx="2"/><path d="M22 22V2"/></svg>',
    };

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
   * Loads an SVG icon by name from inline icons
   */
  async loadIcon(iconName) {
    return this.icons[iconName] || '<svg></svg>';
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
