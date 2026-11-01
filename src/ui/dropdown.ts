/**
 * A Word-style drop-down: the list opens directly below the button (the native macOS popup instead
 * centres the current item over the control and scrolls). Keeps the editor's focus and selection.
 */

export interface DropdownOption {
  value: string;
  label: string;
  /** Inline style for the item (e.g. its own font family). */
  style?: string;
  /** Explicit tick (for menus with several groups); otherwise the item matching the value is ticked. */
  checked?: boolean;
  /** Draws a divider above this item. */
  separatorBefore?: boolean;
}

export interface Dropdown {
  el: HTMLElement;
  /** Shows `value` as current; `placeholder` is shown instead when value is '' (e.g. mixed fonts). */
  setValue(value: string, placeholder?: string): void;
}

let openMenu: { close: () => void } | null = null;

/** Builds a drop-down button. `options` is read each time the menu opens. */
export function createDropdown(opts: {
  className: string;
  title: string;
  options: () => DropdownOption[];
  onSelect: (value: string) => void;
  /** Fixed button content (e.g. an icon) instead of the current option's label. */
  buttonHtml?: string;
}): Dropdown {
  const button = document.createElement('button');
  button.className = `tb-dropdown ${opts.className}`;
  button.title = opts.title;
  button.setAttribute('aria-haspopup', 'listbox');
  const label = document.createElement('span');
  label.className = 'tb-dropdown-label';
  button.append(label);
  if (opts.buttonHtml) label.innerHTML = opts.buttonHtml;
  let current = '';

  const open = () => {
    const menu = document.createElement('div');
    menu.className = 'dropdown-menu';
    menu.setAttribute('role', 'listbox');
    const items = opts.options();
    let active = Math.max(0, items.findIndex((o) => o.value === current));
    const rows = items.map((o, i) => {
      const row = document.createElement('div');
      row.className = 'dropdown-item';
      row.setAttribute('role', 'option');
      if (o.checked ?? o.value === current) row.classList.add('checked');
      if (o.separatorBefore) row.classList.add('separated');
      const text = document.createElement('span');
      text.textContent = o.label;
      if (o.style) text.setAttribute('style', o.style);
      row.append(text);
      row.addEventListener('mousemove', () => highlight(i));
      row.addEventListener('mousedown', (e) => {
        e.preventDefault();
        choose(i);
      });
      menu.append(row);
      return row;
    });
    const highlight = (i: number) => {
      active = (i + rows.length) % rows.length;
      rows.forEach((r, k) => r.classList.toggle('active', k === active));
    };
    const choose = (i: number) => {
      close();
      opts.onSelect(items[i].value);
    };
    let typed = '';
    let typedAt = 0;
    const onKey = (e: KeyboardEvent) => {
      // While open, the menu owns the keyboard (the editor keeps focus for its selection).
      e.preventDefault();
      e.stopPropagation();
      if (e.key === 'Escape') close();
      else if (e.key === 'ArrowDown') highlight(active + 1);
      else if (e.key === 'ArrowUp') highlight(active - 1);
      else if (e.key === 'Enter' || e.key === ' ') choose(active);
      else if (e.key.length === 1) {
        // Type-ahead: "ti" jumps to Times New Roman.
        typed = performance.now() - typedAt > 700 ? e.key.toLowerCase() : typed + e.key.toLowerCase();
        typedAt = performance.now();
        const hit = items.findIndex((o) => o.label.toLowerCase().startsWith(typed));
        if (hit >= 0) highlight(hit);
      }
      rows[active]?.scrollIntoView({ block: 'nearest' });
    };
    const onOutside = (e: MouseEvent) => {
      if (!menu.contains(e.target as Node) && e.target !== button && !button.contains(e.target as Node)) close();
    };
    const close = () => {
      menu.remove();
      button.classList.remove('open');
      window.removeEventListener('keydown', onKey, true);
      window.removeEventListener('mousedown', onOutside, true);
      window.removeEventListener('blur', close);
      window.removeEventListener('resize', close);
      if (openMenu?.close === close) openMenu = null;
    };
    openMenu?.close();
    openMenu = { close };
    document.body.append(menu);
    const r = button.getBoundingClientRect();
    menu.style.left = `${Math.round(r.left)}px`;
    menu.style.top = `${Math.round(r.bottom + 4)}px`;
    menu.style.minWidth = `${Math.round(r.width)}px`;
    menu.style.maxHeight = `${Math.max(160, window.innerHeight - r.bottom - 16)}px`;
    button.classList.add('open');
    highlight(active);
    rows[active]?.scrollIntoView({ block: 'nearest' });
    window.addEventListener('keydown', onKey, true);
    window.addEventListener('mousedown', onOutside, true);
    window.addEventListener('blur', close);
    window.addEventListener('resize', close);
  };

  button.addEventListener('mousedown', (e) => {
    e.preventDefault(); // keep the editor's focus and selection
    if (button.classList.contains('open')) openMenu?.close();
    else open();
  });

  return {
    el: button,
    setValue(value, placeholder = '') {
      current = value;
      if (opts.buttonHtml) return;
      const opt = opts.options().find((o) => o.value === value);
      label.textContent = value ? (opt?.label ?? value) : placeholder;
      label.classList.toggle('placeholder', !value);
    },
  };
}
