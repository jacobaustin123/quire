/** A single small floating input used for links, equations and go-to-page. */

export interface PopoverOptions {
  /** Where to place the popover; null centers it near the top of the window. */
  anchor: DOMRect | null;
  label: string;
  value: string;
  placeholder?: string;
  multiline?: boolean;
  mono?: boolean;
  hint?: string;
  /** Losing focus commits (equations) rather than cancels (links, go-to). */
  commitOnBlur?: boolean;
  /** Called on every edit; may render into the preview element. */
  onInput?: (value: string, preview: HTMLElement) => void;
  onCommit: (value: string) => void;
  onCancel?: () => void;
}

const el = (): HTMLElement => document.getElementById('popover')!;
let active: PopoverOptions | null = null;

/** Whether a popover is currently open. */
export function popoverOpen(): boolean {
  return active !== null;
}

/** Closes the popover, invoking onCancel unless `committed`. */
export function closePopover(committed = false): void {
  const opts = active;
  if (!opts) return;
  active = null;
  el().hidden = true;
  el().replaceChildren();
  if (!committed) opts.onCancel?.();
}

/** Shows the popover with the given options, replacing any open one. */
export function showPopover(opts: PopoverOptions): void {
  closePopover();
  active = opts;
  const root = el();
  const label = document.createElement('div');
  label.className = 'pop-label';
  label.textContent = opts.label;
  const input = document.createElement(opts.multiline ? 'textarea' : 'input') as HTMLInputElement | HTMLTextAreaElement;
  input.value = opts.value;
  input.placeholder = opts.placeholder ?? '';
  input.spellcheck = false;
  if (opts.mono) input.classList.add('mono');
  if (input instanceof HTMLTextAreaElement) input.rows = Math.min(8, Math.max(2, opts.value.split('\n').length));
  const preview = document.createElement('div');
  preview.className = 'pop-preview';
  root.replaceChildren(label, input);
  if (opts.onInput) root.append(preview);
  if (opts.hint) {
    const hint = document.createElement('div');
    hint.className = 'pop-hint';
    hint.textContent = opts.hint;
    root.append(hint);
  }
  root.hidden = false;
  position(root, opts.anchor);

  const commit = () => {
    const v = input.value;
    closePopover(true);
    opts.onCommit(v);
  };
  input.addEventListener('input', () => {
    opts.onInput?.(input.value, preview);
    if (input instanceof HTMLTextAreaElement) input.rows = Math.min(8, Math.max(2, input.value.split('\n').length));
  });
  input.addEventListener('keydown', (e) => {
    const ev = e as KeyboardEvent;
    if (ev.key === 'Escape') {
      ev.preventDefault();
      closePopover();
    } else if (ev.key === 'Enter' && (!opts.multiline || !ev.shiftKey)) {
      ev.preventDefault();
      commit();
    }
    ev.stopPropagation();
  });
  input.addEventListener('blur', () => {
    // Deferred so an Escape/Enter keydown that closed the popover wins.
    setTimeout(() => {
      if (active !== opts) return;
      if (opts.commitOnBlur) commit();
      else closePopover();
    }, 0);
  });
  opts.onInput?.(input.value, preview);
  input.focus();
  input.select();
}

function position(root: HTMLElement, anchor: DOMRect | null): void {
  const w = root.offsetWidth;
  const h = root.offsetHeight;
  let left: number;
  let top: number;
  if (!anchor) {
    left = (window.innerWidth - w) / 2;
    top = 110;
  } else {
    left = anchor.left + anchor.width / 2 - w / 2;
    top = anchor.bottom + 8;
    if (top + h > window.innerHeight - 12) top = anchor.top - h - 8;
  }
  root.style.left = `${Math.max(12, Math.min(left, window.innerWidth - w - 12))}px`;
  root.style.top = `${Math.max(12, top)}px`;
}
