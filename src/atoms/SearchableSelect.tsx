import React, { forwardRef, useEffect, useId, useMemo, useState } from 'react';

export interface SearchableSelectOption {
  value: string;
  label: string;
}

interface SearchableSelectProps extends Omit<
  React.InputHTMLAttributes<HTMLInputElement>,
  'onChange' | 'value' | 'type' | 'list'
> {
  label?: string;
  error?: string;
  options: SearchableSelectOption[];
  placeholder?: string;
  value: string;
  onChange: (value: string) => void;
  fullWidth?: boolean;
}

function normalizeForSnap(s: string): string {
  return s
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replaceAll(/\p{M}/gu, '');
}

function findCanonicalOption(
  options: SearchableSelectOption[],
  raw: string
): SearchableSelectOption | undefined {
  const t = raw.trim();
  if (!t) {
    return undefined;
  }
  const n = normalizeForSnap(t);
  return options.find(
    (o) => normalizeForSnap(o.value) === n || normalizeForSnap(o.label) === n
  );
}

/** Si no hay coincidencia exacta pero solo una opción contiene el texto (p. ej. prefijo), se acepta al salir del campo. */
function findUniquePartialMatch(
  options: SearchableSelectOption[],
  raw: string
): SearchableSelectOption | undefined {
  const n = normalizeForSnap(raw);
  if (!n) {
    return undefined;
  }
  const candidates = options.filter(
    (o) =>
      normalizeForSnap(o.label).includes(n) || normalizeForSnap(o.value).includes(n)
  );
  return candidates.length === 1 ? candidates[0] : undefined;
}

function resolveOptionOnBlur(
  options: SearchableSelectOption[],
  raw: string
): SearchableSelectOption | undefined {
  return findCanonicalOption(options, raw) ?? findUniquePartialMatch(options, raw);
}

function displayLabelForValue(
  options: SearchableSelectOption[],
  value: string
): string {
  if (!value) {
    return '';
  }
  return options.find((o) => o.value === value)?.label ?? value;
}

/**
 * Campo de texto con &lt;datalist&gt; nativo: el navegador filtra sugerencias al escribir (p. ej. tipo / marca / modelo).
 * En blur se normaliza el valor a una opción válida o se vacía para que falle la validación requerida.
 */
export const SearchableSelect = forwardRef<HTMLInputElement, SearchableSelectProps>(
  (
    {
      label,
      error,
      options,
      placeholder,
      value,
      onChange,
      disabled,
      fullWidth = true,
      className = '',
      id,
      onBlur,
      ...inputProps
    },
    ref
  ) => {
    const generatedId = useId();
    const inputId = id ?? generatedId;
    const datalistId = `${inputId}-datalist`;

    const [text, setText] = useState(() => displayLabelForValue(options, value));

    const canonicalFromForm = useMemo(
      () => displayLabelForValue(options, value),
      [options, value]
    );

    useEffect(() => {
      setText(canonicalFromForm);
    }, [canonicalFromForm]);

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
      const next = e.target.value;
      setText(next);
      onChange(next);
    };

    const handleBlur = (e: React.FocusEvent<HTMLInputElement>) => {
      const raw = e.target.value;
      if (!raw.trim()) {
        onChange('');
        setText('');
        onBlur?.(e);
        return;
      }
      const canon = resolveOptionOnBlur(options, raw);
      if (canon) {
        onChange(canon.value);
        setText(canon.label);
      } else {
        onChange('');
        setText('');
      }
      onBlur?.(e);
    };

    const inputClasses = [
      'block px-3 py-2 border rounded-lg text-gray-900 placeholder-gray-500 focus:outline-none focus:ring-2 focus:ring-red-500 focus:border-red-500 transition-all duration-200 bg-white',
      error ? 'border-red-500' : 'border-gray-300',
      disabled ? 'cursor-not-allowed bg-gray-50 text-gray-500' : '',
      fullWidth ? 'w-full' : '',
      className,
    ]
      .filter(Boolean)
      .join(' ');

    return (
      <div className={fullWidth ? 'w-full' : ''}>
        {label ? (
          <label htmlFor={inputId} className="block text-sm font-medium text-gray-700 mb-1">
            {label}
          </label>
        ) : null}
        <input
          {...inputProps}
          ref={ref}
          id={inputId}
          type="text"
          list={datalistId}
          disabled={disabled}
          placeholder={placeholder}
          value={text}
          onChange={handleChange}
          onBlur={handleBlur}
          autoComplete="off"
          className={inputClasses}
        />
        <datalist id={datalistId}>
          {options.map((o) => (
            <option key={o.value} value={o.value}>
              {o.label}
            </option>
          ))}
        </datalist>
        {error ? <p className="mt-1 text-sm text-red-600">{error}</p> : null}
      </div>
    );
  }
);

SearchableSelect.displayName = 'SearchableSelect';
