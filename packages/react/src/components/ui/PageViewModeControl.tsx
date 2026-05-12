/**
 * Page view mode control.
 *
 * Compact icon-triggered select for switching how paginated pages are arranged.
 */

import * as React from 'react';
import { Select, SelectContent, SelectItem, SelectTrigger } from './Select';
import { MaterialSymbol } from './MaterialSymbol';
import { cn } from '../../lib/utils';
import { useTranslation } from '../../i18n';
import type { DocumentViewMode } from '../Toolbar';

interface PageViewModeOption {
  value: DocumentViewMode;
  label: string;
  iconName: string;
}

export interface PageViewModeControlProps {
  value?: DocumentViewMode;
  onChange?: (mode: DocumentViewMode) => void;
  onRefocusEditor?: () => void;
  disabled?: boolean;
  className?: string;
}

export function PageViewModeControl({
  value = 'onePage',
  onChange,
  onRefocusEditor,
  disabled = false,
  className,
}: PageViewModeControlProps) {
  const { t } = useTranslation();
  const options = React.useMemo<PageViewModeOption[]>(
    () => [
      {
        value: 'onePage',
        label: t('pageView.onePage'),
        iconName: 'article',
      },
      {
        value: 'multiplePages',
        label: t('pageView.multiplePages'),
        iconName: 'view_column',
      },
      {
        value: 'pageWidth',
        label: t('pageView.pageWidth'),
        iconName: 'fit_width',
      },
    ],
    [t]
  );
  const selected = options.find((option) => option.value === value) ?? options[0];

  const handleValueChange = React.useCallback(
    (newValue: string) => {
      if (newValue === 'onePage' || newValue === 'multiplePages' || newValue === 'pageWidth') {
        onChange?.(newValue);
        requestAnimationFrame(() => onRefocusEditor?.());
      }
    },
    [onChange, onRefocusEditor]
  );

  return (
    <Select value={value} onValueChange={handleValueChange} disabled={disabled}>
      <SelectTrigger
        className={cn('h-7 w-[54px] min-w-0 justify-center px-1.5', className)}
        aria-label={t('pageView.ariaLabel', { label: selected.label })}
        title={t('pageView.title', { label: selected.label })}
        data-testid="page-view-mode-control"
      >
        <MaterialSymbol name={selected.iconName} size={18} />
      </SelectTrigger>
      <SelectContent className="min-w-[180px]">
        {options.map((option) => (
          <SelectItem
            key={option.value}
            value={option.value}
            className="pr-8"
            data-testid={`page-view-mode-${option.value}`}
          >
            <span className="flex items-center gap-2">
              <MaterialSymbol name={option.iconName} size={18} />
              <span>{option.label}</span>
            </span>
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}

export default PageViewModeControl;
