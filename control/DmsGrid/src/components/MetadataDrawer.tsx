import * as React from 'react';
import {
  OverlayDrawer,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerBody,
  DrawerFooter,
  Field,
  Input,
  Textarea,
  Dropdown,
  Option,
  Switch,
  Combobox,
  Toolbar,
  ToolbarButton,
  Text,
  MessageBar,
  MessageBarBody,
  Spinner,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { DismissRegular } from '@fluentui/react-icons';
import { DatePicker, CalendarStrings } from '@fluentui/react-datepicker-compat';
import { DocumentRow, RawEntity, ClientConfig, T, AttributeValue } from '../types/Documents';
import { DocumentRepository, parallelLimit } from '../services/DocumentRepository';
import { MetadataService, AttributeMetadata } from '../services/MetadataService';
import { ApiError } from '../services/DataverseClient';
function calendarStrings(t: T): CalendarStrings {
  const months = Array.from({ length: 12 }, (_, month) => new Date(2024, month, 1));
  const days = Array.from({ length: 7 }, (_, day) => new Date(2024, 0, 7 + day));
  const format = (dates: Date[], options: Intl.DateTimeFormatOptions) =>
    dates.map((date) => new Intl.DateTimeFormat(undefined, options).format(date));
  return {
    months: format(months, { month: 'long' }),
    shortMonths: format(months, { month: 'short' }),
    days: format(days, { weekday: 'long' }),
    shortDays: format(days, { weekday: 'narrow' }),
    goToToday: t('calendar.goToToday'),
    prevMonthAriaLabel: t('calendar.prevMonthAriaLabel'),
    nextMonthAriaLabel: t('calendar.nextMonthAriaLabel'),
    prevYearAriaLabel: t('calendar.prevYearAriaLabel'),
    nextYearAriaLabel: t('calendar.nextYearAriaLabel'),
    prevYearRangeAriaLabel: t('calendar.prevYearRangeAriaLabel'),
    nextYearRangeAriaLabel: t('calendar.nextYearRangeAriaLabel'),
    monthPickerHeaderAriaLabel: t('calendar.monthPickerHeaderAriaLabel'),
    yearPickerHeaderAriaLabel: t('calendar.yearPickerHeaderAriaLabel'),
    closeButtonAriaLabel: t('calendar.closeButtonAriaLabel'),
    weekNumberFormatString: t('calendar.weekNumberFormatString'),
    selectedDateFormatString: t('calendar.selectedDateFormatString'),
    todayDateFormatString: t('calendar.todayDateFormatString'),
    dayMarkedAriaLabel: t('calendar.dayMarkedAriaLabel'),
  };
}
const useStyles = makeStyles({
  drawer: { width: '420px', maxWidth: '100vw' },
  fields: { display: 'flex', flexDirection: 'column', gap: '16px', paddingTop: '12px' },
  secondary: { fontSize: '12px', color: tokens.colorNeutralForeground2 },
  footer: { display: 'flex', gap: '8px' },
  message: { marginTop: '12px' },
  full: { width: '100%' },
});
function LookupEditor({
  field,
  value,
  onChange,
  metadata,
  t,
  mixed,
}: {
  field: AttributeMetadata;
  value: AttributeValue;
  onChange: (value: AttributeValue) => void;
  metadata: MetadataService;
  t: T;
  mixed: boolean;
}) {
  const [query, setQuery] = React.useState(String(value || '')),
    [options, setOptions] = React.useState<{ id: string; name: string }[]>([]);
  React.useEffect(() => {
    const controller = new AbortController();
    const timer = setTimeout(() => {
      void metadata
        .lookup(query, controller.signal)
        .then(setOptions)
        .catch(() => undefined);
    }, 300);
    return () => {
      clearTimeout(timer);
      controller.abort();
    };
  }, [query, metadata]);
  return (
    <Combobox
      aria-label={field.DisplayName.UserLocalizedLabel?.Label}
      placeholder={mixed ? t('mixed') : t('lookup.search')}
      value={query}
      selectedOptions={value ? [String(value)] : []}
      onChange={(event) => setQuery(event.target.value)}
      onOptionSelect={(_, data) => {
        onChange(data.optionValue || null);
        setQuery(data.optionText || '');
      }}
    >
      {options.map((option) => (
        <Option key={option.id} value={option.id}>
          {option.name}
        </Option>
      ))}
    </Combobox>
  );
}
export function MetadataDrawer({
  rows,
  repo,
  config,
  t,
  onClose,
  onSaved,
}: {
  rows: DocumentRow[];
  repo: DocumentRepository;
  config: ClientConfig;
  t: T;
  onClose: () => void;
  onSaved: (rows: DocumentRow[]) => void;
}) {
  const s = useStyles(),
    metadata = React.useMemo(() => new MetadataService(repo.client), [repo]),
    [fields, setFields] = React.useState<AttributeMetadata[]>([]),
    [values, setValues] = React.useState<RawEntity>({}),
    [error, setError] = React.useState<ApiError | null>(null),
    [saving, setSaving] = React.useState(false),
    [overwrite, setOverwrite] = React.useState(false),
    [currentRows, setCurrentRows] = React.useState(rows),
    [loaded, setLoaded] = React.useState(false),
    [failed, setFailed] = React.useState(false);
  React.useEffect(() => {
    let active = true;
    void metadata
      .load()
      .then((all) => {
        if (active) {
          setFields(
            all.filter((field) => config.entity.editableColumns.includes(field.LogicalName)),
          );
          setLoaded(true);
        }
      })
      .catch(() => {
        if (active) {
          setFailed(true);
          setLoaded(true);
        }
      });
    return () => {
      active = false;
    };
  }, [metadata, config]);
  const common = (field: string): AttributeValue | undefined => {
    const first = currentRows[0]?.raw[field] ?? null;
    return currentRows.every((row) => (row.raw[field] ?? null) === first) ? first : undefined;
  };
  const set = (field: string, value: AttributeValue) =>
    setValues((old) => ({ ...old, [field]: value }));
  const invalid = fields.some(
    (field) =>
      field.IsValidForUpdate &&
      field.RequiredLevel.Value !== 'None' &&
      field.LogicalName in values &&
      (values[field.LogicalName] === null || values[field.LogicalName] === ''),
  );
  async function save(force = false) {
    setSaving(true);
    setError(null);
    const results = await parallelLimit(currentRows, 3, (row) => repo.update(row, values, force));
    const saved = results
      .filter(
        (result): result is PromiseFulfilledResult<DocumentRow> => result.status === 'fulfilled',
      )
      .map((result) => result.value);
    if (saved.length) onSaved(saved);
    const failures = results
      .map((result, i) => (result.status === 'rejected' ? currentRows[i] : undefined))
      .filter((row): row is DocumentRow => !!row);
    const failure = results.find(
      (result): result is PromiseRejectedResult => result.status === 'rejected',
    );
    if (failure) {
      setCurrentRows(failures);
      setError(failure.reason instanceof ApiError ? failure.reason : new ApiError('unknown', 0));
    } else onClose();
    setSaving(false);
  }
  async function reload() {
    setSaving(true);
    try {
      const latest = await Promise.all(currentRows.map((row) => repo.get(row.id)));
      setCurrentRows(latest);
      onSaved(latest);
      setValues({});
      setError(null);
    } catch (error) {
      setError(error instanceof ApiError ? error : new ApiError('unknown', 0));
    } finally {
      setSaving(false);
    }
  }
  return (
    <>
      <OverlayDrawer
        open
        position="end"
        className={s.drawer}
        onOpenChange={(_, data) => !data.open && !saving && onClose()}
        aria-label={t('action.edit')}
      >
        <DrawerHeader>
          <DrawerHeaderTitle
            action={
              <Toolbar>
                <ToolbarButton
                  appearance="subtle"
                  aria-label={t('action.close')}
                  icon={<DismissRegular />}
                  onClick={onClose}
                  disabled={saving}
                />
              </Toolbar>
            }
          >
            {t('action.edit')}
          </DrawerHeaderTitle>
          <Text className={s.secondary}>
            {currentRows.length === 1
              ? currentRows[0].name
              : t('selection.count', { count: currentRows.length })}
          </Text>
        </DrawerHeader>
        <DrawerBody>
          {!loaded && <Spinner label={t('loading')} />}{' '}
          {failed && (
            <MessageBar intent="error">
              <MessageBarBody>{t('error.metadata')}</MessageBarBody>
            </MessageBar>
          )}
          {error && (
            <MessageBar className={s.message} intent="error">
              <MessageBarBody>
                <span>
                  {t(error.kind === 'conflict' ? 'edit.conflict' : 'error.' + error.kind)}
                </span>
                {error.kind === 'conflict' && (
                  <Toolbar>
                    <ToolbarButton onClick={() => void reload()} disabled={saving}>
                      {t('edit.reload')}
                    </ToolbarButton>
                    <ToolbarButton onClick={() => setOverwrite(true)} disabled={saving}>
                      {t('edit.overwrite')}
                    </ToolbarButton>
                  </Toolbar>
                )}
              </MessageBarBody>
            </MessageBar>
          )}
          <div className={s.fields}>
            {fields.map((field) => {
              const name = field.LogicalName,
                raw = name in values ? values[name] : common(name),
                mixed = raw === undefined,
                label = field.DisplayName.UserLocalizedLabel?.Label || t('column.' + name),
                readonly = !field.IsValidForUpdate;
              let editor: React.ReactNode;
              if (readonly) editor = <Text>{String(raw ?? t('empty.value'))}</Text>;
              else if (field.AttributeType === 'Memo')
                editor = (
                  <Textarea
                    value={String(raw ?? '')}
                    placeholder={mixed ? t('mixed') : undefined}
                    maxLength={field.MaxLength}
                    resize="vertical"
                    onChange={(_, data) => set(name, data.value)}
                  />
                );
              else if (
                field.AttributeType === 'Picklist' ||
                field.AttributeType === 'MultiSelectPicklist'
              ) {
                const multiple = field.AttributeType === 'MultiSelectPicklist',
                  selected = raw === null || mixed ? [] : String(raw).split(',');
                editor = (
                  <Dropdown
                    multiselect={multiple}
                    placeholder={mixed ? t('mixed') : t('edit.choose')}
                    value={selected
                      .map(
                        (value) =>
                          field.OptionSet?.Options?.find((option) => String(option.Value) === value)
                            ?.Label.UserLocalizedLabel?.Label || value,
                      )
                      .join(', ')}
                    selectedOptions={selected}
                    onOptionSelect={(_, data) =>
                      set(
                        name,
                        multiple
                          ? data.selectedOptions.join(',')
                          : data.optionValue
                            ? Number(data.optionValue)
                            : null,
                      )
                    }
                  >
                    {field.OptionSet?.Options?.map((option) => (
                      <Option value={String(option.Value)} key={option.Value}>
                        {option.Label.UserLocalizedLabel?.Label || String(option.Value)}
                      </Option>
                    ))}
                  </Dropdown>
                );
              } else if (field.AttributeType === 'Boolean')
                editor = (
                  <Switch
                    checked={raw === true}
                    label={mixed ? t('mixed') : t(raw ? 'yes' : 'no')}
                    onChange={(_, data) => set(name, data.checked)}
                  />
                );
              else if (field.AttributeType === 'DateTime')
                editor = (
                  <DatePicker
                    strings={calendarStrings(t)}
                    className={s.full}
                    value={raw ? new Date(String(raw) + 'T12:00:00') : undefined}
                    placeholder={mixed ? t('mixed') : t('edit.date')}
                    allowTextInput
                    onSelectDate={(date) =>
                      set(
                        name,
                        date
                          ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`
                          : null,
                      )
                    }
                    formatDate={(date) =>
                      date
                        ? new Intl.DateTimeFormat(undefined, { dateStyle: 'medium' }).format(date)
                        : ''
                    }
                    aria-label={label}
                  />
                );
              else if (field.AttributeType === 'Lookup')
                editor = (
                  <LookupEditor
                    field={field}
                    value={raw ?? null}
                    metadata={metadata}
                    t={t}
                    mixed={mixed}
                    onChange={(value) => set(name, value)}
                  />
                );
              else if (['Integer', 'Decimal', 'Money'].includes(field.AttributeType))
                editor = (
                  <Input
                    type="number"
                    step={field.AttributeType === 'Integer' ? 1 : 0.01}
                    value={raw === null || mixed ? '' : String(raw)}
                    placeholder={mixed ? t('mixed') : undefined}
                    onChange={(_, data) => {
                      const value = data.value === '' ? null : Number(data.value);
                      if (value === null || Number.isFinite(value)) set(name, value);
                    }}
                  />
                );
              else if (field.AttributeType === 'String')
                editor = (
                  <Input
                    value={String(raw ?? '')}
                    maxLength={field.MaxLength}
                    placeholder={mixed ? t('mixed') : undefined}
                    onChange={(_, data) => set(name, data.value)}
                  />
                );
              else editor = <Text>{t('edit.unsupported')}</Text>;
              return (
                <Field
                  key={name}
                  label={label}
                  required={field.RequiredLevel.Value !== 'None'}
                  hint={readonly ? t('edit.readonly') : undefined}
                  validationState={
                    field.RequiredLevel.Value !== 'None' && name in values && !values[name]
                      ? 'error'
                      : 'none'
                  }
                  validationMessage={
                    field.RequiredLevel.Value !== 'None' && name in values && !values[name]
                      ? t('edit.required')
                      : undefined
                  }
                >
                  {editor}
                </Field>
              );
            })}
          </div>
        </DrawerBody>
        <DrawerFooter>
          <Toolbar className={s.footer}>
            <ToolbarButton
              appearance="primary"
              disabled={saving || failed || invalid || !Object.keys(values).length}
              onClick={() => void save()}
            >
              {t(saving ? 'saving' : 'action.save')}
            </ToolbarButton>
            <ToolbarButton disabled={saving} onClick={onClose}>
              {t('action.cancel')}
            </ToolbarButton>
          </Toolbar>
        </DrawerFooter>
      </OverlayDrawer>
      <Dialog open={overwrite} onOpenChange={(_, data) => setOverwrite(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{t('edit.overwrite')}</DialogTitle>
            <DialogContent>{t('edit.overwriteConfirm')}</DialogContent>
            <DialogActions>
              <Toolbar>
                <ToolbarButton onClick={() => setOverwrite(false)}>
                  {t('action.cancel')}
                </ToolbarButton>
                <ToolbarButton
                  appearance="primary"
                  onClick={() => {
                    setOverwrite(false);
                    void save(true);
                  }}
                >
                  {t('edit.overwrite')}
                </ToolbarButton>
              </Toolbar>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
}
