import * as React from 'react';
import { Field, Input, Textarea, Dropdown, Option, Checkbox } from '@fluentui/react-components';
import { Attribute, Entity, Value } from '../services/documents';
export function MetadataFields({
  attributes,
  values,
  onChange,
  bulk,
}: {
  attributes: Attribute[];
  values: Entity;
  onChange: (key: string, value: Value) => void;
  bulk: boolean;
}) {
  return (
    <>
      {bulk && <p>Only fields you change will be applied to the selected documents.</p>}
      {attributes.map((attribute) => {
        const key = attribute.LogicalName,
          value = values[key],
          title = attribute.DisplayName.UserLocalizedLabel.Label;
        const required = !bulk && attribute.RequiredLevel.Value === 'ApplicationRequired';
        let input: React.ReactNode;
        const options = attribute.OptionSet?.Options;
        if (options) {
          const chosen = String(value ?? '')
            .split(',')
            .filter(Boolean);
          input = (
            <Dropdown
              aria-label={title}
              multiselect={attribute.AttributeType === 'MultiSelectPicklist'}
              value={chosen
                .map(
                  (v) =>
                    options.find((o) => String(o.Value) === v)?.Label.UserLocalizedLabel.Label ||
                    '',
                )
                .join(', ')}
              selectedOptions={chosen}
              placeholder={bulk ? 'Leave unchanged' : 'Select…'}
              onOptionSelect={(_, data) =>
                onChange(
                  key,
                  attribute.AttributeType === 'MultiSelectPicklist'
                    ? data.selectedOptions.join(',')
                    : data.optionValue
                      ? Number(data.optionValue)
                      : null,
                )
              }
            >
              {!required && attribute.AttributeType !== 'MultiSelectPicklist' && (
                <Option value="">None</Option>
              )}
              {options.map((option) => (
                <Option key={option.Value} value={String(option.Value)}>
                  {option.Label.UserLocalizedLabel.Label}
                </Option>
              ))}
            </Dropdown>
          );
        } else if (attribute.AttributeType === 'Boolean')
          input = (
            <Checkbox
              aria-label={title}
              checked={value === undefined ? 'mixed' : Boolean(value)}
              onChange={(_, data) => onChange(key, Boolean(data.checked))}
            />
          );
        else if (attribute.AttributeType === 'Memo')
          input = (
            <Textarea
              aria-label={title}
              value={String(value ?? '')}
              maxLength={attribute.MaxLength}
              resize="vertical"
              onChange={(_, data) => onChange(key, data.value)}
            />
          );
        else {
          const numeric = ['Integer', 'Decimal', 'Money'].includes(attribute.AttributeType);
          input = (
            <Input
              aria-label={title}
              value={String(value ?? '')}
              type={attribute.AttributeType === 'DateTime' ? 'date' : numeric ? 'number' : 'text'}
              step={attribute.AttributeType === 'Integer' ? '1' : 'any'}
              maxLength={attribute.MaxLength}
              required={required}
              placeholder={
                attribute.AttributeType === 'Lookup'
                  ? 'User GUID'
                  : bulk
                    ? 'Leave unchanged'
                    : undefined
              }
              onChange={(_, data) =>
                onChange(key, data.value === '' ? null : numeric ? Number(data.value) : data.value)
              }
            />
          );
        }
        return (
          <Field key={key} label={title} required={required}>
            {input}
          </Field>
        );
      })}
    </>
  );
}
