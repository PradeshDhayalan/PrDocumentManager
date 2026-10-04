import { choices } from './model';
const label = (text: string) => ({
  UserLocalizedLabel: { Label: text },
  LocalizedLabels: [{ Label: text, LanguageCode: 1033 }],
});
const fields: [string, string, string, number?, boolean?][] = [
  ['dms_name', 'String', 'Name', 255, true],
  ['dms_originalfilename', 'String', 'Original filename', 255],
  ['dms_description', 'Memo', 'Description', 4000],
  ['dms_regardingid', 'String', 'Regarding record', 36, true],
  ['dms_regardingtype', 'String', 'Regarding type', 64, true],
  ['dms_provider', 'Picklist', 'Storage provider', undefined, true],
  ['dms_storageref', 'String', 'Storage reference', 2000],
  ['dms_contenttype', 'String', 'Content type', 128],
  ['dms_filesizekb', 'Integer', 'File size (KB)'],
  ['dms_checksum', 'String', 'Checksum', 64],
  ['dms_documenttype', 'Picklist', 'Document type'],
  ['dms_documentstatus', 'Picklist', 'Status'],
  ['dms_expirydate', 'DateTime', 'Expiry date'],
  ['dms_uploadstate', 'Picklist', 'Upload state'],
  ['dms_priority', 'Picklist', 'Priority'],
  ['dms_tags', 'MultiSelectPicklist', 'Tags'],
  ['dms_confidential', 'Boolean', 'Confidential'],
  ['dms_amount', 'Money', 'Amount'],
  ['dms_rating', 'Decimal', 'Rating'],
  ['dms_revision', 'Integer', 'Revision'],
  ['dms_reviewer', 'Lookup', 'Reviewer'],
  ['createdon', 'DateTime', 'Created'],
  ['modifiedon', 'DateTime', 'Modified'],
  ['_createdby_value', 'Lookup', 'Created by'],
  ['_modifiedby_value', 'Lookup', 'Modified by'],
  ['_ownerid_value', 'Lookup', 'Owner'],
];
// TODO(SPIKE-S5): simplified metadata labels, option colours and required levels require real-org mapping.
export const attributes = fields.map(([name, type, title, length, required]) => ({
  LogicalName: name,
  AttributeType: type,
  DisplayName: label(title),
  RequiredLevel: { Value: required ? 'ApplicationRequired' : 'None' },
  MaxLength: length,
  IsValidForUpdate: ![
    'dms_regardingid',
    'dms_regardingtype',
    'dms_provider',
    'dms_storageref',
    'dms_contenttype',
    'dms_filesizekb',
    'dms_checksum',
    'dms_uploadstate',
    'createdon',
    'modifiedon',
    '_createdby_value',
    '_modifiedby_value',
    '_ownerid_value',
  ].includes(name),
  ...(type === 'DateTime' ? { Format: 'DateOnly' } : {}),
  ...(['Picklist', 'MultiSelectPicklist'].includes(type)
    ? {
        OptionSet: {
          Options: Object.entries(
            choices[name] || { 100000000: 'Low', 100000001: 'Normal', 100000002: 'High' },
          ).map(([value, text], i) => ({
            Value: Number(value),
            Label: label(text),
            Color: ['#107c10', '#0078d4', '#8764b8', '#ca5010', '#038387', '#616161'][i % 6],
          })),
        },
      }
    : {}),
  ...(type === 'Boolean'
    ? {
        OptionSet: {
          TrueOption: { Value: 1, Label: label('Yes') },
          FalseOption: { Value: 0, Label: label('No') },
        },
      }
    : {}),
}));
export const entitySets: Record<string, string> = {
  account: 'accounts',
  contact: 'contacts',
  opportunity: 'opportunities',
  lead: 'leads',
  incident: 'incidents',
  dms_document: 'dms_documents',
};
export const visibleColumns = [
  'dms_name',
  'modifiedon',
  '_modifiedby_value',
  'dms_filesizekb',
  'dms_documenttype',
  'dms_documentstatus',
  'dms_expirydate',
];
export const editableColumns = [
  'dms_name',
  'dms_description',
  'dms_documenttype',
  'dms_documentstatus',
  'dms_expirydate',
  'dms_priority',
  'dms_tags',
  'dms_confidential',
  'dms_amount',
  'dms_rating',
  'dms_revision',
  'dms_reviewer',
];
