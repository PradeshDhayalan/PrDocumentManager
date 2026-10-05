import * as React from 'react';
import { FileTypeIcon } from '@fluentui/react-icons-file-type/headless';
import { GlobeRegular, DocumentRegular } from '@fluentui/react-icons';
import { tokens } from '@fluentui/react-components';
import { DocumentRow } from '../services/documents';
// Microsoft's file-type artwork includes the Office application logos.
export function FileIcon({ doc, large = false }: { doc: DocumentRow; large?: boolean }) {
  const [failed, setFailed] = React.useState(false);
  const size = large ? 48 : 24;
  const style = { width: size, height: size, flexShrink: 0, objectFit: 'contain' as const };
  const extension = doc.name.includes('.') ? doc.name.split('.').pop()?.toLowerCase() : '';
  React.useEffect(() => setFailed(false), [extension]);
  if (doc.fileType === 'Link')
    return (
      <GlobeRegular
        aria-label="SharePoint link"
        style={{ ...style, color: tokens.colorBrandForeground1 }}
      />
    );
  if (failed) return <DocumentRegular aria-label={`${doc.fileType} file`} style={style} />;
  return (
    <FileTypeIcon
      extension={extension}
      size={size}
      imageFileType="svg"
      alt={`${doc.fileType} file`}
      style={style}
      onError={() => setFailed(true)}
    />
  );
}
