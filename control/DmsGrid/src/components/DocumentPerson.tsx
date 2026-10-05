import * as React from 'react';
import { Persona, makeStyles, tokens } from '@fluentui/react-components';
const useStyles = makeStyles({
  person: {
    minWidth: 0,
    '& .fui-Persona__primaryText': {
      fontSize: '13px',
      fontWeight: tokens.fontWeightRegular,
      color: tokens.colorNeutralForeground2,
      overflow: 'hidden',
      textOverflow: 'ellipsis',
      whiteSpace: 'nowrap',
    },
  },
});
export function DocumentPerson({ name, secondaryText }: { name: string; secondaryText?: string }) {
  const s = useStyles();
  return (
    <Persona
      className={s.person}
      name={name}
      size={secondaryText ? 'medium' : 'extra-small'}
      avatar={{ color: 'colorful' }}
      secondaryText={secondaryText}
    />
  );
}
