import * as React from 'react';
import {
  Text,
  ProgressBar,
  Toolbar,
  ToolbarButton,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { DismissRegular, ArrowClockwiseRegular, CheckmarkRegular } from '@fluentui/react-icons';
import { UploadQueue } from '../services/UploadQueue';
import { T } from '../types/Documents';
const useStyles = makeStyles({
  root: {
    padding: '12px 20px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  header: { display: 'flex', alignItems: 'center', justifyContent: 'space-between' },
  jobs: {
    display: 'flex',
    flexDirection: 'column',
    gap: '10px',
    maxHeight: '190px',
    overflowY: 'auto',
  },
  job: {
    display: 'grid',
    gridTemplateColumns: 'minmax(130px,1fr) 100px auto',
    gap: '12px',
    alignItems: 'center',
    '@media(max-width:600px)': { gridTemplateColumns: 'minmax(120px,1fr) auto' },
  },
  name: { fontSize: '12px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  state: { fontSize: '12px', color: tokens.colorNeutralForeground2 },
  progress: { marginTop: '4px' },
  bar: { '@media(max-width:600px)': { display: 'none' } },
});
export function UploadTray({ queue, t }: { queue: UploadQueue; t: T }) {
  const s = useStyles(),
    [jobs, setJobs] = React.useState(queue.snapshot());
  React.useEffect(() => {
    setJobs(queue.snapshot());
    return queue.subscribe(() => setJobs(queue.snapshot()));
  }, [queue]);
  if (!jobs.length) return null;
  return (
    <section className={s.root} aria-label={t('upload.tray')}>
      <div className={s.header}>
        <Text weight="semibold">{t('upload.tray')}</Text>
        <Toolbar>
          <ToolbarButton appearance="subtle" onClick={() => queue.removeCompleted()}>
            {t('upload.clearCompleted')}
          </ToolbarButton>
        </Toolbar>
      </div>
      <div className={s.jobs}>
        {jobs.map((job) => (
          <div key={job.id} className={s.job}>
            <div>
              <div className={s.name}>{job.file.name}</div>
              <Text className={s.state}>{t('job.' + job.state)}</Text>
            </div>
            <div className={s.bar}>
              {job.state === 'complete' ? (
                <CheckmarkRegular />
              ) : (
                <ProgressBar className={s.progress} value={job.progress / 100} />
              )}
            </div>
            <Toolbar>
              {['failed', 'cancelled'].includes(job.state) && (
                <ToolbarButton icon={<ArrowClockwiseRegular />} onClick={() => queue.retry(job.id)}>
                  {t('action.retry')}
                </ToolbarButton>
              )}
              {job.state === 'duplicate' ? (
                <>
                  <ToolbarButton appearance="primary" onClick={() => queue.retry(job.id, true)}>
                    {t('upload.anyway')}
                  </ToolbarButton>
                  <ToolbarButton onClick={() => queue.cancel(job.id)}>
                    {t('action.skip')}
                  </ToolbarButton>
                </>
              ) : (
                ['queued', 'hashing', 'uploading'].includes(job.state) && (
                  <ToolbarButton
                    icon={<DismissRegular />}
                    aria-label={t('upload.cancel', { name: job.file.name })}
                    onClick={() => queue.cancel(job.id)}
                  />
                )
              )}
            </Toolbar>
          </div>
        ))}
      </div>
    </section>
  );
}
