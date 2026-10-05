import {existsSync} from 'node:fs';
import {spawnSync} from 'node:child_process';
if (!existsSync('mock-api/data/state.json')) {
  const result = spawnSync('npm', ['run','seed'], {stdio:'inherit'});
  if (result.status !== 0) process.exit(result.status || 1);
}
