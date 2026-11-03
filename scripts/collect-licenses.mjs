import { readFile, readdir, writeFile, mkdir } from 'node:fs/promises';
import path from 'node:path';

const lock = JSON.parse(await readFile('package-lock.json', 'utf8'));
const entries = [];
const sections = ['SHEET CLOUD — THIRD PARTY NOTICES\n\nPackages retain their original licenses and copyrights. JSZip is used under its MIT license option. This list includes installed production dependencies, some of which are not bundled into the browser application.\n'];
for (const [directory, item] of Object.entries(lock.packages)) {
  if (!directory || item.dev) continue;
  let files;
  try { files = await readdir(directory); } catch { continue; }
  const metadata = JSON.parse(await readFile(path.join(directory, 'package.json'), 'utf8'));
  const license = metadata.license || item.license || 'See package license file';
  const notices = files.filter(name => /^(licen[cs]e|copying|notice|copyright)(\.|$|-)/i.test(name));
  entries.push({ name: metadata.name, version: metadata.version, license, notices });
  sections.push(`\n${'='.repeat(70)}\n${metadata.name} ${metadata.version}\nLicense: ${typeof license === 'string' ? license : JSON.stringify(license)}\n`);
  for (const file of notices) { try { sections.push(`\n--- ${file} ---\n${await readFile(path.join(directory, file), 'utf8')}\n`); } catch { /* Ignore non-file notice directories. */ } }
}
await mkdir('docs', { recursive: true });
await mkdir('public', { recursive: true });
await writeFile('docs/DEPENDENCY_LICENSES.json', JSON.stringify(entries, null, 2) + '\n');
await writeFile('public/THIRD_PARTY_NOTICES.txt', sections.join('\n'));
console.log(`Collected license metadata and notices for ${entries.length} installed production packages.`);
